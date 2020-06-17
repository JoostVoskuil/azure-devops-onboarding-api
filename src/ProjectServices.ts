import * as ca from 'azure-devops-node-api/CoreApi';
import * as fa from 'azure-devops-node-api/FeatureManagementApi';

import * as fs from 'fs';

import { TeamProject, ProjectVisibility } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { ContributedFeatureState, ContributedFeatureEnabledValue } from 'azure-devops-node-api/interfaces/FeatureManagementInterfaces';
import { WebApi } from 'azure-devops-node-api';
import { OnboardingServices } from './index';
import { SubjectType, GroupScope } from './interfaces/Enums';
import { IProjectPermission } from './interfaces/IProjectPermission';
import { ISecurityBits } from './SecurityServices';
import { projectLogger, delay } from './logging';

const apiVersion = 'api-version=5.1-preview.1';

export class ProjectServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** gets a teamProject based on the name
   * @param { string } projectName the project name;
   */
  public async getProject(projectName: string): Promise<TeamProject> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    const project: TeamProject = await coreApi.getProject(projectName);
    return project;
  }

  /** Check if project exists
   * @param { string } projectName the project name;
   * @returns { boolean } true if project exists
   */
  public async checkIfProjectExists(projectName: string): Promise<boolean> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    const project: TeamProject = await coreApi.getProject(projectName);
    if (project) return true;
    return false;
  }

  /** deletes a teamProject based on the name
   * @param { string } projectName the project name;
   */
  public async deleteProject(projectName: string): Promise<void> {
    let numLoops = 0;
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    let project: TeamProject = await coreApi.getProject(projectName);
    if (project) await coreApi.queueDeleteProject(project.id!);
    else {
      projectLogger.info(projectName + ' did not exist, could not delete');
      return;
    }

    while (project && numLoops < this.azureDevOpsServices.configuration().MAX_LOOP_FOR_PROJECT_QUEING) {
      project = await coreApi.getProject(projectName);
      numLoops += 1;
      await delay(500);
      if (!project) {
        await delay(2000);
        projectLogger.info("'" + projectName + "' is deleted.");
        return;
      }
    }
  }

  /** creates a teamProject
   * @param { string } projectName the project name;
   * @param { string } projectDescription the project description;
   * @param { boolean } deleteWhenExists when set to true, the project is deleted when already exists;
   * @param { boolean } disableWork when set to true, Azure Boards Feature is disabled;
   * @param { boolean } disableTestPlans when set to true, Azure Testplans Feature is disabled;
   * @param { boolean } disableArtifacts when set to true, Azure Artifacts Feature is disabled;
   * @returns { TeamProject} the created (or already existing) project;
   */
  public async createProject(projectName: string, projectDescription: string, disableWork: boolean = false, disableTestPlans: boolean = false, disableArtifacts: boolean = false): Promise<TeamProject> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();

    let numLoops = 0;
    let project: TeamProject = await coreApi.getProject(projectName);

    if (project && this.azureDevOpsServices.configuration().TESTING) {
      this.deleteProject(projectName);
    } else if (project && !this.azureDevOpsServices.configuration().TESTING) {
      throw Error("'" + projectName + "' already exists. Cannot create Project");
    }

    const projectToCreate: TeamProject = {
      name: projectName,
      description: projectDescription,
      visibility: ProjectVisibility.Private,
      capabilities: {
        versioncontrol: { sourceControlType: 'Git' },
        processTemplate: { templateTypeId: this.azureDevOpsServices.configuration().TEAM_PROJECT_PROCESS_TEMPLATE_ID },
      },
    };

    await coreApi.queueCreateProject(projectToCreate);

    // Poll until project exists
    while (numLoops < this.azureDevOpsServices.configuration().MAX_LOOP_FOR_PROJECT_QUEING) {
      project = await coreApi.getProject(projectName);
      numLoops += 1;
      await delay(500);
      if (project) {
        await delay(10000);
        projectLogger.info("'" + projectName + "' created.");
        await this.toggleFeatures(project, disableWork, disableTestPlans, disableArtifacts);
        return project;
      }
    }

    throw Error("'" + projectName + "' is not created, timeout.");
  }
  /** Deletes creation mess
   * @param { TeamProject } project the project name;
   * @param { string } defaultTeamAdminGroup the team admin group for the default project
   * @param { string } creationAdmin the admin that created the teamproject (PAT token)
   * @returns { TeamProject} the created (or already existing) project;
   */
  public async deleteProjectCreationMess(project: TeamProject, defaultTeamAdminGroup: string, creationAdmin: string = this.azureDevOpsServices.configuration().AZUREDEVOPS_PAT_OWNER!): Promise<TeamProject> {
    // Default team teamadmin is set to the creator, add new admin group
    await this.azureDevOpsServices.team().changeTeamAdmin(project, project.name + ' Team', defaultTeamAdminGroup);
    await this.azureDevOpsServices.group().addMemberToGroup(project, defaultTeamAdminGroup, project.name + ' Team');
    // Remove the creator as a team admin
    await this.azureDevOpsServices.team().deleteTeamAdminUser(project, project.name + ' Team', creationAdmin, SubjectType.User);
    // Remove the creator as a member of the default team
    await this.azureDevOpsServices.group().deleteMemberFromGroup(project, project.name + ' Team', creationAdmin, SubjectType.User);
    // Remove the creator from the Project Administrators group
    await this.azureDevOpsServices.group().deleteMemberFromGroup(project, 'Project Administrators', creationAdmin, SubjectType.User);

    // Remove creator GIT permission
    const namespaceId = '2e9eb7ed-3c0a-47d4-87c1-0ffdd275fd87';
    const token = 'repoV2/' + project.id;
    const descriptor = (await this.azureDevOpsServices.user().getUserProperties(project, creationAdmin)).descriptor!;
    await this.azureDevOpsServices.security().deleteAccessControlEntry(namespaceId, token, descriptor);

    // Create fake Release and Endpoint to set security namespaces and creation of security groups
    await this.azureDevOpsServices.release().createAndDeleteFakeReleaseDefinition(project);
    await this.azureDevOpsServices.serviceConnection().createAndDeleteFakeEndpoint(project);
    await this.azureDevOpsServices.deploymentGroup().createAndDeleteFakeDeploymentGroup(project);
    await this.azureDevOpsServices.environment().createAndDeleteFakeEnvironment(project);

    // Remove default groups
    await this.azureDevOpsServices.group().deleteGroup(project, 'Build Administrators');
    await this.azureDevOpsServices.group().deleteGroup(project, 'Deployment Group Administrators');
    await this.azureDevOpsServices.group().deleteGroup(project, 'Release Administrators');
    await this.azureDevOpsServices.group().deleteGroup(project, 'Endpoint Administrators');
    await this.azureDevOpsServices.group().deleteGroup(project, 'Endpoint Creators');

    return project;
  }

  /** gets the descriptor of a project
   * @param { TeamProject } project the project;
   * @returns { string} the descriptor;
   */
  public async getProjectDescriptor(project: TeamProject): Promise<string> {
    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/descriptors/' + project.id + '?' + apiVersion;
    const response: any = await (await this.connection.rest.get(descriptorsUrl)).result;
    return response.value;
  }

  /** toggles project Features
   * @param { TeamProject } project the project name;
   * @param { boolean } disableWork when set to true, Azure Boards Feature is disabled;
   * @param { boolean } disableTestPlans when set to true, Azure Testplans Feature is disabled;
   * @param { boolean } disableArtifacts when set to true, Azure Artifacts Feature is disabled;
   * @returns { TeamProject} the project itself;
   */
  public async toggleFeatures(project: TeamProject, disableWork: boolean, disableTestPlans: boolean, disableArtifacts: boolean): Promise<TeamProject> {
    const featureManagementApi: fa.FeatureManagementApi = await this.connection.getFeatureManagementApi();
    const feature: ContributedFeatureState = {
      state: ContributedFeatureEnabledValue.Disabled,
    };

    if (disableWork) await featureManagementApi.setFeatureStateForScope(feature, 'ms.vss-work.agile', 'host', 'project', project.id!);
    if (disableTestPlans) await featureManagementApi.setFeatureStateForScope(feature, 'ms.vss-test-web.test', 'host', 'project', project.id!);
    if (disableArtifacts) await featureManagementApi.setFeatureStateForScope(feature, 'ms.feed.feed', 'host', 'project', project.id!);
    projectLogger.info("Set feature settings for '" + project.name + "'");
    return project;
  }

  /** Applies the default security settings for all teamprojects
   * @param { string[] } includedTeamProjects Only these teamprojects are targeted. Ignored when excludedTeamProjects is used
   * @param { string[] } excludedTeamProjects Exclude these teamprojects. Ignored when includedTeamProjects is used
   * @param { string } permissionFile the Permission file
   */
  public async applyDefaultSecurityForAllProjects(includedTeamProjects?: string[], excludedTeamProjects?: string[]): Promise<void> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    let projects: TeamProject[] = await coreApi.getProjects(undefined, 500);

    if (includedTeamProjects) {
      projects = projects.filter((p) => includedTeamProjects.includes(p.name!));
    }
    if (excludedTeamProjects) {
      projects = projects.filter((p) => !excludedTeamProjects.includes(p.name!));
    }

    for (const project of projects) {
      projectLogger.info("Processing Security for project '" + project.name! + "'");
      await this.setDefaultProjectSecurity(project);
    }
  }

  /** sets the default Project Security based on a template file
   * @param { TeamProject } project the project;
   * @returns { TeamProject} the project itself;
   */
  public async setDefaultProjectSecurity(project: TeamProject): Promise<TeamProject> {
    const securitySettings: IProjectPermission[] = JSON.parse(fs.readFileSync('settings/' + this.azureDevOpsServices.configuration().CONFIG_PROJECTPERMISSIONFILE, 'utf8'));
    for (const role of securitySettings) {
      for (const namespace of role.Namespaces) {
        const bits: ISecurityBits = await this.azureDevOpsServices.security().determineAllowAndDenyBits(namespace.NamespaceId!, namespace.Allow!, namespace.Deny!);
        let projectOnly = true;
        let groupName = role!.Group!;
        if (role.GroupScope === GroupScope.OrganisationGroup) projectOnly = false;
        else if (role.GroupScope === GroupScope.ProjectGroup) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_SECURITY_GROUP_PREFIX + role.Group!;
        await this.azureDevOpsServices.security().AddOrChangeAccessControlEntryOnProjectLevel(project, namespace.NamespaceId!, namespace.TokenPrefix!, groupName, bits.allowBit!, bits.denyBit!, true, projectOnly);
      }
    }
    return project;
  }
}
