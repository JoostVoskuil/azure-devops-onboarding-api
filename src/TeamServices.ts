import * as ca from 'azure-devops-node-api/CoreApi';
import * as wit from 'azure-devops-node-api/WorkItemTrackingApi';
import * as wa from 'azure-devops-node-api/WorkApi';

import { TeamProject, WebApiTeam, TeamContext } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { TeamSettingsPatch, BugsBehavior, TeamFieldValuesPatch, TeamSetting } from 'azure-devops-node-api/interfaces/WorkInterfaces';

import { WebApi } from 'azure-devops-node-api';
import { OnboardingServices } from './index';
import { GroupType, SubjectType } from './interfaces/Enums';
import { teamLogger, delay } from './logging';

const apiVersion = 'api-version=6.0-preview.2';
const hierarchyQueryApiVersion = 'api-version=5.0-preview.1';

export class TeamServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Creates the team objects like folders, dashboards and area paths.
   * @param { TeamProject } project The project where the service connection is maintained;
   * @param { WebApiTeam } team The team;
   * @returns { WebApiTeam } the team;
   */
  public async createObjectsAndSetSecurity(project: TeamProject, team: WebApiTeam): Promise<WebApiTeam> {
    const teamContext: TeamContext = { project: project.name, team: team.name };
    const teamProjectContext: TeamContext = { project: project.name, team: project.name + ' Team' };
    const workApi: wa.WorkApi = await this.connection.getWorkApi();

    // Create Build & Release Folder
    await this.azureDevOpsServices.pipeline().createBuildFolderAndSetPermissions(project, team.name!, this.azureDevOpsServices.configuration().CONFIG_BUILDFOLERPERMISSIONFILE);
    await this.azureDevOpsServices.release().createReleaseFolderAndSetPermissions(project, team.name!, this.azureDevOpsServices.configuration().CONFIG_RELEASEFOLERPERMISSIONFILE);

    await this.azureDevOpsServices.dashboard().createTeamDashboard(project, team);
    // Create Area Path and set default
    const workItemTrackingApi: wit.WorkItemTrackingApi = await this.connection.getWorkItemTrackingApi();
    const node = {
      name: team.name,
    };

    const createAreaPathUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + project.name + '/_apis/wit/classificationnodes/Areas?' + apiVersion;
    const createdAreaPathResult = await this.connection.rest.create(createAreaPathUrl, node);
    teamLogger.info('Area path created.');

    // Set Default Work Settings
    const projectTeamSettings: TeamSetting = await workApi.getTeamSettings(teamProjectContext);
    const teamSettingsPatch: TeamSettingsPatch = {
      backlogIteration: projectTeamSettings.backlogIteration.id,
      defaultIterationMacro: '@CurrentIteration',
      bugsBehavior: BugsBehavior.AsTasks,
      backlogVisibilities: {
        'Microsoft.EpicCategory': true,
        'Microsoft.FeatureCategory': true,
        'Microsoft.RequirementCategory': true,
      },
    };
    await workApi.updateTeamSettings(teamSettingsPatch, teamContext);
    teamLogger.info('Team Default Settings Set.');
    const patch: TeamFieldValuesPatch = {
      defaultValue: project.name + '\\' + team.name,
      values: [
        {
          includeChildren: false,
          value: project.name + '\\' + team.name,
        },
      ],
    };
    const result = await workApi.updateTeamFieldValues(patch, teamContext);
    teamLogger.info('Team Default Area Path Set.');
    return team;
  }

  /** Creates a team in a project
   * @param { TeamProject } project The project where the service connection is maintained;
   * @param { string } teamName the team Name;
   * @param { string } teamDescription the team description;
   * @returns { WebApiTeam } the team;
   */
  public async createTeam(project: TeamProject, teamName: string, teamDescription: string): Promise<WebApiTeam> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();

    let team: WebApiTeam = await coreApi.getTeam(project.id!, teamName);
    if (team) teamLogger.info(teamName + ' already exists.');
    else {
      const newTeam: WebApiTeam = {
        name: teamName,
        description: teamDescription,
      };

      team = await coreApi.createTeam(newTeam, project.id!);
      await delay(1000);
      teamLogger.info("'" + teamName + "' created.");
    }
    return team;
  }

  /** Deletes a team from a project
   * @param { TeamProject } project The project where the service connection is maintained;
   * @param { string } teamName the team Name;
   */
  public async deleteTeam(project: TeamProject, teamName: string): Promise<void> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();

    const team: WebApiTeam = await coreApi.getTeam(project.id!, teamName);

    if (!team) teamLogger.warn(teamName + ' did not exist.');
    else {
      await coreApi.deleteTeam(project.id!, teamName);
      teamLogger.info(teamName + ' deleted.');
    }
  }

  /** Gets the team by name
   * @param { TeamProject } project The project where the service connection is maintained;
   * @param { string } teamName the team Name;
   * @returns { WebApiTeam } the team;
   */
  public async getTeam(project: TeamProject, teamName: string): Promise<WebApiTeam | undefined> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    const team: WebApiTeam = await coreApi.getTeam(project.id!, teamName);
    return team;
  }

  /** Set the project administrators as default team admin
   * @param { TeamProject } project the project;
   * @param { string } teamName the team name
   * @param { string } teamAdmin the team admin;
   * @returns { string} the descriptor;
   */
  public async changeTeamAdmin(project: TeamProject, teamName: string, teamAdmin: string): Promise<void> {
    const setTeamAdminUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/Contribution/HierarchyQuery?' + hierarchyQueryApiVersion;
    const addAdmin = {
      contributionIds: ['ms.vss-admin-web.admin-teams-data-provider'],
      dataProviderContext: {
        properties: {
          setTeamAdmins: true,
          teamDescriptor: await this.azureDevOpsServices.group().getGroupDescriptor(project, teamName),
          admins: [await this.azureDevOpsServices.group().getGroupDescriptor(project, teamAdmin)],
        },
      },
    };

    const result = await this.connection.rest.create(setTeamAdminUrl, addAdmin);
    if (result.statusCode! !== 200) throw new Error("Error changing team admin of team '" + teamName + "'");
    teamLogger.info("Changed admin of team '" + teamName + "'");
  }

  /** Removes a team administrator (user)
   * @param { TeamProject } project the project;
   * @param { string } teamName the team name
   * @param { string } currentTeamAdmin the team admin user;
   * @param { SubjectType } teamAdminType if this is an user or group that you want to remove
   */
  public async deleteTeamAdminUser(project: TeamProject, teamName: string, currentTeamAdmin: string, teamAdminType: SubjectType): Promise<void> {
    const setTeamAdminUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/Contribution/HierarchyQuery?' + hierarchyQueryApiVersion;
    let adminDescriptor: string = '';
    if (teamAdminType === SubjectType.Group) adminDescriptor = await this.azureDevOpsServices.group().getGroupDescriptor(project, currentTeamAdmin);
    else if (teamAdminType === SubjectType.User) adminDescriptor = (await this.azureDevOpsServices.user().getUserProperties(project, currentTeamAdmin)).descriptor!;
    const removeAdmin = {
      contributionIds: ['ms.vss-admin-web.admin-teams-data-provider'],
      dataProviderContext: {
        properties: {
          removeAdmins: true,
          teamDescriptor: await this.azureDevOpsServices.group().getGroupDescriptor(project, teamName),
          adminDescriptor,
        },
      },
    };

    const result = await this.connection.rest.create(setTeamAdminUrl, removeAdmin);
    if (result.statusCode !== 200) throw new Error("Error deleting team admin of team '" + teamName + "'");
    teamLogger.info("Removed adminUser of team '" + teamName + "'");
  }
}
