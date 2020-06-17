import { WebApi } from 'azure-devops-node-api';
import * as ta from 'azure-devops-node-api/TaskAgentApi';
import * as ca from 'azure-devops-node-api/CoreApi';
import * as fs from 'fs';
import { TeamProject, TeamProjectReference } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { VariableGroupParameters, VariableGroupProjectReference, VariableGroup } from 'azure-devops-node-api/interfaces/TaskAgentInterfaces';

import { GroupType } from './interfaces/Enums';
import { OnboardingServices } from './index';
import { libraryLogger } from './logging';
import { LIBRARY_NAMESPACE } from './SecurityNamespaces';

export class LibraryServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Use this function to create/update the specified variablegroup to all the projects.
   * This can be used to push a organisation wide variablegroup. It should not be used to update variablegroups that can be altered by project members since it will overwrite it.
   * @returns { TeamProject} the project itself;
   */
  public async createOrUpdateVariableGroupForAllProjects(): Promise<void> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    const allProjects: TeamProjectReference[] = await coreApi.getProjects();

    for (const project of allProjects) {
      await this.createOrUpdateVariableGroupForProject(project);
    }
  }

  /** Use this function to create/update the specified variablegroup to a specified project.
   * This can be used to push a organisation wide variablegroup. It should not be used to update variablegroup that can be altered by project members since it will overwrite it.
   * @param { TeamProject } project the project where the variablegroup should be pushed to;
   * @returns { TeamProject} the project itself;
   */
  public async createOrUpdateVariableGroupForProject(project: TeamProject): Promise<TeamProject> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const variableGroupParameters: VariableGroupParameters = JSON.parse(fs.readFileSync('settings/' + this.azureDevOpsServices.configuration().CONFIG_PROJECTVARIABLEGROUPFILE, 'utf8'));

    const variableGroupProjectReferences: VariableGroupProjectReference[] = [];
    variableGroupParameters.name = variableGroupParameters.name;
    // only add specified project
    variableGroupProjectReferences.push({
      description: variableGroupParameters.description,
      name: variableGroupParameters.name,
      projectReference: {
        id: project.id,
      },
    });

    variableGroupParameters.variableGroupProjectReferences = variableGroupProjectReferences;

    // check if variableGroup already exists
    const variableGroups: VariableGroup[] = await taskAgentApi.getVariableGroups(project.name!, variableGroupParameters.name!);
    if (variableGroups.length > 0) {
      const projectIds: string[] = [];
      projectIds.push(project.id!);
      await taskAgentApi.updateVariableGroup(variableGroupParameters, variableGroups[0].id!);
      libraryLogger.info("Updated variable group '" + variableGroupParameters.name! + "'");
    } else {
      const variableGroup: VariableGroup = await taskAgentApi.addVariableGroup(variableGroupParameters);
      await this.azureDevOpsServices.authorizeResource().authorizeResource(project, variableGroup.id!.toString(), variableGroup.name!, 'variablegroup');
      libraryLogger.info("Created variable group '" + variableGroupParameters.name + "'");
    }
    return project;
  }

  /** Use this function to create a template Product variablegroup to a specified project.
   * @param { string } project the project where the Library should be pushed to;
   * @param { string } productName the product Name;
   * @returns { TeamProject} the project itself;
   */
  public async createProductVariableGroup(project: TeamProject, productName: string): Promise<TeamProject> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const variableGroupParameters: VariableGroupParameters = JSON.parse(fs.readFileSync('settings/' + this.azureDevOpsServices.configuration().CONFIG_PRODUCTVARIABLEGROUPFILE, 'utf8'));

    const variableGroupProjectReferences: VariableGroupProjectReference[] = [];
    const variableGroupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + productName + ' ' + variableGroupParameters.name;
    variableGroupParameters.name = variableGroupName;
    variableGroupProjectReferences.push({
      description: this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + productName + ' - ' + variableGroupParameters.description,
      name: variableGroupName.trim(),
      projectReference: {
        id: project.id,
      },
    });

    variableGroupParameters.variableGroupProjectReferences = variableGroupProjectReferences;

    const variableGroups: VariableGroup[] = await taskAgentApi.getVariableGroups(project.name!, variableGroupName);
    if (variableGroups.length === 0) {
      const variableGroup: VariableGroup = await taskAgentApi.addVariableGroup(variableGroupParameters);

      const token = 'Library/' + project.id + '/VariableGroup/' + variableGroup.id;

      const security = this.azureDevOpsServices.security();
      await security.setPermissionOnSimpleEntity(project, LIBRARY_NAMESPACE, token, this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + productName, (await security.getSimpleObjectPermissions()).Library!.OwnerRights!);
      await security.setPermissionOnSimpleEntity(project, LIBRARY_NAMESPACE, token, 'Contributors', (await security.getSimpleObjectPermissions()).Library!.ContributorRights!);
      await this.azureDevOpsServices.authorizeResource().deAuthorizeResource(project, variableGroup.id!.toString(), variableGroup.name!, 'variablegroup');
      libraryLogger.info("Created variable group '" + variableGroupParameters.name + "'");
      return project;
    }
    libraryLogger.info("Did not create variable group '" + variableGroupParameters.name + "' because it already exists.");
    return project;
  }

  /** Use this function to harden a Variablegroup
   * @param { string } project the project where the Library should be pushed to;
   * @param { string } variableGroupName the variablegroup Name;
   * @param { string } groupName the group Name;
   * @param { GroupType } groupType the group type of the groupname (team or product)
   * @returns { TeamProject} the project itself;
   */
  public async hardenVariableGroup(project: TeamProject, variableGroupName: string, groupName: string, groupType: GroupType): Promise<void> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const variableGroups: VariableGroup[] = await taskAgentApi.getVariableGroups(project.name!, variableGroupName);
    const thisVariableGroup: VariableGroup | undefined = variableGroups.find((v) => v.name! === variableGroupName);
    if (!thisVariableGroup) throw Error("Variable group '" + variableGroupName + "'  not found");
    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;
    const token = 'Library/' + project.id + '/VariableGroup/' + thisVariableGroup.id;

    const security = this.azureDevOpsServices.security();
    await security.setPermissionOnSimpleEntity(project, LIBRARY_NAMESPACE, token, groupName, (await security.getSimpleObjectPermissions()).Library!.OwnerRights!);
    await security.setPermissionOnSimpleEntity(project, LIBRARY_NAMESPACE, token, 'Contributors', (await security.getSimpleObjectPermissions()).Library!.ContributorRights!);
    await this.azureDevOpsServices.authorizeResource().deAuthorizeResource(project, thisVariableGroup.id!.toString(), thisVariableGroup.name!, 'variablegroup');
    libraryLogger.info("Hardened variable group '" + variableGroupName + "'.");
  }

  public async createVariableGroup(project: TeamProject, variableGroupName: string, variableGroupDescription: string, groupName: string, groupType: GroupType): Promise<void> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const variableGroupParameters: VariableGroupParameters = {
      description: variableGroupDescription,
      name: variableGroupName,
      variables: {
        parameter: {
          isReadOnly: true,
          isSecret: false,
          value: 'value',
        },
      },
      variableGroupProjectReferences: [
        {
          description: variableGroupDescription,
          name: variableGroupName,
          projectReference: {
            id: project.id,
          },
        },
      ],
    };
    const variableGroups: VariableGroup = await taskAgentApi.addVariableGroup(variableGroupParameters);
    await this.hardenVariableGroup(project, variableGroupName, groupName, groupType);
  }
}
