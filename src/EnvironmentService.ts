import { WebApi } from 'azure-devops-node-api';

import * as ta from 'azure-devops-node-api/TaskAgentApi';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { EnvironmentInstance, EnvironmentCreateParameter } from 'azure-devops-node-api/interfaces/TaskAgentInterfaces';

import { OnboardingServices } from './index';
import { GroupType } from './interfaces/Enums';
import { environmentLogger } from './logging';
import { ENVIRONMENT_NAMESPACE } from './SecurityNamespaces';

export class EnvironmentServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Creates an environment
   * @param { TeamProject } project: the teamProject;
   * @param { string } name name;
   * @param { string } description description;
   * @returns { EnvironmentInstance } the environment;
   */
  public async createEnvironment(project: TeamProject, name: string, description?: string): Promise<EnvironmentInstance> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const environmentCreateParameter: EnvironmentCreateParameter = {
      description,
      name,
    };
    const environment: EnvironmentInstance = await taskAgentApi.addEnvironment(environmentCreateParameter, project.name!);
    environmentLogger.info("Created environment '" + name + "'");
    return environment;
  }

  /** Creates a fake environment and deletes it.
   * @param { TeamProject } project the teamproject
   * @returns { TeamProject} the project itself;
   */
  public async createAndDeleteFakeEnvironment(project: TeamProject): Promise<void> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const environment = await this.createEnvironment(project, 'Fake Environment');
    await taskAgentApi.deleteEnvironment(project.name!, environment.id!);
    environmentLogger.info('Created an deleted fake Environment');
  }

  /** Use this function to harden an environment
   * @param { string } project the project where the environment should be pushed to;
   * @param { string } environmentName the environment Name;
   * @param { string } groupName the group Name;
   * @param { GroupType } groupType the group type of the groupname (team or product);
   * @returns { TeamProject} the project itself;
   */
  public async hardenEnvironment(project: TeamProject, environmentName: string, groupName: string, groupType: GroupType): Promise<void> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const environments: EnvironmentInstance[] = await taskAgentApi.getEnvironments(project.name!, environmentName);
    const thisEnvironment = environments.find((e) => e.name === environmentName);

    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;

    const childToken = 'Environments/' + project.id + '/' + thisEnvironment!.id!;
    const securityServices = this.azureDevOpsServices.security();
    await securityServices.setPermissionOnSimpleEntity(project, ENVIRONMENT_NAMESPACE, childToken, groupName, (await securityServices.getSimpleObjectPermissions()).Environment!.OwnerRights!);
    await securityServices.setPermissionOnSimpleEntity(project, ENVIRONMENT_NAMESPACE, childToken, 'Contributors', (await securityServices.getSimpleObjectPermissions()).Environment!.ContributorRights!);
    environmentLogger.info("'" + environmentName + "' hardened.");
  }

  /** Use this function to create and harden an environment
   * @param { string } project the project where the environment should be pushed to;
   * @param { string } environmentName the environment Name;
   * @param { string } groupName the group Name;
   * @param { GroupType } groupType the group type of the groupname (team or product)
   * @returns { TeamProject} the project itself;
   */
  public async createHardenedEnvironment(project: TeamProject, environmentName: string, groupName: string, groupType: GroupType): Promise<void> {
    await this.createEnvironment(project, environmentName);
    await this.hardenEnvironment(project, environmentName, groupName, groupType);
  }
}
