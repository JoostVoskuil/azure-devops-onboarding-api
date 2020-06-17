import { WebApi } from 'azure-devops-node-api';

import * as ta from 'azure-devops-node-api/TaskAgentApi';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { DeploymentMachineGroup, DeploymentGroup, DeploymentGroupCreateParameter } from 'azure-devops-node-api/interfaces/TaskAgentInterfaces';

import { OnboardingServices } from './index';
import { GroupType } from './interfaces/Enums';
import { deploymentGroupLogger } from './logging';
import { AGENTQUEUE_NAMESPACE } from './SecurityNamespaces';

export class DeploymentGroupServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Use this function to harden a deploymentgroup
   * @param { string } project the project where the Library should be pushed to;
   * @param { string } variableGroupName the variablegroup Name;
   * @param { string } groupName the group Name;
   * @param { GroupType } groupType the group type of the groupname (team or product)
   * @returns { TeamProject} the project itself;
   */
  public async hardenDeploymentGroup(project: TeamProject, deploymentGroupName: string, groupName: string, groupType: GroupType): Promise<void> {
    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const allDeploymentMachineGroup: DeploymentMachineGroup[] = await taskAgentApi.getDeploymentMachineGroups(project.name!, deploymentGroupName);

    const token = 'MachineGroups/' + project.id + '/' + allDeploymentMachineGroup[0].id;
    const securityServices = this.azureDevOpsServices.security();
    await securityServices.setPermissionOnSimpleEntity(project, AGENTQUEUE_NAMESPACE, token, groupName, (await securityServices.getSimpleObjectPermissions()).DeploymentGroup!.OwnerRights!);
    await securityServices.setPermissionOnSimpleEntity(project, AGENTQUEUE_NAMESPACE, token, 'Contributors', (await securityServices.getSimpleObjectPermissions()).DeploymentGroup!.ContributorRights!);
    deploymentGroupLogger.info("Hardened Deployment group '" + deploymentGroupName + "'.");
  }

  /** Creates a fake deployment group and deletes it
   * @param { TeamProject } project: the teamProject;
   * @param { string } deploymentGroupName the name of the agentQueue;
   * @param { string } allowedGroup the group that is allowed;
   * @param { string } permission the permission that the allowedGroup is given (see 'Library' security namespace );
   * @param { boolean } inheritPermissions: flag to inherritPermissions (defaults to true)
   * @returns { DeploymentMachineGroup } the deploymentGroup;
   */
  public async createAndDeleteFakeDeploymentGroup(project: TeamProject): Promise<void> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const deploymentGroupParameters: DeploymentGroupCreateParameter = {
      name: 'Fake Deployment Group',
    };

    const deploymentGroup: DeploymentGroup = await taskAgentApi.addDeploymentGroup(deploymentGroupParameters, project.name!);
    await taskAgentApi.deleteDeploymentGroup(project.name!, deploymentGroup.id!);

    deploymentGroupLogger.info('Created and Deleted fake deployment group.');
  }
}
