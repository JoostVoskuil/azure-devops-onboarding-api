import { WebApi } from 'azure-devops-node-api';
import * as ta from 'azure-devops-node-api/TaskAgentApi';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { TaskAgentQueue } from 'azure-devops-node-api/interfaces/TaskAgentInterfaces';
import { OnboardingServices } from './index';
import { agentPoolLogger } from './logging';
import { AGENTQUEUE_NAMESPACE } from './SecurityNamespaces';

export class AgentPoolServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Sets the 'Grand access to all pipelines' flag for all agent Queues and add contributors.
   * @param { TeamProject } project: the teamProject;
   * @returns { TeamProject} the project itself;
   */
  public async setDefaultSecurityForAllPipelines(project: TeamProject): Promise<TeamProject> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const allAgentQueues: TaskAgentQueue[] = await taskAgentApi.getAgentQueues(project.id);

    for (const agentQueue of allAgentQueues) {
      if (agentQueue.name !== "Azure Pipelines") {
        await this.grandAccessToAllPipelinesForAgentQueue(project, agentQueue.name!);
      }
      await this.addContributorsToAgentQueue(project, agentQueue.name!);
    }
    return project;
  }

  /** Sets the 'Grand access to all pipelines' flag for specified agent Queues and add contributors.
   * @param { TeamProject } project: the teamProject;
   * @param { string } agentQueueNames agentQueue names, ',' seperated 
   * @param { boolean } onlyGrandPermissions if true, the pipelines are only granted, if false also the contributors will be added.
   * @returns { TeamProject} the project itself;
   */
  public async setDefaultSecurityForSpecifiedPipelines(project: TeamProject, agentQueueNames: string, onlyGrandPermissions: boolean): Promise<TeamProject> {
    const agentQueueNamesArray = agentQueueNames.split(',');
    for (const agentQueueName of agentQueueNamesArray) {
      await this.grandAccessToAllPipelinesForAgentQueue(project, agentQueueName);
      if (onlyGrandPermissions) {
        await this.addContributorsToAgentQueue(project, agentQueueName);
      }
    }
    return project;
  }

  /** Sets the 'Grand access to all pipelines' flag for the specified agent Queue.
   * @param { TeamProject } project: the teamProject;
   * @returns { TeamProject} the project itself;
   */
  public async grandAccessToAllPipelinesForAgentQueue(project: TeamProject, agentQueueName: string): Promise<TeamProject> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const allAgentQueues: TaskAgentQueue[] = await taskAgentApi.getAgentQueuesByNames([agentQueueName], project.name);
    if (allAgentQueues.length === 0) throw Error("'" + agentQueueName + ' cannot be found.');
    await this.azureDevOpsServices.authorizeResource().authorizeResource(project, allAgentQueues[0].id!.toString(), allAgentQueues[0].name!, 'queue');
    agentPoolLogger.info("Authorized agent queue '" + agentQueueName + "'");
    return project;
  }

  /** Use this function to harden an environment
   * @param { string } project the project where the environment should be pushed to;
   * @param { string } environmentName the environment Name;
   * @param { string } groupName the group Name;
   * @param { GroupType } groupType the group type of the groupname (team or product);
   * @returns { TeamProject} the project itself;
   */
  public async addContributorsToAgentQueue(project: TeamProject, agentQueueName: string): Promise<void> {
    const taskAgentApi: ta.ITaskAgentApi = await this.connection.getTaskAgentApi();
    const allAgentQueues: TaskAgentQueue[] = await taskAgentApi.getAgentQueuesByNames([agentQueueName], project.name);
    if (allAgentQueues.length === 0) throw Error("'" + agentQueueName + ' cannot be found.');

    const token = 'AgentQueues/' + project.id + '/' + allAgentQueues[0].id;
    const securityServices = this.azureDevOpsServices.security();
    await securityServices.setPermissionOnSimpleEntity(project, AGENTQUEUE_NAMESPACE, token, 'Contributors', ["View", "Use"]);
    agentPoolLogger.info("Add Contributors as User to Agent queue '" + agentQueueName + "'");
  }
}
