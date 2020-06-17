import { WebApi } from 'azure-devops-node-api';
import * as ca from 'azure-devops-node-api/CoreApi';
import { TeamProject, TeamProjectReference } from 'azure-devops-node-api/interfaces/CoreInterfaces';

import { OnboardingServices } from './index';
import { GroupType } from './interfaces/Enums';
import { serviceConnectionLogger } from './logging';
import { SERVICE_ENDPOINT_NAMESPACE } from './SecurityNamespaces';

const apiVersion = 'api-version=6.0-preview.4';

export class ServiceConnectionServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Use this function to share a public service connection to all project and authorize it with 'Grand access permission to all pipelines'
   * This can be used to share a organisation wide Service Connection.
   * @param { TeamProject } communityProject The project where the service connection is maintained;
   * @param { string } serviceConnectionName The name of the service connection that needs to be shared;
   */
  public async shareAndAuthorizeServiceConnectionWithAllProjects(communityProject: TeamProject, serviceConnectionName: string): Promise<void> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    const allProjects: TeamProjectReference[] = await coreApi.getProjects();

    // Construct array of ServiceEndpointProjectReference to hold all the projects
    let serviceEndpointProjectReferences: IServiceEndpointProjectReference[] = [];
    for (const teamProject of allProjects) {
      const serviceEndpointProjectReference: IServiceEndpointProjectReference = {
        name: serviceConnectionName,
        description: serviceConnectionName,
        projectReference: {
          id: teamProject.id!,
          name: teamProject.name!,
        },
      };
      serviceEndpointProjectReferences.push(serviceEndpointProjectReference);
    }
    const queryUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + communityProject.name + '/_apis/serviceendpoint/endpoints?endpointNames=' + serviceConnectionName + '&' + apiVersion;
    const queryResult = await this.connection.rest.get(queryUrl);
    // Get the teamprojects that already have this and filter those out.
    // @ts-ignore
    const knownReferences: IServiceEndpointProjectReference[] = queryResult.result.value[0].serviceEndpointProjectReferences;
    for (const ref of knownReferences) {
      serviceEndpointProjectReferences = serviceEndpointProjectReferences.filter((obj) => obj.projectReference.id !== ref.projectReference.id);
    }
    if (serviceEndpointProjectReferences.length > 0) {
      // @ts-ignore
      const endPointId = endPoint.result.value[0].id;
      const updateUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + communityProject.name + '/_apis/serviceendpoint/endpoints/' + endPointId + '?' + apiVersion;
      const result = await this.connection.rest.update(updateUrl, serviceEndpointProjectReferences);

      for (const ref of serviceEndpointProjectReferences) {
        const targetProject: TeamProject = await coreApi.getProject(ref.projectReference.id);
        // @ts-ignore
        await authorizeResource(targetProject, result.result.value[0].id, result.result.value[0].name, 'endpoint');
        await this.hardenSharedServiceConnection(targetProject, serviceConnectionName);
        serviceConnectionLogger.info("Shared and authorized resource '" + serviceConnectionName + "' to project '" + ref.projectReference.name + "'");
      }
    } else {
      serviceConnectionLogger.info("'" + serviceConnectionName + "' was already shared to all projects");
    }
  }

  /** Use this function to share a public service connection to a project and authorize it with 'Grand access to all pipelines'
   * This can be used to share a organisation wide Service Connection.
   * @param { TeamProject } sharedResourcesProject The project where the service connection is maintained;
   * @param { string } serviceConnectionName The name of the service connection that needs to be shared;
   * @param { TeamProject } targetProject The project where the service connections needs to be shared with;
   * @returns { TeamProject} the targetproject itself;
   */
  public async shareAndAuthorizeServiceConnection(sharedResourcesProject: TeamProject, serviceConnectionName: string, targetProject: TeamProject): Promise<TeamProject> {
    // Construct array of ServiceEndpointProjectReference and add only the target Project
    const serviceEndpointProjectReferences: IServiceEndpointProjectReference[] = [];
    const serviceEndpointProjectReference: IServiceEndpointProjectReference = {
      name: serviceConnectionName,
      description: serviceConnectionName,
      projectReference: {
        id: targetProject.id!,
        name: targetProject.name!,
      },
    };
    serviceEndpointProjectReferences.push(serviceEndpointProjectReference);

    const queryUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + sharedResourcesProject.name + '/_apis/serviceendpoint/endpoints?endpointNames=' + serviceConnectionName + '&' + apiVersion;
    const queryResult = await this.connection.rest.get(queryUrl);
    // Check if this service connection is already shared with the targetProject, if not it must be added
    if (!serviceEndpointProjectReferences.find((obj) => obj.projectReference.id !== serviceEndpointProjectReference.projectReference.id)) {
      // @ts-ignore
      const endPointId = queryResult.result.value[0].id;
      // @ts-ignore
      const endPointName = queryResult.result.value[0].name;

      const updateUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + sharedResourcesProject.name + '/_apis/serviceendpoint/endpoints/' + endPointId + '?' + apiVersion;
      const result = await this.connection.rest.update(updateUrl, serviceEndpointProjectReferences);
      // @ts-ignore
      await this.azureDevOpsServices.authorizeResource().authorizeResource(targetProject, endPointId, endPointName, 'endpoint');
      await this.hardenSharedServiceConnection(targetProject, serviceConnectionName);
      serviceConnectionLogger.info("Shared and authorized resource '" + serviceConnectionName + "' to project '" + targetProject.name + "'");
    } else {
      serviceConnectionLogger.info("'" + serviceConnectionName + "' was already shared to project '" + targetProject.name + "'");
    }
    return targetProject;
  }

  /** Use this function to harden a shared Service Connection. The 'Project Valid Users' can use this service connection
   * @param { TeamProject } project The project where the service connection is maintained;
   * @param { string } serviceConnectionName The name of the service connection that needs to be altered;
   * @returns { TeamProject} the project itself;
   */
  public async hardenSharedServiceConnection(project: TeamProject, serviceConnectionName: string): Promise<TeamProject> {
    const queryUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + project.name + '/_apis/serviceendpoint/endpoints?endpointNames=' + serviceConnectionName + '&' + apiVersion;
    const queryResult = await this.connection.rest.get(queryUrl);
    // @ts-ignore
    const endPointId = queryResult.result.value[0].id;

    const token = 'endpoints/' + project.id + '/' + endPointId;
    await this.azureDevOpsServices.security().setPermissionOnSimpleEntity(project, SERVICE_ENDPOINT_NAMESPACE, token, 'Project Valid Users', ['Use', 'ViewEndpoint']);
    return project;
  }

  /** Use this function to harden a Product Service Connection. The specified group is allowed to Administer the connection.
   * @param { TeamProject } project The project where the service connection is maintained;
   * @param { string } serviceConnectionName The name of the service connection that needs to be altered;
   * @param { string } groupName The name of group (Product or Team)
   * @param { GroupType } groupType The type of group (Product or Team)
   * @returns { TeamProject} the project itself;
   */
  public async hardenServiceConnection(project: TeamProject, serviceConnectionName: string, groupName: string, groupType: GroupType): Promise<TeamProject> {
    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;

    const queryUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + project.name + '/_apis/serviceendpoint/endpoints?endpointNames=' + serviceConnectionName + '&' + apiVersion;
    const queryResult = await this.connection.rest.get(queryUrl);
    // @ts-ignore
    const endPointId = queryResult.result.value[0].id;

    const token = 'endpoints/' + project.id + '/' + endPointId;
    const securityServices = this.azureDevOpsServices.security();
    await securityServices.setPermissionOnSimpleEntity(project, SERVICE_ENDPOINT_NAMESPACE, token, groupName, (await securityServices.getSimpleObjectPermissions()).ServiceConnection!.OwnerRights!);
    await securityServices.setPermissionOnSimpleEntity(project, SERVICE_ENDPOINT_NAMESPACE, token, 'Contributors', (await securityServices.getSimpleObjectPermissions()).ServiceConnection!.ContributorRights!);
    await this.azureDevOpsServices.authorizeResource().deAuthorizeResource(project, endPointId, serviceConnectionName, 'endpoint');

    serviceConnectionLogger.info("'" + serviceConnectionName + "' is hardened for group '" + groupName + "'");
    return project;
  }

  /** Creates a fake service connection in order to create the endpoint creator/administrator group
   * @param { TeamProject } project the teamproject
   * @returns { TeamProject} the project itself;
   */
  public async createAndDeleteFakeEndpoint(project: TeamProject): Promise<TeamProject> {
    const createUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/serviceendpoint/endpoints?' + apiVersion;
    fakeEndPoint.serviceEndpointProjectReferences[0].projectReference.id = project.id!;
    fakeEndPoint.serviceEndpointProjectReferences[0].projectReference.name = project.name!;

    const createResult = await this.connection.rest.create(createUrl, fakeEndPoint);
    // @ts-ignore
    const deleteUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/serviceendpoint/endpoints/' + createResult.result.id + '?projectIds=' + project.id + '&deep=true&' + apiVersion;
    const deleteResult = await this.connection.rest.del(deleteUrl);
    serviceConnectionLogger.info('Created and deleted fake Service Connection');
    return project;
  }
}

interface IServiceEndpointProjectReference {
  name: string;
  description: string;
  projectReference: {
    id: string;
    name: string;
  };
}

const fakeEndPoint = {
  data: {
    subscriptionId: '1272a66f-e2e8-4e88-ab43-487409186c3f',
    subscriptionName: 'subscriptionName',
    environment: 'AzureCloud',
    scopeLevel: 'Subscription',
    creationMode: 'Manual',
  },
  name: 'MyNewARMServiceEndpoint',
  type: 'AzureRM',
  url: 'https://management.azure.com/',
  authorization: {
    parameters: {
      tenantid: '1272a66f-e2e8-4e88-ab43-487409186c3f',
      serviceprincipalid: '1272a66f-e2e8-4e88-ab43-487409186c3f',
      authenticationType: 'spnKey',
      serviceprincipalkey: 'SomePassword',
    },
    scheme: 'ServicePrincipal',
  },
  isShared: false,
  isReady: true,
  serviceEndpointProjectReferences: [
    {
      projectReference: {
        id: '',
        name: '',
      },
      name: 'FakeEndPoint2',
    },
  ],
};
