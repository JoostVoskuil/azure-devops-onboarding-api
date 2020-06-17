import { WebApi } from 'azure-devops-node-api';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { OnboardingServices } from './index';
import { resourceLogger } from './logging';

const apiVersion = 'api-version=6.0-preview.1';

export class AuthorizeResourcesServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }
  /** Use this function to authorize a resource and set the 'Grand access permission to all pipelines' flag to true.
   * This is great for shared resources like Lbraries, Agent Pools and Service Connections that are shared accross the organisation.
   * This can be used to push a organisation wide Library. It should not be used to update Libraries that can be altered by Contributors since it will overwrite it.
   * @param { TeamProject } project The projectName where the resource is that needs to be authorized;
   * @param { string } resourceId The resourceId of the resource;
   * @param { string } resourceName The resourceName of the resource;
   * @param { string } resourceType The resourceType. Allowed values are: 'endpoint', 'queue', 'variablegroup';
   * @returns { TeamProject} the project itself;
   */
  public async authorizeResource(project: TeamProject, resourceId: string, resourceName: string, resoureType: string): Promise<TeamProject> {
    const authorizedresourcesUri = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + project.name + '/_apis/build/authorizedresources?' + apiVersion;
    const serviceConnectionToBeAuthorized: IAuthorizedResource = {
      authorized: true,
      id: resourceId,
      name: resourceName,
      type: resoureType,
    };
    // Rebuild Authorized resources list
    // Note: This feels strange that we need to replace everything instead of only authorize that one resource.
    // @ts-ignore
    const authorizedresources: IAuthorizedResource[] = await (await this.connection.rest.get(authorizedresourcesUri)).result.value;
    authorizedresources.push(serviceConnectionToBeAuthorized);
    const result = await this.connection.rest.update(authorizedresourcesUri, authorizedresources);
    resourceLogger.info("Authorized resource '" + resourceName + "'");
    return project;
  }

  /** Use this function to de-authorize a resource and set the 'Grand access permission to all pipelines' flag to false.
   * @param { TeamProject } project The projectName where the resource is that needs to be authorized;
   * @param { string } resourceId The resourceId of the resource;
   * @param { string } resourceName The resourceName of the resource;
   * @param { string } resourceType The resourceType. Allowed values are: 'endpoint', 'queue', 'variablegroup';
   * @returns { TeamProject} the project itself;
   */
  public async deAuthorizeResource(project: TeamProject, resourceId: string, resourceName: string, resoureType: string): Promise<TeamProject> {
    const authorizedresourcesUri = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/' + project.name + '/_apis/build/authorizedresources?' + apiVersion;
    // Rebuild Authorized resources list
    // Note: This feels strange that we need to replace everything instead of only authorize that one resource.
    // @ts-ignore
    let authorizedresources: IAuthorizedResource[] = await (await this.connection.rest.get(authorizedresourcesUri)).result.value!;
    // find resourceToBeDeautorized
    const resourceToBeDeauthorized: IAuthorizedResource | undefined = authorizedresources.find((a) => a.id! === resourceId && a.type! === resoureType);
    if (resourceToBeDeauthorized) {
      // remove resoure
      authorizedresources = authorizedresources.filter((a) => a !== resourceToBeDeauthorized);
      // Change flag and push
      resourceToBeDeauthorized.authorized = false;
      authorizedresources.push(resourceToBeDeauthorized);
      const result = await this.connection.rest.update(authorizedresourcesUri, authorizedresources);
      resourceLogger.info("De-authorized resource '" + resourceName + "'");
    } else {
      resourceLogger.info("De-authorized resource '" + resourceName + "' could not be found.");
    }
    return project;
  }
}

interface IAuthorizedResource {
  authorized: boolean;
  id: string;
  name?: string;
  type: string;
}
