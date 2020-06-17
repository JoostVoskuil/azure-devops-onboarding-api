import { WebApi } from 'azure-devops-node-api';
import { OnboardingServices } from './index';
import { graphLogger, delay } from './logging';
import axios = require('axios');
import qs = require('qs');

export class MicrosoftGraphServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** gets a bearer token to use for Microsoft Graph
   * @returns { string } the Microsoft Graph bearer token;
   */
  public async getMicrosoftGraphToken(): Promise<string> {
    const postData = {
      client_id: this.azureDevOpsServices.configuration().MS_GRAPH_APP_ID,
      scope: this.azureDevOpsServices.configuration().MS_GRAPH_SCOPE,
      client_secret: this.azureDevOpsServices.configuration().MS_GRAPH_APP_SECRET,
      grant_type: 'client_credentials',
    };

    try {
      axios.default.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

      const response = await axios.default.post(this.azureDevOpsServices.configuration().MS_GRAPH_TOKEN_ENDPOINT, qs.stringify(postData));
      // @ts-ignore
      const bearerToken = response.data.access_token;
      if (!bearerToken) throw Error('Bearer token is empty.');
      graphLogger.debug('Logged in to Microsoft Graph to fetch AAD Groups');
      return bearerToken;
    } catch (err) {
      throw Error('Could not get Bearer token from Microsoft Graph. ' + err.message);
    }
  }

  /** gets a AAD Group ID based on the group Name
   * @param { string } groupName the name of the group;
   * @returns { string } the id of the AAD group;
   */
  public async getAADGroupIdBasedOnDisplayName(groupName: string): Promise<string | undefined> {
    const bearerToken = await this.getMicrosoftGraphToken();
    if (!bearerToken) throw Error('Did not got Bearer Token.');
    axios.default.defaults.headers.get['Content-Type'] = 'application/x-www-form-urlencoded';
    /* tslint:disable:no-string-literal */
    axios.default.defaults.headers.get['Authorization'] = 'Bearer ' + bearerToken;
    /* tslint:enable:no-string-literal */

    const response = await axios.default.get("https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '" + groupName + "'");
    // @ts-ignore
    if (response.data.value.length > 0) return response.data.value[0].id;
    return undefined;
  }

  /** gets or creates (if not exist) an Azure Active Directory Group
   * @param { string } name the name of the group;
   * @param { string } description the description of the group;
   * @returns { string } the id of the AAD group;
   */
  public async CreateOrGetAADGroup(name: string, description?: string): Promise<string> {
    let groupId = await this.getAADGroupIdBasedOnDisplayName(name);
    description = 'Group automatically created for Azure DevOps. Context is ' + description;
    if (!groupId) {
      groupId = await this.createAADGroup(name, description);
      await delay(5000); // Not so nice, but there is a delay
    }

    return groupId;
  }

  /** creates an Azure Active Directory Group
   * @param { string } name the name of the group;
   * @param { string } description the description of the group;
   * @returns { string } the id of the AAD group;
   */
  public async createAADGroup(name: string, description: string): Promise<string> {
    try {
      const bearerToken = await this.getMicrosoftGraphToken();
      if (!bearerToken) throw Error('Did not got Bearer Token.');

      const postData = {
        displayName: name,
        description,
        MailEnabled: false,
        mailNickname: name,
        SecurityEnabled: true,
      };

      axios.default.defaults.headers.post['Content-Type'] = 'application/json';
      /* tslint:disable:no-string-literal */
      axios.default.defaults.headers.post['Authorization'] = 'Bearer ' + bearerToken;
      /* tslint:enable:no-string-literal */
      const response = await axios.default.post('https://graph.microsoft.com/v1.0/groups', postData);
      // @ts-ignore
      const id = response.data.id;
      if (!id) throw Error('Could not get id for AAD Group ' + name);
      graphLogger.info("AAD Group '" + name + "' created.");
      return id;
    } catch (err) {
      throw Error('Could not get id for AAD Group ' + name + ': ' + err.message);
    }
  }

  /** Determines if an childObjectId is an direct member of the groupObjectId
   * @param { string } groupOriginId the objectId of the AAD group;
   * @param { string } userOriginId the child (user) object Id;
   * @returns { boolean } returns true if the child is a direct member of the group
   */
  public async CheckIfObjectIdIsDirectMemberOfObject(groupOriginId: string, userOriginId: string): Promise<boolean> {
    const bearerToken = await this.getMicrosoftGraphToken();
    if (!bearerToken) throw Error('Did not got Bearer Token.');
    axios.default.defaults.headers.get['Content-Type'] = 'application/x-www-form-urlencoded';
    /* tslint:disable:no-string-literal */
    axios.default.defaults.headers.get['Authorization'] = 'Bearer ' + bearerToken;
    /* tslint:enable:no-string-literal */

    const response = await axios.default.get('https://graph.microsoft.com/v1.0/groups/' + groupOriginId + '/members');
    // If there are members
    if (response.data.value.length > 0) {
      // loop through members and detect if the childObject is there
      for (const object of response.data.value) {
        if (object.id === userOriginId) return true;
      }
      return false;
    }
    return false;
  }
}
