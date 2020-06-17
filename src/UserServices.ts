import { WebApi } from 'azure-devops-node-api';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';

import { userLogger } from './logging';
import { OnboardingServices } from './index';
import { GroupType } from './interfaces/Enums';

const vsspsApiVersion = 'api-version=5.1-preview.1';

export class UserServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Gets the properties of the supplied user
   * @param { string } projectName the project Name
   * @param { string } principleName the principleName of the user
   * @returns { UserProperties} the properties of the user
   */
  public async getUserProperties(project: TeamProject, principleName: string): Promise<UserProperties> {
    // Get some basic data from this user
    const entitlementUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSAEX_URL + "/_apis/userentitlements?$filter=(name eq '" + principleName + "')";
    const userProperties: UserProperties = { principleName };
    const response = await this.connection.rest.get(entitlementUrl);
    // @ts-ignore
    userProperties.id = response.result.members[0].id;
    // @ts-ignore
    userProperties.descriptor = response.result.members[0].user.descriptor;
    // @ts-ignore
    userProperties.originId = response.result.members[0].user.originId;
    userProperties.projectDescriptor = await this.azureDevOpsServices.project().getProjectDescriptor(project);
    userProperties.projectName = project.name;
    userProperties.projectId = project.id;
    userLogger.info("Retreived properties for user '" + principleName + "'");
    return userProperties;
  }

  /** Determines if user is a (indirect) member of a group. This is fully AAD compabible and uses MS Graph
   * It travels down the group memberships
   * @param { string } project the project
   * @param { string } principleName the principleName of the user
   * @param { string } groupName the name of the group
   * @param { GroupType } groupType the group type (team or product)
   * @returns { boolean} if the user is (indirect) member of the group
   */
  public async isUserMemberOfGroup(project: TeamProject, principleName: string, groupName: string, groupType: GroupType): Promise<boolean> {
    // Get user properties and full groupname
    const userProperties: UserProperties = await this.getUserProperties(project, principleName);
    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;

    // get the descriptor of the group
    const groupDescriptor = await this.azureDevOpsServices.group().getGroupDescriptor(project, groupName);

    const isMember: boolean = await this.isObjectMemberOfGroup(project, groupDescriptor, userProperties.originId!);
    userLogger.debug("'" + principleName + "' member of '" + groupName + "' is " + isMember);
    return isMember;
  }

  /** Determines if user is a (indirect) member of a group. This is fully AAD compabible and uses MS Graph
   * It travels down the group memberships
   * NOTE: WILL FAIL WHEN THERE ARE USERS MEMBER OF A AZURE DEVOPS GROUP
   * @param { TeamProject } project the teamproject
   * @param { string } groupDescriptor the descriptor of the group (with prefix like aad)
   * @param { string } originId the originId of the user (decoded descriptor)
   * @returns { boolean} if the user is (indirect) member of the group
   */
  async isObjectMemberOfGroup(project: TeamProject, groupDescriptor: string, originId: string): Promise<boolean> {
    // Get the members of the group
    const membersUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/Memberships/' + groupDescriptor + '?direction=down&' + vsspsApiVersion;
    const response = await this.connection.rest.get(membersUrl);
    let isMember: boolean = false;

    // Iterate through the members
    // @ts-ignore
    if (response.result.value.length > 0) {
      // @ts-ignore
      for (const member of response.result.value) {
        if (!isMember) {
          // @ts-ignore
          const childDescriptor: string = member.memberDescriptor;
          const cutOff: number = childDescriptor.indexOf('.');
          // When we found an AAD group, we call MS graph to detect the users
          if (childDescriptor.substr(0, cutOff) === 'aadgp') {
            const groupObjectId = await this.azureDevOpsServices.group().getGroupOriginIdBasedOnDescriptor(project, childDescriptor)
            if (await this.azureDevOpsServices.microsoftGraphServices().CheckIfObjectIdIsDirectMemberOfObject(groupObjectId, originId)) {
              isMember = true;
            }
          }
          else if (childDescriptor.substr(0, cutOff) === 'vssgp') {
            // Call this function recursive, we hit an Azure DevOps group member
            isMember = await this.isObjectMemberOfGroup(project, childDescriptor, originId);
          }
        }
      }
      return isMember;
    }
    else {
      return isMember;
    }
  }
}

export interface UserProperties {
  principleName?: string;
  id?: string;
  descriptor?: string;
  originId?: string;
  projectName?: string;
  projectId?: string;
  projectDescriptor?: string;
}
