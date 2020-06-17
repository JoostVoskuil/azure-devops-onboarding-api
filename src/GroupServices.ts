import { WebApi } from 'azure-devops-node-api';

import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { OnboardingServices } from './index';
import { SubjectType, GroupType } from './interfaces/Enums';
import { groupLogger, delay } from './logging';

const apiVersion = 'api-version=5.1-preview.1';
const vsspsApiVersion = 'api-version=5.1-preview.1';

export class GroupServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Creates a Group in a project.
   * @param { TeamProject } project: the teamProject;
   * @param { string } name: the name of the group;
   * @param { string } description: the description of the group;
   * @returns { TeamProject} the project itself;
   */
  public async createGroup(project: TeamProject, name: string, description: string): Promise<TeamProject> {
    const projectDescriptor = await this.azureDevOpsServices.project().getProjectDescriptor(project);
    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups?scopeDescriptor=' + projectDescriptor + '&' + apiVersion;
    const request = {
      displayName: name,
      description,
    };
    const response = await this.connection.rest.create(descriptorsUrl, request);
    groupLogger.info("Created Group '" + name + "' for project " + project.name + '.');
    delay(2000);
    return project;
  }

  /** Creates a Security Group in a project and add the group to Contributors
   * @param { TeamProject } project: the teamProject;
   * @param { string } name: the name of the group;
   * @param { string } description: the description of the group;
   * @returns { TeamProject} the project itself;
   */
  public async createSecurityGroupAndAddToContributors(project: TeamProject, name: string, description: string): Promise<TeamProject> {
    await this.createGroup(project, this.azureDevOpsServices.configuration().AZURE_DEVOPS_SECURITY_GROUP_PREFIX + name, description);
    await this.addMemberToGroup(project, this.azureDevOpsServices.configuration().AZURE_DEVOPS_SECURITY_GROUP_PREFIX + name, 'Contributors');
    return project;
  }

  /** Creates a Product Group in a project.
   * @param { TeamProject } project: the teamProject;
   * @param { string } name: the name of the group;
   * @param { string } description: the description of the group;
   * @returns { TeamProject} the project itself;
   */
  public async createProductGroup(project: TeamProject, name: string, description: string): Promise<TeamProject> {
    return await this.createGroup(project, this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + name, description);
  }

  /** Deletes a Group in a project.
   * @param { TeamProject } project: the teamProject;
   * @param { string } groupName: the name of the group;
   * @returns { TeamProject} the project itself;
   */
  public async deleteGroup(project: TeamProject, groupName: string): Promise<TeamProject> {
    const descriptor = await this.getGroupDescriptor(project, groupName);
    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups/' + descriptor + '?' + apiVersion;
    const response = await this.connection.rest.del(descriptorsUrl);
    groupLogger.info("Deleted Group '" + groupName + "' for project '" + project.name + "'");
    return project;
  }

  /** Get the descriptor of a group
   * @param { TeamProject } project: the teamProject;
   * @param { string } name: the name of the group;
   * @param { boolean } projectOnly: When true, search only in current project
   * @returns { string } the descriptor (encryped SID) of the group;
   */
  public async getGroupDescriptor(project: TeamProject, name: string, projectOnly: boolean = true): Promise<string> {
    const projectDescriptor = await this.azureDevOpsServices.project().getProjectDescriptor(project);
    let descriptorsUrl: string;
    if (projectOnly) descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups?scopeDescriptor=' + projectDescriptor + '&' + apiVersion;
    else descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups?' + apiVersion;
    try {
      const response = await (await this.connection.rest.get(descriptorsUrl)).result;
      // @ts-ignore
      return response.value.filter((n) => n.displayName === name)[0].descriptor;
    } catch (err) {
      throw Error("Group '" + name + "' does not exists.");
    }
  }

  /** Get the originId of a group
   * @param { TeamProject } project: the teamProject;
   * @param { string } name: the name of the group;
   * @returns { boolean } true if the group exists;
   */
  public async checkIfGroupExists(project: TeamProject, groupName: string, groupType: GroupType): Promise<boolean> {
    const projectDescriptor = await this.azureDevOpsServices.project().getProjectDescriptor(project);
    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups?scopeDescriptor=' + projectDescriptor + '&' + apiVersion;

    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;

    try {
      const response = await (await this.connection.rest.get(descriptorsUrl)).result;
      // @ts-ignore
      const group = response.value.filter((n) => n.displayName === groupName)[0];
      if (group) return true;
      return false;
    } catch (err) {
      return false;
    }
  }

  /** Get the originId of a group
   * @param { TeamProject } project: the teamProject;
   * @param { string } name: the name of the group;
   * @returns { string } the originId of the group;
   */
  public async getGroupOriginId(project: TeamProject, name: string): Promise<string> {
    const projectDescriptor = await this.azureDevOpsServices.project().getProjectDescriptor(project);
    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups?scopeDescriptor=' + projectDescriptor + '&' + apiVersion;
    try {
      const response = await (await this.connection.rest.get(descriptorsUrl)).result;
      // @ts-ignore
      return response.value.filter((n) => n.displayName === name)[0].originId;
    } catch (err) {
      throw Error("Could not get originId from group '" + name + "'");
    }
  }

  /** Get the originId of a group based on descriptor
   * @param { TeamProject } project: the teamProject;
   * @param { string } descriptor: the descriptor of the group;
   * @returns { string } the originId of the group;
   */
  public async getGroupOriginIdBasedOnDescriptor(project: TeamProject, descriptor: string): Promise<string> {
    const projectDescriptor = await this.azureDevOpsServices.project().getProjectDescriptor(project);
    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups?scopeDescriptor=' + projectDescriptor + '&' + apiVersion;
    try {
      const response = await (await this.connection.rest.get(descriptorsUrl)).result;
      // @ts-ignore
      return response.value.filter((n) => n.descriptor === descriptor)[0].originId;
    } catch (err) {
      throw Error("Could not get originId from group '" + descriptor + "'");
    }
  }

  /** Adds a AAD Group as a Member of a Azure DevOps Group
   * @param { TeamProject } project: the teamProject;
   * @param { string } azureDevOpsGroupName: the name group in Azure DevOps;
   * @param { string } aadGroupName: the name of the Azure Active Directory Group;
   * @returns { TeamProject} the project itself;
   */
  public async addAADGroupMemberToAzureDevOpsGroup(project: TeamProject, azureDevOpsGroupName: string, aadGroupName: string): Promise<TeamProject> {
    const groupDescriptor = await this.getGroupDescriptor(project, azureDevOpsGroupName);
    const aadGroupId = await this.azureDevOpsServices.microsoftGraphServices().getAADGroupIdBasedOnDisplayName(aadGroupName);

    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/groups?groupDescriptors=' + groupDescriptor + '&' + apiVersion;
    const request = {
      originId: aadGroupId,
    };

    const response = await this.connection.rest.create(descriptorsUrl, request);
    groupLogger.info("Added AADGroup '" + aadGroupName + "' to Azure DevOps Group '" + azureDevOpsGroupName + "'.");

    return project;
  }

  /** Adds a AAD Group as a Member of a Azure DevOps Group
   * @param { TeamProject } project: the teamProject;
   * @param { string } memberGroup: the name of the group that will be added to the target group;
   * @param { string } targetGroup: the name of the group that is the target;
   * @returns { TeamProject} the project itself;*
   */
  public async addMemberToGroup(project: TeamProject, memberGroup: string, targetGroup: string): Promise<TeamProject> {
    const memberDescriptor = await this.getGroupDescriptor(project, memberGroup);
    const targetDescriptor = await this.getGroupDescriptor(project, targetGroup);

    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/memberships/' + memberDescriptor + '/' + targetDescriptor + '?' + apiVersion;
    const result = await this.connection.rest.replace(descriptorsUrl, undefined);
    if (result.statusCode !== 201) throw Error("Could not add '" + memberGroup + "' as a member to Group '" + targetGroup + "'.");
    groupLogger.info("'" + memberGroup + "' is now a member of group '" + targetGroup + "'.");
    return project;
  }

  /** Delete a member from a group
   * @param { TeamProject } project: the teamProject;
   * @param { string } groupName: the name of the group for where the subject is member of;
   * @param { string } memberName: the name of the subject;
   * @param { SubjectType } memberType: user or group;
   */
  public async deleteMemberFromGroup(project: TeamProject, groupName: string, memberName: string, memberType: SubjectType): Promise<void> {
    const groupId = await this.getGroupOriginId(project, groupName);
    let memberId: string | undefined;
    if (memberType === SubjectType.User) memberId = await (await this.azureDevOpsServices.user().getUserProperties(project, memberName)).id;
    if (memberType === SubjectType.Group) memberId = await this.azureDevOpsServices.group().getGroupOriginId(project, memberName);
    if (!memberId) throw Error('Could not find member');
    const descriptorsUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSAEX_URL + '/_apis/GroupEntitlements/' + groupId + '/members/' + memberId + '?' + apiVersion;
    const response = await this.connection.rest.del(descriptorsUrl);

    groupLogger.info("'" + memberName + "' is deleted from group '" + groupName + "'");
  }

  /** Check the number of members of a group
   * @param { TeamProject } project: the teamProject;
   * @param { string } groupName: the name of the group for where the subject is member of;
   * @param { GroupType } groupType: the group Type
   * @returns { number } the number of members
   */
  async getNumberOfMembersOfGroup(project: TeamProject, groupName: string, groupType: GroupType): Promise<number> {
    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;

    const groupDescriptor = await this.getGroupDescriptor(project, groupName);
    // Get the members of the group
    const membersUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_VSSPS_URL + '/_apis/graph/Memberships/' + groupDescriptor + '?direction=down&' + vsspsApiVersion;
    const response = await this.connection.rest.get(membersUrl);

    // @ts-ignore
    return response.result.value.length;
  }
}
