import { WebApi } from 'azure-devops-node-api';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';

import * as fs from 'fs';

import { OnboardingServices } from './index';

import { GroupType, GroupScope } from './interfaces/Enums';
import { ISimpleRights, IComplexObjectPermission } from './interfaces/IObjectPermission';
import { securityLogger, delay } from './logging';

const apiVersion = 'api-version=5.0';

let securityNamespaces: ISecurityNamespace[];
let simpleRights: ISimpleRights;

export class SecurityServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Adds or Merges an ACL entry to a descriptor on Project Level
   * @param { string } namespaceId: the security namespaceId;
   * @param { string } tokenprefix: the prefix for the token
   * @param { TeamProject } project: the object that needs to be altered;
   * @param { string } groupName: the member (object), that needs to be altered;
   * @param { number } allowBit: the allow bits for the acl;
   * @param { number } denyBit: the deny bits for the acl;
   * @param { boolean } merge: defaults to true, flag if the acl should be merged (when false, it is overwriten)
   */
  public async AddOrChangeAccessControlEntryOnProjectLevel(project: TeamProject, namespaceId: string, tokenprefix: string, groupName: string, allowBit: number, denyBit: number, merge: boolean = true, projectOnly: boolean = true): Promise<void> {
    const descriptor = await this.azureDevOpsServices.group().getGroupDescriptor(project, groupName, projectOnly);
    await this.AddOrChangeAccessControlEntry(namespaceId, tokenprefix + project.id, descriptor, allowBit, denyBit, merge);
  }

  /** Deletes an ACL for an object
   * @param { string } namespaceId: the security namespaceId;
   * @param { string } token: the token, that needs to be deleted;
   * @param { boolean } recurse: Removes child ACL's
   */
  public async DeleteACL(namespaceId: string, token: string, recurse: boolean = false): Promise<void> {
    const deleteUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/accesscontrollists/' + namespaceId + '?token=' + token + '&recurse=' + recurse + '&' + apiVersion;
    const result = await this.connection.rest.del(deleteUrl);

    securityLogger.debug("Deleted ACL for: '" + token + '.');
  }

  /** Disabled the Inherit Permissions flag for a token
   * @param { string } namespaceId: the security namespaceId;
   * @param { string } token: the  token, that needs to be altered;
   */
  public async disableACLInheritPermissions(namespaceId: string, token: string): Promise<void> {
    const getUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/accesscontrollists/' + namespaceId + '?token=' + token + '&' + apiVersion;
    await delay(10000);
    const acl = await (await this.connection.rest.get(getUrl)).result;
    // @ts-ignore
    acl.value[0].inheritPermissions = false;
    const createUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/accesscontrollists/' + namespaceId + '?' + apiVersion;
    const result = await this.connection.rest.create(createUrl, acl);
    if (result.statusCode !== 204) securityLogger.debug('Did not set permission for' + token + ' httpstatuscode: ' + result.statusCode);
    securityLogger.debug("Disabled Inherit Permissions for '" + token + '.');
  }

  /** Adds or Merges an ACL entry to a descriptor
   * @param { string } namespaceId: the security namespaceId;
   * @param { string } token: the object that needs to be altered;
   * @param { string } descriptor: the member (object), that needs to be altered;
   * @param { number } allowBit: the allow bits for the acl;
   * @param { number } denyBit: the deny bits for the acl;
   * @param { boolean } merge: defaults to true, flag if the acl should be merged (when false, it is overwriten)
   */
  public async AddOrChangeAccessControlEntry(namespaceId: string, token: string, descriptor: string, allowBit: number, denyBit: number, merge: boolean = true): Promise<void> {
    const requestBody = {
      token,
      merge,
      accessControlEntries: [
        {
          descriptor: 'Microsoft.TeamFoundation.Identity;' + this.decodeDescriptor(descriptor),
          allow: allowBit,
          deny: denyBit,
          extendedinfo: {},
        },
      ],
    };
    const url = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/accesscontrolentries/' + namespaceId + '?' + apiVersion;
    const result = await this.connection.rest.create(url, requestBody);
    if (result.statusCode !== 200) securityLogger.debug('Did not set permission for' + requestBody + ' httpstatuscode: ' + result.statusCode);
    securityLogger.debug("Set permission for namespace: '" + (await this.getSecurityNamespaceName(namespaceId)) + "'.");
  }

  /** Deletes an Access Control Entry
   * @param { string } namespaceId: the security namespaceId;
   * @param { string } token: the object that needs to be altered;
   * @param { string } descriptor: the member (object), that needs to be altered;
   */
  public async deleteAccessControlEntry(namespaceId: string, token: string, descriptor: string): Promise<void> {
    const deleteUrl = this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/accesscontrolentries/' + namespaceId + '?token=' + token + '&descriptors=Microsoft.TeamFoundation.Identity;' + this.decodeDescriptor(descriptor) + '&' + apiVersion;
    const result = await this.connection.rest.del(deleteUrl);
    if (result.statusCode !== 200) securityLogger.debug('Did not delete permission for' + descriptor + ' httpstatuscode: ' + result.statusCode);
    securityLogger.debug("Deleted permission for descriptor: '" + descriptor + "'.");
  }

  /** Set folder Security for builds and releases
   * @param { TeamProject } project: the project;
   * @param { string } namespaceId: the security namespaceId;
   * @param { string } repo: the folder that is created;
   * @param { string } objectSecurity: the objectPermission that is loaded
   * @param { string } permissionGroup
   */
  public async setObjectSecurityBasedOnJsonConfig(project: TeamProject, namespaceId: string, childToken: string, objectSecurity: IComplexObjectPermission[], permissionGroup?: string): Promise<void> {
    for (const role of objectSecurity) {
      let descriptor: string | undefined;
      if (role.GroupScope === GroupScope.ProjectGroup) {
        descriptor = await this.azureDevOpsServices.group().getGroupDescriptor(project, role.Group!);
      } else if (role.GroupScope === GroupScope.TeamRole && permissionGroup) {
        descriptor = await this.azureDevOpsServices.group().getGroupDescriptor(project, permissionGroup + role.Group!);
      } else if (role.GroupScope === GroupScope.Group && permissionGroup) {
        descriptor = await this.azureDevOpsServices.group().getGroupDescriptor(project, permissionGroup);
      }
      if (!descriptor) throw Error('Error getting descriptor: ' + role.GroupScope);

      const bits: ISecurityBits = await this.determineAllowAndDenyBits(namespaceId, role.Allow!, role.Deny!);
      await this.AddOrChangeAccessControlEntry(namespaceId, childToken, descriptor, bits.allowBit!, bits.denyBit!, role.Merge!);
    }
  }

  /** sets the default Project Security based on a template file
   * @param { TeamProject } project the project;
   * @returns { TeamProject} the project itself;
   */
  public async determineAllowAndDenyBits(namespaceId: string, alowArray: string[], denyArray: string[]): Promise<ISecurityBits> {
    let allowBit = 0;
    let denyBit = 0;
    if (alowArray) {
      for (const action of alowArray) {
        const bit = (await this.azureDevOpsServices.security().getSecurityAction(namespaceId, action)).bit;
        allowBit = allowBit + bit;
      }
    }
    if (denyArray) {
      for (const action of denyArray) {
        const bit = (await this.azureDevOpsServices.security().getSecurityAction(namespaceId, action)).bit;
        denyBit = denyBit + bit;
      }
    }
    return {
      allowBit,
      denyBit,
    };
  }

  /** Sets the permission of the dashboard
   * @param { TeamProject } project: the teamProject;
   * @param { string } namespaceId The namespaceId
   * @param { string } childToken The token of the subject
   * @param { string } group the group that is allowed;
   * @param { string[] } allowPermission the permission that the group is given;
   * @param { string[] } denyPermissions the permission that the group is denied;
   * @param { boolean } inheritPermissions if permissions should be inherrited, default is true
   * @returns { TeamProject} the project itself;
   */
  public async setPermissionOnSimpleEntity(project: TeamProject, namespaceId: string, childToken: string, group: string, allowPermissions: string[], inheritPermissions: boolean = true): Promise<TeamProject> {
    const descriptor = await this.azureDevOpsServices.group().getGroupDescriptor(project, group);
    let allowBit = 0;

    for (const permission of allowPermissions) {
      allowBit = allowBit + (await this.azureDevOpsServices.security().getSecurityAction(namespaceId, permission)).bit;
    }

    if (!inheritPermissions) await this.azureDevOpsServices.security().disableACLInheritPermissions(namespaceId, childToken);
    await this.azureDevOpsServices.security().AddOrChangeAccessControlEntry(namespaceId, childToken, descriptor, allowBit, 0);
    return project;
  }

  /** Decodes a descriptor to an SID
   * @param { string } encryptedDescriptor: the desciptor;
   * @returns { string } the Security Identifier;
   */
  public decodeDescriptor(encryptedDescriptor: string): string {
    // cut Off vssgp. or aad.
    const cutOff = encryptedDescriptor.indexOf('.');
    encryptedDescriptor = encryptedDescriptor.substring(cutOff, encryptedDescriptor.length);
    const decodedDecriptor = Buffer.from(encryptedDescriptor, 'base64').toString();
    return decodedDecriptor;
  }

  /** Loads the security namespaces for the organisation */
  private async loadSecurityNamespaces(): Promise<void> {
    if (!securityNamespaces) {
      // @ts-ignore
      securityNamespaces = await (await this.connection.rest.get(this.azureDevOpsServices.configuration().AZURE_DEVOPS_ORGANISATION_URL + '/_apis/securitynamespaces?' + apiVersion)).result.value;
      securityLogger.debug('Loaded security namespaces.');
    }
  }

  /** Loads the simple object permissions for the organisation */
  private async loadSimpleObjectPermissions(simpleObjectPermission: string = this.azureDevOpsServices.configuration().CONFIG_SIMPLEOBJECTPERMISSIONFILE): Promise<void> {
    if (!simpleRights) {
      simpleRights = JSON.parse(fs.readFileSync('settings/' + simpleObjectPermission, 'utf8'));
      securityLogger.debug('Loaded simple object permissions.');
    }
  }

  public async getSimpleObjectPermissions(): Promise<ISimpleRights> {
    this.loadSimpleObjectPermissions();
    return simpleRights;
  }

  /** Gets a security action to retreive the bits
   * @param { string } namespaceId: the namespaceId of the security;
   * @param { string } actionName: the action name;
   * @returns { ISecurityAction} the action that contains the security bits;
   */
  public async getSecurityAction(namespaceId: string, actionName: string): Promise<ISecurityAction> {
    await this.loadSecurityNamespaces();
    const namespace: ISecurityNamespace | undefined = securityNamespaces.find((t) => t.namespaceId! === namespaceId) ?? undefined;
    if (!namespace) throw Error("Namespace '" + namespaceId + "' cannot be found.");
    const bit = namespace.actions.find((b) => b.name! === actionName);
    if (!bit) throw Error("Action '" + actionName + "' cannot be found for namespace '" + namespaceId + "'.");
    return bit;
  }

  /** Gets a security namespaceId based on the dataspaceCategory
   * @param { string } namespaceName: the namespaceName (dataspaceCategory);
   * @returns { string} the namespaceId;
   */
  public async getSecurityNamespaceId(namespaceName: string): Promise<string> {
    await this.loadSecurityNamespaces();
    const namespace: ISecurityNamespace | undefined = securityNamespaces.find((t) => t.dataspaceCategory === namespaceName);
    return namespace!.namespaceId!;
  }

  /** Gets a security name based on the namespaceId
   * @param { string } namespaceId: the namespaceName (dataspaceCategory);
   * @returns { string} the name;
   */
  public async getSecurityNamespaceName(namespaceId: string): Promise<string> {
    await this.loadSecurityNamespaces();
    const namespace: ISecurityNamespace | undefined = securityNamespaces.find((t) => t.namespaceId! === namespaceId);
    return namespace!.name!;
  }
}

interface ISecurityNamespace {
  dataspaceCategory: string;
  displayName: string;
  elementLength: number;
  extensionType?: string;
  isRemotable: boolean;
  name: string;
  namespaceId: string;
  readPermission: number;
  separatorValue: string;
  structureValue: number;
  systemBitMask: number;
  useTokenTranslator: boolean;
  writePermission: number;
  actions: ISecurityAction[];
}

interface ISecurityAction {
  bit: number;
  displayName: string;
  name: string;
  namespaceId: string;
}

export interface ISecurityBits {
  allowBit: number;
  denyBit: number;
}
