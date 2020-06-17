import { GroupScope } from './Enums';

export interface IProjectPermission {
  Group?: string;
  GroupScope: GroupScope;
  Namespaces: IProjectPermissionNamespace[];
}

export interface IProjectPermissionNamespace {
  NamespaceId: string;
  NamespaceDescription?: string;
  TokenPrefix?: string;
  Allow?: string[];
  Deny?: string[];
}
