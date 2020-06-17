import { GroupScope } from './Enums';

export interface IComplexObjectPermission {
  Group?: string;
  GroupScope: GroupScope;
  ExtraNotes?: string;
  Merge?: boolean;
  Allow?: string[];
  Deny?: string[];
}

export interface ISimpleRights {
  Library?: ISimpleObjectPermission;
  Dashboard?: ISimpleObjectPermission;
  Environment?: ISimpleObjectPermission;
  ServiceConnection?: ISimpleObjectPermission;
  DeploymentGroup?: ISimpleObjectPermission;
}

export interface ISimpleObjectPermission {
  OwnerRights: string[];
  ContributorRights: string[];
}
