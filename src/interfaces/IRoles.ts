export interface IRoles {
  SecurityRoles: TeamRole[];
}

export interface TeamRole {
  PostFixName: string;
  SecurityGroup: string;
  Description: string;
}
