import * as da from 'azure-devops-node-api/DashboardApi';

import { TeamProject, WebApiTeam, TeamContext } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { Dashboard, DashboardScope } from 'azure-devops-node-api/interfaces/DashboardInterfaces';

import { WebApi } from 'azure-devops-node-api';
import { OnboardingServices } from './index';
import { DASHBOARD_NAMESPACE } from './SecurityNamespaces';
import { dashboardLogger } from './logging';

export class DashboardServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Creates a team Dashboard
   * @param { TeamProject } project The project;
   * @param { WebApiTeam } team The team;
   */
  public async createTeamDashboard(project: TeamProject, team: WebApiTeam): Promise<void> {
    const teamContext: TeamContext = { project: project.name, team: team.name };
    const dashboardApi: da.DashboardApi = await this.connection.getDashboardApi();
    const dashboard: Dashboard = { name: team.name, dashboardScope: DashboardScope.Project_Team, groupId: team.id };
    const createdDashboard = await dashboardApi.createDashboard(dashboard, teamContext);

    const childToken = '$/' + project.id + '/' + team.id;

    const securityServices = this.azureDevOpsServices.security();
    await securityServices.setPermissionOnSimpleEntity(project, DASHBOARD_NAMESPACE, childToken, team.name!, (await securityServices.getSimpleObjectPermissions()).Dashboard!.OwnerRights!);
    await securityServices.setPermissionOnSimpleEntity(project, DASHBOARD_NAMESPACE, childToken, 'Contributors', (await securityServices.getSimpleObjectPermissions()).Dashboard!.ContributorRights!);
    dashboardLogger.info("'" + team.name + "' dashboard created.");
  }

  /** Create a product Dashboard
   * @param { TeamProject } project The project where the service connection is maintained;
   * @param { string } productName The product;
   * @param { string } dashboardName The dashboard name;
   */
  public async createProductDashboard(project: TeamProject, productName: string): Promise<void> {
    const teamContext: TeamContext = { project: project.name };
    const dashboardApi: da.DashboardApi = await this.connection.getDashboardApi();
    const dashboardName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + productName;
    const dashboard: Dashboard = { name: dashboardName, dashboardScope: DashboardScope.Project };
    const createdDashboard = await dashboardApi.createDashboard(dashboard, teamContext);

    const childToken = '$/' + project.id + '/00000000-0000-0000-0000-000000000000/' + createdDashboard.id;

    const securityServices = this.azureDevOpsServices.security();
    await securityServices.setPermissionOnSimpleEntity(project, DASHBOARD_NAMESPACE, childToken, this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + productName, (await securityServices.getSimpleObjectPermissions()).Dashboard!.OwnerRights!);
    await securityServices.setPermissionOnSimpleEntity(project, DASHBOARD_NAMESPACE, childToken, 'Contributors', (await securityServices.getSimpleObjectPermissions()).Dashboard!.ContributorRights!);

    dashboardLogger.info("'" + dashboardName + "' dashboard created.");
  }
}
