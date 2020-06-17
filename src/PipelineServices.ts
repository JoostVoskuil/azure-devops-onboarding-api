import fs from 'fs';

import { WebApi } from 'azure-devops-node-api';
import * as ba from 'azure-devops-node-api/BuildApi';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { Folder as BuildFolder } from 'azure-devops-node-api/interfaces/BuildInterfaces';

import { OnboardingServices } from './index';
import { IComplexObjectPermission } from './interfaces/IObjectPermission';
import { BUILD_NAMESPACE } from './SecurityNamespaces';
import { pipelineLogger } from './logging';

export class PipelineServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }
  /** creates a Build/Pipelines Folder
   * @param { TeamProject } project the project;
   * @param { string } name the name of the folder;
   * @param { string } permissionFile the permission file where the permissions are specified;
   * @returns { TeamProject} the project itself;
   */
  public async createBuildFolderAndSetPermissions(project: TeamProject, name: string, permissionFile: string): Promise<TeamProject> {
    const buildApi: ba.BuildApi = await this.connection.getBuildApi();
    const buildFolder: BuildFolder = { path: name };
    const createdBuildFolder = await buildApi.createFolder(buildFolder, project.name!, name);
    const objectPermission: IComplexObjectPermission[] = JSON.parse(fs.readFileSync('settings/' + permissionFile, 'utf8'));
    const childToken = project.id + '/' + name;

    await this.azureDevOpsServices.security().setObjectSecurityBasedOnJsonConfig(project, BUILD_NAMESPACE, childToken, objectPermission, name);
    pipelineLogger.info("'" + name + "' pipeline folder created.");
    return project;
  }
}
