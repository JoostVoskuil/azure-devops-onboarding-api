import * as ra from 'azure-devops-node-api/ReleaseApi';

import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { Folder as ReleaseFolder, ReleaseDefinition } from 'azure-devops-node-api/interfaces/ReleaseInterfaces';

import fs from 'fs';

import { WebApi } from 'azure-devops-node-api';
import { OnboardingServices } from './index';
import { IComplexObjectPermission } from './interfaces/IObjectPermission';
import { RELEASE_NAMESPACE } from './SecurityNamespaces';
import { releaseLogger } from './logging';

export class ReleaseServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** creates a Releases Folder
   * @param { TeamProject } project the project
   * @param { string } name the name of the folder
   * @param { string } permissionFile the permission file where the permissions are specified
   * @returns { TeamProject} the project itself;
   */
  public async createReleaseFolderAndSetPermissions(project: TeamProject, name: string, permissionFile: string): Promise<TeamProject> {
    const releaseApi: ra.ReleaseApi = await this.connection.getReleaseApi();
    const releaseFolder: ReleaseFolder = { path: name };
    const createdReleaseFolder = await releaseApi.createFolder(releaseFolder, project.name!, name);
    
    const objectPermission: IComplexObjectPermission[] = JSON.parse(fs.readFileSync('settings/' + permissionFile, 'utf8'));
    const childToken = project.id + '/' + name;

    await this.azureDevOpsServices.security().setObjectSecurityBasedOnJsonConfig(project, RELEASE_NAMESPACE, childToken, objectPermission, name);
    releaseLogger.info("'" + name + "' release folder created.");
    return project;
  }

  /** Creates a fake release in order to create the security group and release administrators
   * @param { TeamProject } project the teamproject
   * @returns { TeamProject} the project itself;
   */
  public async createAndDeleteFakeReleaseDefinition(project: TeamProject): Promise<TeamProject> {
    const releaseApi: ra.ReleaseApi = await this.connection.getReleaseApi();
    const releasedef: ReleaseDefinition = await releaseApi.createReleaseDefinition(releaseDefinition, project.name!);
    await releaseApi.deleteReleaseDefinition(project.name!, releasedef.id!, project.name!);
    releaseLogger.info('Created and deleted fake Release');
    return project;
  }
}

const releaseDefinition: ReleaseDefinition = {
  description: 'Fake Release Definition',
  artifacts: undefined,
  name: 'Fake Release',
  environments: [
    {
      name: 'fake',
      retentionPolicy: {
        daysToKeep: 30,
        releasesToKeep: 3,
        retainBuild: true,
      },
      preDeployApprovals: {
        approvals: [
          {
            rank: 1,
            isAutomated: true,
            isNotificationOn: false,
            id: 0,
          },
        ],
      },
      postDeployApprovals: {
        approvals: [
          {
            rank: 1,
            isAutomated: true,
            isNotificationOn: false,
            id: 0,
          },
        ],
      },
      deployPhases: [
        {
          rank: 1,
          phaseType: 2,
          name: 'Agentless job',
          refName: undefined,
          workflowTasks: [
            {
              environment: {},
              taskId: '28782b92-5e8e-4458-9751-a71cd1492bae',
              version: '1.*',
              name: 'Delay by 0 minutes',
              refName: '',
              enabled: true,
              alwaysRun: false,
              continueOnError: false,
              timeoutInMinutes: 0,
              definitionType: 'task',
              overrideInputs: {},
              condition: 'succeeded()',
              inputs: {
                delayForMinutes: '0',
              },
            },
          ],
        },
      ],
    },
  ],
};
