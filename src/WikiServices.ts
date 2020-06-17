import { WebApi } from 'azure-devops-node-api';
import * as wa from 'azure-devops-node-api/WikiApi';
import { WikiV2, WikiCreateParametersV2, WikiType } from 'azure-devops-node-api/interfaces/WikiInterfaces';

import { wikiLogger } from './logging';
import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { OnboardingServices } from './index';

export class WikiServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** initializes a project wiki
   * @param { TeamProject } project the project
   */
  public async initializeProjectWiki(project: TeamProject): Promise<void> {
    const gitRepository = await this.azureDevOpsServices.gitRepo().createGitRepostitory(project, project.name + ' wiki');
    const wikiApi: wa.WikiApi = await this.connection.getWikiApi();
    const wikiCreateParams: WikiCreateParametersV2 = {
      name: project.name + ' wiki',
      type: WikiType.CodeWiki,
      projectId: project.id,
      repositoryId: gitRepository.id,
      mappedPath: '/',
      version: {
        version: 'master',
      },
    };
    const wiki: WikiV2 = await wikiApi.createWiki(wikiCreateParams, project.name);
    wikiLogger.info("Created Wiki for project '" + project.name + "'");
  }
}
