import fs from 'fs';
import { WebApi } from 'azure-devops-node-api';

import { GroupServices } from './GroupServices';
import { AuthorizeResourcesServices } from './AuthorizeResourcesServices';
import { MicrosoftGraphServices } from './MicrosoftGraphService';
import { PipelineServices } from './PipelineServices';
import { PolicyServices } from './PolicyServices';
import { ProjectServices } from './ProjectServices';
import { ReleaseServices } from './ReleaseServices';
import { TeamServices } from './TeamServices';
import { ServiceConnectionServices } from './ServiceConnectionServices';
import { SecurityServices } from './SecurityServices';
import { GitRepoServices } from './GitRepoServices';
import { UserServices } from './UserServices';
import { DeploymentGroupServices } from './DeploymentGroupServices';
import { WikiServices } from './WikiServices';
import { EnvironmentServices } from './EnvironmentService';
import { LibraryServices } from './LibraryServices';
import { IConfiguration } from './interfaces/IConfiguration';
import { DashboardServices } from './DashboardServices';
import { AgentPoolServices } from './AgentPoolServices';
import { AzureDevOpsConnection } from './AzureDevOpsConnection';

export class OnboardingServices {
  private connection: WebApi;
  private thisConfiguration: IConfiguration;

  constructor(AZUREDEVOPS_PAT: string, AZUREDEVOPS_PAT_OWNER: string, MS_GRAPH_APP_SECRET: string, TESTING: boolean = false) {
    this.thisConfiguration = JSON.parse(fs.readFileSync('settings/configuration.json', 'utf8'));
    this.thisConfiguration.AZURE_DEVOPS_ORGANISATION_URL = this.thisConfiguration.AZURE_DEVOPS_ORGANISATION_URL + this.thisConfiguration.AZURE_DEVOPS_ORGANISATION;
    this.thisConfiguration.AZURE_DEVOPS_ORGANISATION_VSSPS_URL = this.thisConfiguration.AZURE_DEVOPS_ORGANISATION_VSSPS_URL + this.thisConfiguration.AZURE_DEVOPS_ORGANISATION;
    this.thisConfiguration.AZURE_DEVOPS_ORGANISATION_VSAEX_URL = this.thisConfiguration.AZURE_DEVOPS_ORGANISATION_VSAEX_URL + this.thisConfiguration.AZURE_DEVOPS_ORGANISATION;
    this.thisConfiguration.MS_GRAPH_TOKEN_ENDPOINT = this.thisConfiguration.MS_GRAPH_TOKEN_ENDPOINT + this.thisConfiguration.MS_GRAPH_DIRECTORY_ID + this.thisConfiguration.MS_GRAPH_TOKEN_POSTFIX;
    this.thisConfiguration.AZUREDEVOPS_PAT = AZUREDEVOPS_PAT;
    this.thisConfiguration.AZUREDEVOPS_PAT_OWNER = AZUREDEVOPS_PAT_OWNER;
    this.thisConfiguration.MS_GRAPH_APP_SECRET = MS_GRAPH_APP_SECRET;
    this.thisConfiguration.TESTING = TESTING;
    this.connection = AzureDevOpsConnection.getConnection(this.thisConfiguration);
  }

  public configuration(): IConfiguration {
    return this.thisConfiguration;
  }
  public group(): GroupServices {
    return new GroupServices(this.connection, this);
  }
  public agentPool(): AgentPoolServices {
    return new AgentPoolServices(this.connection, this);
  }
  public authorizeResource(): AuthorizeResourcesServices {
    return new AuthorizeResourcesServices(this.connection, this);
  }
  public microsoftGraphServices(): MicrosoftGraphServices {
    return new MicrosoftGraphServices(this.connection, this);
  }
  public pipeline(): PipelineServices {
    return new PipelineServices(this.connection, this);
  }
  public policy(): PolicyServices {
    return new PolicyServices(this.connection, this);
  }
  public project(): ProjectServices {
    return new ProjectServices(this.connection, this);
  }
  public release(): ReleaseServices {
    return new ReleaseServices(this.connection, this);
  }
  public team(): TeamServices {
    return new TeamServices(this.connection, this);
  }
  public serviceConnection(): ServiceConnectionServices {
    return new ServiceConnectionServices(this.connection, this);
  }
  public security(): SecurityServices {
    return new SecurityServices(this.connection, this);
  }
  public gitRepo(): GitRepoServices {
    return new GitRepoServices(this.connection, this);
  }
  public user(): UserServices {
    return new UserServices(this.connection, this);
  }
  public deploymentGroup(): DeploymentGroupServices {
    return new DeploymentGroupServices(this.connection, this);
  }
  public dashboard(): DashboardServices {
    return new DashboardServices(this.connection, this);
  }
  public wiki(): WikiServices {
    return new WikiServices(this.connection, this);
  }
  public environment(): EnvironmentServices {
    return new EnvironmentServices(this.connection, this);
  }
  public library(): LibraryServices {
    return new LibraryServices(this.connection, this);
  }
}
