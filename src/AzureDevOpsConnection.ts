import * as azdev from 'azure-devops-node-api';
import { IRequestOptions } from 'azure-devops-node-api/interfaces/common/VsoBaseInterfaces';
import { IConfiguration } from './interfaces/IConfiguration';

export class AzureDevOpsConnection {
  private static instance: AzureDevOpsConnection;
  private connection: azdev.WebApi;

  private constructor(configuration: IConfiguration) {
    const collectionUri = configuration.AZURE_DEVOPS_ORGANISATION_URL;
    const accessTokenHandler = azdev.getPersonalAccessTokenHandler(configuration.AZUREDEVOPS_PAT!);

    const requestOptions: IRequestOptions = {
      socketTimeout: configuration.AZURE_DEVOPS_API_TIMEOUT,
      maxRetries: configuration.AZURE_DEVOPS_API_MAX_RETRIES,
      allowRetries: configuration.AZURE_DEVOPS_API_ALLOWRETRIES,
    };

    this.connection = new azdev.WebApi(collectionUri, accessTokenHandler, requestOptions);
  }

  public static getInstance(configuration: IConfiguration): AzureDevOpsConnection {
    if (!AzureDevOpsConnection.instance) {
      AzureDevOpsConnection.instance = new AzureDevOpsConnection(configuration);
    }

    return AzureDevOpsConnection.instance;
  }

  public static getConnection(configuration: IConfiguration): azdev.WebApi {
    return this.getInstance(configuration).connection;
  }
}
