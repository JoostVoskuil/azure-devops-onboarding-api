import * as pa from 'azure-devops-node-api/PolicyApi';
import * as ca from 'azure-devops-node-api/CoreApi';
import { TeamProject, TeamProjectReference } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { PolicyConfiguration } from 'azure-devops-node-api/interfaces/PolicyInterfaces';

import fs from 'fs';

import { WebApi } from 'azure-devops-node-api';
import { OnboardingServices } from './index';
import { policyLogger } from './logging';

export class PolicyServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Apply policies specified in 'ProjectPolicies.json' to given project.
   * Note: this does not work with revisions, it just replaces policies.
   * @param { TeamProject } project the project;
   * @returns { TeamProject} the project itself;
   */
  public async applyOrReplaceGitPolicies(project: TeamProject): Promise<TeamProject> {
    const policyApi: pa.PolicyApi = await this.connection.getPolicyApi();
    const currentPolicies: PolicyConfiguration[] = await policyApi.getPolicyConfigurations(project.name!);
    for (const currentPolicy of currentPolicies) {
      await policyApi.deletePolicyConfiguration(project.name!, currentPolicy.id!);
    }

    const policies: IPolicy[] = JSON.parse(fs.readFileSync('settings/' + this.azureDevOpsServices.configuration().CONFIG_GITPOLICYFILE, 'utf8'));
    for (const policy of policies) {
      const policyConfiguration: PolicyConfiguration = {
        isEnabled: true,
        isBlocking: true,
        isDeleted: false,
        type: {
          id: policy.type,
        },
        settings: policy.settings,
      };
      const result = await policyApi.createPolicyConfiguration(policyConfiguration, project.name!);
      policyLogger.info("Policy set for project: '" + project.name + "' for type '" + policy.description + "'");
    }
    return project;
  }

  /** apply policies specified in to all projects */
  public async applyOrRefreshPoliciesForAllProjects(): Promise<void> {
    const coreApi: ca.CoreApi = await this.connection.getCoreApi();
    const allProjects: TeamProjectReference[] = await coreApi.getProjects();

    for (const project of allProjects) {
      await this.applyOrReplaceGitPolicies(project);
    }
  }
}

interface IPolicy {
  type: string;
  description: string;
  settings: object;
}
