import { Category, CategoryServiceFactory, CategoryConfiguration, LogLevel } from 'typescript-logging';
CategoryServiceFactory.setDefaultConfiguration(new CategoryConfiguration(LogLevel.Info));

// Create categories, they will autoregister themselves, one category without parent (root) and a child category.
export const projectLogger = new Category('Project');
export const teamLogger = new Category('Team');
export const groupLogger = new Category('Group Membership');
export const graphLogger = new Category('Microsoft Graph Service');
export const securityLogger = new Category('Security');
export const pipelineLogger = new Category('Pipeline');
export const releaseLogger = new Category('Release');
export const gitRepoLogger = new Category('Git Repositories');
export const workLogger = new Category('Work');
export const policyLogger = new Category('Policy');
export const agentPoolLogger = new Category('Agent Pool');
export const libraryLogger = new Category('Library');
export const resourceLogger = new Category('Resources');
export const serviceConnectionLogger = new Category('Service Connection');
export const deploymentGroupLogger = new Category('Deployment Group');
export const userLogger = new Category('User');
export const wikiLogger = new Category('Wiki');
export const environmentLogger = new Category('Environment');
export const dashboardLogger = new Category('Dashboard');

export async function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
