import { WebApi } from 'azure-devops-node-api';
import * as ga from 'azure-devops-node-api/GitApi';
import * as pa from 'azure-devops-node-api/PolicyApi';

import fs from 'fs';

import { TeamProject } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { PolicyConfiguration } from 'azure-devops-node-api/interfaces/PolicyInterfaces';
import { GitRepository, GitRepositoryCreateOptions, GitPush, ItemContentType, VersionControlChangeType } from 'azure-devops-node-api/interfaces/GitInterfaces';

import { OnboardingServices } from './index';
import { IComplexObjectPermission } from './interfaces/IObjectPermission';
import { GroupType } from './interfaces/Enums';
import { GIT_NAMESPACE } from './SecurityNamespaces';
import { gitRepoLogger } from './logging';

export class GitRepoServices {
  private connection: WebApi;
  private azureDevOpsServices: OnboardingServices;

  constructor(connection: WebApi, azureDevOpsServices: OnboardingServices) {
    this.connection = connection;
    this.azureDevOpsServices = azureDevOpsServices;
  }

  /** Sets the permissions to a Git Reposistory
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   * @param { boolean } inheritPermissions: flag to inherrit Permissions (defaults to true)
   * @param { string } permissionFile the file that contains the permissions;
   * @param { string } groupName the group that is allowed;
   * @param { GroupType } groupType the grouptype (team of product)
   * @returns { GitRepository } the git Repository
   */
  public async hardenGitRepository(project: TeamProject, gitRepositoryName: string, groupName: string, groupType): Promise<GitRepository> {
    const thisGitRepository: GitRepository = await this.getGitRepository(project, gitRepositoryName);

    if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;

    const objectPermission: IComplexObjectPermission[] = JSON.parse(fs.readFileSync('settings/' + this.azureDevOpsServices.configuration().CONFIG_GITPERMISSIONFILE, 'utf8'));
    const parentToken = 'repoV2/' + project.id;
    const childToken = 'repoV2/' + project.id + '/' + thisGitRepository.id;

    await this.azureDevOpsServices.security().setObjectSecurityBasedOnJsonConfig(project, GIT_NAMESPACE, childToken, objectPermission, groupName);
    gitRepoLogger.info("'" + gitRepositoryName + "' git repository permissions set.");
    return project;
  }

  /** Creates a git repository and initializes it
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   * @returns { GitRepository } the git Repository
   */
  public async createGitRepostitory(project: TeamProject, gitRepositoryName: string): Promise<GitRepository> {
    const gitApi: ga.GitApi = await this.connection.getGitApi();
    const gitRepositories: GitRepository[] = await gitApi.getRepositories(project.name);
    const tempRepo = gitRepositories.find((g) => g.name === gitRepositoryName);
    if (tempRepo) throw Error("'" + gitRepositoryName + "' already exists");

    const gitRepositoryToCreate: GitRepositoryCreateOptions = {
      name: gitRepositoryName,
      project,
    };
    const thisGitRepository: GitRepository = await gitApi.createRepository(gitRepositoryToCreate);
    let gitPush: GitPush = {
      commits: [
        {
          comment: 'Initial commit.',
          changes: [
            {
              changeType: VersionControlChangeType.Add,
              item: {
                path: '/Readme.md',
              },
              newContent: {
                content: 'Repository Initialized.',
                contentType: ItemContentType.RawText,
              },
            },
          ],
        },
      ],
      refUpdates: [
        {
          name: 'refs/heads/master',
          oldObjectId: '0000000000000000000000000000000000000000',
        },
      ],
    };
    gitPush = await gitApi.createPush(gitPush, thisGitRepository.id!);
    gitRepoLogger.info("Created Git Repository '" + gitRepositoryName + "'");

    return thisGitRepository;
  }

  /** Sets 'Automatically include reviewers' for a git repository
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   * @param { string } refName the refName (like refs/heads/master) that needs to be protected;
   * @param { string } groupName the group Name that is allowed;
   * @returns { GitRepository } the git Repository
   */
  public async setCodeReviewers(project: TeamProject, gitRepositoryName: string, refName: string, groupName: string, groupType: GroupType): Promise<GitRepository> {
    const thisGitRepository: GitRepository = await this.getGitRepository(project, gitRepositoryName);

    if (groupType === GroupType.Product) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_PRODUCT_GROUP_PREFIX + groupName;
    else if (groupType === GroupType.Team) groupName = this.azureDevOpsServices.configuration().AZURE_DEVOPS_TEAM_GROUP_PREFIX + groupName;

    const groupId = await this.azureDevOpsServices.group().getGroupOriginId(project, groupName);
    const policyConfiguration: PolicyConfiguration = {
      isEnabled: true,
      isBlocking: true,
      isDeleted: false,
      type: {
        id: 'fd2167ab-b0be-447a-8ec8-39368250530e',
      },
      settings: {
        requiredReviewerIds: [groupId],
        minimumApproverCount: 1,
        creatorVoteCounts: false,
        scope: [
          {
            refName,
            matchKind: 'Exact',
            repositoryId: thisGitRepository.id,
          },
        ],
      },
    };
    const policyApi: pa.PolicyApi = await this.connection.getPolicyApi();
    const result = await policyApi.createPolicyConfiguration(policyConfiguration, project.name!);
    gitRepoLogger.info("Policy set for git Repository: '" + gitRepositoryName + "' on ref '" + refName + "'");

    return thisGitRepository;
  }

  /** Deletes a git Repository
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   */
  public async deleteGitRepository(project: TeamProject, gitRepositoryName: string): Promise<void> {
    const gitApi: ga.GitApi = await this.connection.getGitApi();
    const thisGitRepository: GitRepository = await this.getGitRepository(project, gitRepositoryName);
    await gitApi.deleteRepository(thisGitRepository.id!, project.name!);
    gitRepoLogger.info("Deleted Git Repository: '" + gitRepositoryName + "'");
  }

  /** Gets a git Repository based on name
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   */
  public async getGitRepository(project: TeamProject, gitRepositoryName: string): Promise<GitRepository> {
    const gitApi: ga.GitApi = await this.connection.getGitApi();
    const gitRepositories: GitRepository[] = await gitApi.getRepositories(project.name);
    const thisGitRepository = gitRepositories.find((g) => g.name === gitRepositoryName);
    if (!thisGitRepository) throw Error("Gitrepostory '" + gitRepositoryName + "' cannot be found.");
    return thisGitRepository;
  }

  /** Creates a project wide Git repository
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   */
  public async createProjectGitRepository(project: TeamProject, gitRepositoryName: string): Promise<TeamProject> {
    const gitRepository: GitRepository = await this.createGitRepostitory(project, gitRepositoryName);
    if (gitRepository) {
      await this.setCodeReviewers(project, gitRepositoryName, 'refs/heads/master', 'Contributors', GroupType.Project);
      return project;
    }
    throw Error('Did not create git Repository.');
  }

  /** Creates a team or project hardened Git repository
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   * @param { string } groupName the name of the group;
   * @param { GroupType } groupType the type of the group (team or product);
   */
  public async createTeamOrProjectGitRepository(project: TeamProject, gitRepositoryName: string, groupName: string, groupType: GroupType): Promise<TeamProject> {
    const gitRepository: GitRepository = await this.createGitRepostitory(project, gitRepositoryName);
    if (gitRepository) {
      await this.setCodeReviewers(project, gitRepositoryName, 'refs/heads/master', groupName, groupType);
      await this.hardenGitRepository(project, gitRepositoryName, groupName, groupType);
      return project;
    }
    throw Error("Error creating git Repository '" + gitRepositoryName + "'");
  }

  /** Hardens an existing git repo and sets the code reviewers so that it is owned.
   * @param { TeamProject } project: the teamProject;
   * @param { string } gitRepositoryName the name of the git repository;
   * @param { string } groupName the name of the group;
   * @param { GroupType } groupType the type of the group (team or product);
   */
  public async setGitRepostoryOwners(project: TeamProject, gitRepositoryName: string, groupName: string, groupType: GroupType): Promise<TeamProject> {
    const gitRepository: GitRepository = await this.getGitRepository(project, gitRepositoryName);
    if (gitRepository) {
      await this.setCodeReviewers(project, gitRepositoryName, 'refs/heads/master', groupName, groupType);
      await this.hardenGitRepository(project, gitRepositoryName, groupName, groupType);
      return project;
    }
    throw Error("Git repository '" + gitRepositoryName + "' does not exist.");
  }
}
