import { LightningElement, track } from 'lwc';
import testBitbucketConnection from '@salesforce/apex/BitbucketMetadataReader.testBitbucketConnection';
import getFileContent from '@salesforce/apex/BitbucketMetadataReader.getFileContent';
import listRepositoryContents from '@salesforce/apex/BitbucketMetadataReader.listRepositoryContents';
import getBranches from '@salesforce/apex/BitbucketMetadataReader.getBranches';
import getCommits from '@salesforce/apex/BitbucketMetadataReader.getCommits';
import getPullRequests from '@salesforce/apex/BitbucketMetadataReader.getPullRequests';

export default class BitbucketConnectionTester extends LightningElement {
    @track connectionStatus = '';
    @track isLoading = false;
    @track error = '';
    @track fileContent = '';
    @track repositoryContents = [];
    @track branches = [];
    @track commits = [];
    @track pullRequests = [];

    // Test connection to Bitbucket
    handleTestConnection() {
        this.isLoading = true;
        this.error = '';
        this.connectionStatus = '';

        testBitbucketConnection()
            .then(result => {
                this.connectionStatus = result;
                if (result.startsWith('ERROR') || result.startsWith('CREDENTIAL_ERROR')) {
                    this.error = result;
                }
            })
            .catch(error => {
                this.error = 'Error: ' + (error.body?.message || error.message || 'Unknown error');
                this.connectionStatus = 'Connection failed';
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    // Get file content from Bitbucket
    handleGetFile() {
        this.isLoading = true;
        this.error = '';

        // Example: Get README.md file
        getFileContent({ filePath: 'README.md' })
            .then(result => {
                if (result.success) {
                    this.fileContent = result.content;
                } else {
                    this.error = result.error;
                }
            })
            .catch(error => {
                this.error = 'Error: ' + (error.body?.message || error.message || 'Unknown error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    // List repository contents
    handleListContents() {
        this.isLoading = true;
        this.error = '';

        listRepositoryContents({ path: '' })
            .then(result => {
                if (result.success) {
                    this.repositoryContents = result.data?.values || [];
                } else {
                    this.error = result.error;
                }
            })
            .catch(error => {
                this.error = 'Error: ' + (error.body?.message || error.message || 'Unknown error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    // Get branches
    handleGetBranches() {
        this.isLoading = true;
        this.error = '';

        getBranches()
            .then(result => {
                if (result.success) {
                    this.branches = result.branches || [];
                } else {
                    this.error = result.error;
                }
            })
            .catch(error => {
                this.error = 'Error: ' + (error.body?.message || error.message || 'Unknown error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    // Get recent commits
    handleGetCommits() {
        this.isLoading = true;
        this.error = '';

        getCommits({ branch: 'main', limitSize: 5 })
            .then(result => {
                if (result.success) {
                    this.commits = result.commits || [];
                } else {
                    this.error = result.error;
                }
            })
            .catch(error => {
                this.error = 'Error: ' + (error.body?.message || error.message || 'Unknown error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    // Get pull requests
    handleGetPullRequests() {
        this.isLoading = true;
        this.error = '';

        getPullRequests({ state: 'OPEN' })
            .then(result => {
                if (result.success) {
                    this.pullRequests = result.pullRequests || [];
                } else {
                    this.error = result.error;
                }
            })
            .catch(error => {
                this.error = 'Error: ' + (error.body?.message || error.message || 'Unknown error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    get hasError() {
        return this.error && this.error.length > 0;
    }

    get hasFileContent() {
        return this.fileContent && this.fileContent.length > 0;
    }

    get hasRepositoryContents() {
        return this.repositoryContents && this.repositoryContents.length > 0;
    }

    get hasBranches() {
        return this.branches && this.branches.length > 0;
    }

    get hasCommits() {
        return this.commits && this.commits.length > 0;
    }

    get hasPullRequests() {
        return this.pullRequests && this.pullRequests.length > 0;
    }
}