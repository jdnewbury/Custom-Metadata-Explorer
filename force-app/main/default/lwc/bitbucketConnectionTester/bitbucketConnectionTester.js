import { LightningElement, track } from 'lwc';
import testConnection from '@salesforce/apex/BitbucketMetadataReader.testConnection';

export default class BitbucketConnectionTester extends LightningElement {
    @track isLoading = false;
    @track testResult = null;
    @track showResult = false;

    async handleTestConnection() {
        this.isLoading = true;
        this.showResult = false;
        this.testResult = null;

        try {
            const result = await testConnection();
            this.testResult = result;
            this.showResult = true;
        } catch (error) {
            this.testResult = {
                success: false,
                message: error.body?.message || error.message || 'Unknown error occurred'
            };
            this.showResult = true;
        } finally {
            this.isLoading = false;
        }
    }

    get resultClass() {
        if (!this.testResult) return '';
        return this.testResult.success 
            ? 'slds-box slds-theme_success slds-m-top_medium'
            : 'slds-box slds-theme_error slds-m-top_medium';
    }

    get resultIcon() {
        return this.testResult?.success ? 'utility:success' : 'utility:error';
    }

    get resultVariant() {
        return this.testResult?.success ? 'success' : 'error';
    }

    get hasAuthMethod() {
        return this.testResult?.authMethod;
    }
}