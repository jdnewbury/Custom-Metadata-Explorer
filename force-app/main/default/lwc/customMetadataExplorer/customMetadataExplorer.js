import { LightningElement, track, api } from 'lwc';
import getExplorerOptions from '@salesforce/apex/CustomMetadataExplorerController.getExplorerOptions';
import getMetadataDetails from '@salesforce/apex/CustomMetadataExplorerController.getMetadataDetails';

export default class CustomMetadataExplorer extends LightningElement {
    @track headers = [];
    @track dataRows = [];
    @track columns = [];
    @track isLoading = false;
    @track error;
    @track rowCount = 0;
    @track options = [];
    @api selectedValue = '';
    @track errorMessage = '';
    @track rawJson = '';
    objectName = '';
    iconUrl = '/resource/CME_Icon/cme_icon.svg';

    async connectedCallback() {
        try {
            console.log('Fetching explorer options...');
            const options = await getExplorerOptions();
            console.log('Received explorer options:', options);
            this.options = options.map(option => ({
                label: option.label,
                value: option.value
            }));
            console.log('Options set successfully');
        } catch (error) {
            console.error('Error loading explorer options:', error);
            this.error = error.body?.message || error.message;
            this.errorMessage = 'Failed to load explorer options: ' + (error.body?.message || error.message);
        }
    }

    handleObjectNameChange(event) {
        this.objectName = event.target.value;
    }

    handleChange(event) {
        this.selectedValue = event.detail.value;
    }

    async handleLoad() {
        if (!this.selectedValue) {
            return;
        }

        this.isLoading = true;
        this.error = null;
        this.errorMessage = '';
        this.rawJson = '';
        this.headers = [];
        this.dataRows = [];

        try {
            // Call controller which normalizes columns/data
            const dto = await getMetadataDetails({ developerName: this.selectedValue });
            console.log('Controller DTO:', dto);

            if (!dto) {
                throw new Error('No data returned from controller');
            }

            // Bind columns and data ready for lightning-datatable
            this.columns = dto.columns || [];
            // Ensure each row has a stable Id for key-field
            const rows = Array.isArray(dto.data) ? dto.data : [];
            this.dataRows = rows.map((r, idx) => {
                // if row has no Id, synthesize one
                if (r && typeof r === 'object' && !('Id' in r)) {
                    return { Id: String(idx + 1), ...r };
                }
                return r;
            });

            this.headers = (dto.columns || []).map(c => c.label);
            this.rowCount = this.dataRows.length;
            this.rawJson = dto.rawJson || '';

        } catch (error) {
            console.error('Error in handleLoad:', error);
            this.error = error.body?.message || error.message;
            this.errorMessage = error.body?.message || error.message;
        } finally {
            this.isLoading = false;
        }
    }

    get isLoadDisabled() {
        return !this.selectedValue;
    }

    get hasData() {
        return this.dataRows && this.dataRows.length > 0;
    }
}