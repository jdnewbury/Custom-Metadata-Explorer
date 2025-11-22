import { LightningElement, track, api } from 'lwc';
import getExplorerOptions from '@salesforce/apex/BitbucketMetadataController.getExplorerOptions';
import getMetadataDetails from '@salesforce/apex/BitbucketMetadataController.getMetadataDetails';

export default class CustomMetadataExplorer extends LightningElement {
    @track columns = [];
    @track dataRows = [];
    @track isLoading = false;
    @track error;
    @track options = [];
    @api selectedValue = '';
    @track errorMessage = '';
    @track rawJson = '';
    @track containerWidth = 0;
    
    // Store original column definitions with width percentages
    originalColumns = [];
    resizeObserver;

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
            
            // Setup resize observer after component mounts
            this.setupResizeObserver();
        } catch (error) {
            console.error('Error loading explorer options:', error);
            this.error = error.body?.message || error.message;
            this.errorMessage = 'Failed to load explorer options: ' + (error.body?.message || error.message);
        }
    }

    disconnectedCallback() {
        // Cleanup resize observer
        if (this.resizeObserver) {
            this.resizeObserver.disconnect();
        }
    }

    /**
     * Setup ResizeObserver to detect container width changes
     */
    setupResizeObserver() {
        // Wait for template to render
        setTimeout(() => {
            const container = this.template.querySelector('.custom-container');
            if (container) {
                this.containerWidth = container.offsetWidth;
                
                // Create resize observer
                this.resizeObserver = new ResizeObserver(entries => {
                    for (let entry of entries) {
                        const newWidth = entry.contentRect.width;
                        if (Math.abs(newWidth - this.containerWidth) > 10) {
                            this.containerWidth = newWidth;
                            this.recalculateColumnWidths();
                        }
                    }
                });
                
                this.resizeObserver.observe(container);
            }
        }, 100);
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
        this.columns = [];
        this.dataRows = [];

        try {
            // Call new controller method
            const response = await getMetadataDetails({ metadataType: this.selectedValue });
            console.log('Controller response:', response);

            if (!response) {
                throw new Error('No data returned from controller');
            }

            // Store original columns with width percentages
            this.originalColumns = response.columns || [];
            
            // Build datatable columns with calculated pixel widths
            this.columns = this.buildDatatableColumns(response.columns || []);
            
            // Ensure each row has a stable Id for key-field
            const rows = Array.isArray(response.data) ? response.data : [];
            this.dataRows = rows.map((r, idx) => {
                if (r && typeof r === 'object' && !('Id' in r)) {
                    return { Id: String(idx + 1), ...r };
                }
                return r;
            });

            this.rawJson = response.rawJson || '';

        } catch (error) {
            console.error('Error in handleLoad:', error);
            this.error = error.body?.message || error.message;
            this.errorMessage = error.body?.message || error.message;
        } finally {
            this.isLoading = false;
        }
    }

    /**
     * Build lightning-datatable columns with calculated pixel widths
     * @param {Array} columns - Column definitions from Apex with widthPercent
     * @returns {Array} Datatable column definitions with initialWidth
     */
    buildDatatableColumns(columns) {
        if (!columns || columns.length === 0) return [];

        // Get current container width, default to 1000px if not yet measured
        const containerWidth = this.containerWidth || 1000;
        
        // Account for padding/margins (approximately 30px on each side)
        const availableWidth = containerWidth - 60;
        
        // Separate columns with and without explicit width
        const columnsWithWidth = columns.filter(col => col.widthPercent != null);
        const columnsWithoutWidth = columns.filter(col => col.widthPercent == null);
        
        // Calculate width used by explicit percentages
        const totalExplicitPercent = columnsWithWidth.reduce((sum, col) => 
            sum + (col.widthPercent || 0), 0);
        
        // Remaining percentage for auto-sized columns
        const remainingPercent = Math.max(0, 100 - totalExplicitPercent);
        const autoWidthPercent = columnsWithoutWidth.length > 0 
            ? remainingPercent / columnsWithoutWidth.length 
            : 0;

        // Build datatable columns
        return columns.map(col => {
            const percent = col.widthPercent != null ? col.widthPercent : autoWidthPercent;
            const pixelWidth = Math.floor((percent / 100) * availableWidth);
            
            return {
                label: col.label,
                fieldName: col.fieldName,
                type: col.type || 'text',
                sortable: col.sortable !== false,
                initialWidth: Math.max(50, pixelWidth) // Minimum 50px
            };
        });
    }

    /**
     * Recalculate column widths when container resizes
     */
    recalculateColumnWidths() {
        if (this.originalColumns && this.originalColumns.length > 0) {
            console.log('Recalculating column widths for container width:', this.containerWidth);
            this.columns = this.buildDatatableColumns(this.originalColumns);
        }
    }

    get isLoadDisabled() {
        return !this.selectedValue;
    }

    get hasData() {
        return this.dataRows && this.dataRows.length > 0;
    }
}