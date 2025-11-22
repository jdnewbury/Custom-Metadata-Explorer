/**
 * Controller for customMetadataExplorer LWC
 * Provides metadata type options and retrieves metadata using BitbucketMetadataReader
 */
public with sharing class BitbucketMetadataController {
    
    /**
     * Get available metadata types for picklist
     */
    @AuraEnabled(cacheable=true)
    public static List<Map<String, String>> getExplorerOptions() {
        try {
            return MetadataTypeConfig.getActiveMetadataTypes();
        } catch (Exception e) {
            throw new AuraHandledException('Error loading metadata types: ' + e.getMessage());
        }
    }
    
    /**
     * Get metadata details including columns with width and data rows
     * @param metadataType - The DeveloperName from Bitbucket_Metadata_Type__mdt
     * @return MetadataResponse with columns (including width) and data
     */
    @AuraEnabled
    public static MetadataResponse getMetadataDetails(String metadataType) {
        try {
            // Load configuration including column widths
            MetadataTypeConfig config = MetadataTypeConfig.loadConfig(metadataType);
            
            // Get metadata from Bitbucket using existing reader
            Map<String, Object> readerResponse = BitbucketMetadataReader.getMetadataByType(metadataType);
            
            // Check for errors in response
            Boolean success = (Boolean) readerResponse.get('success');
            if (success == false) {
                String errorMsg = (String) readerResponse.get('error');
                throw new AuraHandledException(errorMsg != null ? errorMsg : 'Unknown error from metadata reader');
            }
            
            // Build response with width-aware columns
            MetadataResponse response = new MetadataResponse();
            response.columns = buildColumns(config.columns);
            
            // Extract data from reader response
            Object dataObj = readerResponse.get('data');
            if (dataObj instanceof List<Object>) {
                response.data = (List<Map<String, Object>>) dataObj;
            } else {
                response.data = new List<Map<String, Object>>();
            }
            
            // Get raw JSON if available
            String rawJson = (String) readerResponse.get('rawJson');
            response.rawJson = rawJson != null ? rawJson : '';
            
            response.totalWidth = calculateTotalWidth(config.columns);
            
            return response;
            
        } catch (Exception e) {
            throw new AuraHandledException('Error retrieving metadata: ' + e.getMessage());
        }
    }
    
    /**
     * Build column definitions for lightning-datatable with width support
     */
    private static List<ColumnDefinition> buildColumns(List<MetadataTypeConfig.ColumnConfig> configColumns) {
        List<ColumnDefinition> columns = new List<ColumnDefinition>();
        
        for (MetadataTypeConfig.ColumnConfig col : configColumns) {
            if (!col.active) continue;
            
            ColumnDefinition colDef = new ColumnDefinition();
            colDef.fieldName = col.fieldName;
            colDef.label = col.label;
            colDef.type = col.type;
            colDef.sortable = true;
            colDef.widthPercent = col.widthPercent;
            
            columns.add(colDef);
        }
        
        return columns;
    }
    
    /**
     * Calculate total width percentage specified (for validation/auto-sizing)
     */
    private static Decimal calculateTotalWidth(List<MetadataTypeConfig.ColumnConfig> columns) {
        Decimal total = 0;
        Integer columnsWithWidth = 0;
        
        for (MetadataTypeConfig.ColumnConfig col : columns) {
            if (col.active && col.widthPercent != null) {
                total += col.widthPercent;
                columnsWithWidth++;
            }
        }
        
        return total;
    }
    
    /**
     * Response wrapper for LWC
     */
    public class MetadataResponse {
        @AuraEnabled public List<ColumnDefinition> columns { get; set; }
        @AuraEnabled public List<Map<String, Object>> data { get; set; }
        @AuraEnabled public String rawJson { get; set; }
        @AuraEnabled public Decimal totalWidth { get; set; }
        
        public MetadataResponse() {
            this.columns = new List<ColumnDefinition>();
            this.data = new List<Map<String, Object>>();
        }
    }
    
    /**
     * Column definition with width percentage
     */
    public class ColumnDefinition {
        @AuraEnabled public String fieldName { get; set; }
        @AuraEnabled public String label { get; set; }
        @AuraEnabled public String type { get; set; }
        @AuraEnabled public Boolean sortable { get; set; }
        @AuraEnabled public Decimal widthPercent { get; set; } // Percentage of total width
    }
}