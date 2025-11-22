/**
 * Configuration class for metadata types including column definitions with width support
 */
public class MetadataTypeConfig {

    public String metadataType { get; set; }
    public String displayName { get; set; }
    public String folderName { get; set; }
    public String fileExtension { get; set; }
    public String parserClass { get; set; }
    public List<ColumnConfig> columns { get; set; }

    /**
     * Column configuration including width percentage
     */
    public class ColumnConfig {
        public String fieldName { get; set; }
        public String label { get; set; }
        public String type { get; set; }
        public Integer sortOrder { get; set; }
        public Boolean active { get; set; }
        public String xmlPath { get; set; }
        public Boolean isCollection { get; set; }
        public String collectionSeparator { get; set; }
        public Decimal widthPercent { get; set; } // NEW: Width as percentage (0-100)

        public ColumnConfig() {
            this.active = true;
            this.isCollection = false;
            this.collectionSeparator = '; ';
            this.widthPercent = null; // null means auto-size
        }
    }

    /**
     * Load configuration from custom metadata for a specific metadata type
     */
    public static MetadataTypeConfig loadConfig(String metadataTypeName) {
        // Query metadata type configuration
        List<Bitbucket_Metadata_Type__mdt> types = [
            SELECT Id, MasterLabel, DeveloperName, Display_Name__c, 
                   Folder_Name__c, File_Extension__c, Parser_Class__c,
                   Entry_Element__c, Root_Element__c
            FROM Bitbucket_Metadata_Type__mdt
            WHERE DeveloperName = :metadataTypeName 
            AND Active__c = true
            LIMIT 1
        ];

        if (types.isEmpty()) {
            throw new MetadataException('Metadata type not found or inactive: ' + metadataTypeName);
        }

        Bitbucket_Metadata_Type__mdt typeConfig = types[0];

        // Query column configurations
        List<Bitbucket_Metadata_Column__mdt> columnMdts = [
            SELECT Field_Name__c, Column_Label__c, Column_Type__c, 
                   Sort_Order__c, Active__c, XML_Path__c,
                   Is_Collection__c, Collection_Separator__c,
                   Column_Width__c
            FROM Bitbucket_Metadata_Column__mdt
            WHERE Metadata_Type__r.DeveloperName = :metadataTypeName
            AND Active__c = true
            ORDER BY Sort_Order__c ASC
        ];

        // Build config object
        MetadataTypeConfig config = new MetadataTypeConfig();
        config.metadataType = typeConfig.DeveloperName;
        config.displayName = typeConfig.Display_Name__c;
        config.folderName = typeConfig.Folder_Name__c;
        config.fileExtension = typeConfig.File_Extension__c;
        config.parserClass = typeConfig.Parser_Class__c;
        config.columns = new List<ColumnConfig>();

        for (Bitbucket_Metadata_Column__mdt colMdt : columnMdts) {
            ColumnConfig col = new ColumnConfig();
            col.fieldName = colMdt.Field_Name__c;
            col.label = colMdt.Column_Label__c;
            col.type = colMdt.Column_Type__c != null ? colMdt.Column_Type__c : 'text';
            col.sortOrder = colMdt.Sort_Order__c != null ? Integer.valueOf(colMdt.Sort_Order__c) : 999;
            col.active = colMdt.Active__c;
            col.xmlPath = colMdt.XML_Path__c;
            col.isCollection = colMdt.Is_Collection__c != null ? colMdt.Is_Collection__c : false;
            col.collectionSeparator = colMdt.Collection_Separator__c != null ? colMdt.Collection_Separator__c : '; ';
            col.widthPercent = colMdt.Column_Width__c; // NEW: Get width percentage

            config.columns.add(col);
        }

        return config;
    }

    /**
     * Get all active metadata type configurations for picklist
     */
    public static List<Map<String, String>> getActiveMetadataTypes() {
        List<Map<String, String>> options = new List<Map<String, String>>();

        List<Bitbucket_Metadata_Type__mdt> types = [
            SELECT DeveloperName, Display_Name__c
            FROM Bitbucket_Metadata_Type__mdt
            WHERE Active__c = true
            ORDER BY Sort_Order__c ASC, Display_Name__c ASC
        ];

        for (Bitbucket_Metadata_Type__mdt type : types) {
            options.add(new Map<String, String>{
                'label' => type.Display_Name__c,
                'value' => type.DeveloperName
            });
        }

        return options;
    }

    public class MetadataException extends Exception {}
}
