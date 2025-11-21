/**
 * @description Configuration class that reads metadata type definitions from Custom Metadata.
 *              This allows adding/modifying metadata types without code deployment.
 *
 *              Uses two Custom Metadata Types:
 *              - Bitbucket_Metadata_Type__mdt: Defines each metadata type
 *              - Bitbucket_Metadata_Column__mdt: Defines columns for each type
 *
 * @author Jason Newbury
 * @date 2025
 */
public class MetadataTypeConfig {

    public static final String SF_NS = 'http://soap.sforce.com/2006/04/metadata';

    // Cache for loaded configurations
    private static Map<String, TypeConfig> configCache;
    private static List<Map<String, String>> optionsCache;

    /**
     * @description Configuration for a metadata type (loaded from Custom Metadata)
     */
    public class TypeConfig {
        public String developerName { get; set; }
        public String displayName { get; set; }
        public String folderName { get; set; }
        public String fileExtension { get; set; }
        public String parserClass { get; set; }
        public String rootElement { get; set; }
        public String entryElement { get; set; }
        public List<ColumnConfig> columns { get; set; }

        public TypeConfig() {
            this.columns = new List<ColumnConfig>();
        }

        /**
         * @description Convert columns to TableColumnDTO list for LWC datatable
         */
        public List<CustomMetadataExplorerController.TableColumnDTO> getTableColumns() {
            List<CustomMetadataExplorerController.TableColumnDTO> tableCols = new List<CustomMetadataExplorerController.TableColumnDTO>();
            for (ColumnConfig col : this.columns) {
                if (col.active) {
                    CustomMetadataExplorerController.TableColumnDTO dto = new CustomMetadataExplorerController.TableColumnDTO();
                    dto.label = col.columnLabel;
                    dto.fieldName = col.fieldName;
                    dto.type = col.columnType;
                    tableCols.add(dto);
                }
            }
            return tableCols;
        }
    }

    /**
     * @description Column configuration (loaded from Custom Metadata)
     */
    public class ColumnConfig {
        public String columnLabel { get; set; }
        public String fieldName { get; set; }
        public String xmlPath { get; set; }
        public String columnType { get; set; }
        public Integer sortOrder { get; set; }
        public Boolean active { get; set; }
        public Boolean isCollection { get; set; }
        public String collectionSeparator { get; set; }

        public ColumnConfig() {
            this.columnType = 'text';
            this.active = true;
            this.isCollection = false;
            this.collectionSeparator = '; ';
        }
    }

    /**
     * @description Get configuration for a specific metadata type
     * @param metadataType Developer name of the metadata type
     * @return TypeConfig with all settings and columns
     */
    public static TypeConfig getConfig(String metadataType) {
        loadConfigCache();

        if (!configCache.containsKey(metadataType)) {
            throw new MetadataConfigException('Unknown metadata type: ' + metadataType + 
                '. Please create a Bitbucket_Metadata_Type__mdt record with this Developer Name.');
        }
        return configCache.get(metadataType);
    }

    /**
     * @description Get all available metadata type options for dropdown
     * @return List of options with label and value for LWC combobox
     */
    public static List<Map<String, String>> getAllOptions() {
        loadConfigCache();
        return optionsCache;
    }

    /**
     * @description Check if a metadata type is supported
     */
    public static Boolean isSupported(String metadataType) {
        loadConfigCache();
        return configCache.containsKey(metadataType);
    }

    /**
     * @description Get list of all supported metadata type keys
     */
    public static Set<String> getSupportedTypes() {
        loadConfigCache();
        return configCache.keySet();
    }

    /**
     * @description Clear cache to force reload from Custom Metadata
     *              Useful after configuration changes
     */
    public static void clearCache() {
        configCache = null;
        optionsCache = null;
    }

    /**
     * @description Load configurations from Custom Metadata into cache
     */
    private static void loadConfigCache() {
        if (configCache != null) {
            return;
        }

        configCache = new Map<String, TypeConfig>();
        optionsCache = new List<Map<String, String>>();

        // Query all active metadata types with their columns
        List<Bitbucket_Metadata_Type__mdt> metadataTypes = [
            SELECT 
                DeveloperName,
                Display_Name__c,
                Folder_Name__c,
                File_Extension__c,
                Parser_Class__c,
                Root_Element__c,
                Entry_Element__c,
                Sort_Order__c,
                (SELECT 
                    Column_Label__c,
                    Field_Name__c,
                    XML_Path__c,
                    Column_Type__c,
                    Sort_Order__c,
                    Active__c,
                    Is_Collection__c,
                    Collection_Separator__c
                 FROM Columns__r
                 WHERE Active__c = true
                 ORDER BY Sort_Order__c ASC)
            FROM Bitbucket_Metadata_Type__mdt
            WHERE Active__c = true
            ORDER BY Sort_Order__c ASC, Display_Name__c ASC
        ];

        for (Bitbucket_Metadata_Type__mdt mdt : metadataTypes) {
            // Build TypeConfig
            TypeConfig config = new TypeConfig();
            config.developerName = mdt.DeveloperName;
            config.displayName = mdt.Display_Name__c;
            config.folderName = mdt.Folder_Name__c;
            config.fileExtension = mdt.File_Extension__c;
            config.parserClass = mdt.Parser_Class__c;
            config.rootElement = mdt.Root_Element__c;
            config.entryElement = mdt.Entry_Element__c;

            // Build column configs
            for (Bitbucket_Metadata_Column__mdt colMdt : mdt.Columns__r) {
                ColumnConfig col = new ColumnConfig();
                col.columnLabel = colMdt.Column_Label__c;
                col.fieldName = colMdt.Field_Name__c;
                col.xmlPath = colMdt.XML_Path__c;
                col.columnType = String.isNotBlank(colMdt.Column_Type__c) ? colMdt.Column_Type__c : 'text';
                col.sortOrder = colMdt.Sort_Order__c != null ? colMdt.Sort_Order__c.intValue() : 100;
                col.active = colMdt.Active__c;
                col.isCollection = colMdt.Is_Collection__c;
                col.collectionSeparator = String.isNotBlank(colMdt.Collection_Separator__c) 
                    ? colMdt.Collection_Separator__c 
                    : '; ';
                config.columns.add(col);
            }

            // Add to cache
            configCache.put(mdt.DeveloperName, config);

            // Add to options list
            optionsCache.add(new Map<String, String>{
                'label' => mdt.Display_Name__c + ' (Bitbucket)',
                'value' => 'bitbucket_' + mdt.DeveloperName
            });
        }

        // If no custom metadata found, add a helpful message
        if (configCache.isEmpty()) {
            System.debug(LoggingLevel.WARN, 
                'No Bitbucket_Metadata_Type__mdt records found. ' +
                'Please create metadata type configurations in Setup → Custom Metadata Types.'
            );
        }
    }

    /**
     * @description Custom exception for configuration errors
     */
    public class MetadataConfigException extends Exception {}
}