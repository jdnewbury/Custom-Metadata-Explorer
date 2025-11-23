/**
 * Controller for customMetadataExplorer LWC
 * Self-contained solution for retrieving metadata from Bitbucket with column width support
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
            
            // Get parser instance
            IMetadataParser parser = MetadataParserFactory.getParser(
                config.parserClass != null ? config.parserClass : 'GenericMetadataParser'
            );
            
            // Get files from Bitbucket
            List<String> files = getMetadataFiles(config.folderName, config.fileExtension);
            
            // Parse all files
            List<Map<String, Object>> allData = new List<Map<String, Object>>();
            List<String> jsonParts = new List<String>();
            
            for (String filePath : files) {
                String xmlContent = getFileContent(filePath);
                if (String.isNotBlank(xmlContent)) {
                    // Parse using the actual parser interface signature: parse(xmlContent, filePath, config)
                    List<Map<String, Object>> parsedData = parser.parse(xmlContent, filePath, config.typeConfig);
                    if (parsedData != null && !parsedData.isEmpty()) {
                        allData.addAll(parsedData);
                        jsonParts.add(JSON.serializePretty(parsedData));
                    }
                }
            }
            
            // Build response
            MetadataResponse response = new MetadataResponse();
            response.columns = buildColumns(config.columns);
            response.data = allData;
            response.rawJson = '[' + String.join(jsonParts, ',\n') + ']';
            response.totalWidth = calculateTotalWidth(config.columns);
            
            return response;
            
        } catch (Exception e) {
            System.debug('Error in getMetadataDetails: ' + e.getMessage());
            System.debug('Stack trace: ' + e.getStackTraceString());
            throw new AuraHandledException('Error retrieving metadata: ' + e.getMessage());
        }
    }
    
    /**
     * Get list of metadata files from Bitbucket repository
     */
    private static List<String> getMetadataFiles(String folderName, String fileExtension) {
        List<String> filePaths = new List<String>();
        
        try {
            Bitbucket_Config__mdt config = getBitbucketConfig();
            
            String apiUrl = 'https://api.bitbucket.org/2.0/repositories/' 
                + config.Workspace__c + '/' + config.Repository__c 
                + '/src/' + config.Branch__c + '/' + config.Base_Path__c + '/' + folderName;
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(apiUrl);
            req.setMethod('GET');
            
            // Get auth header - try different approaches
            String authHeader = getAuthorizationHeader();
            req.setHeader('Authorization', authHeader);
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> result = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                List<Object> values = (List<Object>) result.get('values');
                
                if (values != null) {
                    for (Object item : values) {
                        Map<String, Object> file = (Map<String, Object>) item;
                        String path = (String) file.get('path');
                        String type = (String) file.get('type');
                        
                        if (type == 'commit_file' && path != null && path.endsWith(fileExtension)) {
                            filePaths.add(path);
                        }
                    }
                }
            } else {
                System.debug('Bitbucket API error: ' + res.getStatusCode() + ' - ' + res.getBody());
            }
            
        } catch (Exception e) {
            System.debug('Error getting metadata files: ' + e.getMessage());
            throw e;
        }
        
        return filePaths;
    }
    
    /**
     * Get file content from Bitbucket
     */
    private static String getFileContent(String filePath) {
        try {
            Bitbucket_Config__mdt config = getBitbucketConfig();
            
            String rawUrl = 'https://api.bitbucket.org/2.0/repositories/' 
                + config.Workspace__c + '/' + config.Repository__c 
                + '/src/' + config.Branch__c + '/' + filePath;
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(rawUrl);
            req.setMethod('GET');
            
            // Get auth header
            String authHeader = getAuthorizationHeader();
            req.setHeader('Authorization', authHeader);
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                return res.getBody();
            } else {
                System.debug('Error fetching file ' + filePath + ': ' + res.getStatusCode());
            }
            
        } catch (Exception e) {
            System.debug('Error getting file content for ' + filePath + ': ' + e.getMessage());
        }
        
        return null;
    }
    
    /**
     * Get Bitbucket configuration
     */
    private static Bitbucket_Config__mdt getBitbucketConfig() {
        List<Bitbucket_Config__mdt> configs = [
            SELECT Workspace__c, Repository__c, Branch__c, Base_Path__c
            FROM Bitbucket_Config__mdt
            WHERE DeveloperName = 'Default'
            LIMIT 1
        ];
        
        if (configs.isEmpty()) {
            throw new ConfigException('No Bitbucket configuration found. Please create a Bitbucket_Config__mdt record.');
        }
        
        return configs[0];
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
     * Calculate total width percentage
     */
    private static Decimal calculateTotalWidth(List<MetadataTypeConfig.ColumnConfig> columns) {
        Decimal total = 0;
        
        for (MetadataTypeConfig.ColumnConfig col : columns) {
            if (col.active && col.widthPercent != null) {
                total += col.widthPercent;
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
        @AuraEnabled public Decimal widthPercent { get; set; }
    }
    
    /**
     * Get authorization header for Bitbucket API
     * Builds Basic Auth header from config
     */
    private static String getAuthorizationHeader() {
        try {
            // Get credentials from config
            Bitbucket_Config__mdt config = [
                SELECT Email__c, API_Token__c
                FROM Bitbucket_Config__mdt
                WHERE DeveloperName = 'Default'
                LIMIT 1
            ];
            
            // Build Basic Auth header
            String credentials = config.Email__c + ':' + config.API_Token__c;
            Blob credentialsBlob = Blob.valueOf(credentials);
            String encodedCredentials = EncodingUtil.base64Encode(credentialsBlob);
            
            return 'Basic ' + encodedCredentials;
            
        } catch (Exception e) {
            throw new ConfigException('Error getting Bitbucket credentials: ' + e.getMessage());
        }
    }
    
    /**
     * Test Bitbucket connection by listing assignmentRules folder
     * @return Success message or error details
     */
    @AuraEnabled
    public static String testConnection() {
        try {
            Bitbucket_Config__mdt config = getBitbucketConfig();
            
            String testUrl = 'https://api.bitbucket.org/2.0/repositories/' 
                + config.Workspace__c + '/' + config.Repository__c 
                + '/src/' + config.Branch__c + '/' + config.Base_Path__c + '/assignmentRules';
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(testUrl);
            req.setMethod('GET');
            req.setHeader('Authorization', getAuthorizationHeader());
            req.setTimeout(30000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> result = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                List<Object> values = (List<Object>) result.get('values');
                Integer fileCount = values != null ? values.size() : 0;
                
                return 'SUCCESS: Connected to Bitbucket! Found ' + fileCount + ' files in assignmentRules folder.\n' +
                       'Workspace: ' + config.Workspace__c + '\n' +
                       'Repository: ' + config.Repository__c + '\n' +
                       'Branch: ' + config.Branch__c;
            } else if (res.getStatusCode() == 401) {
                return 'ERROR: Authentication failed. Check Email and API_Token in Bitbucket_Config__mdt.\n' +
                       'Status: ' + res.getStatusCode();
            } else if (res.getStatusCode() == 404) {
                return 'ERROR: Repository or folder not found. Check Workspace, Repository, Branch, Base_Path.\n' +
                       'URL: ' + testUrl + '\n' +
                       'Status: ' + res.getStatusCode();
            } else {
                return 'ERROR: Status ' + res.getStatusCode() + ': ' + res.getStatus();
            }
            
        } catch (Exception e) {
            return 'ERROR: ' + e.getMessage() + ' (Line ' + e.getLineNumber() + ')';
        }
    }
    
    public class ConfigException extends Exception {}
}