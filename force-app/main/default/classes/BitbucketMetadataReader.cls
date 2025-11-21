/**
 * @description Universal Bitbucket metadata reader that supports multiple metadata types
 *              with dynamic column configuration. Uses BitbucketCredentialHelper for auth.
 * @author Jason Newbury
 * @date 2025
 */
public with sharing class BitbucketMetadataReader {
    
    /**
     * @description Get all available metadata type options for the LWC dropdown
     * @return List of options with label and value
     */
    @AuraEnabled(cacheable=true)
    public static List<Map<String, String>> getMetadataTypeOptions() {
        return MetadataTypeConfig.getAllOptions();
    }
    
    /**
     * @description Test Bitbucket connection with current configuration
     * @return Map with success status and message
     */
    @AuraEnabled
    public static Map<String, Object> testConnection() {
        return BitbucketCredentialHelper.testConnection();
    }
    
    /**
     * @description Main entry point - fetches and parses metadata for any supported type
     * @param metadataTypeKey The key identifying the metadata type (e.g., 'assignment_rules')
     * @return TableResponseDTO formatted for lightning-datatable
     */
    @AuraEnabled(cacheable=false)
    public static CustomMetadataExplorerController.TableResponseDTO getMetadataByType(String metadataTypeKey) {
        try {
            // Strip 'bitbucket_' prefix if present
            String typeKey = metadataTypeKey.startsWith('bitbucket_') 
                ? metadataTypeKey.substringAfter('bitbucket_') 
                : metadataTypeKey;
            
            // Get configuration for this metadata type
            MetadataTypeConfig.TypeConfig typeConfig = MetadataTypeConfig.getConfig(typeKey);
            
            // Get Bitbucket credentials/config
            BitbucketCredentialHelper.BitbucketConfig bbConfig = BitbucketCredentialHelper.getConfig();
            
            // Get list of files for this metadata type
            List<String> files = getFilesForType(typeConfig, bbConfig);
            
            if (files.isEmpty()) {
                CustomMetadataExplorerController.TableResponseDTO emptyResponse = new CustomMetadataExplorerController.TableResponseDTO();
                emptyResponse.columns = typeConfig.getTableColumns();
                emptyResponse.data = new List<Object>();
                emptyResponse.rawJson = '[]';
                return emptyResponse;
            }
            
            // Get the appropriate parser for this type
            IMetadataParser parser = MetadataParserFactory.getParser(typeKey);
            
            // Parse all files and combine results
            List<Map<String, Object>> allRows = new List<Map<String, Object>>();
            Integer rowId = 1;
            
            for (String filePath : files) {
                String xmlContent = fetchFileContent(filePath, bbConfig);
                List<Map<String, Object>> rows = parser.parse(xmlContent, filePath, typeConfig);
                
                // Add unique Id for datatable key-field
                for (Map<String, Object> row : rows) {
                    row.put('Id', String.valueOf(rowId++));
                    allRows.add(row);
                }
            }
            
            // Build response
            CustomMetadataExplorerController.TableResponseDTO response = new CustomMetadataExplorerController.TableResponseDTO();
            response.columns = typeConfig.getTableColumns();
            response.data = allRows;
            response.rawJson = JSON.serializePretty(allRows);
            
            return response;
            
        } catch (MetadataTypeConfig.MetadataConfigException mce) {
            System.debug('MetadataConfigException: ' + mce.getMessage());
            System.debug('Stack trace: ' + mce.getStackTraceString());
            throw new AuraHandledException('Unsupported metadata type: ' + metadataTypeKey);
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            System.debug('CredentialException: ' + bce.getMessage());
            System.debug('Stack trace: ' + bce.getStackTraceString());
            throw new AuraHandledException('Bitbucket configuration error: ' + bce.getMessage());
        } catch (Exception ex) {
            System.debug('Error in getMetadataByType: ' + ex.getMessage());
            System.debug('Stack trace: ' + ex.getStackTraceString());
            throw new AuraHandledException('Failed to load metadata: ' + ex.getMessage());
        }
    }
    
    /**
     * @description Get list of files for a specific metadata type from Bitbucket
     * @param typeConfig The TypeConfig for the metadata type
     * @param bbConfig Bitbucket configuration
     * @return List of file paths
     */
    public static List<String> getFilesForType(
        MetadataTypeConfig.TypeConfig typeConfig, 
        BitbucketCredentialHelper.BitbucketConfig bbConfig
    ) {
        List<String> files = new List<String>();
        
        try {
            String endpoint = bbConfig.getDirectoryEndpoint(typeConfig.folderName);
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setTimeout(60000);
            
            // Set auth header if not using Named Credential
            if (!bbConfig.useNamedCredential) {
                req.setHeader('Authorization', bbConfig.getAuthorizationHeader());
            }
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> result = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                List<Object> values = (List<Object>) result.get('values');
                
                if (values != null) {
                    for (Object val : values) {
                        Map<String, Object> fileObj = (Map<String, Object>) val;
                        String path = (String) fileObj.get('path');
                        
                        // Filter by file extension
                        if (path != null && path.endsWith(typeConfig.fileExtension)) {
                            files.add(path);
                        }
                    }
                }
            } else if (res.getStatusCode() == 404) {
                // Folder doesn't exist - return empty list
                System.debug('Folder not found: ' + typeConfig.folderName);
            } else if (res.getStatusCode() == 401) {
                throw new BitbucketCredentialHelper.CredentialException(
                    'Authentication failed. Please verify your Bitbucket credentials.'
                );
            } else {
                throw new CalloutException('Bitbucket API error (' + res.getStatusCode() + '): ' + res.getStatus());
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            System.debug('CredentialException in getFilesForType: ' + bce.getMessage());
            throw bce;
        } catch (Exception ex) {
            System.debug('Error in getFilesForType: ' + ex.getMessage());
            System.debug('Stack trace: ' + ex.getStackTraceString());
            throw ex;
        }
        
        return files;
    }
    
    /**
     * @description Fetch raw file content from Bitbucket
     * @param filePath Full path to the file
     * @param bbConfig Bitbucket configuration
     * @return Raw file content as string
     */
    public static String fetchFileContent(
        String filePath, 
        BitbucketCredentialHelper.BitbucketConfig bbConfig
    ) {
        try {
            String endpoint = bbConfig.getFileEndpoint(filePath);
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setTimeout(60000);
            
            // Set auth header if not using Named Credential
            if (!bbConfig.useNamedCredential) {
                req.setHeader('Authorization', bbConfig.getAuthorizationHeader());
            }
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                return res.getBody();
            } else if (res.getStatusCode() == 401) {
                throw new BitbucketCredentialHelper.CredentialException(
                    'Authentication failed. Please verify your Bitbucket credentials.'
                );
            } else {
                throw new CalloutException('Failed to fetch file (' + res.getStatusCode() + '): ' + filePath);
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            System.debug('CredentialException in fetchFileContent: ' + bce.getMessage());
            throw bce;
        } catch (Exception ex) {
            System.debug('Error in fetchFileContent for file: ' + filePath);
            System.debug('Error message: ' + ex.getMessage());
            System.debug('Stack trace: ' + ex.getStackTraceString());
            throw ex;
        }
    }
    
    /**
     * @description Get details for a specific file (for drill-down scenarios)
     * @param filePath Path to the specific file
     * @param metadataTypeKey The metadata type key
     * @return TableResponseDTO for single file
     */
    @AuraEnabled(cacheable=false)
    public static CustomMetadataExplorerController.TableResponseDTO getFileDetails(
        String filePath, 
        String metadataTypeKey
    ) {
        try {
            String typeKey = metadataTypeKey.startsWith('bitbucket_') 
                ? metadataTypeKey.substringAfter('bitbucket_') 
                : metadataTypeKey;
            
            MetadataTypeConfig.TypeConfig typeConfig = MetadataTypeConfig.getConfig(typeKey);
            BitbucketCredentialHelper.BitbucketConfig bbConfig = BitbucketCredentialHelper.getConfig();
            IMetadataParser parser = MetadataParserFactory.getParser(typeKey);
            
            String xmlContent = fetchFileContent(filePath, bbConfig);
            List<Map<String, Object>> rows = parser.parse(xmlContent, filePath, typeConfig);
            
            // Add Ids
            Integer rowId = 1;
            for (Map<String, Object> row : rows) {
                row.put('Id', String.valueOf(rowId++));
            }
            
            CustomMetadataExplorerController.TableResponseDTO response = new CustomMetadataExplorerController.TableResponseDTO();
            response.columns = typeConfig.getTableColumns();
            response.data = rows;
            response.rawJson = JSON.serializePretty(rows);
            
            return response;
            
        } catch (MetadataTypeConfig.MetadataConfigException mce) {
            System.debug('MetadataConfigException in getFileDetails: ' + mce.getMessage());
            System.debug('Stack trace: ' + mce.getStackTraceString());
            throw new AuraHandledException('Unsupported metadata type: ' + metadataTypeKey);
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            System.debug('CredentialException in getFileDetails: ' + bce.getMessage());
            System.debug('Stack trace: ' + bce.getStackTraceString());
            throw new AuraHandledException('Bitbucket configuration error: ' + bce.getMessage());
        } catch (Exception ex) {
            System.debug('Error in getFileDetails: ' + ex.getMessage());
            System.debug('Stack trace: ' + ex.getStackTraceString());
            throw new AuraHandledException('Failed to load file: ' + ex.getMessage());
        }
    }
    
    /**
     * @description Get just the file list for a metadata type (for file picker)
     * @param metadataTypeKey The metadata type key
     * @return List of file info maps
     */
    @AuraEnabled(cacheable=false)
    public static List<Map<String, String>> getFileList(String metadataTypeKey) {
        try {
            String typeKey = metadataTypeKey.startsWith('bitbucket_') 
                ? metadataTypeKey.substringAfter('bitbucket_') 
                : metadataTypeKey;
            
            MetadataTypeConfig.TypeConfig typeConfig = MetadataTypeConfig.getConfig(typeKey);
            BitbucketCredentialHelper.BitbucketConfig bbConfig = BitbucketCredentialHelper.getConfig();
            List<String> files = getFilesForType(typeConfig, bbConfig);
            
            List<Map<String, String>> fileInfoList = new List<Map<String, String>>();
            for (String filePath : files) {
                String fileName = filePath.substringAfterLast('/');
                String objectName = fileName.substringBefore('.');
                
                fileInfoList.add(new Map<String, String>{
                    'path' => filePath,
                    'fileName' => fileName,
                    'objectName' => objectName,
                    'label' => objectName + ' (' + fileName + ')'
                });
            }
            
            return fileInfoList;
            
        } catch (MetadataTypeConfig.MetadataConfigException mce) {
            System.debug('MetadataConfigException in getFileList: ' + mce.getMessage());
            System.debug('Stack trace: ' + mce.getStackTraceString());
            throw new AuraHandledException('Unsupported metadata type: ' + metadataTypeKey);
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            System.debug('CredentialException in getFileList: ' + bce.getMessage());
            System.debug('Stack trace: ' + bce.getStackTraceString());
            throw new AuraHandledException('Bitbucket configuration error: ' + bce.getMessage());
        } catch (Exception ex) {
            System.debug('Error in getFileList: ' + ex.getMessage());
            System.debug('Stack trace: ' + ex.getStackTraceString());
            throw new AuraHandledException('Failed to list files: ' + ex.getMessage());
        }
    }
}