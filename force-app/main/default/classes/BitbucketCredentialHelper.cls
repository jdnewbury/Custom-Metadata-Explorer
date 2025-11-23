public with sharing class BitbucketCredentialHelper {
    
    // Inner class for Bitbucket configuration
    public class BitbucketConfig {
        public String workspace { get; set; }
        public String repoSlug { get; set; }
        public String authType { get; set; }
        public String apiToken { get; set; }
        public String repositoryAccessToken { get; set; }
        
        public BitbucketConfig() {
            // Default constructor
        }
    }
    
    // Custom exception class for credential errors
    public class CredentialException extends Exception {}
    
    // Authentication types enum
    public enum AuthType {
        API_TOKEN,
        REPOSITORY_ACCESS_TOKEN
    }
    
    // Static method to get Bitbucket configuration
    public static BitbucketConfig getBitbucketConfig() {
        BitbucketConfig config = new BitbucketConfig();
        
        try {
            // Retrieve from Custom Settings or Custom Metadata
            // For now, using Custom Settings approach
            Bitbucket_Settings__c settings = Bitbucket_Settings__c.getOrgDefaults();
            
            if (settings != null) {
                config.workspace = settings.Workspace__c;
                config.repoSlug = settings.Repository_Slug__c;
                config.authType = settings.Auth_Type__c;
                
                // Get credentials based on auth type
                if (settings.Auth_Type__c == 'API_TOKEN') {
                    config.apiToken = getApiToken();
                } else if (settings.Auth_Type__c == 'REPOSITORY_ACCESS_TOKEN') {
                    config.repositoryAccessToken = getRepositoryAccessToken();
                }
            } else {
                throw new CredentialException('Bitbucket settings not configured');
            }
            
        } catch (Exception e) {
            throw new CredentialException('Error retrieving Bitbucket configuration: ' + e.getMessage());
        }
        
        return config;
    }
    
    // Method to validate configuration
    public static void validateConfig(BitbucketConfig config) {
        if (String.isBlank(config.workspace)) {
            throw new CredentialException('Workspace is required');
        }
        
        if (String.isBlank(config.repoSlug)) {
            throw new CredentialException('Repository slug is required');
        }
        
        if (String.isBlank(config.authType)) {
            throw new CredentialException('Authentication type is required');
        }
        
        // Validate credentials based on auth type
        if (config.authType == 'API_TOKEN' && String.isBlank(config.apiToken)) {
            throw new CredentialException('API Token is required for API_TOKEN authentication');
        } else if (config.authType == 'REPOSITORY_ACCESS_TOKEN' && String.isBlank(config.repositoryAccessToken)) {
            throw new CredentialException('Repository Access Token is required for REPOSITORY_ACCESS_TOKEN authentication');
        }
    }
    
    // Private method to get API token (from Named Credential or Protected Custom Setting)
    private static String getApiToken() {
        // Option 1: Using Named Credential (recommended)
        // This would be handled via Named Credential configuration
        
        // Option 2: Using Protected Custom Setting
        Bitbucket_Credentials__c creds = Bitbucket_Credentials__c.getInstance();
        if (creds != null && !String.isBlank(creds.API_Token__c)) {
            return creds.API_Token__c;
        }
        
        return null;
    }
    
    // Private method to get Repository Access Token
    private static String getRepositoryAccessToken() {
        // Option 1: Using Named Credential (recommended)
        // This would be handled via Named Credential configuration
        
        // Option 2: Using Protected Custom Setting
        Bitbucket_Credentials__c creds = Bitbucket_Credentials__c.getInstance();
        if (creds != null && !String.isBlank(creds.Repository_Access_Token__c)) {
            return creds.Repository_Access_Token__c;
        }
        
        return null;
    }
    
    // Helper method to construct authorization header
    public static String getAuthorizationHeader(BitbucketConfig config) {
        if (config.authType == 'API_TOKEN' && !String.isBlank(config.apiToken)) {
            // For API Token, use Bearer authentication
            return 'Bearer ' + config.apiToken;
        } else if (config.authType == 'REPOSITORY_ACCESS_TOKEN' && !String.isBlank(config.repositoryAccessToken)) {
            // For Repository Access Token, use Bearer authentication
            return 'Bearer ' + config.repositoryAccessToken;
        }
        
        throw new CredentialException('Invalid authentication configuration');
    }
    
    // Static method to test connection
    @AuraEnabled
    public static Map<String, Object> testConnection() {
        Map<String, Object> result = new Map<String, Object>();
        
        try {
            BitbucketConfig config = getBitbucketConfig();
            validateConfig(config);
            
            // Test API call to Bitbucket
            HttpRequest req = new HttpRequest();
            req.setEndpoint('https://api.bitbucket.org/2.0/repositories/' + 
                          config.workspace + '/' + config.repoSlug);
            req.setMethod('GET');
            req.setHeader('Authorization', getAuthorizationHeader(config));
            req.setHeader('Accept', 'application/json');
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                result.put('success', true);
                result.put('message', 'Connection successful');
                result.put('statusCode', res.getStatusCode());
            } else {
                result.put('success', false);
                result.put('message', 'Connection failed: ' + res.getStatus());
                result.put('statusCode', res.getStatusCode());
            }
            
        } catch (Exception e) {
            result.put('success', false);
            result.put('message', 'Error: ' + e.getMessage());
            result.put('statusCode', 0);
        }
        
        return result;
    }
}