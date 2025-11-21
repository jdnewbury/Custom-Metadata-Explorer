/**
 * @description Helper class for managing Bitbucket credentials and configuration.
 *              Supports API Tokens with Scopes (new) and Repository Access Tokens.
 *              
 *              IMPORTANT: As of Sept 9, 2025, App Passwords are deprecated.
 *              Use API Tokens with scopes instead.
 * 
 * @author Jason Newbury
 * @date 2025
 */
public with sharing class BitbucketCredentialHelper {
    
    /**
     * Credential source options
     * - NAMED_CREDENTIAL: Most secure, credentials stored in Salesforce Named Credential
     * - CUSTOM_METADATA: Deployable, credentials in Custom Metadata Type
     * - CUSTOM_SETTING: Legacy support, credentials in Protected Custom Setting
     * - REPOSITORY_TOKEN: Uses Repository Access Token (tied to repo, not user)
     */
    public enum CredentialSource {
        NAMED_CREDENTIAL,
        CUSTOM_METADATA,
        CUSTOM_SETTING,
        REPOSITORY_TOKEN
    }
    
    /**
     * Authentication type for API calls
     * - API_TOKEN: Uses Atlassian email + API token (new standard)
     * - REPOSITORY_ACCESS_TOKEN: Uses x-token-auth + repo token
     */
    public enum AuthType {
        API_TOKEN,
        REPOSITORY_ACCESS_TOKEN
    }
    
    // ═══════════════════════════════════════════════════════════════════════
    // CONFIGURATION - Set your preferred credential source here
    // ═══════════════════════════════════════════════════════════════════════
    private static final CredentialSource ACTIVE_SOURCE = CredentialSource.CUSTOM_METADATA;
    
    // Named Credential name (if using NAMED_CREDENTIAL)
    private static final String NAMED_CREDENTIAL_NAME = 'Bitbucket';
    
    // Custom Metadata record name (if using CUSTOM_METADATA)
    private static final String CUSTOM_METADATA_RECORD = 'Default';
    
    /**
     * @description Configuration object containing all Bitbucket settings
     */
    public class BitbucketConfig {
        public String apiBaseUrl { get; set; }
        public String workspace { get; set; }
        public String repository { get; set; }
        public String branch { get; set; }
        public String basePath { get; set; }
        
        // For API Token auth (email + token)
        public String email { get; set; }
        public String apiToken { get; set; }
        
        // For Repository Access Token auth
        public String repositoryAccessToken { get; set; }
        
        // Auth configuration
        public AuthType authType { get; set; }
        public Boolean useNamedCredential { get; set; }
        public String namedCredentialName { get; set; }
        
        public BitbucketConfig() {
            this.apiBaseUrl = 'https://api.bitbucket.org/2.0';
            this.branch = 'main';
            this.basePath = 'force-app/main/default';
            this.useNamedCredential = false;
            this.authType = BitbucketCredentialHelper.AuthType.API_TOKEN;
        }
        
        /**
         * @description Build the authorization header value based on auth type
         */
        public String getAuthorizationHeader() {
            String credentials;
            
            if (this.authType == AuthType.REPOSITORY_ACCESS_TOKEN) {
                // Repository Access Token uses special username
                if (String.isBlank(this.repositoryAccessToken)) {
                    throw new CredentialException('Repository Access Token is required');
                }
                credentials = 'x-token-auth:' + this.repositoryAccessToken;
            } else {
                // API Token uses Atlassian email + token
                if (String.isBlank(this.email) || String.isBlank(this.apiToken)) {
                    throw new CredentialException('Email and API Token are required for authentication');
                }
                credentials = this.email + ':' + this.apiToken;
            }
            
            return 'Basic ' + EncodingUtil.base64Encode(Blob.valueOf(credentials));
        }
        
        /**
         * @description Build endpoint URL for repository content listing
         */
        public String getDirectoryEndpoint(String folderName) {
            if (this.useNamedCredential) {
                return String.format(
                    'callout:{0}/repositories/{1}/{2}/src/{3}/{4}/{5}/',
                    new List<String>{this.namedCredentialName, this.workspace, this.repository, this.branch, this.basePath, folderName}
                );
            }
            return String.format(
                '{0}/repositories/{1}/{2}/src/{3}/{4}/{5}/',
                new List<String>{this.apiBaseUrl, this.workspace, this.repository, this.branch, this.basePath, folderName}
            );
        }
        
        /**
         * @description Build endpoint URL for specific file content
         */
        public String getFileEndpoint(String filePath) {
            if (this.useNamedCredential) {
                return String.format(
                    'callout:{0}/repositories/{1}/{2}/src/{3}/{4}',
                    new List<String>{this.namedCredentialName, this.workspace, this.repository, this.branch, filePath}
                );
            }
            return String.format(
                '{0}/repositories/{1}/{2}/src/{3}/{4}',
                new List<String>{this.apiBaseUrl, this.workspace, this.repository, this.branch, filePath}
            );
        }
        
        /**
         * @description Build endpoint URL for repository info (for connection test)
         */
        public String getRepositoryEndpoint() {
            if (this.useNamedCredential) {
                return String.format(
                    'callout:{0}/repositories/{1}/{2}',
                    new List<String>{this.namedCredentialName, this.workspace, this.repository}
                );
            }
            return String.format(
                '{0}/repositories/{1}/{2}',
                new List<String>{this.apiBaseUrl, this.workspace, this.repository}
            );
        }
    }
    
    /**
     * @description Get configuration from the active credential source
     */
    public static BitbucketConfig getConfig() {
        switch on ACTIVE_SOURCE {
            when NAMED_CREDENTIAL {
                return getConfigFromNamedCredential();
            }
            when CUSTOM_METADATA {
                return getConfigFromCustomMetadata();
            }
            when CUSTOM_SETTING {
                // Removed CUSTOM_SETTING as it's not needed, using CUSTOM_METADATA instead
                return getConfigFromCustomMetadata();
            }
            when REPOSITORY_TOKEN {
                return getConfigForRepositoryToken();
            }
            when else {
                throw new CredentialException('Invalid credential source configured');
            }
        }
    }
    
    /**
     * @description Get configuration using Named Credential
     */
    private static BitbucketConfig getConfigFromNamedCredential() {
        BitbucketConfig config = new BitbucketConfig();
        config.useNamedCredential = true;
        config.namedCredentialName = NAMED_CREDENTIAL_NAME;
        
        // Get non-sensitive config from Custom Metadata
        try {
            Bitbucket_Config__mdt mdt = Bitbucket_Config__mdt.getInstance(CUSTOM_METADATA_RECORD);
            if (mdt != null) {
                config.workspace = mdt.Workspace__c;
                config.repository = mdt.Repository__c;
                config.branch = String.isNotBlank(mdt.Branch__c) ? mdt.Branch__c : 'main';
                config.basePath = String.isNotBlank(mdt.Base_Path__c) ? mdt.Base_Path__c : 'force-app/main/default';
            }
        } catch (Exception ex) {
            System.debug('Could not load Custom Metadata config: ' + ex.getMessage());
        }
        
        validateConfigBasics(config);
        return config;
    }
    
    /**
     * @description Get configuration from Custom Metadata Type
     */
    private static BitbucketConfig getConfigFromCustomMetadata() {
        BitbucketConfig config = new BitbucketConfig();
        config.useNamedCredential = false;
        config.authType = AuthType.API_TOKEN;
        
        try {
            Bitbucket_Config__mdt mdt = Bitbucket_Config__mdt.getInstance(CUSTOM_METADATA_RECORD);
            
            if (mdt == null) {
                throw new CredentialException(
                    'Bitbucket configuration not found. Please create a Custom Metadata record named "' + 
                    CUSTOM_METADATA_RECORD + '" in Bitbucket_Config__mdt'
                );
            }
            
            config.apiBaseUrl = String.isNotBlank(mdt.API_Base_URL__c) 
                ? mdt.API_Base_URL__c 
                : 'https://api.bitbucket.org/2.0';
            config.workspace = mdt.Workspace__c;
            config.repository = mdt.Repository__c;
            config.branch = String.isNotBlank(mdt.Branch__c) ? mdt.Branch__c : 'main';
            config.basePath = String.isNotBlank(mdt.Base_Path__c) ? mdt.Base_Path__c : 'force-app/main/default';
            
            // API Token authentication (new method)
            config.email = mdt.Email__c;
            config.apiToken = mdt.API_Token__c;
            
            validateConfig(config);
            
        } catch (CredentialException ce) {
            throw ce;
        } catch (Exception ex) {
            throw new CredentialException('Failed to load Bitbucket configuration: ' + ex.getMessage());
        }
        
        return config;
    }
    
    
    /**
     * @description Get configuration for Repository Access Token
     */
    private static BitbucketConfig getConfigForRepositoryToken() {
        BitbucketConfig config = new BitbucketConfig();
        config.useNamedCredential = false;
        config.authType = AuthType.REPOSITORY_ACCESS_TOKEN;
        
        try {
            Bitbucket_Config__mdt mdt = Bitbucket_Config__mdt.getInstance(CUSTOM_METADATA_RECORD);
            
            if (mdt == null) {
                throw new CredentialException('Bitbucket configuration not found');
            }
            
            config.apiBaseUrl = String.isNotBlank(mdt.API_Base_URL__c) 
                ? mdt.API_Base_URL__c 
                : 'https://api.bitbucket.org/2.0';
            config.workspace = mdt.Workspace__c;
            config.repository = mdt.Repository__c;
            config.branch = String.isNotBlank(mdt.Branch__c) ? mdt.Branch__c : 'main';
            config.basePath = String.isNotBlank(mdt.Base_Path__c) ? mdt.Base_Path__c : 'force-app/main/default';
            
            // Repository Access Token (stored in API_Token__c field)
            config.repositoryAccessToken = mdt.API_Token__c;
            
            validateConfigBasics(config);
            if (String.isBlank(config.repositoryAccessToken)) {
                throw new CredentialException('Repository Access Token is required');
            }
            
        } catch (CredentialException ce) {
            throw ce;
        } catch (Exception ex) {
            throw new CredentialException('Failed to load Bitbucket configuration: ' + ex.getMessage());
        }
        
        return config;
    }
    
    /**
     * @description Validate basic configuration (workspace, repo)
     */
    private static void validateConfigBasics(BitbucketConfig config) {
        List<String> missingFields = new List<String>();
        
        if (String.isBlank(config.workspace)) missingFields.add('Workspace');
        if (String.isBlank(config.repository)) missingFields.add('Repository');
        
        if (!missingFields.isEmpty()) {
            throw new CredentialException(
                'Missing required Bitbucket configuration: ' + String.join(missingFields, ', ')
            );
        }
    }
    
    /**
     * @description Validate full configuration including credentials
     */
    private static void validateConfig(BitbucketConfig config) {
        validateConfigBasics(config);
        
        List<String> missingFields = new List<String>();
        
        if (!config.useNamedCredential) {
            if (config.authType == AuthType.API_TOKEN) {
                if (String.isBlank(config.email)) missingFields.add('Email');
                if (String.isBlank(config.apiToken)) missingFields.add('API Token');
            } else if (config.authType == AuthType.REPOSITORY_ACCESS_TOKEN) {
                if (String.isBlank(config.repositoryAccessToken)) missingFields.add('Repository Access Token');
            }
        }
        
        if (!missingFields.isEmpty()) {
            throw new CredentialException(
                'Missing required Bitbucket credentials: ' + String.join(missingFields, ', ')
            );
        }
    }
    
    /**
     * @description Test the Bitbucket connection with current configuration
     * @return Map with success status and message
     */
    public static Map<String, Object> testConnection() {
        Map<String, Object> result = new Map<String, Object>();
        
        try {
            BitbucketConfig config = getConfig();
            String endpoint = config.getRepositoryEndpoint();
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setTimeout(30000);
            
            if (!config.useNamedCredential) {
                req.setHeader('Authorization', config.getAuthorizationHeader());
            }
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> repoInfo = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                result.put('success', true);
                result.put('message', 'Successfully connected to repository: ' + repoInfo.get('full_name'));
                result.put('repoName', repoInfo.get('full_name'));
                
                // Add auth type info
                String authInfo = config.authType == AuthType.REPOSITORY_ACCESS_TOKEN 
                    ? 'Repository Access Token' 
                    : 'API Token (email: ' + config.email + ')';
                result.put('authMethod', authInfo);
                
            } else if (res.getStatusCode() == 401) {
                result.put('success', false);
                result.put('message', 'Authentication failed. For API Tokens, ensure you\'re using your ' +
                    'Atlassian account EMAIL (not username) with the token. ' +
                    'Check: Bitbucket → Personal settings → Email Aliases');
            } else if (res.getStatusCode() == 403) {
                result.put('success', false);
                result.put('message', 'Access forbidden. Ensure your API token has the ' +
                    '"read:repository:bitbucket" scope selected.');
            } else if (res.getStatusCode() == 404) {
                result.put('success', false);
                result.put('message', 'Repository not found. Check workspace ("' + config.workspace + 
                    '") and repository ("' + config.repository + '") slugs are correct.');
            } else {
                result.put('success', false);
                result.put('message', 'Unexpected response: ' + res.getStatusCode() + ' - ' + res.getStatus());
            }
            
        } catch (CredentialException ce) {
            result.put('success', false);
            result.put('message', 'Configuration error: ' + ce.getMessage());
        } catch (Exception ex) {
            result.put('success', false);
            result.put('message', 'Connection error: ' + ex.getMessage());
        }
        
        return result;
    }
    
    /**
     * @description Custom exception for credential errors
     */
    public with sharing class CredentialException extends Exception {
        public CredentialException(String message) {
            super(message);
        }
    }
}
