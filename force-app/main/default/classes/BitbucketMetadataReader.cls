public with sharing class BitbucketMetadataReader {
    
    private static final String BITBUCKET_API_BASE = 'https://api.bitbucket.org/2.0';
    private static final Integer MAX_RETRIES = 3;
    private static final Integer RETRY_DELAY_MS = 1000;
    
    // AuraEnabled method for LWC
    @AuraEnabled
    public static String testBitbucketConnection() {
        try {
            BitbucketCredentialHelper.BitbucketConfig config = BitbucketCredentialHelper.getBitbucketConfig();
            BitbucketCredentialHelper.validateConfig(config);
            
            // Test connection with a simple API call
            String endpoint = BITBUCKET_API_BASE + '/repositories/' + 
                            config.workspace + '/' + config.repoSlug;
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', BitbucketCredentialHelper.getAuthorizationHeader(config));
            req.setHeader('Accept', 'application/json');
            req.setTimeout(30000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                return 'SUCCESS: Connected to Bitbucket repository';
            } else {
                return 'ERROR: ' + res.getStatus() + ' (Code: ' + res.getStatusCode() + ')';
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            System.debug('Credential error: ' + bce.getMessage());
            return 'CREDENTIAL_ERROR: ' + bce.getMessage();
        } catch (Exception e) {
            System.debug('General error: ' + e.getMessage());
            return 'ERROR: ' + e.getMessage();
        }
    }
    
    // Method to fetch file content from Bitbucket
    @AuraEnabled(cacheable=true)
    public static Map<String, Object> getFileContent(String filePath) {
        Map<String, Object> result = new Map<String, Object>();
        
        try {
            BitbucketCredentialHelper.BitbucketConfig config = BitbucketCredentialHelper.getBitbucketConfig();
            BitbucketCredentialHelper.validateConfig(config);
            
            String endpoint = BITBUCKET_API_BASE + '/repositories/' + 
                            config.workspace + '/' + config.repoSlug + 
                            '/src/main/' + filePath;
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', BitbucketCredentialHelper.getAuthorizationHeader(config));
            req.setHeader('Accept', 'application/json');
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = executeWithRetry(http, req);
            
            if (res.getStatusCode() == 200) {
                result.put('success', true);
                result.put('content', res.getBody());
                result.put('statusCode', res.getStatusCode());
            } else {
                result.put('success', false);
                result.put('error', 'Failed to fetch file: ' + res.getStatus());
                result.put('statusCode', res.getStatusCode());
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            result.put('success', false);
            result.put('error', 'Credential error: ' + bce.getMessage());
            result.put('errorType', 'CREDENTIAL_ERROR');
        } catch (Exception e) {
            result.put('success', false);
            result.put('error', 'Error: ' + e.getMessage());
            result.put('errorType', 'GENERAL_ERROR');
        }
        
        return result;
    }
    
    // Method to list repository contents
    @AuraEnabled(cacheable=true)
    public static Map<String, Object> listRepositoryContents(String path) {
        Map<String, Object> result = new Map<String, Object>();
        
        try {
            BitbucketCredentialHelper.BitbucketConfig config = BitbucketCredentialHelper.getBitbucketConfig();
            BitbucketCredentialHelper.validateConfig(config);
            
            String endpoint = BITBUCKET_API_BASE + '/repositories/' + 
                            config.workspace + '/' + config.repoSlug + 
                            '/src/main/' + (String.isNotBlank(path) ? path : '');
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', BitbucketCredentialHelper.getAuthorizationHeader(config));
            req.setHeader('Accept', 'application/json');
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = executeWithRetry(http, req);
            
            if (res.getStatusCode() == 200) {
                result.put('success', true);
                result.put('data', JSON.deserializeUntyped(res.getBody()));
                result.put('statusCode', res.getStatusCode());
            } else {
                result.put('success', false);
                result.put('error', 'Failed to list contents: ' + res.getStatus());
                result.put('statusCode', res.getStatusCode());
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            result.put('success', false);
            result.put('error', 'Credential error: ' + bce.getMessage());
            result.put('errorType', 'CREDENTIAL_ERROR');
        } catch (Exception e) {
            result.put('success', false);
            result.put('error', 'Error: ' + e.getMessage());
            result.put('errorType', 'GENERAL_ERROR');
        }
        
        return result;
    }
    
    // Method to get repository branches
    @AuraEnabled(cacheable=true)
    public static Map<String, Object> getBranches() {
        Map<String, Object> result = new Map<String, Object>();
        
        try {
            BitbucketCredentialHelper.BitbucketConfig config = BitbucketCredentialHelper.getBitbucketConfig();
            BitbucketCredentialHelper.validateConfig(config);
            
            String endpoint = BITBUCKET_API_BASE + '/repositories/' + 
                            config.workspace + '/' + config.repoSlug + '/refs/branches';
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', BitbucketCredentialHelper.getAuthorizationHeader(config));
            req.setHeader('Accept', 'application/json');
            req.setTimeout(30000);
            
            Http http = new Http();
            HttpResponse res = executeWithRetry(http, req);
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> responseData = (Map<String, Object>)JSON.deserializeUntyped(res.getBody());
                result.put('success', true);
                result.put('branches', responseData.get('values'));
                result.put('statusCode', res.getStatusCode());
            } else {
                result.put('success', false);
                result.put('error', 'Failed to fetch branches: ' + res.getStatus());
                result.put('statusCode', res.getStatusCode());
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            result.put('success', false);
            result.put('error', 'Credential error: ' + bce.getMessage());
            result.put('errorType', 'CREDENTIAL_ERROR');
        } catch (Exception e) {
            result.put('success', false);
            result.put('error', 'Error: ' + e.getMessage());
            result.put('errorType', 'GENERAL_ERROR');
        }
        
        return result;
    }
    
    // Method to get commits
    @AuraEnabled(cacheable=true)
    public static Map<String, Object> getCommits(String branch, Integer limitSize) {
        Map<String, Object> result = new Map<String, Object>();
        
        try {
            BitbucketCredentialHelper.BitbucketConfig config = BitbucketCredentialHelper.getBitbucketConfig();
            BitbucketCredentialHelper.validateConfig(config);
            
            Integer commitLimit = (limitSize != null && limitSize > 0) ? limitSize : 10;
            
            String endpoint = BITBUCKET_API_BASE + '/repositories/' + 
                            config.workspace + '/' + config.repoSlug + '/commits';
            
            if (String.isNotBlank(branch)) {
                endpoint += '/' + branch;
            }
            endpoint += '?pagelen=' + commitLimit;
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', BitbucketCredentialHelper.getAuthorizationHeader(config));
            req.setHeader('Accept', 'application/json');
            req.setTimeout(30000);
            
            Http http = new Http();
            HttpResponse res = executeWithRetry(http, req);
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> responseData = (Map<String, Object>)JSON.deserializeUntyped(res.getBody());
                result.put('success', true);
                result.put('commits', responseData.get('values'));
                result.put('statusCode', res.getStatusCode());
            } else {
                result.put('success', false);
                result.put('error', 'Failed to fetch commits: ' + res.getStatus());
                result.put('statusCode', res.getStatusCode());
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            result.put('success', false);
            result.put('error', 'Credential error: ' + bce.getMessage());
            result.put('errorType', 'CREDENTIAL_ERROR');
        } catch (Exception e) {
            result.put('success', false);
            result.put('error', 'Error: ' + e.getMessage());
            result.put('errorType', 'GENERAL_ERROR');
        }
        
        return result;
    }
    
    // Method to get pull requests
    @AuraEnabled(cacheable=true)
    public static Map<String, Object> getPullRequests(String state) {
        Map<String, Object> result = new Map<String, Object>();
        
        try {
            BitbucketCredentialHelper.BitbucketConfig config = BitbucketCredentialHelper.getBitbucketConfig();
            BitbucketCredentialHelper.validateConfig(config);
            
            String endpoint = BITBUCKET_API_BASE + '/repositories/' + 
                            config.workspace + '/' + config.repoSlug + '/pullrequests';
            
            if (String.isNotBlank(state)) {
                endpoint += '?state=' + state;
            }
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', BitbucketCredentialHelper.getAuthorizationHeader(config));
            req.setHeader('Accept', 'application/json');
            req.setTimeout(30000);
            
            Http http = new Http();
            HttpResponse res = executeWithRetry(http, req);
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> responseData = (Map<String, Object>)JSON.deserializeUntyped(res.getBody());
                result.put('success', true);
                result.put('pullRequests', responseData.get('values'));
                result.put('statusCode', res.getStatusCode());
            } else {
                result.put('success', false);
                result.put('error', 'Failed to fetch pull requests: ' + res.getStatus());
                result.put('statusCode', res.getStatusCode());
            }
            
        } catch (BitbucketCredentialHelper.CredentialException bce) {
            result.put('success', false);
            result.put('error', 'Credential error: ' + bce.getMessage());
            result.put('errorType', 'CREDENTIAL_ERROR');
        } catch (Exception e) {
            result.put('success', false);
            result.put('error', 'Error: ' + e.getMessage());
            result.put('errorType', 'GENERAL_ERROR');
        }
        
        return result;
    }
    
    // Helper method to execute HTTP request with retry logic
    private static HttpResponse executeWithRetry(Http http, HttpRequest req) {
        HttpResponse res = null;
        Integer attempts = 0;
        
        while (attempts < MAX_RETRIES) {
            try {
                res = http.send(req);
                
                // If successful or client error (4xx), don't retry
                if (res.getStatusCode() < 500) {
                    return res;
                }
                
                attempts++;
                
                // If server error (5xx) and not last attempt, wait and retry
                if (attempts < MAX_RETRIES) {
                    System.debug('Retrying request. Attempt ' + attempts + ' of ' + MAX_RETRIES);
                    // Sleep for retry delay
                    Long startTime = System.currentTimeMillis();
                    while (System.currentTimeMillis() - startTime < RETRY_DELAY_MS) {
                        // Wait
                    }
                }
                
            } catch (Exception e) {
                attempts++;
                
                if (attempts >= MAX_RETRIES) {
                    throw e;
                }
                
                System.debug('Request failed with exception. Retrying. Attempt ' + attempts + ' of ' + MAX_RETRIES);
            }
        }
        
        return res;
    }
}