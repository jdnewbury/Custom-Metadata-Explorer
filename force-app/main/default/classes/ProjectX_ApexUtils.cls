/*
*********************************************************
Apex Class Name    : ProjectX_ApexUtils
Created Date       : 27 March, 2025
@description       : This class houses utility functions used by QHub Salesforce instance
@author            : Jason Newbury

Modification Log:
Ver   Date         Author                           Modification
1.0   27.03.2025   Jason Newbury                    Initial Version

Modification Log:
Ver   Date         Author                           Modifications
1.1   09.04.2025   Jason Newbury                    - Removed whitespace throughout
                                                    - added error logging method

*********************************************************
*/
public class ProjectX_ApexUtils {


    // 1. Check if String is Null or Empty
    public static Boolean isNullOrEmpty(String str) {
        return String.isBlank(str);
    }

    // 2. Safe String Conversion
    public static String safeToString(Object obj) {
        return obj == null ? '' : String.valueOf(obj);
    }

    // 3. Format Date to String
    public static String formatDate(Date dt, String format) {
        if (dt == null) {
            return '';
        }
        // Date does not support format(String). Convert to DateTime at midnight and format.
        if (!String.isBlank(format)) {
            DateTime dtm = DateTime.newInstance(dt, Time.newInstance(0, 0, 0, 0));
            return dtm.format(format);
        }
        // Fallback to locale-based Date.format() when no pattern provided
        return dt.format();
    }

    // 4. Convert List to Map by Id
    public static Map<Id, SObject> listToMap(List<SObject> records) {
        Map<Id, SObject> result = new Map<Id, SObject>();
        for (SObject rec : records) {
            result.put(rec.Id, rec);
        }
        return result;
    }

    // 5. Log Error to Debug
    public static void logError(Exception ex) {
        System.debug('Error: ' + ex.getMessage() + ' | Stack: ' + ex.getStackTraceString());
    }

    // 6. Get Current User Info
    public static User getCurrentUser() {
        return [SELECT Id, Name, Email FROM User WHERE Id = :UserInfo.getUserId()];
    }

    // 7. Safe DML Insert
    public static void safeInsert(List<SObject> records) {
        try {
            insert records;
        } catch (DmlException e) {
            logError(e);
        }
    }

    // 8a. JSON Serialization
    public static String toJson(Object obj) {
        return JSON.serialize(obj);
    }

    // 8b. JSON Serialization
    public static Object fromJson(String jsonString, Type targetType) {
        return JSON.deserialize(jsonString, targetType);
    }

    // 9. Get SObject by Id dynamically
    public static SObject getRecordById(Id recordId, String sObjectType, List<String> fields) {
        if (recordId == null || String.isBlank(sObjectType) || fields.isEmpty()) {
            return null;
        }
        String fieldList = String.join(fields, ',');
        String query = 'SELECT ' + fieldList + ' FROM ' + sObjectType + ' WHERE Id = :recordId LIMIT 1';
        return Database.query(query);
    }

    // 10. Check if record exists
    public static Boolean recordExists(Id recordId, String sObjectType) {
        if (recordId == null || String.isBlank(sObjectType)) {
            return false;
        }
        String query = 'SELECT Id FROM ' + sObjectType + ' WHERE Id = :recordId LIMIT 1';
        List<SObject> results = Database.query(query);
        return !results.isEmpty();
    }

    // 11. Run dynamic query safely with limit
    public static List<SObject> runQuery(String sObjectType, List<String> fields, String whereClause, Integer limitSize) {
        if (String.isBlank(sObjectType) || fields.isEmpty()) {
            return new List<SObject>();
        }
        String fieldList = String.join(fields, ',');
        String query = 'SELECT ' + fieldList + ' FROM ' + sObjectType;
        if (!String.isBlank(whereClause)) {
            query += ' WHERE ' + whereClause;
        }
        if (limitSize != null && limitSize > 0) {
            query += ' LIMIT ' + limitSize;
        }
        return Database.query(query);
    }

    // 12. Application Error Log
    /*
    *********************************************************
    @Method Name			: logError
    @author					: Jason Newbury - jason.newbury@publicissapient.com
    @description			: Logs an error to the ProjectXXX_Application_Error_Log__c object.
    @param errorMessage     : The error message to log.
    @param className        : The name of the class where the error occurred.
    @param errorType        : The type of error.
    @param recordId         : The record ID of the originating record.
    @return					: void
    ********************************************************
    */
    /*
    public static void logError(String errorMessage, String className, String errorType, String recordId, String userEmail) {

        // Fetch the user ID for the specified email
        User user;
        try {
            user = [SELECT Id FROM User WHERE Email = :userEmail LIMIT 1];
        } catch (QueryException e) {
            System.debug('User not found for email: ' + userEmail);
            return; // Exit if the user is not found
        }

        // Create a new error log record

        ProjectXXX_Application_Error_Log__c errorLog = new ProjectXXX_Application_Error_Log__c();
        errorLog.ProjectXXX_Error_Date__c = Date.today();
        errorLog.ProjectXXX_Error_Message__c = errorMessage;
        errorLog.ProjectXXX_Error_Source__c = className;
        errorLog.ProjectXXX_Error_Type__c = errorType;
        errorLog.ProjectXXX_Record_Id__c = recordId;
        errorLog.ProjectXXX_Running_User__c = UserInfo.getUserId();

        // Insert the error log record
        try {
            insert errorLog;
        } catch (DmlException e) {
            System.debug('Failed to log error: ' + e.getMessage());
        }
    }
    */

    // 13. Random Key Generator
    /*
    *********************************************************
    @Method Name			: generateRandomKey
    @author					: Jason Newbury
    @description			: method is used to generate a random alphanumeric (unless otherwise specified) key
    @param length			: length of returned string
	@param customCharSet	: if var contains a value, use the values in this string to generate returnKey instead of standard alphanumeric
    @return					: Randomly generated key of inputted length; alphanumeric unless customCharSet is provided
    ********************************************************
    */
    public static String generateRandomKey(Integer length, String customCharSet) {

        // Default Alphanumeric character set to generate key
        String chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz';

        // If custom character set is provided, use that instead of default set to generate key
        if (!String.isBlank(customCharSet)) {
            chars = customCharSet;
        }

        final Integer charLength = chars.length();
        String randStr = '';

        // Use character set to generate random key of specified length
        while (randStr.length() < length) {
           Integer i = Math.mod(Math.abs(Crypto.getRandomInteger()), charLength);
           randStr += chars.substring(i, i+1);
        }

        // Return generated key
        return randStr;
	}

    // 14. Random Key Generator
    /*
        A reusable Apex utility method that converts any JSON (headers + rows) into TableResponseDTO automatically
    */
    public static CustomMetadataExplorerController.TableResponseDTO convertJsonToTableResponse(String jsonString) {
        CustomMetadataExplorerController.TableResponseDTO dto = new CustomMetadataExplorerController.TableResponseDTO();
        dto.rawJson = jsonString;

        if (String.isBlank(jsonString)) {
            dto.columns = new List<CustomMetadataExplorerController.TableColumnDTO>();
            dto.data = new List<Object>();
            return dto;
        }

        try {
            Map<String, Object> parsed = (Map<String, Object>) JSON.deserializeUntyped(jsonString);

            // Extract headers
            List<Object> headers = (List<Object>) parsed.get('headers');
            List<CustomMetadataExplorerController.TableColumnDTO> colDtos = new List<CustomMetadataExplorerController.TableColumnDTO>();
            if (headers != null) {
                for (Object hObj : headers) {
                    String header = (String) hObj;
                    CustomMetadataExplorerController.TableColumnDTO col = new CustomMetadataExplorerController.TableColumnDTO();
                    col.label = header;
                    col.fieldName = header;
                    col.type = 'text'; // Default type; can be enhanced later
                    colDtos.add(col);
                }
            }

            // Extract rows
            List<Object> rows = (List<Object>) parsed.get('rows');

            dto.columns = colDtos;
            dto.data = rows != null ? rows : new List<Object>();
        } catch (Exception ex) {
            System.debug('Error converting JSON to TableResponseDTO: ' + ex.getMessage());
            dto.columns = new List<CustomMetadataExplorerController.TableColumnDTO>();
            dto.data = new List<Object>();
        }

        return dto;
    }

 }