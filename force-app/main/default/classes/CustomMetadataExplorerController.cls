/**
 * @description Controller for Custom Metadata Type Explorer LWC.
 * Provides:
 *  - getExplorerOptions: returns picklist options from Custom_Metadata_Explorer__mdt
 *  - getMetadataDetails: invokes an invocable class to retrieve dynamic table JSON (columns, data)
 *
 * Security:
 *  - with sharing
 *  - SOQL WITH SECURITY_ENFORCED (USER_MODE removed per org feedback)
 *
 * Notes:
 *  - MDT API name: Custom_Metadata_Explorer__mdt
 *  - Fields:
 *      - DeveloperName (value for picklist)
 *      - MasterLabel (label for picklist)
 *      - Class_for_Metadata_Retrieval__c (name of invocable class to call)
 *  - No Agent_Alias__c available in org; direct invocable class invocation used instead.
 */
public with sharing class CustomMetadataExplorerController {

    public class OptionDTO {
        @AuraEnabled public String label;
        @AuraEnabled public String value;
    }

    public class ExplorerRecordDTO {
        @AuraEnabled public String developerName;
        @AuraEnabled public String masterLabel;
        @AuraEnabled public String metadataTypeApiName;
        @AuraEnabled public String classNameForRetrieval; // The invocable class name to call
    }

    public class TableColumnDTO {
        @AuraEnabled public String label;
        @AuraEnabled public String fieldName;
        @AuraEnabled public String type;
    }

    public class TableResponseDTO {
        @AuraEnabled public List<TableColumnDTO> columns;
        @AuraEnabled public List<Object> data;
        @AuraEnabled public String rawJson; // optional: echo of agent response or error for UI
    }

    /**
     * @description Get combobox options from Custom_Metadata_Explorer__mdt.
     * @return List of {label: MasterLabel, value: DeveloperName}
     */
    @AuraEnabled(cacheable=true)
    public static List<OptionDTO> getExplorerOptions() {
        System.debug('Executing getExplorerOptions');
        // Query MDT and enforce FLS, including the Class_for_Metadata_Retrieval__c field
        List<Custom_Metadata_Explorer__mdt> rows = [
            SELECT DeveloperName, MasterLabel, Class_for_Metadata_Retrieval__c
            FROM Custom_Metadata_Explorer__mdt
            WITH SECURITY_ENFORCED
            ORDER BY MasterLabel
        ];
        System.debug('Found ' + rows.size() + ' records in Custom_Metadata_Explorer__mdt');
        List<OptionDTO> options = new List<OptionDTO>();
        for (Custom_Metadata_Explorer__mdt r : rows) {
            System.debug('Processing record: ' + r.DeveloperName + ' - ' + r.MasterLabel);
            OptionDTO o = new OptionDTO();
            o.label = r.MasterLabel;
            o.value = r.DeveloperName;
            options.add(o);
        }
        System.debug('Returning ' + options.size() + ' options');
        return options;
    }

    /**
     * @description Retrieve the selected explorer record details by DeveloperName.
     *              Returns the class name for metadata retrieval from Class_for_Metadata_Retrieval__c field.
     */
    @AuraEnabled(cacheable=true)
    public static ExplorerRecordDTO getExplorerRecord(String developerName) {
        if (String.isBlank(developerName)) {
            return null;
        }
        Custom_Metadata_Explorer__mdt rec = [
            SELECT DeveloperName, MasterLabel, Class_for_Metadata_Retrieval__c
            FROM Custom_Metadata_Explorer__mdt
            WHERE DeveloperName = :developerName
            WITH SECURITY_ENFORCED
            LIMIT 1
        ];
        ExplorerRecordDTO dto = new ExplorerRecordDTO();
        dto.developerName = rec.DeveloperName;
        dto.masterLabel = rec.MasterLabel;
        dto.metadataTypeApiName = rec.DeveloperName; // per user: metadata type API name = DeveloperName
        dto.classNameForRetrieval = rec.Class_for_Metadata_Retrieval__c; // The invocable class name
        return dto;
    }

    /**
     * @description Invoke the appropriate invocable class based on the selected metadata type and return a dynamic table model.
     * @param developerName DeveloperName of Custom_Metadata_Explorer__mdt
     * @return TableResponseDTO with columns and data ready for lightning-datatable
     */
    @AuraEnabled
    public static TableResponseDTO getMetadataDetails(String developerName) {
        if (String.isBlank(developerName)) {
            return emptyResponse('Selection is required.');
        }
        ExplorerRecordDTO explorer = getExplorerRecord(developerName);
        if (explorer == null) {
            return emptyResponse('No explorer record found for the provided selection.');
        }

        // Check if we have a class name for retrieval
        if (String.isBlank(explorer.classNameForRetrieval)) {
            return emptyResponse('Not yet supported. No invocable class configured for metadata type: ' + developerName);
        }

        // Map of known invocable classes (can be extended)
        Map<String, String> classMapping = new Map<String, String>{
            'AssignmentRuleMetadataAction' => 'AssignmentRuleMetadataAction'
        };

        // Get the actual class name (in case it's different from the mapping key)
        String invocableClassName = explorer.classNameForRetrieval;

        // Validate that we know how to handle this class
        if (!classMapping.containsKey(invocableClassName)) {
            return emptyResponse('Unsupported invocable class: ' + invocableClassName);
        }

        // Prepare to call the invocable class
        String jsonResponse;
        try {
            // Handle specific invocable classes
            if (invocableClassName == 'AssignmentRuleMetadataAction') {
                // For AssignmentRuleMetadataAction, we need to call it with a request
                List<AssignmentRuleMetadataAction.Request> requests = new List<AssignmentRuleMetadataAction.Request>();
                AssignmentRuleMetadataAction.Request request = new AssignmentRuleMetadataAction.Request();
                request.objectName = null; // null means all objects
                requests.add(request);

                List<AssignmentRuleMetadataAction.Response> responses = AssignmentRuleMetadataAction.getAssignmentRulesJSON(requests);

                if (responses != null && !responses.isEmpty()) {
                    AssignmentRuleMetadataAction.Response response = responses[0];
                    if (response != null && response.success) {
                        jsonResponse = response.jsonString;
                    } else {
                        return emptyResponse('invocable class execution failed: ' + (response != null ? response.message : 'Unknown error'));
                    }
                } else {
                    return emptyResponse('invocable class returned no response');
                }
            } else {
                // For other classes, we would handle them here
                return emptyResponse('Unsupported invocable class: ' + invocableClassName);
            }
        } catch (Exception ex) {
            System.debug('Error in getMetadataDetails: ' + ex.getMessage());
            System.debug('Stack trace: ' + ex.getStackTraceString());
            return emptyResponse('invocable class invocation failed: ' + ex.getMessage());
        }
        if (String.isBlank(jsonResponse)) {
            return emptyResponse('invocable class returned no data.');
        }

        TableResponseDTO dto = new TableResponseDTO();
        dto.rawJson = jsonResponse;

        try {
            Map<String, Object> parsed = (Map<String, Object>) JSON.deserializeUntyped(jsonResponse);

            // Parse headers as columns (for AssignmentRuleMetadataAction)
            List<Object> headers = (List<Object>) parsed.get('headers');
            List<TableColumnDTO> colDtos = new List<TableColumnDTO>();
            if (headers != null) {
                for (Object hObj : headers) {
                    String header = (String) hObj;
                    TableColumnDTO col = new TableColumnDTO();
                    col.label = header;
                    col.fieldName = header;
                    col.type = 'text';
                    colDtos.add(col);
                }
            }
            dto.columns = colDtos;

            // Parse data rows (for AssignmentRuleMetadataAction)
            List<Object> rows = (List<Object>) parsed.get('data');
            dto.data = rows == null ? new List<Object>() : rows;

            // If columns empty, infer from first row keys (fallback)
            if ((dto.columns == null || dto.columns.isEmpty()) && dto.data != null && !dto.data.isEmpty()) {
                Map<String, Object> first = (Map<String, Object>) dto.data[0];
                List<TableColumnDTO> inferred = new List<TableColumnDTO>();
                for (String key : first.keySet()) {
                    TableColumnDTO col = new TableColumnDTO();
                    col.label = key;
                    col.fieldName = key;
                    col.type = 'text';
                    inferred.add(col);
                }
                dto.columns = inferred;
            }
            if (dto.columns == null) dto.columns = new List<TableColumnDTO>();
            if (dto.data == null) dto.data = new List<Object>();

            return dto;
        } catch (Exception ex) {
            System.debug('Error parsing JSON response: ' + ex.getMessage());
            System.debug('Stack trace: ' + ex.getStackTraceString());
            return emptyResponse('Failed to parse invocable class response: ' + ex.getMessage());
        }
    }

    private static TableResponseDTO emptyResponse(String message) {
        TableResponseDTO dto = new TableResponseDTO();
        dto.columns = new List<TableColumnDTO>();
        dto.data = new List<Object>();
        dto.rawJson = message;
        return dto;
    }

}