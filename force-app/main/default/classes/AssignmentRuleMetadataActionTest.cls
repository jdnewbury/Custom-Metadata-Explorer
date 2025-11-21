@isTest
public class AssignmentRuleMetadataActionTest {
    
    @isTest
    static void testGetAssignmentRulesJSONWithValidInput() {
        // Test the method with valid input
        AssignmentRuleMetadataAction.Request req = new AssignmentRuleMetadataAction.Request();
        req.objectName = 'Case';
        
        List<AssignmentRuleMetadataAction.Request> requests = new List<AssignmentRuleMetadataAction.Request>{req};
        
        List<AssignmentRuleMetadataAction.Response> results = AssignmentRuleMetadataAction.getAssignmentRulesJSON(requests);
        
        // Verify results
        System.assertNotEquals(null, results);
        System.assertEquals(1, results.size());
        System.assertEquals(true, results[0].success);
        System.assertNotEquals(null, results[0].jsonString);
        
        // Parse the JSON to verify structure
        Map<String, Object> parsedJson = (Map<String, Object>) JSON.deserializeUntyped(results[0].jsonString);
        System.assert(parsedJson.containsKey('headers'));
        System.assert(parsedJson.containsKey('data'));
        System.assert(parsedJson.containsKey('rowCount'));
    }
    
    @isTest
    static void testGetAssignmentRulesJSONWithNullInput() {
        // Test with null object name (should return all)
        AssignmentRuleMetadataAction.Request req = new AssignmentRuleMetadataAction.Request();
        req.objectName = null;
        
        List<AssignmentRuleMetadataAction.Request> requests = new List<AssignmentRuleMetadataAction.Request>{req};
        
        List<AssignmentRuleMetadataAction.Response> results = AssignmentRuleMetadataAction.getAssignmentRulesJSON(requests);
        
        // Verify results
        System.assertNotEquals(null, results);
        System.assertEquals(1, results.size());
        System.assertEquals(true, results[0].success);
        System.assertNotEquals(null, results[0].jsonString);
    }
    
    @isTest
    static void testGetAssignmentRulesJSONWithErrorHandling() {
        // Test with invalid object name
        AssignmentRuleMetadataAction.Request req = new AssignmentRuleMetadataAction.Request();
        req.objectName = 'NonExistentObject';
        
        List<AssignmentRuleMetadataAction.Request> requests = new List<AssignmentRuleMetadataAction.Request>{req};
        
        List<AssignmentRuleMetadataAction.Response> results = AssignmentRuleMetadataAction.getAssignmentRulesJSON(requests);
        
        // Verify results
        System.assertNotEquals(null, results);
        System.assertEquals(1, results.size());
        // Should still succeed even with no results
        System.assertEquals(true, results[0].success);
    }
    
    @isTest
    static void testGetAssignmentRulesJSONEmptyResult() {
        // Test with a valid object name but no rules
        AssignmentRuleMetadataAction.Request req = new AssignmentRuleMetadataAction.Request();
        req.objectName = 'Account'; // Assuming no assignment rules exist for Account
        
        List<AssignmentRuleMetadataAction.Request> requests = new List<AssignmentRuleMetadataAction.Request>{req};
        
        List<AssignmentRuleMetadataAction.Response> results = AssignmentRuleMetadataAction.getAssignmentRulesJSON(requests);
        
        // Verify results
        System.assertNotEquals(null, results);
        System.assertEquals(1, results.size());
        System.assertEquals(true, results[0].success);
        System.assertNotEquals(null, results[0].jsonString);
    }
}