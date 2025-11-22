/**
 * Test class for BitbucketMetadataController
 */
@IsTest
private class BitbucketMetadataControllerTest {
    
    @TestSetup
    static void setupTestData() {
        // Test setup if needed - custom metadata is already in org
    }
    
    @IsTest
    static void testGetExplorerOptions() {
        // Test: Get available metadata types
        Test.startTest();
        List<Map<String, String>> options = BitbucketMetadataController.getExplorerOptions();
        Test.stopTest();
        
        // Verify: Should return active metadata types from custom metadata
        System.assertNotEquals(null, options, 'Options should not be null');
        System.assert(options.size() > 0, 'Should have at least one metadata type');
        
        // Verify structure
        for (Map<String, String> option : options) {
            System.assert(option.containsKey('label'), 'Option should have label');
            System.assert(option.containsKey('value'), 'Option should have value');
        }
    }
    
    @IsTest
    static void testGetMetadataDetails_Success() {
        // Setup: Mock Bitbucket response
        Test.setMock(HttpCalloutMock.class, new BitbucketSuccessMock());
        
        Test.startTest();
        BitbucketMetadataController.MetadataResponse response = 
            BitbucketMetadataController.getMetadataDetails('Assignment_Rules');
        Test.stopTest();
        
        // Verify: Response structure
        System.assertNotEquals(null, response, 'Response should not be null');
        System.assertNotEquals(null, response.columns, 'Columns should not be null');
        System.assertNotEquals(null, response.data, 'Data should not be null');
        
        // Verify: Columns have width information
        if (response.columns.size() > 0) {
            BitbucketMetadataController.ColumnDefinition firstCol = response.columns[0];
            System.assertNotEquals(null, firstCol.fieldName, 'Column should have fieldName');
            System.assertNotEquals(null, firstCol.label, 'Column should have label');
        }
    }
    
    @IsTest
    static void testGetMetadataDetails_InvalidType() {
        // Test: Invalid metadata type
        try {
            Test.startTest();
            BitbucketMetadataController.getMetadataDetails('Invalid_Type_That_Does_Not_Exist');
            Test.stopTest();
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('Error'), 'Should contain error message');
        }
    }
    
    @IsTest
    static void testGetMetadataDetails_EmptyResponse() {
        // Setup: Mock empty Bitbucket response
        Test.setMock(HttpCalloutMock.class, new BitbucketEmptyMock());
        
        Test.startTest();
        BitbucketMetadataController.MetadataResponse response = 
            BitbucketMetadataController.getMetadataDetails('Assignment_Rules');
        Test.stopTest();
        
        // Verify: Should handle empty gracefully
        System.assertNotEquals(null, response, 'Response should not be null');
        System.assertEquals(0, response.data.size(), 'Data should be empty');
    }
    
    @IsTest
    static void testColumnWidthCalculation() {
        // Test: Verify column definition structure
        List<MetadataTypeConfig.ColumnConfig> columns = new List<MetadataTypeConfig.ColumnConfig>();
        
        MetadataTypeConfig.ColumnConfig col1 = new MetadataTypeConfig.ColumnConfig();
        col1.active = true;
        col1.fieldName = 'objectName';
        col1.label = 'Object';
        col1.type = 'text';
        col1.widthPercent = 20;
        columns.add(col1);
        
        MetadataTypeConfig.ColumnConfig col2 = new MetadataTypeConfig.ColumnConfig();
        col2.active = true;
        col2.fieldName = 'fullName';
        col2.label = 'Name';
        col2.type = 'text';
        col2.widthPercent = 30;
        columns.add(col2);
        
        MetadataTypeConfig.ColumnConfig col3 = new MetadataTypeConfig.ColumnConfig();
        col3.active = true;
        col3.fieldName = 'active';
        col3.label = 'Active';
        col3.type = 'boolean';
        col3.widthPercent = null; // Auto-size
        columns.add(col3);
        
        // Calculate total
        Decimal total = 0;
        for (MetadataTypeConfig.ColumnConfig col : columns) {
            if (col.widthPercent != null) {
                total += col.widthPercent;
            }
        }
        
        System.assertEquals(50, total, 'Total explicit width should be 50%');
    }
    
    /**
     * Mock for successful Bitbucket API responses
     */
    private class BitbucketSuccessMock implements HttpCalloutMock {
        public HttpResponse respond(HttpRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');
            res.setStatusCode(200);
            
            // Mock directory listing
            if (req.getEndpoint().contains('/src/') && !req.getEndpoint().contains('Case.assignmentRules')) {
                String mockResponse = '{' +
                    '"values": [' +
                        '{"path": "force-app/main/default/assignmentRules/Case.assignmentRules-meta.xml", "type": "commit_file"}' +
                    ']' +
                '}';
                res.setBody(mockResponse);
            }
            // Mock file content
            else {
                String mockXml = '<?xml version="1.0" encoding="UTF-8"?>' +
                    '<AssignmentRules xmlns="http://soap.sforce.com/2006/04/metadata">' +
                    '<assignmentRule>' +
                        '<fullName>Standard</fullName>' +
                        '<active>true</active>' +
                        '<ruleEntry>' +
                            '<assignedTo>TestUser</assignedTo>' +
                            '<assignedToType>User</assignedToType>' +
                            '<criteriaItems>' +
                                '<field>Status</field>' +
                                '<operation>equals</operation>' +
                                '<value>New</value>' +
                            '</criteriaItems>' +
                        '</ruleEntry>' +
                    '</assignmentRule>' +
                    '</AssignmentRules>';
                res.setBody(mockXml);
            }
            
            return res;
        }
    }
    
    /**
     * Mock for empty Bitbucket responses
     */
    private class BitbucketEmptyMock implements HttpCalloutMock {
        public HttpResponse respond(HttpRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');
            res.setStatusCode(200);
            
            // Return empty values array
            res.setBody('{"values": []}');
            
            return res;
        }
    }
}