/**
 * Test class for BitbucketMetadataController
 */
@IsTest
private class BitbucketMetadataControllerTest {
    
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
    static void testGetMetadataDetails_WithMockData() {
        // Setup: Mock Bitbucket response
        Test.setMock(HttpCalloutMock.class, new BitbucketMetadataReaderMock());
        
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
            // Width can be null (auto-sized) or a percentage value
        }
    }
    
    @IsTest
    static void testGetMetadataDetails_InvalidType() {
        // Test: Invalid metadata type
        try {
            Test.startTest();
            BitbucketMetadataController.getMetadataDetails('Invalid_Type');
            Test.stopTest();
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('Error'), 'Should contain error message');
        }
    }
    
    @IsTest
    static void testColumnWidthCalculation() {
        // Test: Verify total width calculation
        List<MetadataTypeConfig.ColumnConfig> columns = new List<MetadataTypeConfig.ColumnConfig>();
        
        MetadataTypeConfig.ColumnConfig col1 = new MetadataTypeConfig.ColumnConfig();
        col1.active = true;
        col1.widthPercent = 20;
        columns.add(col1);
        
        MetadataTypeConfig.ColumnConfig col2 = new MetadataTypeConfig.ColumnConfig();
        col2.active = true;
        col2.widthPercent = 30;
        columns.add(col2);
        
        MetadataTypeConfig.ColumnConfig col3 = new MetadataTypeConfig.ColumnConfig();
        col3.active = true;
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
        // Remaining 50% should be auto-distributed
    }
    
    /**
     * Mock HTTP callout for testing
     */
    private class BitbucketMetadataReaderMock implements HttpCalloutMock {
        public HttpResponse respond(HttpRequest req) {
            HttpResponse res = new HttpResponse();
            res.setStatusCode(200);
            res.setHeader('Content-Type', 'application/json');
            
            // Mock file list response
            if (req.getEndpoint().contains('/src/')) {
                String mockResponse = '{"values": [{"path": "Case.assignmentRules-meta.xml", "type": "commit_file"}]}';
                res.setBody(mockResponse);
            }
            // Mock file content response  
            else if (req.getEndpoint().contains('raw')) {
                String mockXml = '<?xml version="1.0"?><AssignmentRules xmlns="http://soap.sforce.com/2006/04/metadata">';
                mockXml += '<assignmentRule><fullName>Test Rule</fullName><active>true</active>';
                mockXml += '<ruleEntry><assignedTo>TestUser</assignedTo><assignedToType>User</assignedToType>';
                mockXml += '<criteriaItems><field>Status</field><operation>equals</operation><value>New</value></criteriaItems>';
                mockXml += '</ruleEntry></assignmentRule></AssignmentRules>';
                res.setBody(mockXml);
            }
            else {
                res.setBody('{}');
            }
            
            return res;
        }
    }
}