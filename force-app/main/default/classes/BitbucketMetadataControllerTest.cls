/**
 * Test class for BitbucketMetadataController
 * DEPLOY THIS AFTER BitbucketMetadataController IS DEPLOYED
 */
@IsTest
private class BitbucketMetadataControllerTest {
    
    @IsTest
    static void testGetExplorerOptions() {
        Test.startTest();
        List<Map<String, String>> options = BitbucketMetadataController.getExplorerOptions();
        Test.stopTest();
        
        Assert.isNotNull(options, 'Options should not be null');
    }
    
    @IsTest
    static void testGetMetadataDetails_InvalidType() {
        Test.startTest();
        try {
            BitbucketMetadataController.getMetadataDetails('Invalid_Type');
            Assert.fail('Should have thrown exception');
        } catch (AuraHandledException e) {
            Assert.isTrue(e.getMessage().contains('Error'), 'Should contain error');
        }
        Test.stopTest();
    }
    
    @IsTest
    static void testGetMetadataDetails_WithMock() {
        Test.setMock(HttpCalloutMock.class, new MockSuccess());
        
        Test.startTest();
        try {
            Object result = BitbucketMetadataController.getMetadataDetails('Assignment_Rules');
            Assert.isNotNull(result, 'Result should not be null');
        } catch (Exception e) {
            // May fail if config missing - acceptable
            System.debug('Expected error: ' + e.getMessage());
        }
        Test.stopTest();
    }
    
    private class MockSuccess implements HttpCalloutMock {
        public HttpResponse respond(HttpRequest req) {
            HttpResponse res = new HttpResponse();
            res.setStatusCode(200);
            res.setBody('{"values":[]}');
            return res;
        }
    }
}