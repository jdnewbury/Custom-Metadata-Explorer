/*
*********************************************************
Apex Class Name    : ProjectX_ApexUtilsTest
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
@isTest
private class ProjectX_ApexUtilsTest {

    @isTest
    static void testIsNullOrEmpty() {
        System.assertEquals(true, ProjectX_ApexUtils.isNullOrEmpty(null));
        System.assertEquals(true, ProjectX_ApexUtils.isNullOrEmpty(''));
        System.assertEquals(false, ProjectX_ApexUtils.isNullOrEmpty('Hello'));
    }

    @isTest
    static void testSafeToString() {
        System.assertEquals('', ProjectX_ApexUtils.safeToString(null));
        System.assertEquals('123', ProjectX_ApexUtils.safeToString(123));
    }

    @isTest
    static void testFormatDate() {
        Date d = Date.today();
        String formatted = ProjectX_ApexUtils.formatDate(d, 'yyyy-MM-dd');
        //System.assertEquals(d.format('yyyy-MM-dd'), formatted);
        System.assertEquals('', ProjectX_ApexUtils.formatDate(null, 'yyyy-MM-dd'));
    }

    @isTest
    static void testListToMap() {
        Account acc = new Account(Name='Test Account');
        insert acc;
        List<SObject> accList = new List<SObject>{acc};
        Map<Id, SObject> accMap = ProjectX_ApexUtils.listToMap(accList);
        System.assertEquals(acc.Id, accMap.get(acc.Id).Id);
    }

    @isTest
    static void testLogError() {
        try {
            Integer i = 1 / 0; // Force exception
        } catch (Exception ex) {
            ProjectX_ApexUtils.logError(ex);
            // No assert needed, just ensures no exception thrown
        }
    }

    @isTest
    static void testGetCurrentUser() {
        User currentUser = ProjectX_ApexUtils.getCurrentUser();
        System.assertEquals(UserInfo.getUserId(), currentUser.Id);
    }

    @isTest
    static void testSafeInsert() {
        List<Account> accList = new List<Account>{
            new Account(Name='Safe Insert Test')
        };
        ProjectX_ApexUtils.safeInsert(accList);
        System.assertEquals(1, [SELECT COUNT() FROM Account WHERE Name='Safe Insert Test']);
    }

    @isTest
    static void testJsonMethods() {
        Account acc = new Account(Name='JSON Test');
        String jsonStr = ProjectX_ApexUtils.toJson(acc);
        System.assert(jsonStr.contains('JSON Test'));

        Account accFromJson = (Account) ProjectX_ApexUtils.fromJson(jsonStr, Account.class);
        System.assertEquals('JSON Test', accFromJson.Name);
    }

    @isTest
    static void testGetRecordById() {
        Account acc = new Account(Name='SOQL Helper Test');
        insert acc;

        SObject result = ProjectX_ApexUtils.getRecordById(acc.Id, 'Account', new List<String>{'Id','Name'});
        System.assertEquals(acc.Id, result.get('Id'));
        System.assertEquals('SOQL Helper Test', result.get('Name'));

        // Negative case
        System.assertEquals(null, ProjectX_ApexUtils.getRecordById(null, 'Account', new List<String>{'Id'}));
    }

    @isTest
    static void testRecordExists() {
        Account acc = new Account(Name='Exists Test');
        insert acc;

        System.assertEquals(true, ProjectX_ApexUtils.recordExists(acc.Id, 'Account'));
        System.assertEquals(false, ProjectX_ApexUtils.recordExists(null, 'Account'));
    }

    @isTest
    static void testRunQuery() {
        Account acc1 = new Account(Name='Query Test 1');
        Account acc2 = new Account(Name='Query Test 2');
        insert new List<Account>{acc1, acc2};

        List<SObject> results = ProjectX_ApexUtils.runQuery('Account', new List<String>{'Id','Name'}, 'Name LIKE \'Query Test%\'', 10);
        System.assertEquals(2, results.size());

        // Negative case: empty fields
        List<SObject> emptyResults = ProjectX_ApexUtils.runQuery('Account', new List<String>(), 'Name LIKE \'Query Test%\'', 10);
        System.assertEquals(0, emptyResults.size());
    }

   @IsTest
    public static void logErrorTest() {

        // Create a test user
        User testUser = new User(
            FirstName = 'Test',
            LastName = 'User',
            Email = 'test@example.com',
            Username = 'test@' + DateTime.now().getTime() + 'example.com',
            Alias = 'tuser',
            TimeZoneSidKey = 'America/Los_Angeles',
            LocaleSidKey = 'en_US',
            EmailEncodingKey = 'UTF-8',
            ProfileId = [SELECT Id FROM Profile WHERE Name = 'Standard User' LIMIT 1].Id,
            LanguageLocaleKey = 'en_US'
        );
        insert testUser;

        // Generate a random record ID
        Id recordId = Id.valueOf('001' + '15'.repeat(12)); // Generate a random Account ID

        // Test logging an error
        Test.startTest();
        // Commenting out the logError while we wait for the object and fields to be created
        // ProjectX_ApexUtils.logError('Test error message', 'TestClassName', 'TestErrorType', recordId, testUser.Email);
        Test.stopTest();

        // Verify that an error log record was created
        // Commenting out the logError while we wait for the object and fields to be created
        // Integer errorLogCount = [SELECT COUNT() FROM QHub_Application_Error_Log__c];
        // System.assertEquals(1, errorLogCount, 'Error log record should be created');

    }

@IsTest
    public static void generateRandomKey_negativeLength() {

        // Generate random keys
        Test.startTest();
        String genKey1 = ProjectX_ApexUtils.generateRandomKey(-10, NULL);
        Test.stopTest();

        // Check key length
        System.assertEquals(genKey1, '');

    }

}