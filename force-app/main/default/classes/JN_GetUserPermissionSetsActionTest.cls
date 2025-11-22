@isTest
private class JN_GetUserPermissionSetsActionTest {
    @isTest
    static void testGetUserPermissionSets() { // Create a test user
        Profile p = [SELECT Id FROM Profile WHERE Name='Standard User' LIMIT 1];
        User u = new User(
            Alias = 'testuser',
            Email = 'testuser@example.com',
            EmailEncodingKey = 'UTF-8',
            LastName = 'Testing', LanguageLocaleKey = 'en_US',
            LocaleSidKey = 'en_US',
            ProfileId = p.Id,
            TimeZoneSidKey = 'America/Los_Angeles',
            UserName = 'testuser' + DateTime.now().getTime() + '@example.com'
        );
        insert u;

        // Create a test permission set
        PermissionSet ps = new PermissionSet(
            Label = 'Test Permission Set',
            Name = 'TestPermissionSet'
        );
        insert ps;

        // Assign the permission set to the user
        PermissionSetAssignment psa = new PermissionSetAssignment(
            AssigneeId = u.Id,
            PermissionSetId = ps.Id
        );
        insert psa;

        // Call the invocable method
        Test.startTest();
        JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput input = new JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput();
        input.userId = u.Id;

        List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsOutput> outputs = JN_GetUserPermissionSetsAction.getUserPermissionSets(new List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput>{input});
        Test.stopTest();

        // Verify results
        System.assertEquals(1, outputs.size());
        System.assert(outputs[0].permissionSets.contains('Test Permission Set'));
        System.assert(outputs[0].outputMessage.contains('Successfully retrieved'));
    }

    @isTest
    static void testUserNotFound() {
        // Call the invocable method with non-existent user
        Test.startTest();
        JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput input = new JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput();
        input.userId = 'nonexistent@example.com';

        List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsOutput> outputs = JN_GetUserPermissionSetsAction.getUserPermissionSets(new List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput>{input});
        Test.stopTest();

        // Verify error handling
        System.assertEquals(1, outputs.size());
        System.assert(outputs[0].outputMessage.contains('User not found'));
    }

    @isTest
    static void testNoPermissionSets() {
        // Create a test user without permission sets
        Profile p = [SELECT Id FROM Profile WHERE Name='Standard User' LIMIT 1];
        User u = new User(
            Alias = 'noperms',
            Email = 'noperms@example.com',
            EmailEncodingKey = 'UTF-8',
            LastName = 'NoPerms',
            LanguageLocaleKey = 'en_US',
            LocaleSidKey = 'en_US',
            ProfileId = p.Id,
            TimeZoneSidKey = 'America/Los_Angeles',
            UserName = 'noperms' + DateTime.now().getTime() + '@example.com'
        );
        insert u;

        // Call the invocable method using the User Id
        Test.startTest();
        JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput input = new JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput();
        input.userId = u.Id;

        // Run as the test user to ensure consistent behavior
        System.runAs(u) {
            List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsOutput> outputs = JN_GetUserPermissionSetsAction.getUserPermissionSets(new List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput>{input});

            // Verification
            System.assertEquals(1, outputs.size(), 'Should have exactly one output');

            // Verify that the output message indicates success in some way
            Boolean messageIndicatesSuccess = outputs[0].outputMessage != null &&
                (
                    outputs[0].outputMessage.contains('Success') ||
                    outputs[0].outputMessage.contains('success') ||
                    outputs[0].outputMessage.contains('Retrieved') ||
                    outputs[0].outputMessage.contains('retrieved')
                );

            System.assert(messageIndicatesSuccess,
                'Expected output message to indicate successful execution. ' +
                'Actual message: ' + outputs[0].outputMessage);
        }
        Test.stopTest();
    }

    @isTest
    static void testUserEmailLookup() {
        // Create a test user without permission sets
        Profile p = [SELECT Id FROM Profile WHERE Name='Standard User' LIMIT 1];
        User u = new User(
            Alias = 'emtest', // shortened to 8 chars or less
            Email = 'emailtest@example.com',
            EmailEncodingKey = 'UTF-8',
            LastName = 'EmailTest',
            LanguageLocaleKey = 'en_US',
            LocaleSidKey = 'en_US',
            ProfileId = p.Id,
            TimeZoneSidKey = 'America/Los_Angeles',
            UserName = 'emailtest' + DateTime.now().getTime() + '@example.com'
        );
        insert u;

        // Call the invocable method using Email instead of Id
        // This will test the lookup by email path (userIdValue = users[0].Id)
        Test.startTest();
        JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput input = new JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput();
        input.userId = u.Email; // Using email to test this code path

        List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsOutput> outputs = JN_GetUserPermissionSetsAction.getUserPermissionSets(new List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput>{input});
        Test.stopTest();

        // Basic verification
        System.assertEquals(1, outputs.size(), 'Should have exactly one output');
        System.assert(String.isNotBlank(outputs[0].outputMessage), 'Should have a non-empty output message');
    }

    @isTest
    static void testExceptionHandling() {
        // Test with null input to trigger exception
        Test.startTest();
        JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput input = new JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput();
        input.userId = null;
        // This should trigger the exception path
        List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsOutput> outputs = JN_GetUserPermissionSetsAction.getUserPermissionSets(new List<JN_GetUserPermissionSetsAction.GetUserPermissionSetsInput>{input});
        Test.stopTest();

        // Verify exception handling
        System.assertEquals(1, outputs.size(), 'Should have exactly one output even with exception');
        System.assert(outputs[0].outputMessage.contains('Error') || outputs[0].outputMessage.contains('Exception'), 'Output message should indicate an error occurred');
    }
}