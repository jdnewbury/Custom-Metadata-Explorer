public class JN_GetUserPermissionSetsAction {
    // Input wrapper class
    public class GetUserPermissionSetsInput {
        @InvocableVariable(label='User ID or Username' required=true)
        public String userId;
    }
    // Output wrapper class
    public class GetUserPermissionSetsOutput {
        @InvocableVariable(label='Permission Sets')
        public String permissionSets;
        @InvocableVariable(label='Output Message')
        public String outputMessage;
    }
    @InvocableMethod(label='Get User Permission Sets' description='Retrieves permission sets assigned to a specific user')
    public static List<GetUserPermissionSetsOutput> getUserPermissionSets(List<GetUserPermissionSetsInput> inputs) {
        // Initialize output
        List<GetUserPermissionSetsOutput> outputs = new List<GetUserPermissionSetsOutput>();
        GetUserPermissionSetsOutput output = new GetUserPermissionSetsOutput();
        try {
            // Get the input
            String userId = inputs[0].userId;
            // Handle username vs ID
            Id userIdValue;
            if (userId.length() == 18 || userId.length() == 15) {
                // It's likely an ID
                userIdValue = userId;
            } else {
                // It's likely a username, look up the ID
                List<User> users = [SELECT Id FROM User WHERE Username = :userId OR Email = :userId LIMIT 1];
                if (!users.isEmpty()) {
                    userIdValue = users[0].Id;
                } else {
                    output.outputMessage = 'User not found: ' + userId;
                    output.permissionSets = 'No user found with the specified username or ID.';
                    outputs.add(output);
                    return outputs;
                }
            }
            // Query for permission set assignments
            List<PermissionSetAssignment> permissionSetAssignments = [
                SELECT Id, PermissionSet.Id, PermissionSet.Name, PermissionSet.Label, PermissionSet.IsCustom, PermissionSet.Type
                FROM PermissionSetAssignment
                WHERE AssigneeId = :userIdValue
            ];
            // Format the results
            if (permissionSetAssignments.isEmpty()) {
                output.outputMessage = 'No permission sets found for this user.';
                output.permissionSets = 'The specified user has no permission sets assigned.';
            } else {
                // Build formatted string of permission sets
                String formattedPermissionSets = '';
                for (PermissionSetAssignment psa : permissionSetAssignments) {
                    String permType = psa.PermissionSet.IsCustom ? 'Custom' : 'Standard';
                    formattedPermissionSets += '• ' + psa.PermissionSet.Label +
                        ' (' + psa.PermissionSet.Name + ')' +
                        ' [' + permType + ']\n';
                }
                output.outputMessage = 'Successfully retrieved ' + permissionSetAssignments.size() + ' permission sets.';
                output.permissionSets = formattedPermissionSets;
            }
        } catch (Exception e) {
            output.outputMessage = 'Error: ' + e.getMessage();
            output.permissionSets = 'An error occurred while retrieving permission sets.';
        }
        outputs.add(output);
        return outputs;
    }
}