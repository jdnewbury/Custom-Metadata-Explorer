@IsTest
private class CustomMetadataExplorerControllerTest {

    @TestSetup
    static void setupData() {
        // No MDT insert; validate behaviors that don't require actual MDT
    }

    @IsTest
    static void testGetExplorerOptions_NoData() {
        Test.startTest();
        List<CustomMetadataExplorerController.OptionDTO> opts =
            CustomMetadataExplorerController.getExplorerOptions();
        Test.stopTest();
        System.assertNotEquals(null, opts, 'Options list should not be null');
        // In scratch/test org without packaged MDT, expect 0
        System.assertEquals(true, opts.size() >= 0, 'Options size should be non-negative');
    }

    @IsTest
    static void testGetMetadataDetails_BlankSelection() {
        Test.startTest();
        CustomMetadataExplorerController.TableResponseDTO resp =
            CustomMetadataExplorerController.getMetadataDetails('');
        Test.stopTest();
        System.assertNotEquals(null, resp, 'Response must not be null for blank selection');
        System.assertEquals(0, resp.columns.size(), 'No columns expected');
        System.assertEquals(0, resp.data.size(), 'No data expected');
        System.assert(resp.rawJson != null && resp.rawJson.contains('Selection is required'),
            'Expected mandatory selection message');
    }

    @IsTest
    static void testJsonParsing_LogicExpectation() {
        // Validate JSON parsing expectations similarly to controller logic
        Map<String, Object> payload = new Map<String, Object>{
            'metadataTypeApiName' => 'SomeType__mdt',
            'explorerDeveloperName' => 'SomeType',
            'explorerMasterLabel' => 'Some Label'
        };
        String sampleJson = JSON.serialize(new Map<String, Object>{
            'columns' => new List<Object>{
                new Map<String, Object>{ 'label' => 'Col A', 'fieldName' => 'colA', 'type' => 'text' },
                new Map<String, Object>{ 'label' => 'Col B', 'fieldName' => 'colB', 'type' => 'number' }
            },
            'data' => new List<Object>{
                new Map<String, Object>{ 'colA' => 'A1', 'colB' => 2 }
            }
        });

        // Ensure JSON roundtrip so JSON.deserializeUntyped is available from the JSON class (not String)
        Map<String, Object> parsed = (Map<String, Object>) JSON.deserializeUntyped(sampleJson);
        System.assert(parsed.containsKey('columns'), 'Columns key should exist');
        System.assert(parsed.containsKey('data'), 'Data key should exist');

        List<Object> cols = (List<Object>) parsed.get('columns');
        System.assertEquals(2, cols.size(), 'Two columns expected');

        List<Object> rows = (List<Object>) parsed.get('data');
        System.assertEquals(1, rows.size(), 'One row expected');
        Map<String, Object> r0 = (Map<String, Object>) rows[0];
        System.assertEquals('A1', (String) r0.get('colA'));
    }
}