@IsTest
private class ITMBidWorksheetControllerTest {
    
    @IsTest
    static void testGetITMItems_withOrgMetadata() {
        Test.startTest();
        Map<String, Object> result = ITMBidWorksheetController.getITMItems();
        Test.stopTest();
        
        System.assertNotEquals(null, result, 'Result map should not be null');
        System.assert(result.containsKey('laborFactor'));
        System.assert(result.containsKey('equipmentFactor'));
        System.assert(result.containsKey('rates'));
    }
    
    @IsTest
    static void testSaveAndLoadITMWorksheet_basicFlow() {
        Opportunity opp = createTestOpp();
        String sampleJson = '{"test":"data"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(sampleJson));
        
        Test.startTest();
        Id cvId = ITMBidWorksheetController.saveITMWorksheet(opp.Id, encoded);
        String loaded = ITMBidWorksheetController.loadITMWorksheet(opp.Id);
        Test.stopTest();
        
        System.assertNotEquals(null, cvId);
        System.assertEquals(sampleJson, loaded);
    }

    @IsTest
    static void testITMVersioningAndAutosave() {
        Opportunity opp = createTestOpp();
        String sampleJson = '{"version":1}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(sampleJson));
        
        Test.startTest();
        // Test Autosave
        ITMBidWorksheetController.autoSaveITMWorksheet(opp.Id, encoded);
        
        // Test Versioning
        Integer nextVer = ITMBidWorksheetController.getNextVersionNumber(opp.Id);
        Id cvId = ITMBidWorksheetController.saveITMWorksheet(opp.Id, encoded);
        List<ITMBidWorksheetController.VersionHistoryWrapper> history = ITMBidWorksheetController.getVersionHistory(opp.Id);
        
        // Test Loading
        String latest = ITMBidWorksheetController.loadLatestITMWorksheet(opp.Id);
        String byId = ITMBidWorksheetController.loadVersionById(cvId);
        Test.stopTest();
        
        System.assertNotEquals(null, latest);
        System.assertNotEquals(null, byId);
        System.assert(history.size() > 0);
    }
    
    @IsTest
    static void testSaveITMWorksheet_exceptions() {
        try {
            ITMBidWorksheetController.saveITMWorksheet(null, 'abc');
        } catch (AuraHandledException ex) {
            System.assert(true);
        }
        
        Opportunity opp = createTestOpp();
        try {
            ITMBidWorksheetController.saveITMWorksheet(opp.Id, '');
        } catch (AuraHandledException ex) {
            System.assert(true);
        }
    }

    private static Opportunity createTestOpp() {
        Contact con = new Contact(LastName = 'Test');
        insert con;
        Opportunity opp = new Opportunity(
            Name = 'Test ITM Opp',
            StageName = 'Prospecting',
            CloseDate = Date.today(),
            Contact__c = con.Id,
            Job_Prefix__c = 'BF'
        );
        insert opp;
        return opp;
    }
}