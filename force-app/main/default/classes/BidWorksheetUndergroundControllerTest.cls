@IsTest
private class BidWorksheetUndergroundControllerTest {
    
    private static Opportunity createOpp() {
        Contact con = new Contact(LastName = 'Test');
        insert con;
        
        Opportunity opp = new Opportunity(
            Name = 'Test Opp',
            StageName = 'Prospecting',
            CloseDate = Date.today(),
            Contact__c = con.Id,
            Job_Prefix__c = 'BF'
        );
        insert opp;
        return opp;
    }
    
    @IsTest
    static void testMetadataFetching() {
        Test.startTest();
        System.assertNotEquals(null, BidWorksheetUndergroundController.getSheet1Items());
        System.assertNotEquals(null, BidWorksheetUndergroundController.getSheet2Items());
        System.assertNotEquals(null, BidWorksheetUndergroundController.getEstimateSheetItems());
        System.assertNotEquals(null, BidWorksheetUndergroundController.getSOVItems());
        Test.stopTest();
    }
    
    @IsTest
    static void testUndergroundSaveLoad() {
        Opportunity opp = createOpp();
        String json = '{"test":"data"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveSheet(opp.Id, encoded);
        BidWorksheetUndergroundController.autoSaveSheet(opp.Id, encoded);
        String loaded = BidWorksheetUndergroundController.loadLatestSheet(opp.Id);
        Integer nextVer = BidWorksheetUndergroundController.getNextVersionNumber(opp.Id);
        List<BidWorksheetUndergroundController.VersionHistoryWrapper> history = BidWorksheetUndergroundController.getVersionHistory(opp.Id);
        Test.stopTest();
        
        System.assertNotEquals(null, loaded);
        System.assert(history.size() > 0);
    }

    @IsTest
    static void testEstimateSaveLoad() {
        Opportunity opp = createOpp();
        String json = '{"estimate":123}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        Id cvId = BidWorksheetUndergroundController.saveEstimateSheet(opp.Id, encoded);
        BidWorksheetUndergroundController.autoSaveEstimateSheet(opp.Id, encoded);
        String loadedBase64 = BidWorksheetUndergroundController.loadLatestEstimateSheet(opp.Id);
        String loadedJson = EncodingUtil.base64Decode(loadedBase64).toString();
        
        BidWorksheetUndergroundController.getNextVersionNumber_Estimate(opp.Id);
        BidWorksheetUndergroundController.getVersionHistory_Estimate(opp.Id);
        BidWorksheetUndergroundController.loadVersionById_Estimate(cvId);
        Test.stopTest();
        
        System.assertEquals(json, loadedJson);
    }

    @IsTest
    static void testSOVSaveLoad() {
        Opportunity opp = createOpp();
        String json = '{"sov":"test"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        Id cvId = BidWorksheetUndergroundController.saveSOVSheet(opp.Id, encoded);
        BidWorksheetUndergroundController.autoSaveSOVSheet(opp.Id, encoded);
        BidWorksheetUndergroundController.loadSOVSheet(opp.Id);
        BidWorksheetUndergroundController.loadLatestSOVSheet(opp.Id);
        BidWorksheetUndergroundController.getNextVersionNumber_SOV(opp.Id);
        BidWorksheetUndergroundController.getVersionHistory_SOV(opp.Id);
        BidWorksheetUndergroundController.loadVersionById_SOV(cvId);
        Test.stopTest();
    }

    @IsTest
    static void testDesignSaveLoad() {
        Opportunity opp = createOpp();
        String json = '{"design":"ok"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        Id cvId = BidWorksheetUndergroundController.saveDesignWorksheet(opp.Id, encoded);
        BidWorksheetUndergroundController.autoSaveDesignWorksheet(opp.Id, encoded);
        BidWorksheetUndergroundController.loadDesignWorksheet(opp.Id);
        BidWorksheetUndergroundController.loadLatestDesignWorksheet(opp.Id);
        BidWorksheetUndergroundController.getNextVersionNumber_Design(opp.Id);
        BidWorksheetUndergroundController.getVersionHistory_Design(opp.Id);
        BidWorksheetUndergroundController.loadVersionById_Design(cvId);
        Test.stopTest();
    }

    @IsTest
    static void testUpdateOpportunityFieldsExtended() {
        Opportunity opp = createOpp();
        Map<String, Object> fieldMap = new Map<String, Object>{
            'Amount' => 5000,
            'Description' => 'Test Desc',
            'CloseDate' => String.valueOf(Date.today()),
            'Probability' => 50.5
        };
        
        Test.startTest();
        BidWorksheetUndergroundController.updateOpportunityFields(opp.Id, JSON.serialize(fieldMap));
        
        // Test error handling
        try {
            BidWorksheetUndergroundController.updateOpportunityFields(null, '{}');
        } catch (Exception e) {}
        Test.stopTest();
        
        Opportunity updated = [SELECT Amount FROM Opportunity WHERE Id = :opp.Id];
        System.assertEquals(5000, updated.Amount);
    }
}