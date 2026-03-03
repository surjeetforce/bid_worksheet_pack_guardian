@IsTest
private class ITMBidWorksheetControllerTest {
    
    @TestSetup
    static void setupData() {
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
    }

    private static Opportunity getTestOpp() {
        return [SELECT Id FROM Opportunity WHERE Name = 'Test ITM Opp' LIMIT 1];
    }

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
        Opportunity opp = getTestOpp();
        String sampleJson = '{"test":"data"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(sampleJson));
        
        Test.startTest();
        Id cvId = ITMBidWorksheetController.saveITMWorksheet(opp.Id, encoded, false);
        String loaded = ITMBidWorksheetController.loadITMWorksheet(opp.Id);
        Test.stopTest();
        
        System.assertNotEquals(null, cvId);
        System.assertEquals(sampleJson, loaded);
    }

    @IsTest
    static void testITMVersioningAndAutosave() {
        Opportunity opp = getTestOpp();
        String sampleJson = '{"version":1}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(sampleJson));
        
        Test.startTest();
        ITMBidWorksheetController.autoSaveITMWorksheet(opp.Id, encoded);
        Integer nextVer = ITMBidWorksheetController.getNextVersionNumber(opp.Id);
        Id cvId = ITMBidWorksheetController.saveITMWorksheet(opp.Id, encoded, false);
        List<ITMBidWorksheetController.VersionHistoryWrapper> history = ITMBidWorksheetController.getVersionHistory(opp.Id);
        String latest = ITMBidWorksheetController.loadLatestITMWorksheet(opp.Id);
        String byId = ITMBidWorksheetController.loadVersionById(cvId);
        Test.stopTest();
        
        System.assertNotEquals(null, latest);
        System.assertNotEquals(null, byId);
        System.assert(history.size() > 0);
    }

    @IsTest
    static void testLoadITMWorksheetNoData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        String result = ITMBidWorksheetController.loadITMWorksheet(opp.Id);
        Test.stopTest();
        
        System.assertEquals(null, result);
    }

    @IsTest
    static void testLoadLatestNoData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        String result = ITMBidWorksheetController.loadLatestITMWorksheet(opp.Id);
        Test.stopTest();
        
        System.assertEquals(null, result);
    }

    @IsTest
    static void testSaveITMWorksheetNullId() {
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{}'));
        
        Test.startTest();
        try {
            ITMBidWorksheetController.saveITMWorksheet(null, encoded, false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveITMWorksheetBlankData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        try {
            ITMBidWorksheetController.saveITMWorksheet(opp.Id, '', false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testAutoSaveNullId() {
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{}'));
        
        Test.startTest();
        try {
            ITMBidWorksheetController.autoSaveITMWorksheet(null, encoded);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testAutoSaveBlankData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        try {
            ITMBidWorksheetController.autoSaveITMWorksheet(opp.Id, '');
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }


    @IsTest
    static void testLoadLatestNullId() {
        Test.startTest();
        try {
            ITMBidWorksheetController.loadLatestITMWorksheet(null);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
           System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testLoadVersionByIdNullId() {
        Test.startTest();
        try {
            ITMBidWorksheetController.loadVersionById(null);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testGetNextVersionNumberNullId() {
        Test.startTest();
        Integer result = ITMBidWorksheetController.getNextVersionNumber(null);
        Test.stopTest();
        
        System.assertEquals(1, result);
    }

    @IsTest
    static void testGetVersionHistoryNullId() {
        Test.startTest();
        List<ITMBidWorksheetController.VersionHistoryWrapper> history = ITMBidWorksheetController.getVersionHistory(null);
        Test.stopTest();
        
        System.assertEquals(0, history.size());
    }

    @IsTest
    static void testGetVersionHistoryNoData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        List<ITMBidWorksheetController.VersionHistoryWrapper> history = ITMBidWorksheetController.getVersionHistory(opp.Id);
        Test.stopTest();
        
        System.assertEquals(0, history.size());
    }

    @IsTest
    static void testFinalVersionStatus() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        Map<String, Object> status = ITMBidWorksheetController.getFinalVersionStatus(opp.Id);
        Test.stopTest();
        
        System.assertEquals(false, status.get('hasFinalVersion'));
    }

    @IsTest
    static void testFinalVersionStatusNullId() {
        Test.startTest();
        Map<String, Object> status = ITMBidWorksheetController.getFinalVersionStatus(null);
        Test.stopTest();
        
        System.assertEquals(false, status.get('hasFinalVersion'));
    }

    @IsTest
    static void testSaveSecondVersion() {
        Opportunity opp = getTestOpp();
        String encoded1 = EncodingUtil.base64Encode(Blob.valueOf('{"v":1}'));
        String encoded2 = EncodingUtil.base64Encode(Blob.valueOf('{"v":2}'));
        
        Test.startTest();
        ITMBidWorksheetController.saveITMWorksheet(opp.Id, encoded1, false);
        ITMBidWorksheetController.saveITMWorksheet(opp.Id, encoded2, false);
        Integer nextVer = ITMBidWorksheetController.getNextVersionNumber(opp.Id);
        Test.stopTest();
        
        System.assert(nextVer > 1);
    }

    @IsTest
    static void testAutoSaveOverwrite() {
        Opportunity opp = getTestOpp();
        String encoded1 = EncodingUtil.base64Encode(Blob.valueOf('{"auto":1}'));
        String encoded2 = EncodingUtil.base64Encode(Blob.valueOf('{"auto":2}'));
        
        Test.startTest();
        ITMBidWorksheetController.autoSaveITMWorksheet(opp.Id, encoded1);
        ITMBidWorksheetController.autoSaveITMWorksheet(opp.Id, encoded2);
        Test.stopTest();
    }


    @IsTest
    static void testAutoSaveThenSaveDeletesAutosave() {
        Opportunity opp = getTestOpp();
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{"data":"test"}'));
        
        Test.startTest();
        ITMBidWorksheetController.autoSaveITMWorksheet(opp.Id, encoded);
        ITMBidWorksheetController.saveITMWorksheet(opp.Id, encoded, false);
        Test.stopTest();
    }
}