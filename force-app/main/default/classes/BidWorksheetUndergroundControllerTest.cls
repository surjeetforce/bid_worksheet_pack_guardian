@IsTest
private class BidWorksheetUndergroundControllerTest {
    
    @TestSetup
    static void setupData() {
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
    }

    private static Opportunity getTestOpp() {
        return [SELECT Id FROM Opportunity WHERE Name = 'Test Opp' LIMIT 1];
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
        Opportunity opp = getTestOpp();
        String json = '{"test":"data"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveSheet(opp.Id, encoded, false);
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
        Opportunity opp = getTestOpp();
        String json = '{"estimate":123}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        Id cvId = BidWorksheetUndergroundController.saveEstimateSheet(opp.Id, encoded, false);
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
    static void testEstimateSheetLoadDirect() {
        Opportunity opp = getTestOpp();
        String json = '{"estimate":"direct"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveEstimateSheet(opp.Id, encoded, false);
        String loadedBase64 = BidWorksheetUndergroundController.loadEstimateSheet(opp.Id);
        Test.stopTest();
        
        System.assertNotEquals(null, loadedBase64);
        String loadedJson = EncodingUtil.base64Decode(loadedBase64).toString();
        System.assertEquals(json, loadedJson);
    }

    @IsTest
    static void testEstimateSheetLoadNoData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        String result = BidWorksheetUndergroundController.loadEstimateSheet(opp.Id);
        Test.stopTest();
        
        System.assertEquals(null, result);
    }

    @IsTest
    static void testSOVSaveLoad() {
        Opportunity opp = getTestOpp();
        String json = '{"sov":"test"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        Id cvId = BidWorksheetUndergroundController.saveSOVSheet(opp.Id, encoded, false);
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
        Opportunity opp = getTestOpp();
        String json = '{"design":"ok"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        Id cvId = BidWorksheetUndergroundController.saveDesignWorksheet(opp.Id, encoded, false);
        BidWorksheetUndergroundController.autoSaveDesignWorksheet(opp.Id, encoded);
        BidWorksheetUndergroundController.loadDesignWorksheet(opp.Id);
        BidWorksheetUndergroundController.loadLatestDesignWorksheet(opp.Id);
        BidWorksheetUndergroundController.getNextVersionNumber_Design(opp.Id);
        BidWorksheetUndergroundController.getVersionHistory_Design(opp.Id);
        BidWorksheetUndergroundController.loadVersionById_Design(cvId);
        Test.stopTest();
    }

    @IsTest
    static void testDesignLoadNoData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        String result = BidWorksheetUndergroundController.loadDesignWorksheet(opp.Id);
        Test.stopTest();
        
        System.assertEquals(null, result);
    }

    @IsTest
    static void testSOVLoadNoData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        String result = BidWorksheetUndergroundController.loadSOVSheet(opp.Id);
        Test.stopTest();
        
        System.assertEquals(null, result);
    }

    @IsTest
    static void testLoadLatestSheetNoData() {
        Opportunity opp = getTestOpp();

        Test.startTest();
        String result = BidWorksheetUndergroundController.loadLatestSheet(opp.Id);
        Test.stopTest();

        System.assertEquals(null, result);
    }

    @IsTest
    static void testFinalVersionStatus() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        Map<String, Object> ugStatus = BidWorksheetUndergroundController.getFinalVersionStatus(opp.Id);
        Map<String, Object> estStatus = BidWorksheetUndergroundController.getFinalVersionStatus_Estimate(opp.Id);
        Map<String, Object> sovStatus = BidWorksheetUndergroundController.getFinalVersionStatus_SOV(opp.Id);
        Map<String, Object> designStatus = BidWorksheetUndergroundController.getFinalVersionStatus_Design(opp.Id);
        Test.stopTest();
        
        System.assertEquals(false, ugStatus.get('hasFinalVersion'));
        System.assertEquals(false, estStatus.get('hasFinalVersion'));
        System.assertEquals(false, sovStatus.get('hasFinalVersion'));
        System.assertEquals(false, designStatus.get('hasFinalVersion'));
    }

    @IsTest
    static void testFinalVersionStatusWithNullId() {
        Test.startTest();
        Map<String, Object> result = BidWorksheetUndergroundController.getFinalVersionStatus(null);
        Test.stopTest();
        
        System.assertEquals(false, result.get('hasFinalVersion'));
    }

    @IsTest
    static void testGetNextVersionNumberNullId() {
        Test.startTest();
        Integer nextVer = BidWorksheetUndergroundController.getNextVersionNumber(null);
        Test.stopTest();
        
        System.assertEquals(1, nextVer);
    }

    @IsTest
    static void testGetVersionHistoryNullId() {
        Test.startTest();
        List<BidWorksheetUndergroundController.VersionHistoryWrapper> history = BidWorksheetUndergroundController.getVersionHistory(null);
        Test.stopTest();
        
        System.assertEquals(0, history.size());
    }

    @IsTest
    static void testLoadVersionByIdDirect() {
        Opportunity opp = getTestOpp();
        String json = '{"version":"test"}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        Id cvId = BidWorksheetUndergroundController.saveSheet(opp.Id, encoded, false);
        String loadedBase64 = BidWorksheetUndergroundController.loadVersionById(cvId);
        Test.stopTest();
        
        System.assertNotEquals(null, loadedBase64);
    }

    @IsTest
    static void testSaveSheetNullId() {
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{}'));
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveSheet(null, encoded, false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveSheetBlankData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveSheet(opp.Id, '', false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveEstimateSheetNullId() {
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{}'));
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveEstimateSheet(null, encoded, false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveEstimateSheetBlankData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveEstimateSheet(opp.Id, '', false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveSOVSheetNullId() {
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{}'));
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveSOVSheet(null, encoded, false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveSOVSheetBlankData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveSOVSheet(opp.Id, '', false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveDesignNullId() {
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{}'));
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveDesignWorksheet(null, encoded, false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testSaveDesignBlankData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.saveDesignWorksheet(opp.Id, '', false);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testAutoSaveSheetNullId() {
        String encoded = EncodingUtil.base64Encode(Blob.valueOf('{}'));
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.autoSaveSheet(null, encoded);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testAutoSaveSheetBlankData() {
        Opportunity opp = getTestOpp();
        
        Test.startTest();
        try {
            BidWorksheetUndergroundController.autoSaveSheet(opp.Id, '');
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testLoadEstimateSheetNullId() {
        Test.startTest();
        try {
            BidWorksheetUndergroundController.loadEstimateSheet(null);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testLoadLatestSheetNullId() {
        Test.startTest();
        try {
            BidWorksheetUndergroundController.loadLatestSheet(null);
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
            BidWorksheetUndergroundController.loadVersionById(null);
            System.assert(false, 'Should have thrown exception');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }

    @IsTest
    static void testUpdateOpportunityFieldsExtended() {
        Opportunity opp = getTestOpp();
        Map<String, Object> fieldMap = new Map<String, Object>{
            'Amount' => 5000,
            'Description' => 'Test Desc',
            'CloseDate' => String.valueOf(Date.today()),
            'Probability' => 50.5
        };
        
        Test.startTest();
        BidWorksheetUndergroundController.updateOpportunityFields(opp.Id, JSON.serialize(fieldMap));
        
        // Test error handling - null Id
        try {
            BidWorksheetUndergroundController.updateOpportunityFields(null, '{}');
        } catch (Exception e) {}

        // Test error handling - blank data
        try {
            BidWorksheetUndergroundController.updateOpportunityFields(opp.Id, '');
        } catch (Exception e) {}
        Test.stopTest();
        
        Opportunity updated = [SELECT Amount FROM Opportunity WHERE Id = :opp.Id];
        System.assertEquals(5000, updated.Amount);
    }

    @IsTest
    static void testSaveSheetSecondVersion() {
        Opportunity opp = getTestOpp();
        String json1 = '{"version":1}';
        String json2 = '{"version":2}';
        String encoded1 = EncodingUtil.base64Encode(Blob.valueOf(json1));
        String encoded2 = EncodingUtil.base64Encode(Blob.valueOf(json2));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveSheet(opp.Id, encoded1, false);
        BidWorksheetUndergroundController.saveSheet(opp.Id, encoded2, false);
        
        Integer nextVer = BidWorksheetUndergroundController.getNextVersionNumber(opp.Id);
        Test.stopTest();
        
        System.assert(nextVer > 1);
    }

    @IsTest
    static void testAutoSaveOverwrite() {
        Opportunity opp = getTestOpp();
        String json1 = '{"autosave":1}';
        String json2 = '{"autosave":2}';
        String encoded1 = EncodingUtil.base64Encode(Blob.valueOf(json1));
        String encoded2 = EncodingUtil.base64Encode(Blob.valueOf(json2));
        
        Test.startTest();
        BidWorksheetUndergroundController.autoSaveSheet(opp.Id, encoded1);
        BidWorksheetUndergroundController.autoSaveSheet(opp.Id, encoded2);
        Test.stopTest();
    }

    @IsTest
    static void testSaveEstimateSecondVersion() {
        Opportunity opp = getTestOpp();
        String encoded1 = EncodingUtil.base64Encode(Blob.valueOf('{"v":1}'));
        String encoded2 = EncodingUtil.base64Encode(Blob.valueOf('{"v":2}'));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveEstimateSheet(opp.Id, encoded1, false);
        BidWorksheetUndergroundController.saveEstimateSheet(opp.Id, encoded2, false);
        Test.stopTest();
    }

    @IsTest
    static void testSaveSOVSecondVersion() {
        Opportunity opp = getTestOpp();
        String encoded1 = EncodingUtil.base64Encode(Blob.valueOf('{"v":1}'));
        String encoded2 = EncodingUtil.base64Encode(Blob.valueOf('{"v":2}'));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveSOVSheet(opp.Id, encoded1, false);
        BidWorksheetUndergroundController.saveSOVSheet(opp.Id, encoded2, false);
        Test.stopTest();
    }

    @IsTest
    static void testSaveDesignSecondVersion() {
        Opportunity opp = getTestOpp();
        String encoded1 = EncodingUtil.base64Encode(Blob.valueOf('{"v":1}'));
        String encoded2 = EncodingUtil.base64Encode(Blob.valueOf('{"v":2}'));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveDesignWorksheet(opp.Id, encoded1, false);
        BidWorksheetUndergroundController.saveDesignWorksheet(opp.Id, encoded2, false);
        Test.stopTest();
    }

    @IsTest
    static void testMarkAsFinalAndBlock() {
        Opportunity opp = getTestOpp();
        String json = '{"isFinal":true}';
        String encoded = EncodingUtil.base64Encode(Blob.valueOf(json));
        
        Test.startTest();
        BidWorksheetUndergroundController.saveSheet(opp.Id, encoded, true);
        
        Map<String, Object> status = BidWorksheetUndergroundController.getFinalVersionStatus(opp.Id);
        System.assertEquals(true, status.get('hasFinalVersion'));
        
        try {
            BidWorksheetUndergroundController.saveSheet(opp.Id, encoded, false);
            System.assert(false, 'Should have thrown exception for final version');
        } catch (AuraHandledException e) {
            System.debug('Exception from Apex');
        }
        Test.stopTest();
    }
}