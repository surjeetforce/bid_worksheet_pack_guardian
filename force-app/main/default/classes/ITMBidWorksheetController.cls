public without sharing class ITMBidWorksheetController {
    
    private static final String ITM_SHEET_TITLE = 'ITM_Bid_Worksheet';
    private static final String ITM_AUTO_SAVE_TITLE = 'ITM_Bid_Worksheet_AutoSave';
    
    // ========================================
    // GENERIC HELPER METHODS (Reusable Pattern)
    // ========================================
    
    /**
     * Generic helper to get next version number for ITM worksheet
     */
    private static Integer getNextVersionNumberGeneric(Id opportunityId, String sheetTitle){
        if (opportunityId == null){
            return 1;
        }

        try{
            List<ContentDocumentLink> links = [SELECT ContentDocumentId
                                               FROM ContentDocumentLink
                                               WHERE LinkedEntityId = :opportunityId 
                                               AND ContentDocument.Title = :sheetTitle];
            
            if (links.isEmpty()){
                return 1;
            }

            Set<Id> docIds = new Set<Id>();
            for (ContentDocumentLink l : links){
                docIds.add(l.ContentDocumentId);
            }

            Integer versionCount = [SELECT COUNT() 
                                    FROM ContentVersion 
                                    WHERE ContentDocumentId IN:docIds];
            
            return versionCount + 1;
        } catch (Exception e){
            System.debug('Error getting next version number for ' + sheetTitle + ': ' + e.getMessage());
            return 1;
        }
    }

    /**
     * Generic helper to get version history for ITM worksheet
     */
    private static List<VersionHistoryWrapper> getVersionHistoryGeneric(Id opportunityId, String sheetTitle){
        List<VersionHistoryWrapper> versionList = new List<VersionHistoryWrapper>();
        
        if (opportunityId == null){
            return versionList;
        }

        try{
            List<ContentDocumentLink> links = [SELECT ContentDocumentId
                                               FROM ContentDocumentLink
                                               WHERE LinkedEntityId = :opportunityId 
                                               AND ContentDocument.Title = :sheetTitle];
            
            if (links.isEmpty()){
                return versionList;
            }

            Set<Id> docIds = new Set<Id>();
            for (ContentDocumentLink l : links){
                docIds.add(l.ContentDocumentId);
            }

            List<ContentVersion> versions = [SELECT Id, VersionNumber, CreatedDate, CreatedBy.Name, ContentSize
                                             FROM ContentVersion
                                             WHERE ContentDocumentId IN:docIds
                                             ORDER BY CreatedDate DESC];

            for (ContentVersion cv : versions){
                VersionHistoryWrapper wrapper = new VersionHistoryWrapper();
                wrapper.versionId = cv.Id;
                wrapper.versionNumber = Integer.valueOf(cv.VersionNumber);
                wrapper.createdDate = cv.CreatedDate;
                wrapper.createdBy = cv.CreatedBy.Name;
                wrapper.fileSize = cv.ContentSize;
                versionList.add(wrapper);
            }
        } catch (Exception e){
            System.debug('Error getting version history for ' + sheetTitle + ': ' + e.getMessage());
        }

        return versionList;
    }

    /**
     * Generic helper to load latest ITM worksheet (checks both version control and autosave)
     */
    private static String loadLatestSheetGeneric(Id opportunityId, String sheetTitle, String autoSaveTitle){
        if (opportunityId == null){
            throw new AuraHandledException('Opportunity Id is required.');
        }

        ContentVersion versionControlLatest = null;
        Attachment autoSaveAttachment = null;
        DateTime versionControlTime = null;
        DateTime autoSaveTime = null;

        // Check version control document
        try{
            List<ContentDocumentLink> versionControlLinks = [SELECT ContentDocumentId
                                                             FROM ContentDocumentLink
                                                             WHERE LinkedEntityId = :opportunityId 
                                                             AND ContentDocument.Title = :sheetTitle];
            
            if (!versionControlLinks.isEmpty()){
                Set<Id> docIds = new Set<Id>();
                for (ContentDocumentLink l : versionControlLinks){
                    docIds.add(l.ContentDocumentId);
                }
                
                List<ContentVersion> versions = [SELECT VersionData, CreatedDate, LastModifiedDate
                                                 FROM ContentVersion
                                                 WHERE ContentDocumentId IN:docIds
                                                 ORDER BY CreatedDate DESC
                                                 LIMIT 1];
                if (!versions.isEmpty()){
                    versionControlLatest = versions[0];
                    versionControlTime = versionControlLatest.LastModifiedDate != null 
                                        ? versionControlLatest.LastModifiedDate 
                                        : versionControlLatest.CreatedDate;
                }
            }
        } catch (Exception e){
            System.debug('Error loading version control document for ' + sheetTitle + ': ' + e.getMessage());
        }

        // Check auto-save attachment
        try{
            List<Attachment> autoSaveAttachments = [SELECT Body, LastModifiedDate
                                                     FROM Attachment
                                                     WHERE ParentId = :opportunityId 
                                                     AND Name = :autoSaveTitle
                                                     LIMIT 1];
            
            if (!autoSaveAttachments.isEmpty()){
                autoSaveAttachment = autoSaveAttachments[0];
                autoSaveTime = autoSaveAttachment.LastModifiedDate;
            }
        } catch (Exception e){
            System.debug('Error loading auto-save attachment for ' + autoSaveTitle + ': ' + e.getMessage());
        }

        // Return whichever is most recent
        Blob dataToReturn = null;
        if (versionControlLatest != null && autoSaveAttachment != null){
            // Compare timestamps and return the most recent
            if (versionControlTime > autoSaveTime){
                dataToReturn = versionControlLatest.VersionData;
            } else {
                dataToReturn = autoSaveAttachment.Body;
            }
        } else if (versionControlLatest != null){
            dataToReturn = versionControlLatest.VersionData;
        } else if (autoSaveAttachment != null){
            dataToReturn = autoSaveAttachment.Body;
        }

        if (dataToReturn == null){
            return null;
        }

        return EncodingUtil.base64Encode(dataToReturn);
    }

    /**
     * Generic helper to autosave ITM worksheet data using Attachment
     */
    private static Id autoSaveSheetGeneric(Id opportunityId, String base64Data, String autoSaveTitle){
        if (opportunityId == null){
            throw new AuraHandledException('Opportunity Id is required.');
        }
        if (String.isBlank(base64Data)){
            throw new AuraHandledException('No sheet data provided.');
        }

        // Decode the incoming JSON payload
        Blob body;
        try{
            body = EncodingUtil.base64Decode(base64Data);
        } catch (Exception e){
            throw new AuraHandledException('Failed to decode sheet data: ' + e.getMessage());
        }

        // Find existing auto-save attachment
        Attachment existing;
        try{
            existing = [SELECT Id, Body 
                        FROM Attachment 
                        WHERE ParentId = :opportunityId 
                        AND Name = :autoSaveTitle 
                        LIMIT 1];
        } catch (QueryException qe){
            existing = null;
        }

        Attachment att;
        if (existing != null){
            // UPDATE existing attachment (prevents version accumulation)
            existing.Body = body;
            update existing;
            return existing.Id;
        } else{
            // Create the first auto-save attachment
            att = new Attachment(
                ParentId = opportunityId,
                Name = autoSaveTitle,
                Body = body,
                ContentType = 'application/json'
            );
            insert att;
            return att.Id;
        }
    }

    /**
     * Generic helper to delete autosave document/attachment
     */
    private static void deleteAutoSaveDocumentGeneric(Id opportunityId, String autoSaveTitle){
        try{
            // Delete auto-save attachment
            List<Attachment> attachmentsToDelete = [SELECT Id 
                                                     FROM Attachment 
                                                     WHERE ParentId = :opportunityId 
                                                     AND Name = :autoSaveTitle];
            
            if (!attachmentsToDelete.isEmpty()){
                delete attachmentsToDelete;
                System.debug('✅ Deleted auto-save attachment: ' + autoSaveTitle);
            }
            
            // Also delete any old ContentDocument-based autosave (for migration/cleanup)
            List<ContentDocumentLink> links = [SELECT ContentDocumentId
                                               FROM ContentDocumentLink
                                               WHERE LinkedEntityId = :opportunityId 
                                               AND ContentDocument.Title = :autoSaveTitle];
            
            if (!links.isEmpty()){
                Set<Id> docIds = new Set<Id>();
                for (ContentDocumentLink l : links){
                    docIds.add(l.ContentDocumentId);
                }
                
                List<ContentDocument> docsToDelete = [SELECT Id FROM ContentDocument WHERE Id IN:docIds];
                if (!docsToDelete.isEmpty()){
                    delete docsToDelete;
                    System.debug('✅ Deleted old auto-save ContentDocument: ' + autoSaveTitle);
                }
            }
        } catch (Exception e){
            System.debug('Error deleting auto-save document/attachment ' + autoSaveTitle + ': ' + e.getMessage());
            // Don't throw - deletion failure shouldn't break the save
        }
    }

    // ========================================
    // PUBLIC METHODS (Using Generic Helpers)
    // ========================================
    
    /**
     * Get all ITM Bid Items from custom metadata
     * Organized by section and category
     */
    @AuraEnabled(cacheable=true)
    public static Map<String, Object> getITMItems() {
        try {
            // Query all metadata records
            List<ITM_Bid_Item__mdt> allItems = [
                SELECT Id, MasterLabel, Row_Number__c, Category__c, 
                       Item_Description__c, Default_Hours__c, Section__c, Column__c
                FROM ITM_Bid_Item__mdt
                ORDER BY Row_Number__c ASC
            ];
            
            // Organize by section
            Map<String, List<Map<String, Object>>> sections = new Map<String, List<Map<String, Object>>>();
            sections.put('Labor Factor', new List<Map<String, Object>>());
            sections.put('Equipment Factor', new List<Map<String, Object>>());
            sections.put('Rates', new List<Map<String, Object>>());
            
            for (ITM_Bid_Item__mdt item : allItems) {
                Map<String, Object> itemData = new Map<String, Object>();
                itemData.put('id', item.MasterLabel);
                itemData.put('rowNumber', item.Row_Number__c);
                itemData.put('category', item.Category__c);
                itemData.put('description', item.Item_Description__c);
                itemData.put('defaultHours', item.Default_Hours__c);
                itemData.put('section', item.Section__c);
                itemData.put('column', item.Column__c);
                
                // Add to appropriate section
                if (sections.containsKey(item.Section__c)) {
                    sections.get(item.Section__c).add(itemData);
                }
            }
            
            // Return organized data
            Map<String, Object> result = new Map<String, Object>();
            result.put('laborFactor', sections.get('Labor Factor'));
            result.put('equipmentFactor', sections.get('Equipment Factor'));
            result.put('rates', sections.get('Rates'));
            
            System.debug('ITM Items loaded: ' + allItems.size() + ' records');
            return result;
            
        } catch (Exception e) {
            System.debug('Error loading ITM items: ' + e.getMessage());
            throw new AuraHandledException('Error loading ITM items: ' + e.getMessage());
        }
    }
    
    /**
     * Save ITM worksheet data to Opportunity
     */
    @AuraEnabled
    public static Id saveITMWorksheet(Id opportunityId, String base64Data) {
        if (opportunityId == null) {
            throw new AuraHandledException('Opportunity Id is required.');
        }
        if (String.isBlank(base64Data)) {
            throw new AuraHandledException('No ITM worksheet data provided.');
        }
        
        // Decode the incoming JSON payload
        Blob body;
        try {
            body = EncodingUtil.base64Decode(base64Data);
        } catch (Exception e) {
            throw new AuraHandledException('Failed to decode ITM data: ' + e.getMessage());
        }
        
        // Use consistent title for ITM worksheet
        String itmTitle = 'ITM_Bid_Worksheet';
        String itmFileName = 'ITM_Bid_Worksheet.json';
        
        // Find an existing ContentDocument linked to this Opportunity with the ITM title
        ContentDocumentLink existingLink;
        try {
            existingLink = [
                SELECT ContentDocumentId, ContentDocument.Title
                FROM ContentDocumentLink
                WHERE LinkedEntityId = :opportunityId 
                AND ContentDocument.Title = :itmTitle
                LIMIT 1
            ];
        } catch (QueryException qe) {
            existingLink = null;
        }
        
        ContentVersion cv;
        if (existingLink != null) {
            // Add a new version to the existing document
            cv = new ContentVersion(
                ContentDocumentId = existingLink.ContentDocumentId,
                Title = itmTitle,
                PathOnClient = itmFileName,
                VersionData = body
            );
            insert cv;
            System.debug('Updated existing ITM worksheet document');
        } else {
            // Create the first version and publish directly to the Opportunity
            cv = new ContentVersion(
                FirstPublishLocationId = opportunityId,
                Title = itmTitle,
                PathOnClient = itmFileName,
                VersionData = body
            );
            insert cv;
            System.debug('Created new ITM worksheet document');
        }
        
        // Delete auto-save attachment after save
        deleteAutoSaveDocumentGeneric(opportunityId, ITM_AUTO_SAVE_TITLE);
        
        return cv.Id;
    }
    
    /**
     * Load saved ITM worksheet data from Opportunity
     */
    @AuraEnabled
    public static String loadITMWorksheet(Id opportunityId) {
        try {
            if (opportunityId == null) {
                throw new AuraHandledException('Opportunity Id is required.');
            }
            
            String itmTitle = 'ITM_Bid_Worksheet';
            
            // First find ContentDocumentIds linked to this Opportunity with the desired title
            List<ContentDocumentLink> links = [
                SELECT ContentDocumentId
                FROM ContentDocumentLink
                WHERE LinkedEntityId = :opportunityId 
                AND ContentDocument.Title = :itmTitle
            ];
            
            if (links.isEmpty()) {
                System.debug('No ITM worksheet found for Opportunity: ' + opportunityId);
                return null;
            }
            
            Set<Id> docIds = new Set<Id>();
            for (ContentDocumentLink l : links) {
                docIds.add(l.ContentDocumentId);
            }
            
            // Get the latest version of the document
            ContentVersion latest;
            try {
                latest = [
                    SELECT VersionData
                    FROM ContentVersion
                    WHERE ContentDocumentId IN :docIds
                    ORDER BY CreatedDate DESC
                    LIMIT 1
                ];
            } catch (QueryException qe) {
                return null;
            }
            
            if (latest == null || latest.VersionData == null) {
                return null;
            }
            
            // Return as plain JSON string
            String jsonData = latest.VersionData.toString();
            System.debug('Loaded ITM worksheet data for Opportunity: ' + opportunityId);
            return jsonData;
            
        } catch (Exception e) {
            System.debug('Error loading ITM worksheet: ' + e.getMessage());
            System.debug('Stack trace: ' + e.getStackTraceString());
            return null;
        }
    }
    
    // ========================================
    // PUBLIC WRAPPER METHODS FOR ITM
    // ========================================

    /**
     * Get next version number for ITM worksheet
     */
    @AuraEnabled(cacheable = false)
    public static Integer getNextVersionNumber(Id opportunityId){
        return getNextVersionNumberGeneric(opportunityId, ITM_SHEET_TITLE);
    }

    /**
     * Get version history for ITM worksheet
     */
    @AuraEnabled(cacheable = false)
    public static List<VersionHistoryWrapper> getVersionHistory(Id opportunityId){
        return getVersionHistoryGeneric(opportunityId, ITM_SHEET_TITLE);
    }

    /**
     * Load latest ITM worksheet (checks autosave + version control)
     */
    @AuraEnabled(cacheable = false)
    public static String loadLatestITMWorksheet(Id opportunityId){
        return loadLatestSheetGeneric(opportunityId, ITM_SHEET_TITLE, ITM_AUTO_SAVE_TITLE);
    }

    /**
     * Load specific ITM version by ID
     */
    @AuraEnabled(cacheable = false)
    public static String loadVersionById(Id versionId){
        if (versionId == null){
            throw new AuraHandledException('Version Id is required.');
        }

        try{
            ContentVersion cv = [SELECT VersionData
                                 FROM ContentVersion
                                 WHERE Id = :versionId
                                 LIMIT 1];

            if (cv == null || cv.VersionData == null){
                return null;
            }

            return EncodingUtil.base64Encode(cv.VersionData);
        } catch (Exception e){
            System.debug('Error loading version by ID: ' + e.getMessage());
            throw new AuraHandledException('Failed to load version: ' + e.getMessage());
        }
    }

    /**
     * Autosave ITM worksheet
     */
    @AuraEnabled
    public static Id autoSaveITMWorksheet(Id opportunityId, String base64Data){
        return autoSaveSheetGeneric(opportunityId, base64Data, ITM_AUTO_SAVE_TITLE);
    }

    // ========================================
    // WRAPPER CLASSES
    // ========================================

    /**
     * Wrapper class for metadata items
     */
    public class ITMItemWrapper {
        @AuraEnabled public String id;
        @AuraEnabled public Decimal rowNumber;
        @AuraEnabled public String category;
        @AuraEnabled public String description;
        @AuraEnabled public Decimal defaultHours;
        @AuraEnabled public String section;
        @AuraEnabled public String column;
    }

    /**
     * Wrapper class for version history
     */
    public class VersionHistoryWrapper{
        @AuraEnabled
        public String versionId { get; set; }
        
        @AuraEnabled
        public Integer versionNumber { get; set; }
        
        @AuraEnabled
        public DateTime createdDate { get; set; }
        
        @AuraEnabled
        public String createdBy { get; set; }
        
        @AuraEnabled
        public Long fileSize { get; set; }
    }
}