public without sharing class BidWorksheetUndergroundController{
    private static final String SHEET_TITLE = 'UndergroundBid';
    private static final String FILE_NAME = 'UndergroundBid.json';
    private static final String AUTO_SAVE_TITLE = 'UndergroundBid_AutoSave';
    private static final String AUTO_SAVE_FILE_NAME = 'UndergroundBid_AutoSave.json';
    /**
     * Get all items for Sheet #1 (Materials)
     * Cacheable for better performance
     */
    @AuraEnabled(cacheable = true)
    public static List<Sheet1ItemWrapper> getSheet1Items(){
        List<Sheet1ItemWrapper> items = new List<Sheet1ItemWrapper>();

        try{
            List<Underground_Estimate_Item_mdt__mdt> metadataRecords = [SELECT Row_Number__c, Column_Side__c, Description__c, Size__c, Is_Indent__c, Default_Unit_Price__c, Unit_Price_Field_Type__c, Gross_Field_Type__c
                                                                        FROM Underground_Estimate_Item_mdt__mdt
                                                                        WHERE Sheet_Number__c = 1 AND Section__c = 1
                                                                        ORDER BY Row_Number__c, Column_Side__c];

            System.debug('Found ' + metadataRecords.size() + ' Sheet 1 metadata records');

            Map<Integer, Sheet1ItemWrapper> rowMap = new Map<Integer, Sheet1ItemWrapper>();

            for (Underground_Estimate_Item_mdt__mdt record : metadataRecords){
                Integer rowNum = Integer.valueOf(record.Row_Number__c);

                if (!rowMap.containsKey(rowNum)){
                    Sheet1ItemWrapper wrapper = new Sheet1ItemWrapper();
                    wrapper.excelRow = rowNum;  // ⭐ SET EXCEL ROW NUMBER
                    rowMap.put(rowNum, wrapper);
                }

                Sheet1ItemWrapper item = rowMap.get(rowNum);

                if (record.Column_Side__c == 'Left'){
                    item.left = new ColumnData(record.Description__c, record.Size__c, record.Is_Indent__c, record.Default_Unit_Price__c, record.Unit_Price_Field_Type__c, record.Gross_Field_Type__c);
                } else if (record.Column_Side__c == 'Right'){
                    item.right = new ColumnData(record.Description__c, record.Size__c, record.Is_Indent__c, record.Default_Unit_Price__c, record.Unit_Price_Field_Type__c, record.Gross_Field_Type__c);
                }
            }

            List<Integer> sortedKeys = new List<Integer>(rowMap.keySet());
            sortedKeys.sort ();

            for (Integer key : sortedKeys){
                items.add(rowMap.get(key));
            }

            System.debug('Returning ' + items.size() + ' Sheet 1 rows');

        } catch (Exception e){
            System.debug('Error in getSheet1Items: ' + e.getMessage());
            throw new AuraHandledException('Error loading Sheet 1 data: ' + e.getMessage());
        }

        return items;
    }

    /**
     * Get all items for Sheet #2 (Labor & Costs)
     * Cacheable for better performance
     */
    @AuraEnabled(cacheable = true)
    public static List<Sheet2ItemWrapper> getSheet2Items(){
        List<Sheet2ItemWrapper> items = new List<Sheet2ItemWrapper>();

        try{
            List<Underground_Estimate_Item_mdt__mdt> metadataRecords = [SELECT Row_Number__c, Column_Side__c, Description__c, Is_Read_Only__c, Default_Amount__c, Default_Unit_Price__c, Is_Comment_Row__c, Is_Total_Row__c, Unit_Price_Field_Type__c, Gross_Field_Type__c
                                                                        FROM Underground_Estimate_Item_mdt__mdt
                                                                        WHERE Sheet_Number__c = 1 AND Section__c = 2
                                                                        ORDER BY Row_Number__c, Column_Side__c];

            System.debug('Found ' + metadataRecords.size() + ' Sheet 2 metadata records');

            Map<Integer, Sheet2ItemWrapper> rowMap = new Map<Integer, Sheet2ItemWrapper>();

            for (Underground_Estimate_Item_mdt__mdt record : metadataRecords){
                Integer rowNum = Integer.valueOf(record.Row_Number__c);

                if (!rowMap.containsKey(rowNum)){
                    Sheet2ItemWrapper wrapper = new Sheet2ItemWrapper();
                    wrapper.excelRow = rowNum;  // ⭐ SET EXCEL ROW NUMBER
                    rowMap.put(rowNum, wrapper);
                }

                Sheet2ItemWrapper item = rowMap.get(rowNum);

                if (record.Column_Side__c == 'Left'){
                    item.left = new Sheet2ColumnData(record.Description__c, record.Is_Read_Only__c, record.Default_Amount__c, record.Default_Unit_Price__c, record.Is_Total_Row__c, false, record.Unit_Price_Field_Type__c, record.Gross_Field_Type__c  // isCommentRow for left is always false
                    );
                } else if (record.Column_Side__c == 'Right'){
                    item.right = new Sheet2ColumnData(record.Description__c, record.Is_Read_Only__c, record.Default_Amount__c, record.Default_Unit_Price__c, false, record.Is_Comment_Row__c, record.Unit_Price_Field_Type__c, record.Gross_Field_Type__c);    // isTotalRow for right
                }
            }

            List<Integer> sortedKeys = new List<Integer>(rowMap.keySet());
            sortedKeys.sort ();

            for (Integer key : sortedKeys){
                items.add(rowMap.get(key));
            }

            System.debug('Returning ' + items.size() + ' Sheet 2 rows');

        } catch (Exception e){
            System.debug('Error in getSheet2Items: ' + e.getMessage());
            throw new AuraHandledException('Error loading Sheet 2 data: ' + e.getMessage());
        }

        return items;
    }

    @AuraEnabled(cacheable = true)
    public static List<EstimateSheetItemWrapper> getEstimateSheetItems(){
        List<EstimateSheetItemWrapper> items = new List<EstimateSheetItemWrapper>();

        try{
            // Query with Section__c field
            List<Underground_Estimate_Item_mdt__mdt> metadataRecords = [SELECT Row_Number__c, Column_Side__c, Section__c, Size__c, Description__c, Default_Amount__c, Default_Unit_Price__c, Is_Read_Only__c, Is_Comment_Row__c, Is_Total_Row__c, Unit_Price_Field_Type__c, Gross_Field_Type__c
                                                                        FROM Underground_Estimate_Item_mdt__mdt
                                                                        WHERE Sheet_Number__c = 2
                                                                        ORDER BY Section__c ASC, Row_Number__c ASC, Column_Side__c ASC];

            System.debug('Found ' + metadataRecords.size() + ' Estimate Sheet metadata records');

            // Group by row number - each row has Left and Right
            Map<String, EstimateSheetItemWrapper> rowMap = new Map<String, EstimateSheetItemWrapper>();

            for (Underground_Estimate_Item_mdt__mdt record : metadataRecords){
                Integer rowNum = Integer.valueOf(record.Row_Number__c);
                Integer section = record.Section__c != null ? Integer.valueOf(record.Section__c) : 1;

                // Create unique key: section + row number
                String key = section + '_' + rowNum;

                // Initialize wrapper if not exists
                if (!rowMap.containsKey(key)){
                    EstimateSheetItemWrapper wrapper = new EstimateSheetItemWrapper();
                    wrapper.section = section;
                    wrapper.excelRow = rowNum;  // ⭐ ADDED
                    rowMap.put(key, wrapper);
                }

                EstimateSheetItemWrapper item = rowMap.get(key);

                // Populate left or right
                if (record.Column_Side__c == 'Left'){
                    item.left = new EstimateColumnData(record.Description__c, record.Size__c, record.Default_Amount__c, record.Default_Unit_Price__c, record.Is_Read_Only__c, record.Is_Comment_Row__c, record.Is_Total_Row__c, record.Unit_Price_Field_Type__c, record.Gross_Field_Type__c);
                } else if (record.Column_Side__c == 'Right'){
                    item.right = new EstimateColumnData(record.Description__c, record.Size__c, record.Default_Amount__c, record.Default_Unit_Price__c, record.Is_Read_Only__c, record.Is_Comment_Row__c, record.Is_Total_Row__c, record.Unit_Price_Field_Type__c, record.Gross_Field_Type__c);
                }

                rowMap.put(key, item);
            }

            // Convert map to list (already sorted by query ORDER BY)
            items = rowMap.values();

            System.debug('Returning ' + items.size() + ' Estimate Sheet rows');

        } catch (Exception e){
            System.debug('Error in getEstimateSheetItems: ' + e.getMessage());
            System.debug('Stack trace: ' + e.getStackTraceString());
            throw new AuraHandledException('Error loading Estimate Sheet data: ' + e.getMessage());
        }

        return items;
    }

    /**
     * Save Design Worksheet data
     */
    @AuraEnabled
    public static Id saveDesignWorksheet(Id opportunityId, String base64Data){
        if (opportunityId == null){
            throw new AuraHandledException('Opportunity Id is required.');
        }
        if (String.isBlank(base64Data)){
            throw new AuraHandledException('No design worksheet data provided.');
        }

        Blob body;
        try{
            body = EncodingUtil.base64Decode(base64Data);
        } catch (Exception e){
            throw new AuraHandledException('Failed to decode data: ' + e.getMessage());
        }

        String designTitle = 'DesignWorksheet';
        String designFileName = 'DesignWorksheet.json';

        ContentDocumentLink existingLink;
        try{
            existingLink = [SELECT ContentDocumentId, ContentDocument.Title
                            FROM ContentDocumentLink
                            WHERE LinkedEntityId = :opportunityId AND ContentDocument.Title = :designTitle
                            LIMIT 1];
        } catch (QueryException qe){
            existingLink = null;
        }

        ContentVersion cv;
        if (existingLink != null){
            cv = new ContentVersion(
                ContentDocumentId = existingLink.ContentDocumentId, 
                Title = designTitle, 
                PathOnClient = designFileName, 
                VersionData = body
            );
            insert cv;
        } else{
            cv = new ContentVersion(
                FirstPublishLocationId = opportunityId, 
                Title = designTitle, 
                PathOnClient = designFileName, 
                VersionData = body
            );
            insert cv;
        }

        // Delete auto-save attachment after save
        deleteAutoSaveDocumentGeneric(opportunityId, 'DesignWorksheet_AutoSave');

        return cv.Id;
    }

    /**
     * Load Design Worksheet data
     */
    @AuraEnabled
    public static String loadDesignWorksheet(Id opportunityId){
        try{
            if (opportunityId == null){
                throw new AuraHandledException('Opportunity Id is required.');
            }

            String designTitle = 'DesignWorksheet';

            // First find ContentDocumentIds linked to this Opportunity with the desired title
            List<ContentDocumentLink> links = [SELECT ContentDocumentId
                                               FROM ContentDocumentLink
                                               WHERE LinkedEntityId = :opportunityId AND ContentDocument.Title = :designTitle];

            if (links.isEmpty()){
                System.debug('No Design Worksheet found for Opportunity: ' + opportunityId);
                return null;
            }

            Set<Id> docIds = new Set<Id>();
            for (ContentDocumentLink l : links){
                docIds.add(l.ContentDocumentId);
            }

            // Get the latest version of the document
            ContentVersion latest;
            try{
                latest = [SELECT VersionData
                          FROM ContentVersion
                          WHERE ContentDocumentId IN:docIds
                          ORDER BY CreatedDate DESC
                          LIMIT 1];
            } catch (QueryException qe){
                return null;
            }

            if (latest == null || latest.VersionData == null){
                return null;
            }

            // Return as plain JSON string (not base64)
            String jsonData = latest.VersionData.toString();
            System.debug('✅ Loaded Design Worksheet data for Opportunity: ' + opportunityId);
            return jsonData;

        } catch (Exception e){
            System.debug('❌ Error loading Design Worksheet: ' + e.getMessage());
            System.debug('Stack trace: ' + e.getStackTraceString());
            return null;
        }
    }

    /**
     * Save the serialized sheet as a Salesforce File (ContentVersion) linked to the Opportunity.
     * If the file already exists, a new version is created on the same ContentDocument.
     */
    @AuraEnabled
    public static Id saveSheet(Id opportunityId, String base64Data){
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

        // Find an existing ContentDocument linked to this Opportunity with the stable title
        ContentDocumentLink existingLink;
        try{
            existingLink = [SELECT ContentDocumentId, ContentDocument.Title
                            FROM ContentDocumentLink
                            WHERE LinkedEntityId = :opportunityId AND ContentDocument.Title = :SHEET_TITLE
                            LIMIT 1];
        } catch (QueryException qe){
            existingLink = null;
        }

        ContentVersion cv;
        if (existingLink != null){
            // Add a new version to the existing document
            cv = new ContentVersion(
                ContentDocumentId = existingLink.ContentDocumentId, 
                Title = SHEET_TITLE, 
                PathOnClient = FILE_NAME, 
                VersionData = body
            );
            insert cv;
        } else{
            // Create the first version and publish directly to the Opportunity
            cv = new ContentVersion(
                FirstPublishLocationId = opportunityId, 
                Title = SHEET_TITLE, 
                PathOnClient = FILE_NAME, 
                VersionData = body
            );
            insert cv;
        }
        
        // Delete auto-save document if it exists
        deleteAutoSaveDocument(opportunityId);
        
		return cv.Id;
    }

    @AuraEnabled
    public static Id saveEstimateSheet(Id opportunityId, String base64Data){
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

        // Use consistent title for Estimate Sheet
        String estimateTitle = 'BidWorksheet_Estimate';
        String estimateFileName = 'BidWorksheet_Estimate.json';

        // Find an existing ContentDocument linked to this Opportunity with the estimate title
        ContentDocumentLink existingLink;
        try{
            existingLink = [SELECT ContentDocumentId, ContentDocument.Title
                            FROM ContentDocumentLink
                            WHERE LinkedEntityId = :opportunityId AND ContentDocument.Title = :estimateTitle
                            LIMIT 1];
        } catch (QueryException qe){
            existingLink = null;
        }

        ContentVersion cv;
        if (existingLink != null){
            // Add a new version to the existing document
            cv = new ContentVersion(
                ContentDocumentId = existingLink.ContentDocumentId, 
                Title = estimateTitle, 
                PathOnClient = estimateFileName, 
                VersionData = body
            );
            insert cv;
            System.debug('✅ Updated existing Estimate Sheet document');
        } else{
            // Create the first version and publish directly to the Opportunity
            cv = new ContentVersion(
                FirstPublishLocationId = opportunityId, 
                Title = estimateTitle, 
                PathOnClient = estimateFileName, 
                VersionData = body
            );
            insert cv;
            System.debug('✅ Created new Estimate Sheet document');
        }

        // Delete auto-save attachment after save
        deleteAutoSaveDocumentGeneric(opportunityId, 'BidWorksheet_Estimate_AutoSave');

        return cv.Id;
    }

    /**
     * Load the latest version of the sheet file for the Opportunity.
     * Checks both version control document and auto-save document, returns whichever is most recent.
     * Returns the data as a base64 string (the LWC will decode and parse JSON).
     */
    @AuraEnabled(cacheable = false)
    public static String loadLatestSheet(Id opportunityId){
        // Use generic helper for Underground worksheet
        return loadLatestSheetGeneric(opportunityId, SHEET_TITLE, AUTO_SAVE_TITLE);
    }

    /**
     * Load saved Estimate Sheet data from ContentVersion
     */
    @AuraEnabled
    public static String loadEstimateSheet(Id opportunityId){
        if (opportunityId == null){
            throw new AuraHandledException('Opportunity Id is required.');
        }

        String estimateTitle = 'BidWorksheet_Estimate';
        
        // First find ContentDocumentIds linked to this Opportunity with the desired title
        List<ContentDocumentLink> links = [SELECT ContentDocumentId
                                           FROM ContentDocumentLink
                                           WHERE LinkedEntityId = :opportunityId AND ContentDocument.Title = :estimateTitle];

        if (links.isEmpty()){
            System.debug('No saved Estimate Sheet found for Opportunity: ' + opportunityId);
            return null;
        }

        Set<Id> docIds = new Set<Id>();
        for (ContentDocumentLink l : links){
            docIds.add(l.ContentDocumentId);
        }

        ContentVersion latest;
        try{
            // Order by LastModifiedDate to get the truly latest version
            latest = [SELECT VersionData
                      FROM ContentVersion
                      WHERE ContentDocumentId IN:docIds
                      ORDER BY LastModifiedDate DESC
                      LIMIT 1];
        } catch (QueryException qe){
            System.debug('❌ Error querying ContentVersion: ' + qe.getMessage());
            latest = null;
        }

        if (latest == null || latest.VersionData == null){
            System.debug('No valid Estimate Sheet version found for Opportunity: ' + opportunityId);
            return null;
        }

        // Return as base64 encoded string (consistent with loadLatestSheet)
        String base64Data = EncodingUtil.base64Encode(latest.VersionData);
        System.debug('✅ Loaded latest Estimate Sheet data for Opportunity: ' + opportunityId);
        return base64Data;
    }

    @AuraEnabled
public static void updateOpportunityFields(Id opportunityId, String fieldDataJson){
    if (opportunityId == null){
        throw new AuraHandledException('Opportunity Id is required.');
    }
    if (String.isBlank(fieldDataJson)){
        throw new AuraHandledException('No field data provided.');
    }

    try{
        // Parse the field data
        Map<String, Object> fieldData = (Map<String, Object>)JSON.deserializeUntyped(fieldDataJson);

        System.debug('📊 Updating Opportunity: ' + opportunityId);
        System.debug('📊 Number of fields to update: ' + fieldData.size());

        // Get field metadata for type checking
        Map<String, Schema.SObjectField> fieldMap = Schema.SObjectType.Opportunity.fields.getMap();

        // Build Opportunity record with field values
        Opportunity opp = new Opportunity(Id = opportunityId);

        // Map each field from the JSON to the Opportunity field
        for (String fieldName : fieldData.keySet()){
            Object value = fieldData.get(fieldName);

            if (value == null){
                System.debug('  ⚠️ ' + fieldName + ' = null (skipped)');
                continue;
            }

            // Get field type
            if (!fieldMap.containsKey(fieldName)){
                System.debug('  ⚠️ Field ' + fieldName + ' does not exist');
                continue;
            }

            Schema.DisplayType fieldType = fieldMap.get(fieldName).getDescribe().getType();

            try{
                // Handle different data types based on Salesforce field type
                if (fieldType == Schema.DisplayType.DATE){
                    // Convert string to Date
                    if (value instanceof String){
                        String dateStr = (String) value;
                        if (String.isNotBlank(dateStr)){
                            opp.put(fieldName, Date.valueOf(dateStr));
                            System.debug('  ✅ ' + fieldName + ' (DATE) = ' + dateStr);
                        }
                    }
                } else if (fieldType == Schema.DisplayType.DATETIME){
                    // Convert string to DateTime
                    if (value instanceof String){
                        String dateTimeStr = (String) value;
                        if (String.isNotBlank(dateTimeStr)){
                            opp.put(fieldName, DateTime.valueOf(dateTimeStr));
                            System.debug('  ✅ ' + fieldName + ' (DATETIME) = ' + dateTimeStr);
                        }
                    }
                } else if (fieldType == Schema.DisplayType.DOUBLE || 
                           fieldType == Schema.DisplayType.CURRENCY || 
                           fieldType == Schema.DisplayType.PERCENT){
                    // Handle numeric fields
                    if (value instanceof Decimal){
                        opp.put(fieldName, (Decimal) value);
                        System.debug('  ✅ ' + fieldName + ' (NUMBER) = ' + value);
                    } else if (value instanceof Integer){
                        opp.put(fieldName, Decimal.valueOf((Integer) value));
                        System.debug('  ✅ ' + fieldName + ' (NUMBER) = ' + value);
                    } else if (value instanceof String){
                        try{
                            Decimal decValue = Decimal.valueOf((String) value);
                            opp.put(fieldName, decValue);
                            System.debug('  ✅ ' + fieldName + ' (NUMBER) = ' + decValue);
                        } catch (Exception numEx){
                            System.debug('  ⚠️ Could not convert ' + fieldName + ' to number: ' + value);
                        }
                    }
                } else if (fieldType == Schema.DisplayType.INTEGER){
                    // Handle integer fields
                    if (value instanceof Integer){
                        opp.put(fieldName, (Integer) value);
                        System.debug('  ✅ ' + fieldName + ' (INTEGER) = ' + value);
                    } else if (value instanceof Decimal){
                        opp.put(fieldName, ((Decimal) value).intValue());
                        System.debug('  ✅ ' + fieldName + ' (INTEGER) = ' + value);
                    }
                } else {
                    // For all other types (String, Picklist, etc.)
                    opp.put(fieldName, value);
                    System.debug('  ✅ ' + fieldName + ' = ' + value);
                }
            } catch (Exception fieldEx){
                System.debug('  ❌ Error setting ' + fieldName + ': ' + fieldEx.getMessage());
                // Continue with other fields
            }
        }

        // Update the Opportunity
        update opp;
        System.debug('✅ Successfully updated Opportunity');

    } catch (DmlException dmlEx){
        String errorMsg = 'DML Error: ';
        for (Integer i = 0; i < dmlEx.getNumDml(); i++){
            errorMsg += dmlEx.getDmlMessage(i) + '; ';
        }
        System.debug('❌ ' + errorMsg);
        throw new AuraHandledException(errorMsg);
    } catch (Exception e){
        System.debug('❌ Error: ' + e.getMessage());
        throw new AuraHandledException('Failed to update Opportunity: ' + e.getMessage());
    }
}

    // ========================================
    // WRAPPER CLASSES for Sheet #1
    // ========================================
    public class Sheet1ItemWrapper{
        @AuraEnabled
        public Integer excelRow{ get; set; }

        // ⭐ ADDED
        @AuraEnabled
        public ColumnData left{ get; set; }

        @AuraEnabled
        public ColumnData right{ get; set; }

        public Sheet1ItemWrapper(){
            this.excelRow = null;
            this.left = new ColumnData();
            this.right = new ColumnData();
        }

    }

    public class ColumnData{
        @AuraEnabled
        public String description{ get; set; }

        @AuraEnabled
        public String size{ get; set; }

        @AuraEnabled
        public Boolean isIndent{ get; set; }

        @AuraEnabled
        public Decimal defaultUnitPrice{ get; set; }

        @AuraEnabled
        public String unitPriceFieldType{ get; set; }

        @AuraEnabled
        public String grossFieldType{ get; set; }

        public ColumnData(){
            this.description = '';
            this.size = '';
            this.isIndent = false;
            this.defaultUnitPrice = null;
            this.unitPriceFieldType = 'Currency';
            this.grossFieldType = 'Currency';
        }

        public ColumnData(String description, String sz, Boolean indent, Decimal price, String unitPriceType, String grossType){
            this.description = description != null ? description : '';
            this.size = sz != null ? sz : '';
            this.isIndent = indent != null ? indent : false;
            this.defaultUnitPrice = price;
            this.unitPriceFieldType = unitPriceType != null ? unitPriceType : 'Currency';
            this.grossFieldType = grossType != null ? grossType : 'Currency';
        }

    }

    // ========================================
    // WRAPPER CLASSES for Sheet #2
    // ========================================
    public class Sheet2ItemWrapper{
        @AuraEnabled
        public Integer excelRow{ get; set; }

        // ⭐ ADDED
        @AuraEnabled
        public Sheet2ColumnData left{ get; set; }

        @AuraEnabled
        public Sheet2ColumnData right{ get; set; }

        public Sheet2ItemWrapper(){
            this.excelRow = null;
            this.left = new Sheet2ColumnData();
            this.right = new Sheet2ColumnData();
        }

    }

    public class Sheet2ColumnData{
        @AuraEnabled
        public String description{ get; set; }

        @AuraEnabled
        public Boolean descriptionReadonly{ get; set; }

        @AuraEnabled
        public Decimal amount{ get; set; }

        @AuraEnabled
        public Decimal unitPrice{ get; set; }

        @AuraEnabled
        public Boolean isTotalRow{ get; set; }

        @AuraEnabled
        public Boolean isCommentRow{ get; set; }

        @AuraEnabled
        public String unitPriceFieldType{ get; set; }

        @AuraEnabled
        public String grossFieldType{ get; set; }

        public Sheet2ColumnData(){
            this.description = '';
            this.descriptionReadonly = false;
            this.amount = null;
            this.unitPrice = null;
            this.isTotalRow = false;
            this.isCommentRow = false;
            this.unitPriceFieldType = 'Currency';
            this.grossFieldType = 'Currency';
        }

        public Sheet2ColumnData(String description, Boolean readonly, Decimal amt, Decimal price, Boolean totalRow, Boolean commentRow, String unitPriceType, String grossType){
            this.description = description != null ? description : '';
            this.descriptionReadonly = readonly != null ? readonly : false;
            this.amount = amt;
            this.unitPrice = price;
            this.isTotalRow = totalRow != null ? totalRow : false;
            this.isCommentRow = commentRow != null ? commentRow : false;
            this.unitPriceFieldType = unitPriceType != null ? unitPriceType : 'Currency';
            this.grossFieldType = grossType != null ? grossType : 'Currency';
        }

    }

    // ========================================
    // WRAPPER CLASSES for Estimate Sheet
    // ========================================
    public class EstimateSheetItemWrapper{
        @AuraEnabled
        public Integer section{ get; set; }

        @AuraEnabled
        public Integer excelRow{ get; set; }

        // ⭐ ADDED
        @AuraEnabled
        public EstimateColumnData left{ get; set; }

        @AuraEnabled
        public EstimateColumnData right{ get; set; }

        public EstimateSheetItemWrapper(){
            this.section = 1;
            this.excelRow = null;
            this.left = new EstimateColumnData();
            this.right = new EstimateColumnData();
        }

    }

    public class EstimateColumnData{
        @AuraEnabled
        public String description{ get; set; }

        @AuraEnabled
        public String size{ get; set; }

        @AuraEnabled
        public Decimal defaultAmount{ get; set; }

        @AuraEnabled
        public Decimal defaultUnitPrice{ get; set; }

        @AuraEnabled
        public Boolean isReadonly{ get; set; }

        @AuraEnabled
        public Boolean isCommentRow{ get; set; }

        @AuraEnabled
        public Boolean isTotalRow{ get; set; }

        @AuraEnabled
        public String unitPriceFieldType{ get; set; }

        @AuraEnabled
        public String grossFieldType{ get; set; }

        public EstimateColumnData(){
            this.description = '';
            this.size = '';
            this.defaultAmount = null;
            this.defaultUnitPrice = null;
            this.isReadonly = false;
            this.isCommentRow = false;
            this.isTotalRow = false;
            this.unitPriceFieldType = 'Currency';
            this.grossFieldType = 'Currency';
        }

        public EstimateColumnData(String description, String sz, Decimal amt, Decimal price, Boolean readonly, Boolean commentRow, Boolean totalRow, String unitPriceType, String grossType){
            this.description = description != null ? description : '';
            this.size = sz != null ? sz : '';
            this.defaultAmount = amt;
            this.defaultUnitPrice = price;
            this.isReadonly = readonly != null ? readonly : false;
            this.isCommentRow = commentRow != null ? commentRow : false;
            this.isTotalRow = totalRow != null ? totalRow : false;
            this.unitPriceFieldType = unitPriceType != null ? unitPriceType : 'Currency';
            this.grossFieldType = grossType != null ? grossType : 'Currency';
        }

    }

    // ========================================
    // SCHEDULE OF VALUES METHODS
    // ========================================
    /**
     * Get metadata for Schedule of Values
     * Returns both Summary Rows (Section 1) and Building Rows (Section 2)
     */
    @AuraEnabled(cacheable = true)
    public static SOVMetadataWrapper getSOVItems(){
        SOVMetadataWrapper wrapper = new SOVMetadataWrapper();

        try{
            // Query all SOV metadata (Sheet_Number__c = 4)
            List<Underground_Estimate_Item_mdt__mdt> metadataRecords = [SELECT Row_Number__c, Section__c, Description__c, Default_Amount__c
                                                                        FROM Underground_Estimate_Item_mdt__mdt
                                                                        WHERE Sheet_Number__c = 3
                                                                        ORDER BY Section__c ASC, Row_Number__c ASC];

            System.debug('Found ' + metadataRecords.size() + ' SOV metadata records');

            // Section 1 = Summary Rows
            List<SOVRowData> summaryRows = new List<SOVRowData>();
            // Section 2 = Building Category Rows
            List<SOVRowData> buildingRows = new List<SOVRowData>();

            for (Underground_Estimate_Item_mdt__mdt record : metadataRecords){
                Integer section = record.Section__c != null ? Integer.valueOf(record.Section__c) : 1;

                SOVRowData rowData = new SOVRowData();
                rowData.id = String.valueOf(record.Row_Number__c.intValue());
                rowData.label = record.Description__c != null ? record.Description__c : '';
                rowData.mirrorsBuilding = true; // Default to true
                rowData.defaultValue = record.Default_Amount__c != null ? record.Default_Amount__c : 0;

                if (section == 1){
                    summaryRows.add(rowData);
                } else if (section == 2){
                    buildingRows.add(rowData);
                }
            }

            wrapper.summaryRows = summaryRows;
            wrapper.buildingRows = buildingRows;

            System.debug('Returning ' + summaryRows.size() + ' summary rows and ' + buildingRows.size() + ' building rows');

        } catch (Exception e){
            System.debug('Error in getSOVItems: ' + e.getMessage());
            throw new AuraHandledException('Error loading SOV data: ' + e.getMessage());
        }

        return wrapper;
    }

    /**
     * Save Schedule of Values data
     */
    @AuraEnabled
    public static Id saveSOVSheet(Id opportunityId, String base64Data){
        if (opportunityId == null){
            throw new AuraHandledException('Opportunity Id is required.');
        }
        if (String.isBlank(base64Data)){
            throw new AuraHandledException('No SOV data provided.');
        }

        // Decode the incoming JSON payload
        Blob body;
        try{
            body = EncodingUtil.base64Decode(base64Data);
        } catch (Exception e){
            throw new AuraHandledException('Failed to decode SOV data: ' + e.getMessage());
        }

        // Use consistent title for SOV
        String sovTitle = 'BidWorksheet_SOV';
        String sovFileName = 'BidWorksheet_SOV.json';

        // Find an existing ContentDocument linked to this Opportunity with the SOV title
        ContentDocumentLink existingLink;
        try{
            existingLink = [SELECT ContentDocumentId, ContentDocument.Title
                            FROM ContentDocumentLink
                            WHERE LinkedEntityId = :opportunityId AND ContentDocument.Title = :sovTitle
                            LIMIT 1];
        } catch (QueryException qe){
            existingLink = null;
        }

        ContentVersion cv;
        if (existingLink != null){
            // Add a new version to the existing document
            cv = new ContentVersion(
                ContentDocumentId = existingLink.ContentDocumentId, 
                Title = sovTitle, 
                PathOnClient = sovFileName, 
                VersionData = body
            );
            insert cv;
            System.debug('✅ Updated existing SOV document');
        } else{
            // Create the first version and publish directly to the Opportunity
            cv = new ContentVersion(
                FirstPublishLocationId = opportunityId, 
                Title = sovTitle, 
                PathOnClient = sovFileName, 
                VersionData = body
            );
            insert cv;
            System.debug('✅ Created new SOV document');
        }

        // Delete auto-save attachment after save
        deleteAutoSaveDocumentGeneric(opportunityId, 'BidWorksheet_SOV_AutoSave');

        return cv.Id;
    }

    /**
     * Load saved Schedule of Values data
     */
    @AuraEnabled
    public static String loadSOVSheet(Id opportunityId){
        try{
            if (opportunityId == null){
                throw new AuraHandledException('Opportunity Id is required.');
            }

            String sovTitle = 'BidWorksheet_SOV';

            // First find ContentDocumentIds linked to this Opportunity with the desired title
            List<ContentDocumentLink> links = [SELECT ContentDocumentId
                                               FROM ContentDocumentLink
                                               WHERE LinkedEntityId = :opportunityId AND ContentDocument.Title = :sovTitle];

            if (links.isEmpty()){
                System.debug('No SOV file found for Opportunity: ' + opportunityId);
                return null;
            }

            Set<Id> docIds = new Set<Id>();
            for (ContentDocumentLink l : links){
                docIds.add(l.ContentDocumentId);
            }

            // Get the latest version of the document
            ContentVersion latest;
            try{
                latest = [SELECT VersionData
                          FROM ContentVersion
                          WHERE ContentDocumentId IN:docIds
                          ORDER BY CreatedDate DESC
                          LIMIT 1];
            } catch (QueryException qe){
                return null;
            }

            if (latest == null || latest.VersionData == null){
                return null;
            }

            // Return as plain JSON string
            String jsonData = latest.VersionData.toString();
            System.debug('✅ Loaded SOV data for Opportunity: ' + opportunityId);
            return jsonData;

        } catch (Exception e){
            System.debug('❌ Error loading SOV: ' + e.getMessage());
            System.debug('Stack trace: ' + e.getStackTraceString());
            return null;
        }
    }

    // ========================================
    // WRAPPER CLASSES FOR SCHEDULE OF VALUES
    // ========================================
    public class SOVMetadataWrapper{
        @AuraEnabled
        public List<SOVRowData> summaryRows{ get; set; }

        @AuraEnabled
        public List<SOVRowData> buildingRows{ get; set; }

        public SOVMetadataWrapper(){
            this.summaryRows = new List<SOVRowData>();
            this.buildingRows = new List<SOVRowData>();
        }

    }

    public class SOVRowData{
        @AuraEnabled
        public String id{ get; set; }

        @AuraEnabled
        public String label{ get; set; }

        @AuraEnabled
        public Boolean mirrorsBuilding{ get; set; }

        @AuraEnabled
        public Decimal defaultValue{ get; set; }

        public SOVRowData(){
            this.id = '';
            this.label = '';
            this.mirrorsBuilding = true;
            this.defaultValue = 0;
        }

    }

    // ========================================
    // GENERIC HELPER METHODS (Reusable Pattern)
    // ========================================
    
    /**
     * Generic helper to get next version number for any worksheet type
     * @param opportunityId The Opportunity ID
     * @param sheetTitle The title of the ContentDocument
     * @return Next version number (count + 1, or 1 if no document exists)
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
     * Generic helper to get version history for any worksheet type
     * @param opportunityId The Opportunity ID
     * @param sheetTitle The title of the ContentDocument
     * @return List of version history wrappers
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
     * Generic helper to load latest sheet (checks both version control and autosave)
     * @param opportunityId The Opportunity ID
     * @param sheetTitle The title of the ContentDocument for version control
     * @param autoSaveTitle The name of the Attachment for autosave
     * @return Base64 encoded data, or null if no data exists
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
     * Generic helper to autosave sheet data using Attachment
     * @param opportunityId The Opportunity ID
     * @param base64Data Base64 encoded data
     * @param autoSaveTitle The name for the autosave Attachment
     * @return Attachment ID
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
     * @param opportunityId The Opportunity ID
     * @param autoSaveTitle The name of the autosave Attachment/ContentDocument
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
     * Auto-save sheet data to auto-save document
     * Creates new document if it doesn't exist, or adds new version to existing document
     */
    @AuraEnabled
    public static Id autoSaveSheet(Id opportunityId, String base64Data){
        // Use generic helper for Underground worksheet
        return autoSaveSheetGeneric(opportunityId, base64Data, AUTO_SAVE_TITLE);
    }

    /**
     * Get the next version number that will be created on manual save
     * Returns count of existing versions + 1, or 1 if no version control document exists
     */
    @AuraEnabled(cacheable = false)
    public static Integer getNextVersionNumber(Id opportunityId){
        // Use generic helper for Underground worksheet
        return getNextVersionNumberGeneric(opportunityId, SHEET_TITLE);
    }

    /**
     * Get version history for the version control document
     * Returns list of all versions with metadata
     */
    @AuraEnabled(cacheable = false)
    public static List<VersionHistoryWrapper> getVersionHistory(Id opportunityId){
        // Use generic helper for Underground worksheet
        return getVersionHistoryGeneric(opportunityId, SHEET_TITLE);
    }

    /**
     * Load a specific version by ContentVersion ID
     * This method is already generic - works for any ContentVersion
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

    // ========================================
    // WRAPPER METHODS FOR ESTIMATE WORKSHEET
    // ========================================

    /**
     * Get next version number for Estimate worksheet
     */
    @AuraEnabled(cacheable = false)
    public static Integer getNextVersionNumber_Estimate(Id opportunityId){
        return getNextVersionNumberGeneric(opportunityId, 'BidWorksheet_Estimate');
    }

    /**
     * Get version history for Estimate worksheet
     */
    @AuraEnabled(cacheable = false)
    public static List<VersionHistoryWrapper> getVersionHistory_Estimate(Id opportunityId){
        return getVersionHistoryGeneric(opportunityId, 'BidWorksheet_Estimate');
    }

    /**
     * Load latest Estimate sheet (checks autosave + version control)
     */
    @AuraEnabled(cacheable = false)
    public static String loadLatestEstimateSheet(Id opportunityId){
        return loadLatestSheetGeneric(opportunityId, 'BidWorksheet_Estimate', 'BidWorksheet_Estimate_AutoSave');
    }

    /**
     * Load specific Estimate version by ID
     */
    @AuraEnabled(cacheable = false)
    public static String loadVersionById_Estimate(Id versionId){
        return loadVersionById(versionId); // Reuse existing generic method
    }

    /**
     * Autosave Estimate worksheet
     */
    @AuraEnabled
    public static Id autoSaveEstimateSheet(Id opportunityId, String base64Data){
        return autoSaveSheetGeneric(opportunityId, base64Data, 'BidWorksheet_Estimate_AutoSave');
    }

    // ========================================
    // WRAPPER METHODS FOR SOV WORKSHEET
    // ========================================

    /**
     * Get next version number for SOV worksheet
     */
    @AuraEnabled(cacheable = false)
    public static Integer getNextVersionNumber_SOV(Id opportunityId){
        return getNextVersionNumberGeneric(opportunityId, 'BidWorksheet_SOV');
    }

    /**
     * Get version history for SOV worksheet
     */
    @AuraEnabled(cacheable = false)
    public static List<VersionHistoryWrapper> getVersionHistory_SOV(Id opportunityId){
        return getVersionHistoryGeneric(opportunityId, 'BidWorksheet_SOV');
    }

    /**
     * Load latest SOV sheet (checks autosave + version control)
     */
    @AuraEnabled(cacheable = false)
    public static String loadLatestSOVSheet(Id opportunityId){
        return loadLatestSheetGeneric(opportunityId, 'BidWorksheet_SOV', 'BidWorksheet_SOV_AutoSave');
    }

    /**
     * Load specific SOV version by ID
     */
    @AuraEnabled(cacheable = false)
    public static String loadVersionById_SOV(Id versionId){
        return loadVersionById(versionId); // Reuse existing generic method
    }

    /**
     * Autosave SOV worksheet
     */
    @AuraEnabled
    public static Id autoSaveSOVSheet(Id opportunityId, String base64Data){
        return autoSaveSheetGeneric(opportunityId, base64Data, 'BidWorksheet_SOV_AutoSave');
    }

    // ========================================
    // WRAPPER METHODS FOR DESIGN WORKSHEET
    // ========================================

    /**
     * Get next version number for Design worksheet
     */
    @AuraEnabled(cacheable = false)
    public static Integer getNextVersionNumber_Design(Id opportunityId){
        return getNextVersionNumberGeneric(opportunityId, 'DesignWorksheet');
    }

    /**
     * Get version history for Design worksheet
     */
    @AuraEnabled(cacheable = false)
    public static List<VersionHistoryWrapper> getVersionHistory_Design(Id opportunityId){
        return getVersionHistoryGeneric(opportunityId, 'DesignWorksheet');
    }

    /**
     * Load latest Design worksheet (checks autosave + version control)
     */
    @AuraEnabled(cacheable = false)
    public static String loadLatestDesignWorksheet(Id opportunityId){
        return loadLatestSheetGeneric(opportunityId, 'DesignWorksheet', 'DesignWorksheet_AutoSave');
    }

    /**
     * Load specific Design version by ID
     */
    @AuraEnabled(cacheable = false)
    public static String loadVersionById_Design(Id versionId){
        return loadVersionById(versionId); // Reuse existing generic method
    }

    /**
     * Autosave Design worksheet
     */
    @AuraEnabled
    public static Id autoSaveDesignWorksheet(Id opportunityId, String base64Data){
        return autoSaveSheetGeneric(opportunityId, base64Data, 'DesignWorksheet_AutoSave');
    }

    /**
     * Helper method to delete auto-save document (Underground)
     */
    private static void deleteAutoSaveDocument(Id opportunityId){
        // Use generic helper for Underground worksheet
        deleteAutoSaveDocumentGeneric(opportunityId, AUTO_SAVE_TITLE);
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