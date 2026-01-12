import { LightningElement, api, track, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getITMItems from '@salesforce/apex/ITMBidWorksheetController.getITMItems';
import saveITMWorksheet from '@salesforce/apex/ITMBidWorksheetController.saveITMWorksheet';
import loadITMWorksheet from '@salesforce/apex/ITMBidWorksheetController.loadITMWorksheet';
import updateOpportunityFields from '@salesforce/apex/BidWorksheetUndergroundController.updateOpportunityFields';
import getNextVersionNumber from '@salesforce/apex/ITMBidWorksheetController.getNextVersionNumber';
import getVersionHistory from '@salesforce/apex/ITMBidWorksheetController.getVersionHistory';
import loadVersionById from '@salesforce/apex/ITMBidWorksheetController.loadVersionById';
import loadLatestITMWorksheet from '@salesforce/apex/ITMBidWorksheetController.loadLatestITMWorksheet';
import autoSaveITMWorksheet from '@salesforce/apex/ITMBidWorksheetController.autoSaveITMWorksheet';

export default class ItmBidWorksheet extends LightningElement {
    @api recordId; // Opportunity ID - automatically set by Quick Action
    @track isLoading = true;
    @track isSaving = false;

    // Fallback recordId for testing/development
    fallbackRecordId = '006VF00000I9RJaYAN';

    // Data structures
    @track laborFactorRows = [];
    @track equipmentFactorRows = [];
    @track ratesRows = [];

    // Totals
    @track totalAlarmLaborHours = '0.00';
    @track totalSprinklerLaborHours = '0.00';
    @track totalAlarmEquipmentCost = '0.00';
    @track totalSprinklerEquipmentCost = '0.00';

    @track alarmSubtotal = '0.00';
    @track sprinklerSubtotal = '0.00';
    @track alarmGainPercent = 0.2; // 20% default
    @track sprinklerGainPercent = 0.1; // 10% default
    @track alarmTotalQuote = '0.00';
    @track sprinklerTotalQuote = '0.00';

    // Version control state
    @track versionList = [];
    @track selectedVersionId = '';
    @track nextVersionNumber = 1;
    @track isLoadingVersion = false;
    
    // Auto-save state
    autoSaveTimeout = null;
    @track autoSaveStatus = ''; // 'saving', 'saved', ''
    _isInitializing = true;
    _isLoadingData = false;
    _isUserEditing = false;
    _editingTimeout = null;

    // Wire to load metadata
    @wire(getITMItems)
    wiredItems({ error, data }) {
        if (data) {
            console.log('✅ Metadata loaded', data);
            this.initializeRows(data);
            // Wait for recordId before loading saved data
            this.waitForRecordIdAndLoad();
        } else if (error) {
            console.error('❌ Error loading metadata:', error);
            this.showToast('Error', 'Failed to load ITM items: ' + error.body.message, 'error');
            this.isLoading = false;
        }
    }

    /**
     * Wait for recordId to be available, then load saved data
     */
    async waitForRecordIdAndLoad() {
        // Try up to 10 times with 100ms delay
        let attempts = 0;
        const maxAttempts = 10;

        while (attempts < maxAttempts) {
            if (this.recordId) {
                console.log('✅ recordId available:', this.recordId);
                // Load version data first, then saved data
                await this.loadVersionData();
                await this.loadSavedData();
                return;
            }

            console.log(`⏳ Waiting for recordId... attempt ${attempts + 1}/${maxAttempts}`);
            await new Promise(resolve => setTimeout(resolve, 100));
            attempts++;
        }

        // If still no recordId after all attempts, use fallback
        console.warn('⚠️ recordId not set after', maxAttempts, 'attempts, using fallback');
        await this.loadVersionData();
        await this.loadSavedData();
    }

    connectedCallback() {
        this._isInitializing = true;
        setTimeout(() => {
            this._isInitializing = false;
        }, 3000);
    }

    /**
     * Initialize rows from metadata
     */
    initializeRows(data) {
        console.log('Initializing rows...');

        // Group items by row number
        const laborByRow = this.groupByRow(data.laborFactor);
        const equipmentByRow = this.groupByRow(data.equipmentFactor);
        const ratesByRow = this.groupByRow(data.rates);

        // Create row objects with left (Alarm) and right (Sprinkler) columns
        this.laborFactorRows = this.createRowPairs(laborByRow);
        this.equipmentFactorRows = this.createRowPairs(equipmentByRow);
        this.ratesRows = this.createRowPairs(ratesByRow);

        console.log('Labor Factor rows:', this.laborFactorRows.length);
        console.log('Labor Factor rows:', this.laborFactorRows);
        console.log('Equipment Factor rows:', this.equipmentFactorRows.length);
        console.log('Rates rows:', this.ratesRows.length);
    }

    /**
     * Group items by row number
     */
    groupByRow(items) {
        const grouped = {};
        items.forEach(item => {
            const rowNum = item.rowNumber;
            if (!grouped[rowNum]) {
                grouped[rowNum] = { left: null, right: null };
            }
            if (item.column === 'Left') {
                grouped[rowNum].left = item;
            } else {
                grouped[rowNum].right = item;
            }
        });
        return grouped;
    }

    /**
     * Create row pair objects
     */
    createRowPairs(groupedItems) {
        const rows = [];
        Object.keys(groupedItems).sort((a, b) => a - b).forEach(rowNum => {
            const pair = groupedItems[rowNum];
            const rowNumber = parseInt(rowNum);
            rows.push({
                rowNumber: rowNumber,
                isLaborMixRate: rowNumber === 41, // Used for disabling quantity input
                left: {
                    id: pair.left?.id || '',
                    description: pair.left?.description || '',
                    quantity: 0,
                    hours: pair.left?.defaultHours || 0,
                    total: '0.00'
                },
                right: {
                    id: pair.right?.id || '',
                    description: pair.right?.description || '',
                    quantity: 0,
                    hours: pair.right?.defaultHours || 0,
                    total: '0.00'
                }
            });
        });
        return rows;
    }

    /**
     * Load version data (version list and next version number)
     */
    async loadVersionData() {
        const targetId = this.recordId || this.fallbackRecordId;
        if (!targetId) return;

        try {
            await this.loadNextVersionNumber();
            await this.loadVersionList();
        } catch (error) {
            console.error('Error loading version data:', error);
        }
    }

    async loadNextVersionNumber() {
        const targetId = this.recordId || this.fallbackRecordId;
        if (!targetId) return;
        
        try {
            this.nextVersionNumber = await getNextVersionNumber({ opportunityId: targetId });
        } catch (error) {
            console.error('Error loading next version number:', error);
            this.nextVersionNumber = 1;
        }
    }

    async loadVersionList() {
        const targetId = this.recordId || this.fallbackRecordId;
        if (!targetId) return;
        
        try {
            if (!this.nextVersionNumber) {
                await this.loadNextVersionNumber();
            }
            
            const versions = await getVersionHistory({ opportunityId: targetId });
            
            const savedVersions = versions.map(v => ({
                label: `Version ${v.versionNumber} - ${this.formatDate(v.createdDate)} - ${v.createdBy}`,
                value: v.versionId,
                versionNumber: v.versionNumber,
                isDraft: false
            }));
            
            const draftOption = {
                label: `Draft - Version ${this.nextVersionNumber}`,
                value: 'draft',
                versionNumber: this.nextVersionNumber,
                isDraft: true
            };
            
            this.versionList = [draftOption, ...savedVersions];
            
            if (!this.selectedVersionId || this.selectedVersionId === 'draft') {
                this.selectedVersionId = 'draft';
            }
        } catch (error) {
            console.error('Error loading version list:', error);
            this.versionList = [{
                label: `Draft - Version ${this.nextVersionNumber || 1}`,
                value: 'draft',
                versionNumber: this.nextVersionNumber || 1,
                isDraft: true
            }];
            this.selectedVersionId = 'draft';
        }
    }

    formatDate(dateTime) {
        if (!dateTime) return '';
        const date = new Date(dateTime);
        return date.toLocaleString('en-US', {
            month: 'short',
            day: 'numeric',
            year: 'numeric',
            hour: 'numeric',
            minute: '2-digit'
        });
    }

    /**
     * Load saved worksheet data
     */
    async loadSavedData() {
        if (!this.recordId && !this.fallbackRecordId) {
            console.log('❌ [LOAD ITM] No recordId, skipping load');
            return;
        }

        // Don't load if rows haven't been initialized
        if (this.laborFactorRows.length === 0 && this.equipmentFactorRows.length === 0 && this.ratesRows.length === 0) {
            console.log('⚠️ [LOAD ITM] Rows not initialized yet');
            return;
        }

        // Don't load if user is actively editing
        if (this._isUserEditing) {
            console.log('📍 [LOAD ITM] User is editing, skipping load');
            return;
        }

        this._isLoadingData = true;

        try {
            const targetId = this.recordId || this.fallbackRecordId;
            console.log('🔵 loadSavedData called');
            console.log('🔵 recordId:', this.recordId);
            console.log('🔵 targetId (with fallback):', targetId);

            let savedData;
            
            // If selectedVersionId is set and not draft, load that specific version
            if (this.selectedVersionId && this.selectedVersionId !== 'draft') {
                console.log('🔍 [LOAD ITM] Loading specific version:', this.selectedVersionId);
                const base64Data = await loadVersionById({ versionId: this.selectedVersionId });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                }
            } else {
                // Otherwise, load latest (autosave or most recent)
                console.log('🔍 [LOAD ITM] Loading latest (draft)');
                const base64Data = await loadLatestITMWorksheet({ opportunityId: targetId });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                } else {
                    // Fallback to old method
                    const oldData = await loadITMWorksheet({ opportunityId: targetId });
                    if (oldData) {
                        savedData = oldData;
                    }
                }
            }

            console.log('🔵 savedData received:', savedData ? 'YES' : 'NO');

            if (savedData) {
                console.log('Loading saved data...');
                const data = typeof savedData === 'string' ? JSON.parse(savedData) : savedData;

                console.log('🔵 Parsed data keys:', Object.keys(data));
                console.log('🔵 Parsed data keys:', data);
                console.log('🔵 laborFactorRows count:', data.laborFactorRows?.length);
                console.log('🔵 equipmentFactorRows count:', data.equipmentFactorRows?.length);
                console.log('🔵 ratesRows count:', data.ratesRows?.length);

                // Restore quantities and hours
                if (data.laborFactorRows) {
                    this.restoreRows(this.laborFactorRows, data.laborFactorRows);
                }
                if (data.equipmentFactorRows) {
                    this.restoreRows(this.equipmentFactorRows, data.equipmentFactorRows);
                }
                if (data.ratesRows) {
                    this.restoreRows(this.ratesRows, data.ratesRows);
                }

                // Restore gain percentages
                if (data.alarmGainPercent !== undefined) {
                    this.alarmGainPercent = data.alarmGainPercent;
                }
                if (data.sprinklerGainPercent !== undefined) {
                    this.sprinklerGainPercent = data.sprinklerGainPercent;
                }

                console.log('✅ Saved data restored');
            } else {
                console.log('⚠️ No saved data found');
            }

        } catch (error) {
            console.error('❌ Error loading saved data:', error);
        } finally {
            this.isLoading = false;
            this.calculateAllTotals();
            setTimeout(() => {
                this._isLoadingData = false;
            }, 500);
        }
    }

    decodeData(base64Data) {
        try {
            const binaryString = atob(base64Data);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            const decoder = new TextDecoder('utf-8');
            return decoder.decode(bytes);
        } catch (err) {
            console.error('Decode data failed:', err);
            throw err;
        }
    }

    async handleVersionChange(event) {
        const newVersionId = event.detail.value;
        if (!newVersionId) return;
        
        // If switching away from draft and there are unsaved changes, save first
        if (this.selectedVersionId === 'draft' && newVersionId !== 'draft') {
            if (this.autoSaveStatus === 'saving') {
                let waitCount = 0;
                while (this.autoSaveStatus === 'saving' && waitCount < 25) {
                    await new Promise(resolve => setTimeout(resolve, 200));
                    waitCount++;
                }
            }
            
            if (this.autoSaveTimeout) {
                clearTimeout(this.autoSaveTimeout);
                this.autoSaveTimeout = null;
                await this.performAutoSave();
            }
        }
        
        this.selectedVersionId = newVersionId;
        
        if (this.selectedVersionId !== 'draft') {
            this.isLoadingVersion = true;
            setTimeout(() => {
                this.isLoadingVersion = false;
            }, 2000);
        }
        
        await this.loadSavedData();
    }

    handleCellChange(event) {
        console.log('📍 [ITM] handleCellChange called', {
            isInitializing: this._isInitializing,
            isLoadingVersion: this.isLoadingVersion,
            isLoadingData: this._isLoadingData,
            recordId: this.recordId
        });
        
        if (this._isInitializing || this.isLoadingVersion || this._isLoadingData) {
            console.log('📍 [ITM] Skipping autosave - initialization/loading in progress');
            return;
        }
        
        this._isUserEditing = true;
        if (this._editingTimeout) {
            clearTimeout(this._editingTimeout);
        }
        this._editingTimeout = setTimeout(() => {
            this._isUserEditing = false;
        }, 1000);
        
        if (this.autoSaveTimeout) {
            clearTimeout(this.autoSaveTimeout);
        }
        
        console.log('📍 [ITM] Setting autosave timeout (2 seconds)');
        this.autoSaveTimeout = setTimeout(() => {
            console.log('📍 [ITM] Autosave timeout fired, calling performAutoSave');
            this.performAutoSave();
        }, 2000);
    }

    async performAutoSave() {
        const targetId = this.recordId || this.fallbackRecordId;
        console.log('📍 [ITM] performAutoSave called', { targetId, recordId: this.recordId });
        
        if (!targetId) {
            console.warn('⚠️ [ITM] No targetId available for autosave');
            return;
        }

        try {
            console.log('📍 [ITM] Starting autosave...');
            this.autoSaveStatus = 'saving';
            
            const payload = await this.saveSheet();
            console.log('📍 [ITM] Payload created, encoding...');
            const base64Payload = this.encodeData(payload);

            console.log('📍 [ITM] Calling autoSaveITMWorksheet Apex method...');
            await autoSaveITMWorksheet({
                opportunityId: targetId,
                base64Data: base64Payload
            });

            // ⭐ Update Opportunity fields during autosave
            console.log('📍 [ITM] Updating Opportunity fields...');
            const fieldData = {
                Total_Alarm_Labor_Hours__c: parseFloat(this.totalAlarmLaborHours) || 0,
                Total_Sprinkler_Labor_Hours__c: parseFloat(this.totalSprinklerLaborHours) || 0,
                Total_Alarm_Cost__c: parseFloat(this.totalAlarmEquipmentCost) || 0,
                Total_Sprinkler_Cost__c: parseFloat(this.totalSprinklerEquipmentCost) || 0,
                TOTAL_Alarm_QUOTE_PRICE__c: parseFloat(this.alarmTotalQuote) || 0,
                TOTAL_Sprinkler_QUOTE_PRICE__c: parseFloat(this.sprinklerTotalQuote) || 0
            };

            try {
                await updateOpportunityFields({
                    opportunityId: targetId,
                    fieldDataJson: JSON.stringify(fieldData)
                });
                console.log('✅ [ITM] Opportunity fields updated during autosave');
            } catch (fieldError) {
                // Don't fail autosave if field update fails, just log it
                console.warn('⚠️ [ITM] Failed to update Opportunity fields during autosave:', fieldError);
            }

            this.autoSaveStatus = 'saved';
            console.log('✅ [ITM] Auto-saved ITM worksheet successfully');

            setTimeout(() => {
                this.autoSaveStatus = '';
            }, 2000);
        } catch (error) {
            console.error('❌ [ITM] Error during auto-save:', error);
            const errorMessage = error?.body?.message || error?.message || String(error);
            console.error('❌ [ITM] Error details:', errorMessage);
            this.autoSaveStatus = '';
        }
    }

    /**
     * Collect and return worksheet data for saving
     */
    async saveSheet() {
        const targetId = this.recordId || this.fallbackRecordId;
        
        return {
            worksheetType: 'ITM',
            version: '1.0',
            savedDate: new Date().toISOString(),
            opportunityId: targetId,
            laborFactorRows: this.laborFactorRows,
            equipmentFactorRows: this.equipmentFactorRows,
            ratesRows: this.ratesRows,
            alarmGainPercent: this.alarmGainPercent,
            sprinklerGainPercent: this.sprinklerGainPercent,
            totals: {
                totalAlarmLaborHours: this.totalAlarmLaborHours,
                totalSprinklerLaborHours: this.totalSprinklerLaborHours,
                totalAlarmEquipmentCost: this.totalAlarmEquipmentCost,
                totalSprinklerEquipmentCost: this.totalSprinklerEquipmentCost,
                alarmSubtotal: this.alarmSubtotal,
                sprinklerSubtotal: this.sprinklerSubtotal,
                alarmTotalQuote: this.alarmTotalQuote,
                sprinklerTotalQuote: this.sprinklerTotalQuote
            }
        };
    }

    encodeData(data) {
        try {
            const json = JSON.stringify(data);
            return btoa(unescape(encodeURIComponent(json)));
        } catch (err) {
            console.error('Encode data failed', err);
            throw err;
        }
    }

    get isAutoSaving() {
        return this.autoSaveStatus === 'saving';
    }

    get isAutoSaved() {
        return this.autoSaveStatus === 'saved';
    }

    get versionListDisabled() {
        return this.versionList.length === 0;
    }

    /**
     * Restore row data from saved state
     */
    restoreRows(targetRows, savedRows) {
        savedRows.forEach((savedRow, index) => {
            if (index < targetRows.length) {
                const row = targetRows[index];

                // Restore left column
                if (savedRow.left) {
                    row.left.quantity = savedRow.left.quantity || 0;
                    row.left.hours = savedRow.left.hours || row.left.hours;
                    row.left.total = savedRow.left.total || '0.00';
                }

                // Restore right column
                if (savedRow.right) {
                    row.right.quantity = savedRow.right.quantity || 0;
                    row.right.hours = savedRow.right.hours || row.right.hours;
                    row.right.total = savedRow.right.total || '0.00';
                }
            }
        });

        // Force reactivity update
        if (targetRows === this.laborFactorRows) {
            this.laborFactorRows = [...this.laborFactorRows];
        } else if (targetRows === this.equipmentFactorRows) {
            this.equipmentFactorRows = [...this.equipmentFactorRows];
        } else if (targetRows === this.ratesRows) {
            this.ratesRows = [...this.ratesRows];
        }
    }

    /**
     * Handle quantity input change
     */
    handleQuantityChange(event) {
        this.handleCellChange(event);
        
        const rowNumber = parseInt(event.target.dataset.row);
        const column = event.target.dataset.column;
        const section = event.target.dataset.section;
        const value = parseFloat(event.target.value) || 0;

        console.log(`Quantity changed: Row ${rowNumber}, Column ${column}, Value ${value}`);

        this.updateRowValue(section, rowNumber, column, 'quantity', value);
        this.calculateAllTotals();
    }

    /**
     * Handle hours input change
     */
    handleHoursChange(event) {
        this.handleCellChange(event);
        
        const rowNumber = parseInt(event.target.dataset.row);
        const column = event.target.dataset.column;
        const section = event.target.dataset.section;
        const value = parseFloat(event.target.value) || 0;

        console.log(`Hours changed: Row ${rowNumber}, Column ${column}, Value ${value}`);

        this.updateRowValue(section, rowNumber, column, 'hours', value);
        this.calculateAllTotals();
    }

    /**
     * Update row value helper
     */
    updateRowValue(section, rowNumber, column, field, value) {
        let rows;
        if (section === 'labor') {
            rows = this.laborFactorRows;
        } else if (section === 'equipment') {
            rows = this.equipmentFactorRows;
        } else if (section === 'rates') {
            rows = this.ratesRows;
        }

        const row = rows.find(r => r.rowNumber === rowNumber);
        if (row) {
            row[column][field] = value;
            // Calculate row total (2 decimal places)
            const total = row[column].quantity * row[column].hours;
            row[column].total = total > 0 ? total.toFixed(2) : '0.00';
        }
    }

    /**
     * Handle gain percent change
     */
    handleGainPercentChange(event) {
        this.handleCellChange(event);
        
        const column = event.target.dataset.column;
        const value = parseFloat(event.target.value) || 0;

        if (column === 'left') {
            this.alarmGainPercent = value;
        } else {
            this.sprinklerGainPercent = value;
        }

        this.calculateAllTotals();
    }

    /**
     * Calculate all totals
     */
    calculateAllTotals() {
        console.log('Calculating all totals...');

        // Section 1: Labor Factor totals
        this.totalAlarmLaborHours = this.sumColumn(this.laborFactorRows, 'left').toFixed(2);
        this.totalSprinklerLaborHours = this.sumColumn(this.laborFactorRows, 'right').toFixed(2);

        console.log('Total Alarm Labor Hours:', this.totalAlarmLaborHours);
        console.log('Total Sprinkler Labor Hours:', this.totalSprinklerLaborHours);

        // Section 2: Equipment Factor totals
        this.totalAlarmEquipmentCost = this.sumColumn(this.equipmentFactorRows, 'left').toFixed(2);
        this.totalSprinklerEquipmentCost = this.sumColumn(this.equipmentFactorRows, 'right').toFixed(2);

        console.log('Total Alarm Equipment Cost:', this.totalAlarmEquipmentCost);
        console.log('Total Sprinkler Equipment Cost:', this.totalSprinklerEquipmentCost);

        // Section 3: Rates - Auto-populate Labor Mix Rate quantities
        if (this.ratesRows.length > 0) {
            // First row (41) is Labor Mix Rate - should auto-populate from labor hours
            this.ratesRows[0].left.quantity = parseFloat(this.totalAlarmLaborHours);
            this.ratesRows[0].left.total = (this.ratesRows[0].left.quantity * this.ratesRows[0].left.hours).toFixed(2);

            this.ratesRows[0].right.quantity = parseFloat(this.totalSprinklerLaborHours);
            this.ratesRows[0].right.total = (this.ratesRows[0].right.quantity * this.ratesRows[0].right.hours).toFixed(2);
        }

        // Calculate rates section totals
        const alarmRatesTotal = this.sumColumn(this.ratesRows, 'left');
        const sprinklerRatesTotal = this.sumColumn(this.ratesRows, 'right');

        // Subtotal = Equipment + Rates
        this.alarmSubtotal = (parseFloat(this.totalAlarmEquipmentCost) + alarmRatesTotal).toFixed(2);
        this.sprinklerSubtotal = (parseFloat(this.totalSprinklerEquipmentCost) + sprinklerRatesTotal).toFixed(2);

        console.log('Alarm Subtotal:', this.alarmSubtotal);
        console.log('Sprinkler Subtotal:', this.sprinklerSubtotal);

        // Total Quote = Subtotal * (1 + Gain%)
        this.alarmTotalQuote = (parseFloat(this.alarmSubtotal) * (1 + this.alarmGainPercent)).toFixed(2);
        this.sprinklerTotalQuote = (parseFloat(this.sprinklerSubtotal) * (1 + this.sprinklerGainPercent)).toFixed(2);

        console.log('Alarm Total Quote:', this.alarmTotalQuote);
        console.log('Sprinkler Total Quote:', this.sprinklerTotalQuote);
    }

    /**
     * Sum column helper
     */
    sumColumn(rows, column) {
        return rows.reduce((sum, row) => {
            const total = parseFloat(row[column].total) || 0;
            return sum + total;
        }, 0);
    }

    /**
     * Format currency
     */
    formatCurrency(value) {
        return new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(value || 0);
    }

    /**
     * Format number
     */
    formatNumber(value) {
        return new Intl.NumberFormat('en-US', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(value || 0);
    }

    /**
     * Save worksheet
     */
    async handleSave() {
        const targetId = this.recordId || this.fallbackRecordId;
        console.log('🔵 handleSave called');
        console.log('🔵 recordId:', this.recordId);
        console.log('🔵 targetId (with fallback):', targetId);
        console.log('🔵 isSaving:', this.isSaving);

        this.isSaving = true;

        try {
            console.log('Saving ITM worksheet...');

            // Prepare data to save
            const worksheetData = {
                worksheetType: 'ITM',
                version: '1.0',
                savedDate: new Date().toISOString(),
                opportunityId: targetId,
                laborFactorRows: this.laborFactorRows,
                equipmentFactorRows: this.equipmentFactorRows,
                ratesRows: this.ratesRows,
                alarmGainPercent: this.alarmGainPercent,
                sprinklerGainPercent: this.sprinklerGainPercent,
                totals: {
                    totalAlarmLaborHours: this.totalAlarmLaborHours,
                    totalSprinklerLaborHours: this.totalSprinklerLaborHours,
                    totalAlarmEquipmentCost: this.totalAlarmEquipmentCost,
                    totalSprinklerEquipmentCost: this.totalSprinklerEquipmentCost,
                    alarmSubtotal: this.alarmSubtotal,
                    sprinklerSubtotal: this.sprinklerSubtotal,
                    alarmTotalQuote: this.alarmTotalQuote,
                    sprinklerTotalQuote: this.sprinklerTotalQuote
                }
            };

            // Encode to base64
            const jsonString = JSON.stringify(worksheetData);
            const base64Data = btoa(unescape(encodeURIComponent(jsonString)));

            // Save to Salesforce
            await saveITMWorksheet({
                opportunityId: targetId,
                base64Data: base64Data
            });

            console.log('✅ ITM worksheet saved successfully');

            // ⭐ Refresh version data AFTER save (so nextVersionNumber is updated)
            await this.loadVersionData();
            
            // Set draft as selected after save
            this.selectedVersionId = 'draft';

            // Prepare Opportunity field updates
            const fieldData = {
                Total_Alarm_Labor_Hours__c: parseFloat(this.totalAlarmLaborHours) || 0,
                Total_Sprinkler_Labor_Hours__c: parseFloat(this.totalSprinklerLaborHours) || 0,
                Total_Alarm_Cost__c: parseFloat(this.totalAlarmEquipmentCost) || 0,
                Total_Sprinkler_Cost__c: parseFloat(this.totalSprinklerEquipmentCost) || 0,
                TOTAL_Alarm_QUOTE_PRICE__c: parseFloat(this.alarmTotalQuote) || 0,
                TOTAL_Sprinkler_QUOTE_PRICE__c: parseFloat(this.sprinklerTotalQuote) || 0
            };

            // Call existing Apex to update the Opportunity fields
            await updateOpportunityFields({
                opportunityId: targetId,
                fieldDataJson: JSON.stringify(fieldData)
            });

            console.log('✅ Opportunity fields updated for ITM worksheet');
            this.showToast('Success', 'ITM Bid Worksheet saved successfully!', 'success');

        } catch (error) {
            console.error('❌ Error saving worksheet:', error);
            this.showToast('Error', 'Failed to save worksheet: ' + error.body.message, 'error');
        } finally {
            this.isSaving = false;
        }
    }

    /**
     * Show toast notification
     */
    showToast(title, message, variant) {
        const event = new ShowToastEvent({
            title: title,
            message: message,
            variant: variant
        });
        this.dispatchEvent(event);
    }

    /**
     * Get current date for header
     */
    get currentDate() {
        return new Date().toLocaleDateString('en-US');
    }
}