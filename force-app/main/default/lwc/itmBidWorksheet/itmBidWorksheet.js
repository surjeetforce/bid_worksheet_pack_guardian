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
import getFinalVersionStatus from '@salesforce/apex/ITMBidWorksheetController.getFinalVersionStatus';

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
    @track totalAlarmLaborHours = 0;
    @track totalSprinklerLaborHours = 0;
    @track totalAlarmEquipmentCost = 0;
    @track totalSprinklerEquipmentCost = 0;

    @track alarmSubtotal = 0;
    @track sprinklerSubtotal = 0;
    @track alarmGainPercent = 20; // 20% default (whole number)
    @track sprinklerGainPercent = 10; // 10% default (whole number)
    @track alarmTotalQuote = 0;
    @track sprinklerTotalQuote = 0;

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

    // Final version state
    @track hasFinalVersion = false;
    @track finalVersionId = null;
    @track isReadOnly = false;
    @track showMarkAsFinalCheckbox = false;

    // Wire to load metadata
    @wire(getITMItems)
    wiredItems({ error, data }) {
        if (data) {
            this.initializeRows(data);
            // Wait for recordId before loading saved data
            this.waitForRecordIdAndLoad();
        } else if (error) {
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
                // Load version data first, then saved data
                await this.loadVersionData();
                await this.loadSavedData();
                return;
            }

            await new Promise(resolve => setTimeout(resolve, 100));
            attempts++;
        }

        // If still no recordId after all attempts, use fallback
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

        // Group items by row number
        const laborByRow = this.groupByRow(data.laborFactor);
        const equipmentByRow = this.groupByRow(data.equipmentFactor);
        const ratesByRow = this.groupByRow(data.rates);

        // Create row objects with left (Alarm) and right (Sprinkler) columns
        this.laborFactorRows = this.createRowPairs(laborByRow);
        this.equipmentFactorRows = this.createRowPairs(equipmentByRow);
        this.ratesRows = this.createRowPairs(ratesByRow);

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
            const isLaborMixRate = rowNumber === 41;
            rows.push({
                rowNumber: rowNumber,
                isLaborMixRate: isLaborMixRate, // Used for disabling quantity input
                ratesQtyDisabled: isLaborMixRate, // Updated in updateReadOnlyState when isReadOnly changes
                left: {
                    id: pair.left?.id || '',
                    description: pair.left?.description || '',
                    quantity: 0,
                    hours: pair.left?.defaultHours || 0,
                    total: 0
                },
                right: {
                    id: pair.right?.id || '',
                    description: pair.right?.description || '',
                    quantity: 0,
                    hours: pair.right?.defaultHours || 0,
                    total: 0
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
            // Check final version status first
            const finalStatus = await getFinalVersionStatus({ opportunityId: targetId });
            this.hasFinalVersion = finalStatus.hasFinalVersion;
            this.finalVersionId = finalStatus.finalVersionId;
            
            await this.loadNextVersionNumber();
            await this.loadVersionList();
        } catch (error) {
        }
    }

    async loadNextVersionNumber() {
        const targetId = this.recordId || this.fallbackRecordId;
        if (!targetId) return;
        
        try {
            this.nextVersionNumber = await getNextVersionNumber({ opportunityId: targetId });
        } catch (error) {
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
                label: `Version ${v.versionNumber}${v.isFinal ? ' - FINAL' : ''} - ${this.formatDate(v.createdDate)} - ${v.createdBy}`,
                value: v.versionId,
                versionNumber: v.versionNumber,
                isDraft: false,
                isFinal: v.isFinal
            }));
            
            // Only show draft option if no final version exists
            const versionListOptions = [];
            
            if (!this.hasFinalVersion) {
                versionListOptions.push({
                    label: `Draft - Version ${this.nextVersionNumber}`,
                    value: 'draft',
                    versionNumber: this.nextVersionNumber,
                    isDraft: true,
                    isFinal: false
                });
            }
            
            versionListOptions.push(...savedVersions);
            this.versionList = versionListOptions;
            
            // Set selected version
            if (!this.selectedVersionId || this.selectedVersionId === 'draft') {
                this.selectedVersionId = this.hasFinalVersion ? this.finalVersionId : 'draft';
            }
            
            // Update read-only state based on final version
            this.updateReadOnlyState();
            
        } catch (error) {
            
            if (!this.hasFinalVersion) {
                this.versionList = [{
                    label: `Draft - Version ${this.nextVersionNumber || 1}`,
                    value: 'draft',
                    versionNumber: this.nextVersionNumber || 1,
                    isDraft: true,
                    isFinal: false
                }];
                this.selectedVersionId = 'draft';
            } else {
                this.versionList = [];
                this.selectedVersionId = this.finalVersionId;
            }
            
            this.updateReadOnlyState();
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
     * Update read-only state based on final version status
     */
    updateReadOnlyState() {
        // Worksheet is read-only if a final version exists
        this.isReadOnly = this.hasFinalVersion;
        
        // Show "Mark as Final" checkbox only when viewing draft and no final version exists
        this.showMarkAsFinalCheckbox = !this.hasFinalVersion && this.selectedVersionId === 'draft';
        
        // Rates section: Hours (quantity) column is disabled when read-only OR Labor Mix Rate row
        this.ratesRows.forEach(row => {
            row.ratesQtyDisabled = this.isReadOnly || row.isLaborMixRate;
        });
        this.ratesRows = [...this.ratesRows];
    }

    /**
     * Load saved worksheet data
     */
    async loadSavedData() {
        if (!this.recordId && !this.fallbackRecordId) {
            return;
        }

        // Don't load if rows haven't been initialized
        if (this.laborFactorRows.length === 0 && this.equipmentFactorRows.length === 0 && this.ratesRows.length === 0) {
            return;
        }

        // Don't load if user is actively editing
        if (this._isUserEditing) {
            return;
        }

        this._isLoadingData = true;

        try {
            const targetId = this.recordId || this.fallbackRecordId;

            let savedData;
            
            // If selectedVersionId is set and not draft, load that specific version
            if (this.selectedVersionId && this.selectedVersionId !== 'draft') {
                const base64Data = await loadVersionById({ versionId: this.selectedVersionId });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                }
            } else {
                // Otherwise, load latest (autosave or most recent)
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


            if (savedData) {
                const data = typeof savedData === 'string' ? JSON.parse(savedData) : savedData;


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
                    const numValue = parseFloat(data.alarmGainPercent);
                    if (!isNaN(numValue)) {
                        this.alarmGainPercent = numValue.toFixed(5);
                    } else {
                        this.alarmGainPercent = '0';
                    }
                }
                if (data.sprinklerGainPercent !== undefined) {
                    const numValue = parseFloat(data.sprinklerGainPercent);
                    if (!isNaN(numValue)) {
                        this.sprinklerGainPercent = numValue.toFixed(5);
                    } else {
                        this.sprinklerGainPercent = '0';
                    }
                }

            } else {
            }

        } catch (error) {
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
        this._isUserEditing = false;
        
        // Update read-only state
        this.updateReadOnlyState();
        
        if (this.selectedVersionId !== 'draft') {
            this.isLoadingVersion = true;
            setTimeout(() => {
                this.isLoadingVersion = false;
            }, 2000);
        }
        
        await this.loadSavedData();
    }

    handleCellChange(event) {
        
        if (this._isInitializing || this.isLoadingVersion || this._isLoadingData) {
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
        
        this.autoSaveTimeout = setTimeout(() => {
            this.performAutoSave();
        }, 2000);
    }

    async performAutoSave() {
        // Don't auto-save if final version exists
        if (this.hasFinalVersion) {
            return;
        }
        
        const targetId = this.recordId || this.fallbackRecordId;
        
        if (!targetId) {
            return;
        }

        try {
            this.autoSaveStatus = 'saving';
            
            const payload = await this.saveSheet();
            const base64Payload = this.encodeData(payload);

            await autoSaveITMWorksheet({
                opportunityId: targetId,
                base64Data: base64Payload
            });

            this.autoSaveStatus = 'saved';

            setTimeout(() => {
                this.autoSaveStatus = '';
            }, 2000);
        } catch (error) {
            const errorMessage = error?.body?.message || error?.message || String(error);
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

    get showSaveButton() {
        return !this.hasFinalVersion;
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
                    if (savedRow.left.description !== undefined) row.left.description = savedRow.left.description;
                    row.left.quantity = savedRow.left.quantity; // Preserve as string if it was saved as one
                    row.left.hours = savedRow.left.hours;       // Preserve as string if it was saved as one
                    row.left.total = Math.round(( (parseFloat(savedRow.left.total) || 0) + Number.EPSILON) * 100) / 100;
                }

                // Restore right column
                if (savedRow.right) {
                    if (savedRow.right.description !== undefined) row.right.description = savedRow.right.description;
                    row.right.quantity = savedRow.right.quantity; // Preserve as string if it was saved as one
                    row.right.hours = savedRow.right.hours;       // Preserve as string if it was saved as one
                    row.right.total = Math.round(( (parseFloat(savedRow.right.total) || 0) + Number.EPSILON) * 100) / 100;
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
     * Handle description input change
     */
    handleDescriptionChange(event) {
        this.handleCellChange(event);
        
        const rowNumber = parseInt(event.target.dataset.row);
        const column = event.target.dataset.column;
        const section = event.target.dataset.section;
        const value = event.target.value;


        this.updateRowValue(section, rowNumber, column, 'description', value);
    }

    /**
     * Handle quantity input change
     */
    handleQuantityChange(event) {
        this.handleCellChange(event);
        
        const rowNumber = parseInt(event.target.dataset.row);
        const column = event.target.dataset.column;
        const section = event.target.dataset.section;
        const value = event.target.value; // Store as string to preserve decimal point while typing

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
        const value = event.target.value; // Store as string to preserve decimal point while typing

        this.updateRowValue(section, rowNumber, column, 'hours', value);
        this.calculateAllTotals();
    }

    /**
     * Handle blur to perform final rounding
     */
    handleBlur(event) {
        const rowNumber = parseInt(event.target.dataset.row);
        const column = event.target.dataset.column;
        const section = event.target.dataset.section;
        const field = event.target.dataset.field || (event.target.classList.contains('qty-input') ? 'quantity' : 'hours');
        const rawValue = parseFloat(event.target.value) || 0;
        const roundedValue = Math.round((rawValue + Number.EPSILON) * 100) / 100;

        if (section && rowNumber && column) {
            this.updateRowValue(section, rowNumber, column, field, roundedValue);
        } else if (event.target.classList.contains('gain-input')) {
            // Percentage in 0-1 scale (0.1 = 10%, 0.15 = 15%)
            // Validate: value must be between 0 and 1
            if (rawValue < 0 || rawValue > 1) {
                this.showToast('Error', 'Value should be between 0 and 1', 'error');
                // Reset to 0
                event.target.value = '0';
                if (column === 'left') {
                    this.alarmGainPercent = 0;
                } else {
                    this.sprinklerGainPercent = 0;
                }
                return;
            }
            // Round to 5 decimal places for percentage values
            const percentValue = Math.round((rawValue + Number.EPSILON) * 100000) / 100000;
            if (column === 'left') {
                this.alarmGainPercent = percentValue;
            } else {
                this.sprinklerGainPercent = percentValue;
            }
        }
        
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
            
            // Only recalculate total if quantity or hours changed
            if (field === 'quantity' || field === 'hours') {
                const q = Number(row[column].quantity) || 0;
                const h = Number(row[column].hours) || 0;
                const total = q * h;
                // Store as rounded number, but we'll use toFixed(2) for display in HTML where needed
                row[column].total = Math.round((total + Number.EPSILON) * 100) / 100;
            }
        }
    }

    /**
     * Handle gain percent change
     */
    handleGainPercentChange(event) {
        this.handleCellChange(event);
        
        const column = event.target.dataset.column;
        const value = event.target.value; // Store as string to preserve decimal point while typing

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

        // Section 1: Labor Factor totals
        this.totalAlarmLaborHours = this.sumColumn(this.laborFactorRows, 'left');
        this.totalSprinklerLaborHours = this.sumColumn(this.laborFactorRows, 'right');

        // Section 2: Equipment Factor totals
        this.totalAlarmEquipmentCost = this.sumColumn(this.equipmentFactorRows, 'left');
        this.totalSprinklerEquipmentCost = this.sumColumn(this.equipmentFactorRows, 'right');

        // Section 3: Rates - Auto-populate Labor Mix Rate quantities
        if (this.ratesRows.length > 0) {
            // First row (41) is Labor Mix Rate - should auto-populate from labor hours
            // Ensure we use a clean number for the mix rate row quantity
            const aHours = Number(this.totalAlarmLaborHours) || 0;
            this.ratesRows[0].left.quantity = aHours.toFixed(2);
            const alarmRateTotal = Number(this.ratesRows[0].left.quantity) * (Number(this.ratesRows[0].left.hours) || 0);
            this.ratesRows[0].left.total = Math.round((alarmRateTotal + Number.EPSILON) * 100) / 100;

            const sHours = Number(this.totalSprinklerLaborHours) || 0;
            this.ratesRows[0].right.quantity = sHours.toFixed(2);
            const sprinklerRateTotal = Number(this.ratesRows[0].right.quantity) * (Number(this.ratesRows[0].right.hours) || 0);
            this.ratesRows[0].right.total = Math.round((sprinklerRateTotal + Number.EPSILON) * 100) / 100;
        }

        // Calculate rates section totals
        const alarmRatesTotal = this.sumColumn(this.ratesRows, 'left');
        const sprinklerRatesTotal = this.sumColumn(this.ratesRows, 'right');

        // Subtotal = Equipment + Rates
        const aSub = Number(this.totalAlarmEquipmentCost) + alarmRatesTotal;
        this.alarmSubtotal = Math.round((aSub + Number.EPSILON) * 100) / 100;
        
        const sSub = Number(this.totalSprinklerEquipmentCost) + sprinklerRatesTotal;
        this.sprinklerSubtotal = Math.round((sSub + Number.EPSILON) * 100) / 100;

        // Total Quote = Subtotal * (1 + Gain%)
        // Gain percentage in 0-1 scale (0.15 = 15%)
        const aQuote = Number(this.alarmSubtotal) * (1 + (Number(this.alarmGainPercent) || 0));
        this.alarmTotalQuote = Math.round((aQuote + Number.EPSILON) * 100) / 100;
        
        const sQuote = Number(this.sprinklerSubtotal) * (1 + (Number(this.sprinklerGainPercent) || 0));
        this.sprinklerTotalQuote = Math.round((sQuote + Number.EPSILON) * 100) / 100;
    }

    /**
     * Sum column helper
     */
    sumColumn(rows, column) {
        let sum = 0;
        rows.forEach(row => {
            const val = parseFloat(row[column].total) || 0;
            sum += val;
        });
        return Number(Math.round((sum + Number.EPSILON) * 100) / 100);
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

        // Check if final version exists
        if (this.hasFinalVersion) {
            this.showToast('Error', 'Cannot save: A final version already exists for this worksheet.', 'error');
            return;
        }

        this.isSaving = true;

        try {

            // Get mark as final checkbox value
            const markAsFinalCheckbox = this.template.querySelector('.mark-as-final-checkbox');
            const markAsFinal = markAsFinalCheckbox ? markAsFinalCheckbox.checked : false;

            // Prepare data to save
            const worksheetData = {
                worksheetType: 'ITM',
                version: '1.0',
                isFinal: markAsFinal,
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

            // Save to Salesforce with markAsFinal flag
            await saveITMWorksheet({
                opportunityId: targetId,
                base64Data: base64Data,
                markAsFinal: markAsFinal
            });


            // If marked as final, update state
            if (markAsFinal) {
                this.hasFinalVersion = true;
                this.isReadOnly = true;
                this.showMarkAsFinalCheckbox = false;
            }

            // ⭐ Refresh version data AFTER save (so nextVersionNumber is updated)
            await this.loadVersionData();
            
            // Set draft as selected after save (if not final)
            if (!markAsFinal) {
                this.selectedVersionId = 'draft';
            }

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

            this.showToast('Success', markAsFinal ? 'ITM Bid Worksheet saved and marked as FINAL!' : 'ITM Bid Worksheet saved successfully!', 'success');

        } catch (error) {
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

    /**
     * Handle Download CSV for full worksheet
     */
    handleDownloadCSVFull() {
        const lines = [];
        const dateStr = new Date().toISOString().split('T')[0];
        
        lines.push('ITM BID WORKSHEET FULL EXPORT');
        lines.push(`DATE,${this.currentDate}`);
        lines.push('');

        lines.push('LABOR FACTOR');
        lines.push(this.convertToCSVContent(this.laborFactorRows, 'QTY,HOURS,TOTAL HOURS', 'labor'));
        lines.push('');

        lines.push('EQUIPMENT FACTOR');
        lines.push(this.convertToCSVContent(this.equipmentFactorRows, 'QTY,COST,TOTAL COST', 'equipment'));
        lines.push('');

        lines.push('RATES');
        lines.push(this.convertToCSVContent(this.ratesRows, 'HOURS,RATE,TOTAL', 'rates'));
        lines.push('');

        lines.push('SUMMARY');
        lines.push(this.convertToCSVSummary());

        this.downloadCSVFile(lines.join('\n'), `ITM_Worksheet_Full_${dateStr}.csv`);
    }

    downloadCSV(title, rows, colHeaders) {
        const csvContent = this.convertToCSVContent(rows, colHeaders);
        const dateStr = new Date().toISOString().split('T')[0];
        this.downloadCSVFile(csvContent, `ITM_Worksheet_${title}_${dateStr}.csv`);
    }

    convertToCSVContent(rows, colHeaders, sectionType) {
        const headers = [`FIRE ALARM,,,,,,FIRE SPRINKLER`, `DESCRIPTION,${colHeaders},,DESCRIPTION,${colHeaders}`];
        const lines = [headers[0], headers[1]];

        rows.forEach(row => {
            const rowValues = [];
            // Left
            rowValues.push(this.formatCSVValue(row.left.description));
            rowValues.push(row.left.quantity);
            rowValues.push(row.left.hours);
            rowValues.push(row.left.total);
            // Spacer
            rowValues.push('');
            // Right
            rowValues.push(this.formatCSVValue(row.right.description));
            rowValues.push(row.right.quantity);
            rowValues.push(row.right.hours);
            rowValues.push(row.right.total);
            
            lines.push(rowValues.join(','));
        });

        // Add total rows based on section type
        if (sectionType === 'labor') {
            const totalRow = [];
            totalRow.push('Total Alarm Labor Hours');
            totalRow.push('');
            totalRow.push('');
            totalRow.push(this.totalAlarmLaborHours);
            totalRow.push('');
            totalRow.push('Total Sprinkler Labor Hours');
            totalRow.push('');
            totalRow.push('');
            totalRow.push(this.totalSprinklerLaborHours);
            lines.push(totalRow.join(','));
        } else if (sectionType === 'equipment') {
            const totalRow = [];
            totalRow.push('Total Cost');
            totalRow.push('');
            totalRow.push('');
            totalRow.push(this.totalAlarmEquipmentCost);
            totalRow.push('');
            totalRow.push('Total Cost');
            totalRow.push('');
            totalRow.push('');
            totalRow.push(this.totalSprinklerEquipmentCost);
            lines.push(totalRow.join(','));
        }

        return lines.join('\n');
    }

    convertToCSVSummary() {
        const lines = [];
        lines.push('FIRE ALARM,,,FIRE SPRINKLER');
        lines.push(`SUBTOTAL,${this.alarmSubtotal},,SUBTOTAL,${this.sprinklerSubtotal}`);
        lines.push(`GAIN %,${this.alarmGainPercent}%,,GAIN %,${this.sprinklerGainPercent}%`);
        lines.push(`TOTAL QUOTE PRICE,${this.alarmTotalQuote},,TOTAL QUOTE PRICE,${this.sprinklerTotalQuote}`);
        return lines.join('\n');
    }

    formatCSVValue(val) {
        if (val === null || val === undefined) return '';
        let stringVal = String(val);
        stringVal = stringVal.replace(/"/g, '""');
        if (stringVal.includes(',') || stringVal.includes('\n') || stringVal.includes('"')) {
            stringVal = `"${stringVal}"`;
        }
        return stringVal;
    }

    downloadCSVFile(csvContent, filename) {
        // Use text/plain to avoid LWS security issues with text/csv
        const blob = new Blob([csvContent], { type: 'text/plain' });
        const link = document.createElement('a');
        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', filename);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            
            // Cleanup
            setTimeout(() => {
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            }, 100);
        }
    }
}