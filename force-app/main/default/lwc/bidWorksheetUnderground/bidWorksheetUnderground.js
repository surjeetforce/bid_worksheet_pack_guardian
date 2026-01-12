import { LightningElement, api, track, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getSheet1Items from '@salesforce/apex/BidWorksheetUndergroundController.getSheet1Items';
import saveSheet from '@salesforce/apex/BidWorksheetUndergroundController.saveSheet';
import loadLatestSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestSheet';
import loadVersionById from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById';
import autoSaveSheet from '@salesforce/apex/BidWorksheetUndergroundController.autoSaveSheet';

export default class BidWorksheetUnderground extends LightningElement {
    @api recordId; // Opportunity Id
    @api sheetNumber = 1;
    
    _versionIdToLoad = null;
    _lastLoadedVersionId = null; // Track last loaded version to avoid reloading same version
    _isLoadingData = false; // Flag to prevent autosave during data loading
    
    @api
    get versionIdToLoad() {
        return this._versionIdToLoad;
    }
    
    set versionIdToLoad(value) {
        const oldValue = this._versionIdToLoad;
        // Normalize empty string to null for comparison
        const normalizedOldValue = oldValue === '' ? null : oldValue;
        const normalizedNewValue = value === '' ? null : value;
        
        this._versionIdToLoad = value;
        
        console.log('🔔 Sheet #1 versionIdToLoad changed:', { oldValue, newValue: value, lastLoaded: this._lastLoadedVersionId });
        
        // Always reload if:
        // 1. lastLoaded is null (first time load) - ALWAYS reload on first load
        // 2. OR value actually changed (normalized comparison)
        // 3. OR value is different from lastLoaded
        const isFirstLoad = this._lastLoadedVersionId === null;
        const valueChanged = normalizedNewValue !== normalizedOldValue;
        const isDifferentVersion = normalizedNewValue !== this._lastLoadedVersionId;
        
        // On first load, always reload regardless of valueChanged
        // Otherwise, reload if value changed AND it's a different version
        const shouldReload = isFirstLoad || (valueChanged && isDifferentVersion);
        
        if (shouldReload) {
            this._lastLoadedVersionId = normalizedNewValue;
            
            console.log('🔔 Sheet #1: Version changed, checking if tableRows are ready...', {
                hasTableRows: !!this.tableRows,
                tableRowsLength: this.tableRows?.length || 0
            });
            
            // Only load if tableRows are initialized (metadata loaded)
            if (this.tableRows && this.tableRows.length > 0) {
                console.log('✅ Sheet #1: tableRows ready, triggering loadSavedSheet()');
                // Small delay to ensure DOM is stable
                setTimeout(() => {
                    this.loadSavedSheet();
                }, 100);
            } else {
                console.log('⏳ Sheet #1: tableRows not ready yet, will load when metadata is ready');
            }
        } else {
            console.log('⏸️ Sheet #1: Version not changed or already loaded, skipping reload', {
                valueChanged,
                isFirstLoad,
                isDifferentVersion
            });
        }
    }

    @track tableRows = [];
    @track subTotal = '0.00';
    @track revisionDate = '5/4/00';
    @track isSaving = false;
    @track isLoading = true;

    nextRowId = 0;

    calculationTimeout = null;
    autoSaveTimeout = null;

    get currentDate() {
        const today = new Date();
        return today.toLocaleDateString('en-US');
    }

    connectedCallback() {
        if (!this.recordId) {
            console.error('❌ No recordId provided to Underground Sheet 1');
            this.showToast('Error', 'Record ID is required', 'error');
            this.isLoading = false;
            return;
        }
    }


    /**
     * Wire to Apex method to fetch metadata
     */
    @wire(getSheet1Items)
    wiredItems({ error, data }) {
        if (data) {
            console.log('📍 Sheet #1 metadata loaded successfully');
            console.log('📍 Number of rows:', data.length);
            console.log('Sample data:', JSON.stringify(data[0], null, 2));

            // ⭐ Set flag FIRST to prevent autosave during initialization
            this._isLoadingData = true;

            this.initializeDataFromMetadata(data);
            this.isLoading = false;
            // Reload saved data (if any) now that base rows are ready
            // Use setTimeout to ensure DOM is updated and tableRows are fully initialized
            setTimeout(() => {
                console.log('📍 Sheet #1: Attempting to load saved data...');
                console.log('📍 Sheet #1: Current versionIdToLoad:', this._versionIdToLoad);
                console.log('📍 Sheet #1: tableRows ready:', this.tableRows && this.tableRows.length > 0);
                
                // Always ensure versionIdToLoad is set - if null/empty, set to 'draft'
                // This ensures the setter fires and loads data
                const versionToLoad = (this._versionIdToLoad && this._versionIdToLoad !== '') 
                    ? this._versionIdToLoad 
                    : 'draft';
                
                // Reset to force load and trigger setter
                this._lastLoadedVersionId = null;
                this.versionIdToLoad = versionToLoad;
            }, 100);
            
            // Clear flag after initialization completes (loadSavedSheet sets its own flag, so this is a backup)
            // The flag will be cleared by loadSavedSheet's applyLoadedData, but we set a timeout as backup
            setTimeout(() => {
                if (this._isLoadingData) {
                    console.log('📍 Sheet #1: Clearing _isLoadingData flag after initialization');
                    this._isLoadingData = false;
                }
            }, 1500); // Give enough time for loadSavedSheet to complete

        } else if (error) {
            console.error('❌ Sheet #1: Error loading metadata:', error);
            this.showToast('Error', 'Failed to load sheet configuration: ' + (error.body ? error.body.message : error.message), 'error');
            this.isLoading = false;
            this._isLoadingData = false; // Clear flag on error
        }
    }

    /**
     * Initialize table rows from Custom Metadata
     */
    initializeDataFromMetadata(metadataItems) {
        this.tableRows = metadataItems.map((item, index) =>
            this.createRowFromMetadata(item, index)
        );
        this.nextRowId = this.tableRows.length;

        // ⭐ ADDED: Debug logging to verify row mapping
        console.log('=== Underground Sheet 1 (Section 1) Row Mapping ===');
        console.log('Total rows loaded:', this.tableRows.length);
        console.log('Row range:',
            this.tableRows[0]?.excelRow,
            'to',
            this.tableRows[this.tableRows.length - 1]?.excelRow
        );

        // Show a few sample rows
        const samples = [0, Math.floor(this.tableRows.length / 2), this.tableRows.length - 1];
        samples.forEach(idx => {
            if (this.tableRows[idx]) {
                console.log(`Sample row ${idx}:`, {
                    excelRow: this.tableRows[idx].excelRow,
                    leftDesc: this.tableRows[idx].left.description?.substring(0, 30)
                });
            }
        });
    }

    /**
     * Create row structure from metadata item
     */
    createRowFromMetadata(data, id) {
        const rowId = id !== null ? id : this.nextRowId++;

        // Determine if fields should be readonly based on description
        const leftHasDescription = !!(data.left.description && data.left.description.trim());
        const rightHasDescription = !!(data.right.description && data.right.description.trim());

        return {
            id: rowId,
            excelRow: data.excelRow || null,  // ✅ From Apex wrapper
            left: {
                description: data.left.description || '',
                descriptionReadonly: !!data.left.description,
                size: data.left.size || '',  // ✅ FIXED: Was missing
                sizeReadonly: true,
                amount: '',  // Empty - user enters
                amountReadonly: !leftHasDescription,
                unitPrice: data.left.defaultUnitPrice || '',  // ✅ Pre-populate from defaultUnitPrice
                unitPriceReadonly: false,  // ✅ ALL unit prices are editable
                defaultUnitPrice: data.left.defaultUnitPrice || null,  // ✅ STORE for reference
                gross: '',
                unitPriceFieldType: data.left.unitPriceFieldType || 'Currency',
                grossFieldType: data.left.grossFieldType || 'Currency',
                isTotalRow: data.left.isTotalRow || false,
                isCommentRow: false,
                descriptionClass: data.left.isIndent ? 'description-cell indent' : 'description-cell'
            },
            right: {
                description: data.right.description || '',
                descriptionReadonly: !!data.right.description,
                size: data.right.size || '',  // ✅ FIXED: Was missing
                sizeReadonly: true,
                amount: '',  // Empty - user enters
                amountReadonly: !rightHasDescription,
                unitPrice: data.right.defaultUnitPrice || '',  // ✅ Pre-populate from defaultUnitPrice
                unitPriceReadonly: false,  // ✅ ALL unit prices are editable
                defaultUnitPrice: data.right.defaultUnitPrice || null,  // ✅ STORE for reference
                gross: '',
                unitPriceFieldType: data.right.unitPriceFieldType || 'Currency',
                grossFieldType: data.right.grossFieldType || 'Currency',
                isTotalRow: false,
                isCommentRow: data.right.isCommentRow || false,
                descriptionClass: data.right.isIndent ? 'description-cell indent' : 'description-cell'
            }
        };
    }

    calculateGross(amount, unitPrice) {
        const amountNum = parseFloat(amount) || 0;
        const priceNum = parseFloat(unitPrice) || 0;
        const gross = amountNum * priceNum;
        return gross > 0 ? gross.toFixed(2) : '';
    }

    handleCellChange(event) {
        const rowId = parseInt(event.target.dataset.row);
        const col = event.target.dataset.col;
        const field = event.target.dataset.field;
        const value = event.target.value;

        // ✅ VALIDATION
        if (field === 'amount' || field === 'unitPrice') {
            const numValue = parseFloat(value);

            // Block negative values
            if (numValue < 0) {
                this.showToast('Warning', 'Negative values not allowed', 'warning');
                event.target.value = ''; // Clear the field
                return;
            }
        }

        console.log(`Cell changed: Row ${rowId}, Col ${col}, Field ${field}, Value: ${value}`);

        const rowIndex = this.tableRows.findIndex(row => row.id === rowId);
        if (rowIndex !== -1) {
            const updatedRow = { ...this.tableRows[rowIndex] };
            updatedRow[col] = { ...updatedRow[col], [field]: value };

            // When AMOUNT changes:
            if (field === 'amount') {
                // If amount is cleared, clear gross (but keep unitPrice as it's pre-populated)
                if (!value || value.trim() === '') {
                    updatedRow[col].gross = '';
                }
            }

            // Calculate GROSS $ whenever amount or unitPrice changes
            if (field === 'amount' || field === 'unitPrice') {
                updatedRow[col].gross = this.calculateGross(
                    updatedRow[col].amount,
                    updatedRow[col].unitPrice
                );
                console.log(`Calculated gross: ${updatedRow[col].gross}`);
            }

            // Update the array
            this.tableRows = [
                ...this.tableRows.slice(0, rowIndex),
                updatedRow,
                ...this.tableRows.slice(rowIndex + 1)
            ];

            // Recalculate subtotal
            if (this.calculationTimeout) {
                clearTimeout(this.calculationTimeout);
            }
            this.calculationTimeout = setTimeout(() => {
                this.calculateSubTotal();
            }, 300); // Wait 300ms after last keystroke
            
            // Notify parent for auto-save (only if not loading data)
            if (!this._isLoadingData) {
                this.notifyParentForAutoSave();
            }
        }
    }

    notifyParentForAutoSave() {
        const event = new CustomEvent('cellchange', {
            detail: {
                sheetNumber: this.sheetNumber
            }
        });
        this.dispatchEvent(event);
    }

    calculateSubTotal() {
        let total = 0;

        this.tableRows.forEach(row => {
            const leftGross = parseFloat(row.left.gross) || 0;
            const rightGross = parseFloat(row.right.gross) || 0;
            total += leftGross + rightGross;
        });

        this.subTotal = total.toFixed(2);
        console.log(`Sheet #1 Subtotal: $${this.subTotal}`);

        this.notifyParent();
    }

    notifyParent() {
        const event = new CustomEvent('sheetupdate', {
            detail: {
                sheetNumber: this.sheetNumber,
                subtotal: this.subTotal
            }
        });
        this.dispatchEvent(event);
    }

    /**
     * Create empty row (for manually added rows) - kept for potential future use
     */
    createRow(data = null, id = null) {
        const rowId = id !== null ? id : this.nextRowId++;

        return {
            id: rowId,
            left: {
                description: '',
                descriptionReadonly: false,
                size: '',
                sizeReadonly: false,
                amount: '',
                unitPrice: '',
                defaultUnitPrice: null, // No default price for manual rows
                gross: '',
                descriptionClass: 'description-cell'
            },
            right: {
                description: '',
                descriptionReadonly: false,
                size: '',
                sizeReadonly: false,
                amount: '',
                unitPrice: '',
                defaultUnitPrice: null, // No default price for manual rows
                gross: '',
                descriptionClass: 'description-cell'
            }
        };
    }

    @api
    async saveSheet() {
        return new Promise((resolve, reject) => {
            try {
                const sheetData = this.collectFormData();
                console.log('💾 Sheet #1 Data collected:', Object.keys(sheetData));
                console.log('💾 Sheet #1 LineItems count:', sheetData.lineItems?.length);

                resolve(sheetData);
            } catch (error) {
                console.error('❌ Sheet #1 saveSheet error:', error);
                reject(error);
            }
        });
    }

    handleSave() {
        this.isSaving = true;

        this.saveSheet()
            .then(async (sheetData) => {
                const base64Payload = this.encodeData(sheetData);
                await saveSheet({
                    opportunityId: this.recordId,
                    base64Data: base64Payload
                });
                this.showToast('Success', 'Estimate sheet saved to Opportunity Files', 'success');
            })
            .catch(error => {
                const message = error && error.body && error.body.message ? error.body.message : error.message;
                this.logError('Save failed', error);
                this.showToast('Error', 'Error saving sheet: ' + message, 'error');
            })
            .finally(() => {
                this.isSaving = false;
            });
    }

    // ⭐ FIXED: Match Sheet 2's structure exactly
    collectFormData() {
        return {
            sheetNumber: this.sheetNumber,
            subTotal: this.subTotal,
            revisionDate: this.revisionDate,
            lineItems: this.tableRows.map(row => ({
                id: row.id,
                excelRow: row.excelRow,
                left: { ...row.left },
                right: { ...row.right }
            }))
        };
    }

    async loadSavedSheet() {
        if (!this.recordId) {
            console.log('❌ [LOAD Sheet #1] No recordId, skipping load');
            return;
        }

        // Don't load if tableRows haven't been initialized from metadata yet
        if (!this.tableRows || this.tableRows.length === 0) {
            console.log('⚠️ [LOAD Sheet #1] TableRows not initialized yet, will load after metadata');
            return;
        }

        // Don't set isLoading here as it might interfere with metadata loading
        try {
            let base64Data;
            
            // If versionIdToLoad is set, load that specific version
            if (this.versionIdToLoad && this.versionIdToLoad !== 'draft') {
                console.log('🔍 [LOAD Sheet #1] Loading specific version:', this.versionIdToLoad);
                base64Data = await loadVersionById({ versionId: this.versionIdToLoad });
            } else {
                // Otherwise, load latest (autosave or most recent)
                console.log('🔍 [LOAD Sheet #1] Starting load for Opportunity:', this.recordId);
                base64Data = await loadLatestSheet({ opportunityId: this.recordId });
            }

            console.log('🔍 [LOAD Sheet #1] Received base64Data:', base64Data ? 'YES (length: ' + base64Data.length + ')' : 'NO');

            if (!base64Data) {
                console.log('⚠️ [LOAD Sheet #1] No saved data found - using default metadata values');
                return;
            }

            const jsonString = this.decodeData(base64Data);
            console.log('🔍 [LOAD Sheet #1] Decoded JSON length:', jsonString.length);

            const savedState = JSON.parse(jsonString);

            if (savedState.sheet1) {
                console.log('🔍 [LOAD Sheet #1] Sheet1 keys:', Object.keys(savedState.sheet1));
                console.log('🔍 [LOAD Sheet #1] Sheet1 has lineItems?', !!savedState.sheet1.lineItems);
                console.log('🔍 [LOAD Sheet #1] Sheet1 lineItems count:', savedState.sheet1.lineItems?.length);
            }

            this.applyLoadedData(savedState);
            console.log('✅ [LOAD Sheet #1] Loaded saved Sheet #1 state from file');

        } catch (error) {
            // Don't show error toast if file doesn't exist (first time use)
            const errorMessage = error?.body?.message || error?.message || String(error);
            console.error('❌ [LOAD Sheet #1] Error details:', {
                message: errorMessage,
                stack: error?.stack,
                body: error?.body
            });

            if (errorMessage.includes('not found') || errorMessage.includes('No ContentVersion') || errorMessage.includes('List has no rows')) {
                console.log('⚠️ [LOAD Sheet #1] No saved file found yet (first time use) - using default metadata values');
                return;
            }
            this.logError('Load saved sheet failed', error);
            // Only show error toast for actual errors, not missing files
            if (!errorMessage.includes('not found') && !errorMessage.includes('No ContentVersion')) {
                this.showToast('Error', 'Failed to load saved sheet: ' + errorMessage, 'error');
            }
        }
    }

    // ⭐ FIXED: Use same merging approach as Sheet 2
    applyLoadedData(data) {
        if (!data) {
            console.log('❌ [APPLY Sheet #1] No data provided to applyLoadedData');
            return;
        }

        // Don't apply if tableRows haven't been initialized from metadata yet
        if (!this.tableRows || this.tableRows.length === 0) {
            console.log('❌ [APPLY Sheet #1] TableRows not initialized yet, skipping data application');
            return;
        }

        // Set flag to prevent autosave during data loading
        this._isLoadingData = true;

        try {
            // Handle unified payload structure (from parent save) or direct sheet data
            const sheetData = data.sheet1 || data;
            console.log('🔧 [APPLY Sheet #1] Applying Sheet #1 data');
            console.log('🔧 [APPLY Sheet #1] Data has sheet1?', !!data.sheet1);
            console.log('🔧 [APPLY Sheet #1] Extracted sheetData keys:', Object.keys(sheetData));
            console.log('🔧 [APPLY Sheet #1] Has lineItems?', !!sheetData.lineItems);
            console.log('🔧 [APPLY Sheet #1] LineItems count:', sheetData.lineItems?.length);

            // Helper function to safely update values
            const updateIfExists = (savedValue, currentValue) => {
                return (savedValue !== undefined && savedValue !== null) ? savedValue : currentValue;
            };

            // Only update if value exists in saved data (preserve existing values if not present)
            this.sheetNumber = updateIfExists(sheetData.sheetNumber, this.sheetNumber);
            this.revisionDate = updateIfExists(sheetData.revisionDate, this.revisionDate);
            this.subTotal = updateIfExists(sheetData.subTotal, this.subTotal);

            // Only inflate rows if we have lineItems
            if (sheetData.lineItems && Array.isArray(sheetData.lineItems) && sheetData.lineItems.length > 0) {
                console.log('🔧 [APPLY Sheet #1] Merging saved lineItems with metadata rows');
                console.log('🔧 [APPLY Sheet #1] Current tableRows count:', this.tableRows.length);
                console.log('🔧 [APPLY Sheet #1] Saved lineItems count:', sheetData.lineItems.length);

                // ⭐ NEW: Simple merging approach (same as Sheet 2)
                const mergedRows = this.tableRows.map((existingRow, index) => {
                    const savedRow = sheetData.lineItems[index];
                    if (!savedRow) {
                        return existingRow;
                    }

                    return {
                        ...existingRow,
                        id: savedRow.id !== undefined ? savedRow.id : existingRow.id,
                        left: {
                            ...existingRow.left,
                            description: (savedRow.left?.description !== undefined && savedRow.left?.description !== null)
                                ? savedRow.left.description : existingRow.left.description,
                            size: (savedRow.left?.size !== undefined && savedRow.left?.size !== null)
                                ? savedRow.left.size : existingRow.left.size,
                            amount: (savedRow.left?.amount !== undefined && savedRow.left?.amount !== null)
                                ? savedRow.left.amount : existingRow.left.amount,
                            unitPrice: (savedRow.left?.unitPrice !== undefined && savedRow.left?.unitPrice !== null)
                                ? savedRow.left.unitPrice : existingRow.left.unitPrice,
                            gross: (savedRow.left?.gross !== undefined && savedRow.left?.gross !== null)
                                ? savedRow.left.gross : existingRow.left.gross
                        },
                        right: {
                            ...existingRow.right,
                            description: (savedRow.right?.description !== undefined && savedRow.right?.description !== null)
                                ? savedRow.right.description : existingRow.right.description,
                            size: (savedRow.right?.size !== undefined && savedRow.right?.size !== null)
                                ? savedRow.right.size : existingRow.right.size,
                            amount: (savedRow.right?.amount !== undefined && savedRow.right?.amount !== null)
                                ? savedRow.right.amount : existingRow.right.amount,
                            unitPrice: (savedRow.right?.unitPrice !== undefined && savedRow.right?.unitPrice !== null)
                                ? savedRow.right.unitPrice : existingRow.right.unitPrice,
                            gross: (savedRow.right?.gross !== undefined && savedRow.right?.gross !== null)
                                ? savedRow.right.gross : existingRow.right.gross
                        }
                    };
                });

                this.tableRows = mergedRows;
                this.nextRowId = this.tableRows.length;

                console.log('🔧 [APPLY Sheet #1] Merged rows successfully');
                console.log('🔧 [APPLY Sheet #1] Sample merged row 0:', JSON.stringify(this.tableRows[0], null, 2));

                // Force reactivity update
                this.tableRows = [...this.tableRows];
            } else {
                console.log('⚠️ [APPLY Sheet #1] No lineItems found in saved data, keeping metadata rows');
            }

            // Recalculate to ensure totals match restored rows
            // Use setTimeout to ensure DOM has updated
            setTimeout(() => {
                console.log('🔧 [APPLY Sheet #1] Triggering recalculation...');
                this.calculateSubTotal();
                console.log('✅ [APPLY Sheet #1] Applied loaded Sheet #1 data successfully');
                
                // Clear flag after a delay to allow DOM to settle and prevent autosave
                setTimeout(() => {
                    this._isLoadingData = false;
                }, 500);
            }, 50);

        } catch (err) {
            console.error('❌ [APPLY Sheet #1] Error applying loaded data:', err);
            this.logError('Apply loaded data failed', err);
            this._isLoadingData = false; // Clear flag on error
            this.showToast('Error', 'Failed to apply loaded sheet data: ' + err.message, 'error');
        }
    }

    encodeData(data) {
        try {
            const json = JSON.stringify(data);
            return btoa(unescape(encodeURIComponent(json)));
        } catch (err) {
            this.logError('Encode data failed', err);
            throw err;
        }
    }

    decodeData(base64Data) {
        try {
            return decodeURIComponent(escape(atob(base64Data)));
        } catch (err) {
            this.logError('Decode data failed', err);
            throw err;
        }
    }

    logError(context, error) {
        const safeMsg = error && error.body && error.body.message ? error.body.message : error && error.message ? error.message : String(error);
        console.error(`❌ ${context}:`, safeMsg, error);
    }

    showToast(title, message, variant) {
        const event = new ShowToastEvent({
            title: title,
            message: message,
            variant: variant
        });
        this.dispatchEvent(event);
    }
}