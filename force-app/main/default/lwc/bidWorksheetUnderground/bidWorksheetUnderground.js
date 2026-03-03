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
    @api isReadOnly = false; // Set by parent when final version exists (Underground tab only)

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
            
            
            // Only load if tableRows are initialized (metadata loaded)
            if (this.tableRows && this.tableRows.length > 0) {
                // Small delay to ensure DOM is stable
                setTimeout(() => {
                    this.loadSavedSheet();
                }, 100);
            } else {
            }
        } else {
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

    get sheetWrapperClass() {
        return 'estimate-table-wrapper' + (this.isReadOnly ? ' readonly-sheet' : '');
    }

    connectedCallback() {
        if (!this.recordId) {
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

            // ⭐ Set flag FIRST to prevent autosave during initialization
            this._isLoadingData = true;

            this.initializeDataFromMetadata(data);
            this.isLoading = false;
            // Reload saved data (if any) now that base rows are ready
            // Use setTimeout to ensure DOM is updated and tableRows are fully initialized
            setTimeout(() => {
                
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
                    this._isLoadingData = false;
                }
            }, 1500); // Give enough time for loadSavedSheet to complete

        } else if (error) {
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

        // Show a few sample rows
        const samples = [0, Math.floor(this.tableRows.length / 2), this.tableRows.length - 1];
        samples.forEach(idx => {
            if (this.tableRows[idx]) {
            }
        });
    }

    /**
     * Create row structure from metadata item
     */
    createRowFromMetadata(data, id) {
        const rowId = id !== null ? id : this.nextRowId++;

        return {
            id: rowId,
            excelRow: data.excelRow || null,  // ✅ From Apex wrapper
            left: {
                description: data.left.description || '',
                descriptionReadonly: false, // ⭐ Make editable
                size: data.left.size || '',
                sizeReadonly: false, // ⭐ Make editable
                amount: '',
                amountReadonly: false, // ⭐ Make editable
                unitPrice: (() => {
                    // Format percentage defaultUnitPrice to 5 decimals
                    if (data.left.unitPriceFieldType === 'Percentage' && data.left.defaultUnitPrice) {
                        const numValue = parseFloat(data.left.defaultUnitPrice);
                        if (!isNaN(numValue)) {
                            return numValue.toFixed(5);
                        }
                        return '0';
                    }
                    return data.left.defaultUnitPrice || '';
                })(),
                unitPriceReadonly: false,
                defaultUnitPrice: (() => {
                    // Format percentage defaultUnitPrice to 5 decimals
                    if (data.left.unitPriceFieldType === 'Percentage' && data.left.defaultUnitPrice) {
                        const numValue = parseFloat(data.left.defaultUnitPrice);
                        if (!isNaN(numValue)) {
                            return numValue.toFixed(5);
                        }
                        return '0';
                    }
                    return data.left.defaultUnitPrice || null;
                })(),
                gross: '',
                unitPriceFieldType: data.left.unitPriceFieldType || 'Currency',
                grossFieldType: data.left.grossFieldType || 'Currency',
                isCurrencyGross: (data.left.grossFieldType || 'Currency') === 'Currency',
                isTotalRow: data.left.isTotalRow || false,
                isCommentRow: false,
                descriptionClass: data.left.isIndent ? 'description-cell indent' : 'description-cell'
            },
            right: {
                description: data.right.description || '',
                descriptionReadonly: false, // ⭐ Make editable
                size: data.right.size || '',
                sizeReadonly: false, // ⭐ Make editable
                amount: '',
                amountReadonly: false, // ⭐ Make editable
                unitPrice: (() => {
                    // Format percentage defaultUnitPrice to 5 decimals
                    if (data.right.unitPriceFieldType === 'Percentage' && data.right.defaultUnitPrice) {
                        const numValue = parseFloat(data.right.defaultUnitPrice);
                        if (!isNaN(numValue)) {
                            return numValue.toFixed(5);
                        }
                        return '0';
                    }
                    return data.right.defaultUnitPrice || '';
                })(),
                unitPriceReadonly: false,
                defaultUnitPrice: (() => {
                    // Format percentage defaultUnitPrice to 5 decimals
                    if (data.right.unitPriceFieldType === 'Percentage' && data.right.defaultUnitPrice) {
                        const numValue = parseFloat(data.right.defaultUnitPrice);
                        if (!isNaN(numValue)) {
                            return numValue.toFixed(5);
                        }
                        return '0';
                    }
                    return data.right.defaultUnitPrice || null;
                })(),
                gross: '',
                unitPriceFieldType: data.right.unitPriceFieldType || 'Currency',
                grossFieldType: data.right.grossFieldType || 'Currency',
                isCurrencyGross: (data.right.grossFieldType || 'Currency') === 'Currency',
                isTotalRow: false,
                isCommentRow: data.right.isCommentRow || false,
                descriptionClass: data.right.isIndent ? 'description-cell indent' : 'description-cell'
            }
        };
    }

    calculateGross(amount, unitPrice, isPercentage = false) {
        const amountNum = parseFloat(amount) || 0;
        let priceNum = parseFloat(unitPrice) || 0;
        // Percentage fields use 0-1 scale (0.1 = 10%), no division by 100
        const gross = amountNum * priceNum;
        return gross > 0 ? gross.toFixed(2) : '';
    }

    handleCellChange(event) {
        const rowId = parseInt(event.target.dataset.row);
        const col = event.target.dataset.col;
        const field = event.target.dataset.field;
        const value = event.target.value;

        // ✅ VALIDATION - Moved to handleBlur for smoother typing, 
        // but we still want to block some characters if needed.
        // For now, we'll allow all typing and clean up on blur.


        const rowIndex = this.tableRows.findIndex(row => row.id === rowId);
        if (rowIndex !== -1) {
            // Validate percentage fields (0-1 scale)
            if (field === 'unitPrice' && this.tableRows[rowIndex][col].unitPriceFieldType === 'Percentage') {
                const numValue = parseFloat(value);
                if (value.trim() !== '' && (isNaN(numValue) || numValue < 0 || numValue > 1)) {
                    this.showToast('Error', 'Value should be between 0 and 1', 'error');
                    // Reset to 0
                    this.tableRows[rowIndex][col].unitPrice = '0';
                    event.target.value = '0';
                    this.tableRows = [...this.tableRows]; // Force re-render
                    return;
                }
            }
            
            // Update the model directly without re-assigning the whole array
            // This prevents full re-render and cursor jumping
            this.tableRows[rowIndex][col][field] = value;

            // When AMOUNT changes:
            if (field === 'amount') {
                // If amount is cleared, clear gross (but keep unitPrice as it's pre-populated)
                if (!value || value.trim() === '') {
                    this.tableRows[rowIndex][col].gross = '';
                }
            }

            // Calculate GROSS $ whenever amount or unitPrice changes
            if (field === 'amount' || field === 'unitPrice') {
                const isPercentage = this.tableRows[rowIndex][col].unitPriceFieldType === 'Percentage';
                this.tableRows[rowIndex][col].gross = this.calculateGross(
                    this.tableRows[rowIndex][col].amount,
                    this.tableRows[rowIndex][col].unitPrice,
                    isPercentage
                );
            }

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

    /**
     * Handle blur to perform final normalization and rounding
     */
    handleBlur(event) {
        const rowId = parseInt(event.target.dataset.row);
        const col = event.target.dataset.col;
        const field = event.target.dataset.field;
        let value = event.target.value;

        const rowIndex = this.tableRows.findIndex(row => row.id === rowId);
        if (rowIndex === -1) return;

        // If the field is empty, keep it empty instead of forcing a 0
        if (!value || value.trim() === '') {
            this.tableRows[rowIndex][col][field] = '';
            this.tableRows = [...this.tableRows];
            this.calculateSubTotal();
            return;
        }

        if (field === 'amount' || field === 'unitPrice') {
            let numValue = parseFloat(value) || 0;
            const isPercentage = this.tableRows[rowIndex][col].unitPriceFieldType === 'Percentage';
            
            // Validate percentage fields (0-1 scale)
            if (field === 'unitPrice' && isPercentage) {
                if (numValue < 0 || numValue > 1) {
                    this.showToast('Error', 'Value should be between 0 and 1', 'error');
                    // Reset to 0
                    value = '0';
                    this.tableRows[rowIndex][col][field] = value;
                    this.tableRows = [...this.tableRows]; // Force re-render
                    this.calculateSubTotal(); // Recalculate
                    return;
                }
            }
            
            // Block negative values
            if (numValue < 0) {
                this.showToast('Warning', 'Negative values not allowed', 'warning');
                value = '0';
            } else {
                // Round percentage fields to 5 decimals, others to 2 decimals
                const precision = (field === 'unitPrice' && isPercentage) ? 5 : 2;
                value = Math.round((numValue + Number.EPSILON) * Math.pow(10, precision)) / Math.pow(10, precision);
                // Format unitPrice based on type: percentage with 5 decimals, currency with 2 decimals
                if (field === 'unitPrice') {
                    value = isPercentage ? value.toFixed(5) : value.toFixed(2);
                } else if (value === 0 && field === 'amount') {
                    value = '';
                } else {
                    value = value.toString();
                }
            }
            
            this.tableRows[rowIndex][col][field] = value;
            
            // Recalculate gross with normalized values
            this.tableRows[rowIndex][col].gross = this.calculateGross(
                this.tableRows[rowIndex][col].amount,
                this.tableRows[rowIndex][col].unitPrice,
                isPercentage
            );
        }

        // Force a re-render on blur to ensure the input field shows the normalized value
        this.tableRows = [...this.tableRows];
        this.calculateSubTotal();
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

                resolve(sheetData);
            } catch (error) {
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
            return;
        }

        // Don't load if tableRows haven't been initialized from metadata yet
        if (!this.tableRows || this.tableRows.length === 0) {
            return;
        }

        // Don't set isLoading here as it might interfere with metadata loading
        try {
            let base64Data;
            
            // If versionIdToLoad is set, load that specific version
            if (this.versionIdToLoad && this.versionIdToLoad !== 'draft') {
                base64Data = await loadVersionById({ versionId: this.versionIdToLoad });
            } else {
                // Otherwise, load latest (autosave or most recent)
                base64Data = await loadLatestSheet({ opportunityId: this.recordId });
            }


            if (!base64Data) {
                return;
            }

            const jsonString = this.decodeData(base64Data);

            const savedState = JSON.parse(jsonString);

            if (savedState.sheet1) {
            }

            this.applyLoadedData(savedState);

        } catch (error) {
            // Don't show error toast if file doesn't exist (first time use)
            const errorMessage = error?.body?.message || error?.message || String(error);

            if (errorMessage.includes('not found') || errorMessage.includes('No ContentVersion') || errorMessage.includes('List has no rows')) {
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
            return;
        }

        // Don't apply if tableRows haven't been initialized from metadata yet
        if (!this.tableRows || this.tableRows.length === 0) {
            return;
        }

        // Set flag to prevent autosave during data loading
        this._isLoadingData = true;

        try {
            // Handle unified payload structure (from parent save) or direct sheet data
            const sheetData = data.sheet1 || data;

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
                            unitPrice: (() => {
                                const savedValue = (savedRow.left?.unitPrice !== undefined && savedRow.left?.unitPrice !== null)
                                    ? savedRow.left.unitPrice : existingRow.left.unitPrice;
                                if (!savedValue || savedValue === '') return savedValue;
                                const numValue = parseFloat(savedValue);
                                if (isNaN(numValue)) return savedValue;
                                // Format percentage unitPrice to 5 decimals, currency to 2 decimals
                                if (existingRow.left.unitPriceFieldType === 'Percentage') {
                                    return numValue.toFixed(5);
                                } else {
                                    return numValue.toFixed(2);
                                }
                            })(),
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
                            unitPrice: (() => {
                                const savedValue = (savedRow.right?.unitPrice !== undefined && savedRow.right?.unitPrice !== null)
                                    ? savedRow.right.unitPrice : existingRow.right.unitPrice;
                                if (!savedValue || savedValue === '') return savedValue;
                                const numValue = parseFloat(savedValue);
                                if (isNaN(numValue)) return savedValue;
                                // Format percentage unitPrice to 5 decimals, currency to 2 decimals
                                if (existingRow.right.unitPriceFieldType === 'Percentage') {
                                    return numValue.toFixed(5);
                                } else {
                                    return numValue.toFixed(2);
                                }
                            })(),
                            gross: (savedRow.right?.gross !== undefined && savedRow.right?.gross !== null)
                                ? savedRow.right.gross : existingRow.right.gross
                        }
                    };
                });

                this.tableRows = mergedRows;
                this.nextRowId = this.tableRows.length;


                // Force reactivity update
                this.tableRows = [...this.tableRows];
            } else {
            }

            // Recalculate to ensure totals match restored rows
            // Use setTimeout to ensure DOM has updated
            setTimeout(() => {
                this.calculateSubTotal();
                
                // Clear flag after a delay to allow DOM to settle and prevent autosave
                setTimeout(() => {
                    this._isLoadingData = false;
                }, 500);
            }, 50);

        } catch (err) {
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