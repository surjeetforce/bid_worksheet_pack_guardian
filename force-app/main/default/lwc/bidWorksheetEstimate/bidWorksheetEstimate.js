import { LightningElement, api, track, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getEstimateSheetItems from '@salesforce/apex/BidWorksheetUndergroundController.getEstimateSheetItems';
import loadEstimateSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadEstimateSheet';
import loadLatestEstimateSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestEstimateSheet';
import loadVersionById_Estimate from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById_Estimate';

export default class BidWorksheetEstimate extends LightningElement {
    @api recordId;

    @track section1Rows = [];
    @track section2Rows = [];
    @track section3Rows = [];

    @track section1Subtotal = '0.00';
    @track section2Subtotal = '0.00';
    @track section3Subtotal = '0.00';
    @track grandTotal = '0.00';

    @track revisionDate = '5/4/00';
    @track isLoading = true;
    @track activeSections = ['section1'];

    nextRowId = 0;

    // Version control properties
    _versionIdToLoad = null;
    _lastLoadedVersionId = null;
    _isLoadingData = false;

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
        
        console.log(`📍 [Estimate] versionIdToLoad changed from ${oldValue} to ${value}, lastLoaded: ${this._lastLoadedVersionId}`);
        
        // Check if rows are initialized
        const rowsReady = this.section1Rows.length > 0 || this.section2Rows.length > 0 || this.section3Rows.length > 0;
        
        if (!rowsReady) {
            console.log('⏳ [Estimate] Rows not ready yet, will load when metadata is ready');
            return;
        }
        
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
            console.log(`📍 [Estimate] Version changed, triggering loadSavedData()`, {
                valueChanged,
                isFirstLoad,
                isDifferentVersion
            });
            // Don't reload if user is actively editing
            if (!this._isUserEditing) {
                this.loadSavedData();
            } else {
                console.log('📍 [Estimate] User is editing, skipping load to prevent data loss');
            }
        } else {
            console.log('⏸️ [Estimate] Version not changed or already loaded, skipping reload', {
                valueChanged,
                isFirstLoad,
                isDifferentVersion
            });
        }
    }

    // Flag to track if user is actively editing
    _isUserEditing = false;
    _editingTimeout = null;
    // Track the last edited cell to skip recalculation for it
    _lastEditedCell = null;

    // ========================================
    // WHOLE NUMBER FIELD CONFIGURATION
    // ========================================
    // Format: { rowNumber: { side: { field: true } } }
    // Example: { 182: { left: { gross: true } } } means row 182, left side, gross field is whole number
    // To add more fields, just add entries here:
    static WHOLE_NUMBER_FIELDS = {
        182: { left: { gross: true } }, // TOTAL LABOR HRS. - gross field is whole number
        184: { left: { quantity: true } }, // LABOR (FM+ 7TH PERIOD) - quantity field is whole number
        // Add more rows here as needed:
        // 185: { right: { unitPrice: true } }, // Example: Another row
    };

    // ========================================
    // EDITABLE FIELD OVERRIDE CONFIGURATION
    // ========================================
    // Format: { rowNumber: { side: { field: true } } }
    // Example: { 186: { left: { quantity: true } } } means row 186, left side, quantity field is editable (overrides calculated readonly)
    // Example: { 187: { left: { size: true } } } means row 187, left side, size field is editable
    // To add more editable fields, just add entries here:
    static EDITABLE_FIELD_OVERRIDES = {
        186: { left: { quantity: true } }, // BIM - quantity field is editable (overrides calculated readonly)
        187: { left: { size: true } }, // FABRICATION QUARTER HOUR PER - size field is editable
        // Add more rows here as needed:
        // 188: { right: { quantity: true, size: true } }, // Example: Another row with multiple editable fields
    };

    get currentDate() {
        const today = new Date();
        return today.toLocaleDateString('en-US');
    }

    connectedCallback() {
        if (!this.recordId) {
            this.recordId = '006VF00000I9RJaYAN';
        }
        // Don't load saved data here - wait for metadata to load first
    }

    async loadSavedData() {
        if (!this.recordId) {
            console.log('❌ [LOAD Estimate] No recordId, skipping load');
            return;
        }

        // Don't load if rows haven't been initialized from metadata yet
        if (this.section1Rows.length === 0 && this.section2Rows.length === 0 && this.section3Rows.length === 0) {
            console.log('⚠️ [LOAD Estimate] Rows not initialized yet, will load after metadata');
            return;
        }

        // Don't load if user is actively editing
        if (this._isUserEditing) {
            console.log('📍 [LOAD Estimate] User is editing, skipping load to prevent data loss');
            return;
        }

        // Set loading flag to prevent autosave during load
        this._isLoadingData = true;

        try {
            let base64Data;
            
            // If versionIdToLoad is set, load that specific version
            if (this.versionIdToLoad && this.versionIdToLoad !== 'draft') {
                console.log('🔍 [LOAD Estimate] Loading specific version:', this.versionIdToLoad);
                base64Data = await loadVersionById_Estimate({ versionId: this.versionIdToLoad });
                this._lastLoadedVersionId = this.versionIdToLoad;
            } else {
                // Otherwise, load latest (autosave or most recent)
                console.log('🔍 [LOAD Estimate] Loading latest (draft)');
                base64Data = await loadLatestEstimateSheet({ opportunityId: this.recordId });
                this._lastLoadedVersionId = 'draft';
            }

            if (!base64Data) {
                console.log('⚠️ [LOAD Estimate] No saved data found - using defaults');
                this._isLoadingData = false;
                return;
            }

            // Decode base64 data
            const jsonString = this.decodeData(base64Data);
            const data = JSON.parse(jsonString);

            console.log('✅ [LOAD Estimate] Loaded estimate data');
            this.applySavedData(data);
        } catch (error) {
            const errorMessage = error?.body?.message || error?.message || String(error);
            if (errorMessage.includes('not found') || errorMessage.includes('No ContentVersion') || errorMessage.includes('List has no rows')) {
                console.log('⚠️ [LOAD Estimate] No saved file found yet (first time use)');
            } else {
                console.error('❌ [LOAD Estimate] Error loading estimate data:', error);
            }
        } finally {
            // Clear loading flag after a delay to allow DOM to settle
            setTimeout(() => {
                this._isLoadingData = false;
            }, 500);
        }
    }

    decodeData(base64Data) {
        // Decode base64 string to get the JSON
        const binaryString = atob(base64Data);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        const decoder = new TextDecoder('utf-8');
        return decoder.decode(bytes);
    }

    applySavedData(data) {
        // Set loading flag during data application
        this._isLoadingData = true;

        if (data.section1) {
            this.restoreSectionData(this.section1Rows, data.section1);
        }
        if (data.section2) {
            this.restoreSectionData(this.section2Rows, data.section2);
        }
        if (data.section3) {
            this.restoreSectionData(this.section3Rows, data.section3);
        }

        this.calculateTotals();

        console.log('Saved data restored successfully');
        
        // Clear loading flag after a delay
        setTimeout(() => {
            this._isLoadingData = false;
        }, 500);
    }

    restoreSectionData(rows, savedItems) {
        savedItems.forEach(item => {
            const row = rows.find(r =>
                r.excelRow === item.excelRow ||
                (rows.indexOf(r) === item.rowNumber - 1)
            );

            if (row) {
                const side = item.column.toLowerCase();

                if (row[side]) {
                    // Restore field values
                    if (item.size !== undefined && item.size !== null) {
                        row[side].size = item.size || '';
                    }
                    row[side].quantity = item.quantity || '';
                    row[side].unitPrice = item.unitPrice || '';
                    row[side].gross = item.gross || '';

                    if (!row[side].descriptionReadonly && item.description) {
                        row[side].description = item.description;

                        // ⭐ Update readonly states based on description
                        const hasDescription = item.description && item.description.trim() !== '';
                        if (!row[side].isTotalRow) {
                            row[side].quantityReadonly = !hasDescription;
                            row[side].unitPriceReadonly = !hasDescription;
                            row[side].grossReadonly = !hasDescription;

                            // Update class properties to reflect readonly state
                            row[side].quantityClass = !hasDescription ? 'readonly-cell' : '';
                            row[side].unitPriceClass = !hasDescription ? 'col-unit readonly-cell' : 'col-unit';
                            row[side].descriptionClass = hasDescription ? 'description-cell readonly-cell' : 'description-cell';
                        }
                    } else {
                        // Update class properties even if description didn't change
                        const hasDescription = !!(row[side].description && row[side].description.trim());
                        const isTotalRow = row[side].isTotalRow || false;
                        const isReadonly = isTotalRow || !hasDescription;

                        row[side].quantityClass = isReadonly ? 'readonly-cell' : '';
                        row[side].unitPriceClass = isReadonly ? 'col-unit readonly-cell' : 'col-unit';
                    }
                }
            }
        });
    }

    @wire(getEstimateSheetItems)
    wiredItems({ error, data }) {
        if (data) {
            console.log('Estimate metadata loaded:', data.length, 'items');
            // Set loading flag first to prevent autosave during initialization
            this._isLoadingData = true;
            
            this.initializeDataFromMetadata(data);
            this.isLoading = false;
            
            // Load saved data after metadata is loaded and rows are initialized
            setTimeout(() => {
                console.log('📍 Estimate: Attempting to load saved data...');
                console.log('📍 Estimate: Current versionIdToLoad:', this._versionIdToLoad);
                console.log('📍 Estimate: Rows ready:', this.section1Rows.length > 0 || this.section2Rows.length > 0 || this.section3Rows.length > 0);
                
                // Always ensure versionIdToLoad is set - if null/empty, set to 'draft'
                // This ensures the setter fires and loads data
                const versionToLoad = (this._versionIdToLoad && this._versionIdToLoad !== '') 
                    ? this._versionIdToLoad 
                    : 'draft';
                
                // Reset to force load and trigger setter
                this._lastLoadedVersionId = null;
                this.versionIdToLoad = versionToLoad;
                
                // Clear flag after initialization completes
                setTimeout(() => {
                    this._isLoadingData = false;
                }, 1500);
            }, 100);
        } else if (error) {
            console.error('Error loading metadata:', error);
            this.showToast('Error', 'Failed to load estimate data', 'error');
            this.isLoading = false;
            this._isLoadingData = false;
        }
    }

    initializeDataFromMetadata(metadataItems) {
        const section1Items = [];
        const section2Items = [];
        const section3Items = [];

        console.log('metadataItems :- ', metadataItems);

        metadataItems.forEach((item, index) => {
            const section = item.section || this.inferSection(item);

            if (section === 1) {
                section1Items.push(this.createRowFromMetadata(item, index));
            } else if (section === 2) {
                section2Items.push(this.createRowFromMetadata(item, index));
            } else if (section === 3) {
                section3Items.push(this.createRowFromMetadata(item, index));
            }
        });

        this.section1Rows = section1Items;
        this.section2Rows = section2Items;
        this.section3Rows = section3Items;

        console.log('section3Items :- ', section3Items);

        console.log(`Initialized sections: S1=${section1Items.length}, S2=${section2Items.length}, S3=${section3Items.length}`);

        // Run calculations to set readonly states for calculated rows (like row 178)
        this.calculateTotals();
    }

    inferSection(item) {
        return 1;
    }

    createRowFromMetadata(data, id) {
        const leftDescEmpty = !data.left.description || data.left.description.trim() === '';
        const rightDescEmpty = !data.right.description || data.right.description.trim() === '';

        // Check if this is a total/calculated row
        const leftIsTotalOrCalculated = data.left.isTotalRow;
        const rightIsTotalOrCalculated = data.right.isTotalRow;

        console.log('Row', data.excelRow, 'Right desc:', data.right.description,
            'isEmpty:', rightDescEmpty,
            'isTotalRow:', rightIsTotalOrCalculated,
            'unitPriceReadonly will be:', (rightIsTotalOrCalculated || rightDescEmpty));        // Check if descriptions are empty

        // Check whole number configuration for this row
        const rowWholeNumberConfig = BidWorksheetEstimate.WHOLE_NUMBER_FIELDS[data.excelRow] || {};
        const leftWholeNumber = rowWholeNumberConfig.left || {};
        const rightWholeNumber = rowWholeNumberConfig.right || {};

        // Check editable field overrides for this row
        const rowEditableOverrides = BidWorksheetEstimate.EDITABLE_FIELD_OVERRIDES[data.excelRow] || {};
        const leftEditableOverrides = rowEditableOverrides.left || {};
        const rightEditableOverrides = rowEditableOverrides.right || {};

        // Determine readonly states with overrides
        const leftQuantityReadonly = leftEditableOverrides.quantity ? false : (leftIsTotalOrCalculated || leftDescEmpty);
        const leftUnitPriceReadonly = leftEditableOverrides.unitPrice ? false : (leftIsTotalOrCalculated || leftDescEmpty);
        const leftSizeReadonly = leftEditableOverrides.size ? false : true; // Default: size is readonly

        const rightQuantityReadonly = rightEditableOverrides.quantity ? false : (rightIsTotalOrCalculated || rightDescEmpty);
        const rightUnitPriceReadonly = rightEditableOverrides.unitPrice ? false : (rightIsTotalOrCalculated || rightDescEmpty);
        const rightSizeReadonly = rightEditableOverrides.size ? false : true; // Default: size is readonly

        return {
            id: id,
            excelRow: data.excelRow || null,
            rowClass: data.left.isTotalRow || data.right.isTotalRow ? 'total-row' : '',
            left: {
                description: data.left.description || '',
                descriptionReadonly: !!data.left.description || data.left.isReadonly,
                size: data.left.size || '',
                sizeReadonly: leftSizeReadonly, // ⭐ Configurable: can be overridden by EDITABLE_FIELD_OVERRIDES
                quantity: '',
                unitPrice: data.left.defaultUnitPrice || '',
                unitPriceReadonly: leftUnitPriceReadonly,
                defaultUnitPrice: data.left.defaultUnitPrice,
                gross: '',
                grossReadonly: true, // ⭐ ALWAYS readonly - it's calculated
                unitPriceFieldType: data.left.unitPriceFieldType || 'Currency',
                grossFieldType: data.left.grossFieldType || 'Currency',
                quantityReadonly: leftQuantityReadonly,
                quantityUserEntered: false, // ⭐ Track if user manually entered quantity
                sizeUserEntered: false, // ⭐ Track if user manually entered size
                isTotalRow: data.left.isTotalRow,
                descriptionClass: (!!data.left.description || data.left.isReadonly) ? 'description-cell readonly-cell' : 'description-cell',
                sizeClass: leftSizeReadonly ? 'readonly-cell' : '',
                quantityClass: leftQuantityReadonly ? 'readonly-cell' : '',
                unitPriceClass: leftUnitPriceReadonly ? 'col-unit readonly-cell' : 'col-unit',
                // Whole number flags for formatting
                isWholeNumberQuantity: leftWholeNumber.quantity || false,
                isWholeNumberUnitPrice: leftWholeNumber.unitPrice || false,
                isWholeNumberGross: leftWholeNumber.gross || false,
                // Currency flag for formatting
                isCurrencyGross: (data.left.grossFieldType || 'Currency') === 'Currency',
                // Computed fraction digits for HTML (expressions not allowed in HTML)
                grossMinFractionDigits: leftWholeNumber.gross ? 0 : 2,
                grossMaxFractionDigits: leftWholeNumber.gross ? 0 : 2
            },
            right: {
                description: data.right.description || '',
                descriptionReadonly: !!data.right.description || data.right.isReadonly,
                size: data.right.size || '',
                sizeReadonly: rightSizeReadonly, // ⭐ Configurable: can be overridden by EDITABLE_FIELD_OVERRIDES
                quantity: '',
                unitPrice: data.right.defaultUnitPrice || '',
                unitPriceReadonly: rightUnitPriceReadonly,
                defaultUnitPrice: data.right.defaultUnitPrice,
                gross: '',
                grossReadonly: true, // ⭐ ALWAYS readonly - it's calculated
                unitPriceFieldType: data.right.unitPriceFieldType || 'Currency',
                grossFieldType: data.right.grossFieldType || 'Currency',
                quantityReadonly: rightQuantityReadonly,
                quantityUserEntered: false, // ⭐ Track if user manually entered quantity
                sizeUserEntered: false, // ⭐ Track if user manually entered size
                isTotalRow: data.right.isTotalRow,
                isCommentRow: data.right.isCommentRow,
                descriptionClass: (!!data.right.description || data.right.isReadonly) ? 'description-cell readonly-cell' : 'description-cell',
                sizeClass: rightSizeReadonly ? 'readonly-cell' : '',
                quantityClass: rightQuantityReadonly ? 'readonly-cell' : '',
                unitPriceClass: rightUnitPriceReadonly ? 'col-unit readonly-cell' : 'col-unit',
                // Whole number flags for formatting
                isWholeNumberQuantity: rightWholeNumber.quantity || false,
                isWholeNumberUnitPrice: rightWholeNumber.unitPrice || false,
                isWholeNumberGross: rightWholeNumber.gross || false,
                // Currency flag for formatting (check grossFieldType)
                isCurrencyGross: (data.right.grossFieldType || 'Currency') === 'Currency',
                // Computed fraction digits for HTML (expressions not allowed in HTML)
                grossMinFractionDigits: (rightWholeNumber.gross ? 0 : 2),
                grossMaxFractionDigits: (rightWholeNumber.gross ? 0 : 2)
            }
        };
    }

    handleCellChange(event) {
        // Set flag to indicate user is actively editing
        this._isUserEditing = true;
        
        // Clear any existing timeout
        if (this._editingTimeout) {
            clearTimeout(this._editingTimeout);
        }
        
        // Clear the flag after 1 second of no activity
        this._editingTimeout = setTimeout(() => {
            this._isUserEditing = false;
        }, 1000);

        const section = parseInt(event.target.dataset.section);
        const rowId = parseInt(event.target.dataset.row);
        const col = event.target.dataset.col;
        const field = event.target.dataset.field;
        const value = event.target.value;

        const sectionKey = `section${section}Rows`;
        const rows = this[sectionKey];

        const rowIndex = rows.findIndex(row => row.id === rowId);
        if (rowIndex !== -1) {
            const updatedRow = { ...rows[rowIndex] };
            updatedRow[col] = { ...updatedRow[col], [field]: value };

            // Handle description changes - toggle readonly for other fields
            if (field === 'description') {
                const isEmpty = !value || value.trim() === '';

                // Only toggle readonly if it's NOT a total/calculated row
                if (!updatedRow[col].isTotalRow && !updatedRow[col].descriptionReadonly) {
                    updatedRow[col].quantityReadonly = isEmpty;
                    updatedRow[col].unitPriceReadonly = isEmpty;
                    // ⭐ Gross ALWAYS stays readonly (it's calculated)
                    updatedRow[col].grossReadonly = true;

                    // Update class properties to reflect readonly state
                    updatedRow[col].quantityClass = isEmpty ? 'readonly-cell' : '';
                    updatedRow[col].unitPriceClass = isEmpty ? 'col-unit readonly-cell' : 'col-unit';
                    updatedRow[col].descriptionClass = isEmpty ? 'description-cell' : 'description-cell readonly-cell';

                    // Clear values if description is cleared
                    if (isEmpty) {
                        updatedRow[col].quantity = '';
                        updatedRow[col].unitPrice = updatedRow[col].defaultUnitPrice || '';
                        updatedRow[col].gross = '';
                    }
                }
            }

            if (field === 'quantity' && value) {
                if (updatedRow[col].defaultUnitPrice && !updatedRow[col].unitPrice) {
                    updatedRow[col].unitPrice = updatedRow[col].defaultUnitPrice;
                }
            }

            if (field === 'quantity' || field === 'unitPrice') {
                updatedRow[col].gross = this.calculateGross(
                    updatedRow[col].quantity,
                    updatedRow[col].unitPrice
                );
            }

            // ⭐ Track manual edits for editable override fields
            // Row 186 (BIM) - quantity field
            if (field === 'quantity' && updatedRow.excelRow === 186 && col === 'left') {
                const row186EditableOverrides = BidWorksheetEstimate.EDITABLE_FIELD_OVERRIDES[186] || {};
                const leftOverrides = row186EditableOverrides.left || {};
                if (leftOverrides.quantity) {
                    // If user enters a value, mark as user-entered
                    // If user clears the field (empty), reset flag to allow auto-calculation
                    updatedRow[col].quantityUserEntered = !!(value && value.trim() !== '');
                }
            }

            // Row 187 (Fabrication) - size field
            if (field === 'size' && updatedRow.excelRow === 187 && col === 'left') {
                const row187EditableOverrides = BidWorksheetEstimate.EDITABLE_FIELD_OVERRIDES[187] || {};
                const leftOverrides = row187EditableOverrides.left || {};
                if (leftOverrides.size) {
                    // If user enters a value, mark as user-entered
                    // If user clears the field (empty), reset flag to allow auto-calculation
                    updatedRow[col].sizeUserEntered = !!(value && value.trim() !== '');
                }
            }

            this[sectionKey] = [
                ...rows.slice(0, rowIndex),
                updatedRow,
                ...rows.slice(rowIndex + 1)
            ];

            // Track which cell was just edited (to skip its recalculation)
            this._lastEditedCell = {
                excelRow: updatedRow.excelRow,
                col: col,
                field: field
            };

            this.calculateTotals();
            
            // Clear the flag after calculations
            setTimeout(() => {
                this._lastEditedCell = null;
            }, 100);
            
            // Notify parent for autosave (only if not loading data)
            if (!this._isLoadingData) {
                this.notifyParentForAutoSave();
            }
        }
    }

    /**
     * Notify parent component of cell change for autosave
     */
    notifyParentForAutoSave() {
        const event = new CustomEvent('cellchange', {
            bubbles: true,
            composed: true
        });
        this.dispatchEvent(event);
    }

    calculateGross(quantity, unitPrice) {
        const qty = parseFloat(quantity) || 0;
        const price = parseFloat(unitPrice) || 0;
        const gross = qty * price;
        return gross > 0 ? gross.toFixed(2) : '';
    }

    calculateTotals() {
        this.section1Subtotal = this.calculateSectionTotal(this.section1Rows);
        this.section2Subtotal = this.calculateSectionTotal(this.section2Rows);

        this.applySection3Calculations();

        this.section3Subtotal = this.calculateSectionTotal(this.section3Rows);
        console.log('section3Subtotal :- ', this.section3Subtotal);

        const s1 = parseFloat(this.section1Subtotal) || 0;
        const s2 = parseFloat(this.section2Subtotal) || 0;
        const s3 = parseFloat(this.section3Subtotal) || 0;
        this.grandTotal = (s1 + s2 + s3).toFixed(2);

        console.log(`Totals: S1=$${this.section1Subtotal}, S2=$${this.section2Subtotal}, S3=$${this.section3Subtotal}, Grand=$${this.grandTotal}`);

        this.notifyParent();
    }

    applySection1Calculations() {
        const s1Rows = this.section1Rows;

        const findRow = (excelRowNum) => {
            return s1Rows.find(r => r.excelRow === excelRowNum);
        };

        const getGross = (row, side) => {
            if (!row) return 0;
            return parseFloat(row[side].gross) || 0;
        };

        const getQty = (row, side) => {
            if (!row) return 0;
            return parseFloat(row[side].quantity) || 0;
        };

        // ROW 22: TOTAL HEADS
        const row22 = findRow(22);
        if (row22) {
            let totalQty = 0;
            let totalGross = 0;

            for (let i = 5; i <= 21; i++) {
                const r = findRow(i);
                if (r) {
                    totalQty += getQty(r, 'left');
                    totalGross += getGross(r, 'left');
                }
            }

            row22.left.quantity = totalQty > 0 ? totalQty.toFixed(2) : '';
            row22.left.gross = totalGross > 0 ? totalGross.toFixed(2) : '';
            row22.left.quantityReadonly = true;
            row22.left.unitPriceReadonly = true; // ⭐ ADDED
            row22.left.grossReadonly = true;
        }

        // ROW 69: SUB TOTAL SHT #1
        // Formula: SUM(I22)+SUM(I23:I67)+SUM(R5:R67)
        const row69 = findRow(69);
        if (row69) {
            let totalGross = 0;

            // Left column: SUM(I22)+SUM(I23:I67) = SUM(I22:I67)
            for (let i = 22; i <= 67; i++) {
                const r = findRow(i);
                if (r) {
                    totalGross += getGross(r, 'left');
                }
            }

            // Right column: SUM(R5:R67)
            for (let i = 5; i <= 67; i++) {
                const r = findRow(i);
                if (r) {
                    totalGross += getGross(r, 'right');
                }
            }

            row69.left.gross = totalGross > 0 ? totalGross.toFixed(2) : '';
            row69.left.quantityReadonly = true; // ⭐ ADDED
            row69.left.unitPriceReadonly = true; // ⭐ ADDED
            row69.left.grossReadonly = true;
        }

        this.section1Rows = [...s1Rows];
    }

    applySection2Calculations() {
        const s2Rows = this.section2Rows;

        const findRow = (excelRowNum) => {
            return s2Rows.find(r => r.excelRow === excelRowNum);
        };

        const getGross = (row, side) => {
            if (!row) return 0;
            return parseFloat(row[side].gross) || 0;
        };

        // ROW 125: SUB TOTAL SHT #2
        const row125 = findRow(125);
        if (row125) {
            let totalGross = 0;

            for (let i = 77; i <= 123; i++) {
                const r = findRow(i);
                if (r) {
                    totalGross += getGross(r, 'left');
                    totalGross += getGross(r, 'right');
                }
            }

            row125.left.gross = totalGross > 0 ? totalGross.toFixed(2) : '';
            row125.left.quantityReadonly = true; // ⭐ ADDED
            row125.left.unitPriceReadonly = true; // ⭐ ADDED
            row125.left.grossReadonly = true;
        }

        this.section2Rows = [...s2Rows];
    }

    applySection3Calculations() {
        this.applySection1Calculations();
        this.applySection2Calculations();

        const s3Rows = this.section3Rows;

        const findRow = (excelRowNum) => {
            return s3Rows.find(r => r.excelRow === excelRowNum);
        };

        const getGross = (row, side) => {
            if (!row) return 0;
            return parseFloat(row[side].gross) || 0;
        };

        const getQty = (row, side) => {
            if (!row) return 0;
            return parseFloat(row[side].quantity) || 0;
        };

        const getUnit = (row, side) => {
            if (!row) return 0;
            return parseFloat(row[side].unitPrice) || 0;
        };

        // ROW 133: TOTAL MAT'L SHT #1 & 2
        const row133 = findRow(133);
        if (row133) {
            const row69 = this.section1Rows.find(r => r.excelRow === 69);
            const row125 = this.section2Rows.find(r => r.excelRow === 125);

            const row69Total = row69 ? (parseFloat(row69.left.gross) || 0) : 0;
            const row125Total = row125 ? (parseFloat(row125.left.gross) || 0) : 0;

            row133.left.gross = (row69Total + row125Total).toFixed(2);
            row133.left.quantityReadonly = true; // ⭐ ADDED
            row133.left.unitPriceReadonly = true; // ⭐ ADDED
            row133.left.grossReadonly = true;
        }

        // ROW 135: SALES TAX 10%
        const row135 = findRow(135);
        if (row135 && row133) {
            const materialTotal = parseFloat(row133.right.gross) || 0;
            const taxRate = getUnit(row135, 'right') || 0;
            row135.right.gross = (materialTotal * taxRate).toFixed(2);
            row135.right.quantityReadonly = true; // ⭐ ADDED
            // ⭐ Unit price is editable (tax rate can be changed)
            row135.right.grossReadonly = true;
        }

        // ROW 155: GRAND TOTAL MATERIAL COST
        const row155 = findRow(155);
        if (row155) {
            let sum = 0;
            for (let i = 133; i <= 153; i++) {
                const r = findRow(i);
                if (r) sum += getGross(r, 'left');
            }
            row155.left.gross = sum.toFixed(2);
            row155.left.quantityReadonly = true; // ⭐ ADDED
            row155.left.unitPriceReadonly = true; // ⭐ ADDED
            row155.left.grossReadonly = true;

            if (row133) {
                row133.right.gross = row155.left.gross;
                row133.right.quantityReadonly = true; // ⭐ ADDED
                row133.right.unitPriceReadonly = true; // ⭐ ADDED
                row133.right.grossReadonly = true;
            }
        }

        // ROW 152: MATERIAL, PRMT., EQUIP., ....
        const row152 = findRow(152);
        if (row152) {
            let sum = 0;
            for (let i = 133; i <= 150; i++) {
                const r = findRow(i);
                console.log('Adding row', i, 'gross:', r ? getGross(r, 'right') : 0);
                if (r) sum += getGross(r, 'right');
            }
            row152.right.gross = sum.toFixed(2);
            row152.right.quantityReadonly = true; // ⭐ ADDED
            row152.right.unitPriceReadonly = true; // ⭐ ADDED
            row152.right.grossReadonly = true;
        }

        // ROW 153: FIELD, ENGR., FAB TOTAL
        const row153 = findRow(153);
        if (row153) {
            const row189 = findRow(189);
            const laborTotal = row189 ? (parseFloat(row189.left.gross) || 0) : 0;
            row153.right.gross = laborTotal > 0 ? laborTotal.toFixed(2) : '';
            row153.right.quantityReadonly = true; // ⭐ ADDED
            row153.right.unitPriceReadonly = true; // ⭐ ADDED
            row153.right.grossReadonly = true;
        }

        // ROW 157: HEADCOUNT
        const row157 = findRow(157);
        if (row157) {
            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            row157.left.size = headcount > 0 ? headcount.toFixed(2) : '';
            row157.left.quantity = '';
            row157.left.quantityReadonly = true;
            row157.left.unitPriceReadonly = true; // ⭐ ADDED
            row157.left.grossReadonly = true; // ⭐ ADDED
        }


        // ROW 160: TOTAL DIRECT COST
        const row160 = findRow(160);
        if (row160) {
            let sum = 0;
            for (let i = 152; i <= 158; i++) {
                const r = findRow(i);
                if (r) sum += getGross(r, 'right');
            }
            row160.right.gross = sum > 0 ? sum.toFixed(2) : '';
            row160.right.quantityReadonly = true; // ⭐ ADDED
            row160.right.unitPriceReadonly = true; // ⭐ ADDED
            row160.right.grossReadonly = true;
        }

        // ROW 161: %OVERHEAD
        const row161 = findRow(161);
        if (row161 && row160) {
            const directCost = parseFloat(row160.right.gross) || 0;
            row161.right.quantity = directCost > 0 ? directCost.toFixed(2) : '';
            row161.right.quantityReadonly = true;

            // ⭐ Unit price is editable (overhead % can be changed)
            const unitPrice = getUnit(row161, 'right') || 0.15;
            row161.right.gross = directCost > 0 ? (directCost * unitPrice).toFixed(2) : '';
            row161.right.grossReadonly = true;
        }

        // ROW 163: SUBTOTAL
        const row163 = findRow(163);
        if (row163 && row160 && row161) {
            const directCost = parseFloat(row160.right.gross) || 0;
            const overhead = parseFloat(row161.right.gross) || 0;
            row163.right.gross = (directCost + overhead).toFixed(2);
            row163.right.quantityReadonly = true; // ⭐ ADDED
            row163.right.unitPriceReadonly = true; // ⭐ ADDED
            row163.right.grossReadonly = true;
        }

        // ROW 164: %GAIN
        const row164 = findRow(164);
        if (row164 && row163) {
            const subtotal = parseFloat(row163.right.gross) || 0;
            row164.right.quantity = subtotal > 0 ? subtotal.toFixed(2) : '';
            row164.right.quantityReadonly = true;

            // ⭐ Unit price is editable (gain % can be changed)
            const unitPrice = getUnit(row164, 'right') || 0.15;
            row164.right.gross = subtotal > 0 ? (subtotal * unitPrice).toFixed(2) : '';
            row164.right.grossReadonly = true;
        }

        // ROW 166: TOTAL QUOTE PRICE
        const row166 = findRow(166);
        if (row166 && row163 && row164) {
            const overhead = getGross(row163, 'right');
            const gain = getGross(row164, 'right');
            row166.right.gross = (overhead + gain).toFixed(2);
            row166.right.quantityReadonly = true; // ⭐ ADDED
            row166.right.unitPriceReadonly = true; // ⭐ ADDED
            row166.right.grossReadonly = true;
        }

        // ROW 167: PRICE MINUS SP, PUMP, BF, MFLEX
        const row167 = findRow(167);
        if (row167 && row166) {
            const quotePrice = parseFloat(row166.right.gross) || 0;
            const row144 = findRow(144);
            const row145 = findRow(145);
            const row150 = findRow(150);
            const row151 = findRow(151);
            const row152 = findRow(152);

            const i144 = row144 ? (parseFloat(row144.left.gross) || 0) : 0;
            const i145 = row145 ? (parseFloat(row145.left.gross) || 0) : 0;
            const i150 = row150 ? (parseFloat(row150.left.gross) || 0) : 0;
            const i151 = row151 ? (parseFloat(row151.left.gross) || 0) : 0;
            const i152 = row152 ? (parseFloat(row152.left.gross) || 0) : 0;

            const result = quotePrice - i144 - i145 - i150 - i151 - i152;
            row167.right.gross = result.toFixed(2);
            row167.right.quantityReadonly = true; // ⭐ ADDED
            row167.right.unitPriceReadonly = true; // ⭐ ADDED
            row167.right.grossReadonly = true;
        }

        // ROW 168: GROSS MARGIN
        const row168 = findRow(168);
        if (row168 && row161 && row164 && row166) {
            const overhead = parseFloat(row161.right.gross) || 0;
            const gain = parseFloat(row164.right.gross) || 0;
            const quotePrice = parseFloat(row166.right.gross) || 0;

            if (quotePrice > 0) {
                const margin = (overhead + gain) / quotePrice;
                row168.right.quantity = margin.toFixed(4);
                row168.right.quantityReadonly = true;
            } else {
                row168.right.quantity = '';
            }
            row168.right.unitPriceReadonly = true; // ⭐ ADDED
            row168.right.grossReadonly = true; // ⭐ ADDED
        }

        // ROW 169: BOND AMOUNT
        const row169 = findRow(169);
        if (row169 && row166) {
            const totalQuote = parseFloat(row166.right.gross) || 0;

            const bondCalc = (totalQuote / 1000) * 1.5;
            row169.right.quantity = (bondCalc < 100 ? 100 : bondCalc).toFixed(2);
            row169.right.quantityReadonly = true;

            row169.right.gross = ((totalQuote / 1000) * 12).toFixed(2);
            row169.right.unitPriceReadonly = true; // ⭐ ADDED
            row169.right.grossReadonly = true;
        }

        // ROW 173: MATERIAL PER HEAD
        const row173 = findRow(173);
        if (row173) {
            const row155 = findRow(155);
            const row144 = findRow(144);
            const row145 = findRow(145);
            const row150 = findRow(150);
            const row151 = findRow(151);
            const row152 = findRow(152);

            const i155 = row155 ? (parseFloat(row155.left.gross) || 0) : 0;
            const i144 = row144 ? (parseFloat(row144.left.gross) || 0) : 0;
            const i145 = row145 ? (parseFloat(row145.left.gross) || 0) : 0;
            const i150 = row150 ? (parseFloat(row150.left.gross) || 0) : 0;
            const i151 = row151 ? (parseFloat(row151.left.gross) || 0) : 0;
            const i152 = row152 ? (parseFloat(row152.left.gross) || 0) : 0;

            const materialCost = i155 - i151 - i144 - i152 - i145 - i150;
            row173.right.quantity = materialCost.toFixed(2);
            row173.right.quantityReadonly = true;

            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            row173.right.unitPrice = headcount > 0 ? headcount.toFixed(2) : '';
            row173.right.unitPriceReadonly = true;

            if (headcount > 0) {
                row173.right.gross = (materialCost / headcount).toFixed(2);
                row173.right.grossReadonly = true;
            }
        }

        // ROW 174: DIRECT COST PER HEAD
        const row174 = findRow(174);
        if (row174 && row160) {
            const directCost = parseFloat(row160.right.gross) || 0;
            row174.right.quantity = directCost > 0 ? directCost.toFixed(2) : '';
            row174.right.quantityReadonly = true;

            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            row174.right.unitPrice = headcount > 0 ? headcount.toFixed(2) : '';
            row174.right.unitPriceReadonly = true;

            if (headcount > 0) {
                row174.right.gross = (directCost / headcount).toFixed(2);
                row174.right.grossReadonly = true;
            }
        }

        // ROW 175: BUILDING SQ. FOOTAGE
        const row175 = findRow(175);
        if (row175) {
            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            row175.right.unitPrice = headcount > 0 ? headcount.toFixed(2) : '';
            row175.right.unitPriceReadonly = true;

            // ⭐ Quantity is editable (user enters sq footage)
            const qty = getQty(row175, 'right');
            const unit = parseFloat(row175.right.unitPrice) || 0;
            if (qty > 0 && unit > 0) {
                row175.right.gross = (qty / unit).toFixed(2);
                row175.right.grossReadonly = true;
            }
        }

        // ROW 176: SALES COST PER HEAD
        const row176 = findRow(176);
        if (row176 && row166) {
            const quotePrice = parseFloat(row166.right.gross) || 0;
            const row144 = findRow(144);
            const row145 = findRow(145);
            const row150 = findRow(150);
            const row151 = findRow(151);
            const row152 = findRow(152);

            const i144 = row144 ? (parseFloat(row144.left.gross) || 0) : 0;
            const i145 = row145 ? (parseFloat(row145.left.gross) || 0) : 0;
            const i150 = row150 ? (parseFloat(row150.left.gross) || 0) : 0;
            const i151 = row151 ? (parseFloat(row151.left.gross) || 0) : 0;
            const i152 = row152 ? (parseFloat(row152.left.gross) || 0) : 0;

            const salesCost = quotePrice - i144 - i145 - i150 - i151 - i152;
            row176.right.quantity = salesCost.toFixed(2);
            row176.right.quantityReadonly = true;

            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            row176.right.unitPrice = headcount > 0 ? headcount.toFixed(2) : '';
            row176.right.unitPriceReadonly = true;

            if (headcount > 0) {
                row176.right.gross = (salesCost / headcount).toFixed(2);
                row176.right.grossReadonly = true;
            }
        }

        // ROW 178: COST PER SQUARE FOOT
        const row178 = findRow(178);
        if (row178 && row167 && row175) {
            const netPrice = parseFloat(row167.right.gross) || 0;
            const sqFootage = parseFloat(row175.right.quantity) || 0;

            if (sqFootage > 0) {
                row178.right.gross = (netPrice / sqFootage).toFixed(2);
                row178.right.grossReadonly = true;
            } else {
                row178.right.gross = '';
            }
            // Set quantity and unitPrice to readonly (calculated row)
            row178.right.quantityReadonly = true;
            row178.right.unitPriceReadonly = true;
            // Update class properties to reflect readonly state
            row178.right.quantityClass = 'readonly-cell';
            row178.right.unitPriceClass = 'col-unit readonly-cell';
        }

        // ROW 182: TOTAL LABOR HRS.
        const row182 = findRow(182);
        if (row182) {
            let sum = 0;
            for (let i = 159; i <= 180; i++) {
                const r = findRow(i);
                if (r) sum += getGross(r, 'left');
            }
            // Always store with 2 decimal places for backend accuracy
            // UI will display as whole number if isWholeNumberGross is true
            row182.left.gross = sum > 0 ? sum.toFixed(2) : '';
            row182.left.quantityReadonly = true; // ⭐ ADDED
            row182.left.unitPriceReadonly = true; // ⭐ ADDED
            row182.left.grossReadonly = true;
        }

        // ROW 184: LABOR (FM+ 7TH PERIOD)
        const row184 = findRow(184);
        if (row184 && row182) {
            console.log('new test row184 :- ', JSON.stringify(row184));
            console.log('new test row182 :- ', JSON.stringify(row182));

            const laborHrs = parseFloat(row182.left.gross) || 0;
            // Store with 2 decimals for backend, but display as whole number if configured
            if (row184.left.isWholeNumberQuantity) {
                // Display as whole number (rounded), but keep decimal precision in value
                row184.left.quantity = laborHrs > 0 ? Math.round(laborHrs).toString() : '';
            } else {
                row184.left.quantity = laborHrs > 0 ? laborHrs.toFixed(2) : '';
            }
            row184.left.quantityReadonly = true;

            // ⭐ Unit price is editable (labor rate can be changed)
            const unitPrice = parseFloat(row184.left.unitPrice) || 0;
            if (laborHrs > 0 && unitPrice > 0) {
                row184.left.gross = (laborHrs * unitPrice).toFixed(2);
            } else {
                row184.left.gross = '';
            }
            row184.left.grossReadonly = true;

            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            if (headcount > 0 && laborHrs > 0) {
                row184.left.size = (laborHrs / headcount).toFixed(2);
            } else {
                row184.left.size = '';
            }

            console.log('new test row184 :- ', row184);
            console.log('new test row182 :- ', row182);
        }

        // ROW 185: ENGINEERING HALF HOUR HEAD
        const row185 = findRow(185);
        if (row185) {
            // ⭐ Quantity is editable (user can enter engineering hours)
            const qty = getQty(row185, 'left') || 40;
            // ⭐ Unit price is editable (engineering rate can be changed)
            const unitPrice = getUnit(row185, 'left');

            if (qty > 0 && unitPrice > 0) {
                row185.left.gross = (qty * unitPrice).toFixed(2);
            } else {
                row185.left.gross = '';
            }
            row185.left.grossReadonly = true; // ⭐ Gross is calculated

            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            if (headcount > 0 && qty > 0) {
                row185.left.size = (qty / headcount).toFixed(2);
            } else {
                row185.left.size = '';
            }
        }

        // ROW 186: BIM
        const row186 = findRow(186);
        if (row186) {
            const row157 = findRow(157);
            if (row157) {
                // ⭐ Check if quantity should be editable (override configuration)
                const row186EditableOverrides = BidWorksheetEstimate.EDITABLE_FIELD_OVERRIDES[186] || {};
                const leftOverrides = row186EditableOverrides.left || {};
                const isQuantityEditable = leftOverrides.quantity;
                const isSizeEditable = leftOverrides.size;
                
                const headcount = parseFloat(row157.left.size) || 0;
                const bimQty = headcount / 2;
                
                // ⭐ Recalculate quantity based on source (row 157)
                // Skip recalculation if user is currently editing this field
                const isEditingThisField = this._lastEditedCell && 
                    this._lastEditedCell.excelRow === 186 && 
                    this._lastEditedCell.col === 'left' && 
                    this._lastEditedCell.field === 'quantity';
                
                if (!isQuantityEditable) {
                    // Always calculate if readonly
                    row186.left.quantity = bimQty > 0 ? bimQty.toFixed(2) : '';
                    row186.left.quantityUserEntered = false;
                } else if (!isEditingThisField) {
                    // Recalculate if user is NOT currently editing this field
                    // This allows manual edits to stay, but recalculates when source (row 157) changes
                    row186.left.quantity = bimQty > 0 ? bimQty.toFixed(2) : '';
                }
                // If user is editing this field, preserve their input (don't recalculate)
                
                row186.left.quantityReadonly = !isQuantityEditable; // If in config, make editable
                
                // ⭐ Use calculated quantity value
                const currentQty = parseFloat(row186.left.quantity) || bimQty;

                // ⭐ Unit price is editable (BIM rate can be changed)
                const unitPrice = getUnit(row186, 'left');
                if (currentQty > 0 && unitPrice > 0) {
                    row186.left.gross = (currentQty * unitPrice).toFixed(2);
                } else {
                    row186.left.gross = '';
                }
                row186.left.grossReadonly = true;

                const row22 = this.section1Rows.find(r => r.excelRow === 22);
                const totalHeadcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
                
                // ⭐ Recalculate size based on currentQty and totalHeadcount
                // Skip recalculation if user is currently editing this field
                const isEditingThisSizeField = this._lastEditedCell && 
                    this._lastEditedCell.excelRow === 186 && 
                    this._lastEditedCell.col === 'left' && 
                    this._lastEditedCell.field === 'size';
                
                if (!isSizeEditable) {
                    // Always calculate if readonly
                    if (totalHeadcount > 0 && currentQty > 0) {
                        row186.left.size = (currentQty / totalHeadcount).toFixed(2);
                    } else {
                        row186.left.size = '';
                    }
                    row186.left.sizeUserEntered = false;
                } else if (!isEditingThisSizeField) {
                    // Recalculate if user is NOT currently editing this field
                    if (totalHeadcount > 0 && currentQty > 0) {
                        row186.left.size = (currentQty / totalHeadcount).toFixed(2);
                    } else {
                        row186.left.size = '';
                    }
                }
                // If user is editing this field, preserve their input (don't recalculate)
                
                // ⭐ Check if size should be editable (override configuration)
                row186.left.sizeReadonly = !isSizeEditable; // If in config, make editable
                row186.left.sizeClass = row186.left.sizeReadonly ? 'readonly-cell' : ''; // Update class
            }
        }

        // ROW 187: FABRICATION QUARTER HOUR PER
        const row187 = findRow(187);
        if (row187) {
            // ⭐ Check if size should be editable (override configuration)
            const row187EditableOverrides = BidWorksheetEstimate.EDITABLE_FIELD_OVERRIDES[187] || {};
            const leftOverrides = row187EditableOverrides.left || {};
            const isSizeEditable = leftOverrides.size;
            
            // ⭐ Quantity is editable (user can enter fabrication hours)
            const qty = getQty(row187, 'left');
            // ⭐ Unit price is editable (fabrication rate can be changed)
            const unitPrice = getUnit(row187, 'left');

            if (qty > 0 && unitPrice > 0) {
                row187.left.gross = (qty * unitPrice).toFixed(2);
            } else {
                row187.left.gross = '';
            }
            row187.left.grossReadonly = true; // ⭐ Gross is calculated

            const row22 = this.section1Rows.find(r => r.excelRow === 22);
            const headcount = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;
            
            // ⭐ Recalculate size based on qty and headcount
            // Skip recalculation if user is currently editing this field
            const isEditingThisSizeField = this._lastEditedCell && 
                this._lastEditedCell.excelRow === 187 && 
                this._lastEditedCell.col === 'left' && 
                this._lastEditedCell.field === 'size';
            
            if (!isSizeEditable) {
                // Always calculate if readonly
                if (headcount > 0 && qty > 0) {
                    row187.left.size = (qty / headcount).toFixed(2);
                } else {
                    row187.left.size = '';
                }
                row187.left.sizeUserEntered = false;
            } else if (!isEditingThisSizeField) {
                // Recalculate if user is NOT currently editing this field
                if (headcount > 0 && qty > 0) {
                    row187.left.size = (qty / headcount).toFixed(2);
                } else {
                    row187.left.size = '';
                }
            }
            // If user is editing this field, preserve their input (don't recalculate)
            
            // ⭐ Check if size should be editable (override configuration)
            row187.left.sizeReadonly = !isSizeEditable; // If in config, make editable
            row187.left.sizeClass = row187.left.sizeReadonly ? 'readonly-cell' : ''; // Update class
        }

        // ROW 189: FIELD,ENG,FAB, TOTAL
        const row189 = findRow(189);
        if (row189) {
            let sum = 0;
            for (let i = 184; i <= 188; i++) {
                const r = findRow(i);
                if (r) sum += getGross(r, 'left');
            }
            row189.left.gross = sum > 0 ? sum.toFixed(2) : '';
            row189.left.quantityReadonly = true; // ⭐ ADDED
            row189.left.unitPriceReadonly = true; // ⭐ ADDED
            row189.left.grossReadonly = true;
        }

        this.section3Rows = [...s3Rows];
    }

    calculateSectionTotal(rows) {
        let total = 0;
        rows.forEach(row => {
            const leftGross = parseFloat(row.left.gross) || 0;
            const rightGross = parseFloat(row.right.gross) || 0;
            total += leftGross + rightGross;
        });
        return total.toFixed(2);
    }

    notifyParent() {
        this.dispatchEvent(new CustomEvent('sheetupdate', {
            detail: { grandTotal: this.grandTotal }
        }));
    }

    @api
    async saveSheet() {
        return {
            section1: this.collectSectionData(this.section1Rows, 1),
            section2: this.collectSectionData(this.section2Rows, 2),
            section3: this.collectSectionData(this.section3Rows, 3),
            grandTotal: this.grandTotal
        };
    }

    collectSectionData(rows, sectionNum) {
        const lineItems = [];
        rows.forEach((row, index) => {
            // Save left side if it has description or any data
            if (row.left.description || row.left.quantity || row.left.unitPrice || row.left.size || row.left.gross) {
                lineItems.push({
                    section: sectionNum,
                    rowNumber: index + 1,
                    excelRow: row.excelRow,
                    column: 'Left',
                    description: row.left.description || '',
                    size: row.left.size || '',
                    quantity: parseFloat(row.left.quantity) || 0,
                    unitPrice: parseFloat(row.left.unitPrice) || 0,
                    gross: parseFloat(row.left.gross) || 0
                });
            }
            // Save right side if it has description or any data
            if (row.right.description || row.right.quantity || row.right.unitPrice || row.right.size || row.right.gross) {
                lineItems.push({
                    section: sectionNum,
                    rowNumber: index + 1,
                    excelRow: row.excelRow,
                    column: 'Right',
                    description: row.right.description || '',
                    size: row.right.size || '',
                    quantity: parseFloat(row.right.quantity) || 0,
                    unitPrice: parseFloat(row.right.unitPrice) || 0,
                    gross: parseFloat(row.right.gross) || 0
                });
            }
        });
        return lineItems;
    }

    showToast(title, message, variant) {
        this.dispatchEvent(new ShowToastEvent({ title, message, variant }));
    }
}