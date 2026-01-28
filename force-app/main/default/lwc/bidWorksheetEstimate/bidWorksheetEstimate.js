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
        
        // Check if rows are initialized
        const rowsReady = this.section1Rows.length > 0 || this.section2Rows.length > 0 || this.section3Rows.length > 0;
        
        if (!rowsReady) {
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
            // Don't reload if user is actively editing
            if (!this._isUserEditing) {
                this.loadSavedData();
            } else {
            }
        } else {
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
        186: { left: { quantity: true } }, // BIM - quantity field is whole number
    };

    // ========================================
    // EDITABLE FIELD OVERRIDE CONFIGURATION
    // ========================================
    // Format: { rowNumber: { side: { field: true } } }
    // Example: { 186: { left: { quantity: true } } } means row 186, left side, quantity field is editable (overrides calculated readonly)
    // Example: { 187: { left: { size: true } } } means row 187, left side, size field is editable
    // To add more editable fields, just add entries here:
    static EDITABLE_FIELD_OVERRIDES = {
        153: { left: { description: true } }, // MISC. - description field is editable
        180: { left: { description: true } }, // MISC. - description field is editable
        186: { left: { quantity: true } }, // BIM - quantity field is editable (overrides calculated readonly)
        187: { left: { size: true } }, // FABRICATION QUARTER HOUR PER - size field is editable
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
            return;
        }

        // Don't load if rows haven't been initialized from metadata yet
        if (this.section1Rows.length === 0 && this.section2Rows.length === 0 && this.section3Rows.length === 0) {
            return;
        }

        // Don't load if user is actively editing
        if (this._isUserEditing) {
            return;
        }

        // Set loading flag to prevent autosave during load
        this._isLoadingData = true;

        try {
            let base64Data;
            
            // If versionIdToLoad is set, load that specific version
            if (this.versionIdToLoad && this.versionIdToLoad !== 'draft') {
                base64Data = await loadVersionById_Estimate({ versionId: this.versionIdToLoad });
                this._lastLoadedVersionId = this.versionIdToLoad;
            } else {
                // Otherwise, load latest (autosave or most recent)
                base64Data = await loadLatestEstimateSheet({ opportunityId: this.recordId });
                this._lastLoadedVersionId = 'draft';
            }

            if (!base64Data) {
                this._isLoadingData = false;
                return;
            }

            // Decode base64 data
            const jsonString = this.decodeData(base64Data);
            const data = JSON.parse(jsonString);

            this.applySavedData(data);
        } catch (error) {
            const errorMessage = error?.body?.message || error?.message || String(error);
            if (errorMessage.includes('not found') || errorMessage.includes('No ContentVersion') || errorMessage.includes('List has no rows')) {
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
                    row[side].quantityRaw = item.quantity || '';
                    if (row[side].isWholeNumberQuantity && item.quantity) {
                        row[side].quantity = Math.round(item.quantity).toString();
                    } else {
                        row[side].quantity = item.quantity || '';
                    }
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
            // Set loading flag first to prevent autosave during initialization
            this._isLoadingData = true;
            
            this.initializeDataFromMetadata(data);
            this.isLoading = false;
            
            // Load saved data after metadata is loaded and rows are initialized
            setTimeout(() => {
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
        const leftDescriptionReadonly = leftEditableOverrides.description ? false : (!!data.left.description || data.left.isReadonly);

        const rightQuantityReadonly = rightEditableOverrides.quantity ? false : (rightIsTotalOrCalculated || rightDescEmpty);
        const rightUnitPriceReadonly = rightEditableOverrides.unitPrice ? false : (rightIsTotalOrCalculated || rightDescEmpty);
        const rightSizeReadonly = rightEditableOverrides.size ? false : true; // Default: size is readonly
        const rightDescriptionReadonly = rightEditableOverrides.description ? false : (!!data.right.description || data.right.isReadonly);

        return {
            id: id,
            excelRow: data.excelRow || null,
            rowClass: data.left.isTotalRow || data.right.isTotalRow ? 'total-row' : '',
            left: {
                description: data.left.description || '',
                descriptionReadonly: leftDescriptionReadonly,
                size: data.left.size || '',
                sizeReadonly: leftSizeReadonly, // ⭐ Configurable: can be overridden by EDITABLE_FIELD_OVERRIDES
                quantity: '',
                quantityRaw: '', // ⭐ Precise decimal value for calculations
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
                descriptionClass: leftDescriptionReadonly ? 'description-cell readonly-cell' : 'description-cell',
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
                descriptionReadonly: rightDescriptionReadonly,
                size: data.right.size || '',
                sizeReadonly: rightSizeReadonly, // ⭐ Configurable: can be overridden by EDITABLE_FIELD_OVERRIDES
                quantity: '',
                quantityRaw: '', // ⭐ Precise decimal value for calculations
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
                descriptionClass: rightDescriptionReadonly ? 'description-cell readonly-cell' : 'description-cell',
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

            // Handle quantityRaw for quantity field
            if (field === 'quantity') {
                updatedRow[col].quantityRaw = value;
            }

            // Handle description changes - toggle readonly for other fields
            if (field === 'description') {
                const isEmpty = !value || value.trim() === '';

                // Only toggle readonly if it's NOT a total/calculated row
                // OR if description is NOT readonly (meaning it's one of our editable overrides like MISC.)
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
                        updatedRow[col].quantityRaw = '';
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
                    updatedRow[col].quantityRaw || updatedRow[col].quantity,
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
        // 1. Run formula logic first to update calculated rows
        this.applySection3Calculations();

        // 2. NOW calculate the subtotals based on the updated rows
        this.section1Subtotal = this.calculateSectionTotal(this.section1Rows);
        this.section2Subtotal = this.calculateSectionTotal(this.section2Rows);
        this.section3Subtotal = this.calculateSectionTotal(this.section3Rows);

        // 3. Update grand total
        const s1 = parseFloat(this.section1Subtotal) || 0;
        const s2 = parseFloat(this.section2Subtotal) || 0;
        const s3 = parseFloat(this.section3Subtotal) || 0;
        this.grandTotal = (s1 + s2 + s3).toFixed(2);

        console.log(`Totals: S1=$${this.section1Subtotal}, S2=$${this.section2Subtotal}, S3=$${this.section3Subtotal}, Grand=$${this.grandTotal}`);

        this.notifyParent();
    }

    applySection1Calculations() {
        const s1Rows = [...this.section1Rows];
        let sectionChanged = false;

        const findRowIndex = (excelRowNum) => {
            return s1Rows.findIndex(r => r.excelRow === excelRowNum);
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
        const row22Idx = findRowIndex(22);
        if (row22Idx !== -1) {
            let totalQty = 0;
            let totalGross = 0;

            for (let i = 5; i <= 21; i++) {
                const rIdx = findRowIndex(i);
                if (rIdx !== -1) {
                    totalQty += getQty(s1Rows[rIdx], 'left');
                    totalGross += getGross(s1Rows[rIdx], 'left');
                }
            }

            const row22 = s1Rows[row22Idx];
            const newQty = totalQty > 0 ? totalQty.toFixed(2) : '';
            const newGross = totalGross > 0 ? totalGross.toFixed(2) : '';

            if (row22.left.quantity !== newQty || row22.left.gross !== newGross) {
                s1Rows[row22Idx] = {
                    ...row22,
                    left: {
                        ...row22.left,
                        quantity: newQty,
                        gross: newGross,
                        quantityReadonly: true,
                        unitPriceReadonly: true,
                        grossReadonly: true
                    }
                };
                sectionChanged = true;
            }
        }

        // ROW 69: SUB TOTAL SHT #1
        const row69Idx = findRowIndex(69);
        if (row69Idx !== -1) {
            let totalGross = 0;

            // Left column: SUM(I22:I67)
            for (let i = 22; i <= 67; i++) {
                const rIdx = findRowIndex(i);
                if (rIdx !== -1) {
                    totalGross += getGross(s1Rows[rIdx], 'left');
                }
            }

            // Right column: SUM(R5:R67)
            for (let i = 5; i <= 67; i++) {
                const rIdx = findRowIndex(i);
                if (rIdx !== -1) {
                    totalGross += getGross(s1Rows[rIdx], 'right');
                }
            }

            const row69 = s1Rows[row69Idx];
            const newGross = totalGross > 0 ? totalGross.toFixed(2) : '';

            if (row69.left.gross !== newGross) {
                s1Rows[row69Idx] = {
                    ...row69,
                    left: {
                        ...row69.left,
                        gross: newGross,
                        quantityReadonly: true,
                        unitPriceReadonly: true,
                        grossReadonly: true
                    }
                };
                sectionChanged = true;
            }
        }

        if (sectionChanged) {
            this.section1Rows = s1Rows;
        }
    }

    applySection2Calculations() {
        const s2Rows = [...this.section2Rows];
        let sectionChanged = false;

        const findRowIndex = (excelRowNum) => {
            return s2Rows.findIndex(r => r.excelRow === excelRowNum);
        };

        const getGross = (row, side) => {
            if (!row) return 0;
            return parseFloat(row[side].gross) || 0;
        };

        // ROW 125: SUB TOTAL SHT #2
        const row125Idx = findRowIndex(125);
        if (row125Idx !== -1) {
            let totalGross = 0;

            for (let i = 77; i <= 123; i++) {
                const rIdx = findRowIndex(i);
                if (rIdx !== -1) {
                    totalGross += getGross(s2Rows[rIdx], 'left');
                    totalGross += getGross(s2Rows[rIdx], 'right');
                }
            }

            const row125 = s2Rows[row125Idx];
            const newGross = totalGross > 0 ? totalGross.toFixed(2) : '';

            if (row125.left.gross !== newGross) {
                s2Rows[row125Idx] = {
                    ...row125,
                    left: {
                        ...row125.left,
                        gross: newGross,
                        quantityReadonly: true,
                        unitPriceReadonly: true,
                        grossReadonly: true
                    }
                };
                sectionChanged = true;
            }
        }

        if (sectionChanged) {
            this.section2Rows = s2Rows;
        }
    }

    applySection3Calculations() {
        this.applySection1Calculations();
        this.applySection2Calculations();

        const s3Rows = [...this.section3Rows];
        let sectionChanged = false;

        const findIdx = (excelRowNum) => {
            return s3Rows.findIndex(r => r.excelRow === excelRowNum);
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

        const updateRowSide = (idx, side, updates) => {
            const row = s3Rows[idx];
            let changed = false;
            for (let key in updates) {
                if (row[side][key] !== updates[key]) {
                    changed = true;
                    break;
                }
            }
            if (changed) {
                s3Rows[idx] = {
                    ...row,
                    [side]: { ...row[side], ...updates }
                };
                sectionChanged = true;
            }
        };

        // Get common dependencies
        const row22 = this.section1Rows.find(r => r.excelRow === 22);
        const headcountFromS1 = row22 ? (parseFloat(row22.left.quantity) || 0) : 0;

        // 1. ROW 182: TOTAL LABOR HRS.
        const row182Idx = findIdx(182);
        if (row182Idx !== -1) {
            let sum = 0;
            for (let i = 159; i <= 180; i++) {
                const rIdx = findIdx(i);
                if (rIdx !== -1) sum += getGross(s3Rows[rIdx], 'left');
            }
            updateRowSide(row182Idx, 'left', {
                gross: sum > 0 ? sum.toFixed(2) : '',
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 2. ROW 184: LABOR (FM+ 7TH PERIOD) - depends on 182
        const row184Idx = findIdx(184);
        if (row184Idx !== -1 && row182Idx !== -1) {
            const laborHrs = getGross(s3Rows[row182Idx], 'left');
            const unitPrice = getUnit(s3Rows[row184Idx], 'left');
            
            const newQtyRaw = laborHrs > 0 ? laborHrs.toFixed(2) : '';
            const newQty = (laborHrs > 0 && s3Rows[row184Idx].left.isWholeNumberQuantity) ? Math.round(laborHrs).toString() : newQtyRaw;
            
            updateRowSide(row184Idx, 'left', {
                quantityRaw: newQtyRaw,
                quantity: newQty,
                quantityReadonly: true,
                gross: (laborHrs > 0 && unitPrice > 0) ? (laborHrs * unitPrice).toFixed(2) : '',
                grossReadonly: true,
                size: (headcountFromS1 > 0 && laborHrs > 0) ? (laborHrs / headcountFromS1).toFixed(2) : ''
            });
        }

        // 3. ROW 185: ENGINEERING HALF HOUR HEAD
        const row185Idx = findIdx(185);
        if (row185Idx !== -1) {
            const qty = getQty(s3Rows[row185Idx], 'left');
            const unitPrice = getUnit(s3Rows[row185Idx], 'left');
            
            updateRowSide(row185Idx, 'left', {
                gross: (qty > 0 && unitPrice > 0) ? (qty * unitPrice).toFixed(2) : '',
                grossReadonly: true,
                size: (headcountFromS1 > 0 && qty > 0) ? (qty / headcountFromS1).toFixed(2) : ''
            });
        }

        // 4. ROW 157: HEADCOUNT
        const row157Idx = findIdx(157);
        if (row157Idx !== -1) {
            updateRowSide(row157Idx, 'left', {
                size: headcountFromS1 > 0 ? headcountFromS1.toFixed(2) : '',
                quantity: '',
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 5. ROW 186: BIM - depends on 157
        const row186Idx = findIdx(186);
        if (row186Idx !== -1 && row157Idx !== -1) {
            const row186 = s3Rows[row186Idx];
            const row186EditableOverrides = BidWorksheetEstimate.EDITABLE_FIELD_OVERRIDES[186] || {};
            const leftOverrides = row186EditableOverrides.left || {};
            const isQuantityEditable = leftOverrides.quantity;
            const isSizeEditable = leftOverrides.size;
            
            const headcountVal = parseFloat(s3Rows[row157Idx].left.size) || 0;
            const bimQty = headcountVal / 2;
            
            const isEditingThisField = this._lastEditedCell && 
                this._lastEditedCell.excelRow === 186 && 
                this._lastEditedCell.col === 'left' && 
                this._lastEditedCell.field === 'quantity';
            
            let updatedLeft = { ...row186.left };
            
            if (!isQuantityEditable) {
                updatedLeft.quantityRaw = bimQty > 0 ? bimQty.toFixed(2) : '';
                updatedLeft.quantity = (bimQty > 0 && updatedLeft.isWholeNumberQuantity) ? Math.round(bimQty).toString() : updatedLeft.quantityRaw;
                updatedLeft.quantityUserEntered = false;
            } else if (!isEditingThisField) {
                updatedLeft.quantityRaw = bimQty > 0 ? bimQty.toFixed(2) : '';
                updatedLeft.quantity = (bimQty > 0 && updatedLeft.isWholeNumberQuantity) ? Math.round(bimQty).toString() : updatedLeft.quantityRaw;
            }
            
            updatedLeft.quantityReadonly = !isQuantityEditable;
            const currentQty = parseFloat(updatedLeft.quantityRaw) || bimQty;
            const unitPrice = getUnit(row186, 'left');
            updatedLeft.gross = (currentQty > 0 && unitPrice > 0) ? (currentQty * unitPrice).toFixed(2) : '';
            updatedLeft.grossReadonly = true;

            const isEditingThisSizeField = this._lastEditedCell && 
                this._lastEditedCell.excelRow === 186 && 
                this._lastEditedCell.col === 'left' && 
                this._lastEditedCell.field === 'size';
            
            if (!isSizeEditable) {
                updatedLeft.size = (headcountFromS1 > 0 && currentQty > 0) ? (currentQty / headcountFromS1).toFixed(2) : '';
                updatedLeft.sizeUserEntered = false;
            } else if (!isEditingThisSizeField) {
                updatedLeft.size = (headcountFromS1 > 0 && currentQty > 0) ? (currentQty / headcountFromS1).toFixed(2) : '';
            }
            
            updatedLeft.sizeReadonly = !isSizeEditable;
            updatedLeft.sizeClass = updatedLeft.sizeReadonly ? 'readonly-cell' : '';
            
            updateRowSide(row186Idx, 'left', updatedLeft);
        }

        // 6. ROW 187: FABRICATION QUARTER HOUR PER
        const row187Idx = findIdx(187);
        if (row187Idx !== -1) {
            const row187 = s3Rows[row187Idx];
            const row187EditableOverrides = BidWorksheetEstimate.EDITABLE_FIELD_OVERRIDES[187] || {};
            const leftOverrides = row187EditableOverrides.left || {};
            const isSizeEditable = leftOverrides.size;
            
            const qty = getQty(row187, 'left');
            const unitPrice = getUnit(row187, 'left');
            
            let updatedLeft = { ...row187.left };
            updatedLeft.gross = (qty > 0 && unitPrice > 0) ? (qty * unitPrice).toFixed(2) : '';
            updatedLeft.grossReadonly = true;

            const isEditingThisSizeField = this._lastEditedCell && 
                this._lastEditedCell.excelRow === 187 && 
                this._lastEditedCell.col === 'left' && 
                this._lastEditedCell.field === 'size';
            
            if (!isSizeEditable) {
                updatedLeft.size = (headcountFromS1 > 0 && qty > 0) ? (qty / headcountFromS1).toFixed(2) : '';
                updatedLeft.sizeUserEntered = false;
            } else if (!isEditingThisSizeField) {
                updatedLeft.size = (headcountFromS1 > 0 && qty > 0) ? (qty / headcountFromS1).toFixed(2) : '';
            }
            
            updatedLeft.sizeReadonly = !isSizeEditable;
            updatedLeft.sizeClass = updatedLeft.sizeReadonly ? 'readonly-cell' : '';
            
            updateRowSide(row187Idx, 'left', updatedLeft);
        }

        // 7. ROW 189: FIELD,ENG,FAB, TOTAL - depends on 184, 185, 186, 187
        const row189Idx = findIdx(189);
        if (row189Idx !== -1) {
            let sum = 0;
            for (let i = 184; i <= 188; i++) {
                const rIdx = findIdx(i);
                if (rIdx !== -1) sum += getGross(s3Rows[rIdx], 'left');
            }
            updateRowSide(row189Idx, 'left', {
                gross: sum > 0 ? sum.toFixed(2) : '',
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 8. ROW 153: FIELD, ENGR., FAB TOTAL - depends on 189
        const row153Idx = findIdx(153);
        if (row153Idx !== -1 && row189Idx !== -1) {
            const laborTotal = getGross(s3Rows[row189Idx], 'left');
            updateRowSide(row153Idx, 'right', {
                gross: laborTotal > 0 ? laborTotal.toFixed(2) : '',
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 9. ROW 133 (LEFT): TOTAL MAT'L SHT #1 & 2
        const row133Idx = findIdx(133);
        if (row133Idx !== -1) {
            const row69 = this.section1Rows.find(r => r.excelRow === 69);
            const row125 = this.section2Rows.find(r => r.excelRow === 125);

            const row69Total = row69 ? (parseFloat(row69.left.gross) || 0) : 0;
            const row125Total = row125 ? (parseFloat(row125.left.gross) || 0) : 0;

            updateRowSide(row133Idx, 'left', {
                gross: (row69Total + row125Total).toFixed(2),
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 10. ROW 155 (LEFT): GRAND TOTAL MATERIAL COST - depends on 133-153 (LEFT)
        const row155Idx = findIdx(155);
        if (row155Idx !== -1) {
            let sum = 0;
            for (let i = 133; i <= 153; i++) {
                const rIdx = findIdx(i);
                if (rIdx !== -1) sum += getGross(s3Rows[rIdx], 'left');
            }
            const newGross = sum.toFixed(2);
            updateRowSide(row155Idx, 'left', {
                gross: newGross,
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });

            // 11. ROW 133 (RIGHT) - depends on 155 (LEFT)
            if (row133Idx !== -1) {
                updateRowSide(row133Idx, 'right', {
                    gross: newGross,
                    quantityReadonly: true,
                    unitPriceReadonly: true,
                    grossReadonly: true
                });
            }
        }

        // 12. ROW 135 (RIGHT): SALES TAX 10% - depends on 133 (RIGHT)
        const row135Idx = findIdx(135);
        if (row135Idx !== -1 && row133Idx !== -1) {
            const materialTotal = getGross(s3Rows[row133Idx], 'right');
            const taxRate = getUnit(s3Rows[row135Idx], 'right') || 0;
            updateRowSide(row135Idx, 'right', {
                gross: (materialTotal * taxRate).toFixed(2),
                quantityReadonly: true,
                grossReadonly: true
            });
        }

        // 13. ROW 152 (RIGHT): MATERIAL, PRMT., EQUIP... - depends on 133-150 (RIGHT)
        const row152Idx = findIdx(152);
        if (row152Idx !== -1) {
            let sum = 0;
            for (let i = 133; i <= 150; i++) {
                const rIdx = findIdx(i);
                if (rIdx !== -1) sum += getGross(s3Rows[rIdx], 'right');
            }
            updateRowSide(row152Idx, 'right', {
                gross: sum.toFixed(2),
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 14. ROW 160 (RIGHT): TOTAL DIRECT COST - depends on 152-158 (RIGHT)
        const row160Idx = findIdx(160);
        if (row160Idx !== -1) {
            let sum = 0;
            for (let i = 152; i <= 158; i++) {
                const rIdx = findIdx(i);
                if (rIdx !== -1) sum += getGross(s3Rows[rIdx], 'right');
            }
            updateRowSide(row160Idx, 'right', {
                gross: sum > 0 ? sum.toFixed(2) : '',
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 15. ROW 161 (RIGHT): %OVERHEAD - depends on 160
        const row161Idx = findIdx(161);
        if (row161Idx !== -1 && row160Idx !== -1) {
            const directCost = getGross(s3Rows[row160Idx], 'right');
            const unitPrice = getUnit(s3Rows[row161Idx], 'right') || 0.15;
            updateRowSide(row161Idx, 'right', {
                quantity: directCost > 0 ? directCost.toFixed(2) : '',
                quantityReadonly: true,
                gross: directCost > 0 ? (directCost * unitPrice).toFixed(2) : '',
                grossReadonly: true
            });
        }

        // 16. ROW 163 (RIGHT): SUBTOTAL - depends on 160, 161
        const row163Idx = findIdx(163);
        if (row163Idx !== -1 && row160Idx !== -1 && row161Idx !== -1) {
            const directCost = getGross(s3Rows[row160Idx], 'right');
            const overhead = getGross(s3Rows[row161Idx], 'right');
            updateRowSide(row163Idx, 'right', {
                gross: (directCost + overhead).toFixed(2),
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 17. ROW 164 (RIGHT): %GAIN - depends on 163
        const row164Idx = findIdx(164);
        if (row164Idx !== -1 && row163Idx !== -1) {
            const subtotal = getGross(s3Rows[row163Idx], 'right');
            const unitPrice = getUnit(s3Rows[row164Idx], 'right') || 0.15;
            updateRowSide(row164Idx, 'right', {
                quantity: subtotal > 0 ? subtotal.toFixed(2) : '',
                quantityReadonly: true,
                gross: subtotal > 0 ? (subtotal * unitPrice).toFixed(2) : '',
                grossReadonly: true
            });
        }

        // 18. ROW 166 (RIGHT): TOTAL QUOTE PRICE - depends on 163, 164
        const row166Idx = findIdx(166);
        if (row166Idx !== -1 && row163Idx !== -1 && row164Idx !== -1) {
            const subtotal = getGross(s3Rows[row163Idx], 'right');
            const gain = getGross(s3Rows[row164Idx], 'right');
            updateRowSide(row166Idx, 'right', {
                gross: (subtotal + gain).toFixed(2),
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 19. ROW 167 (RIGHT): PRICE MINUS SP... - depends on 166, 144, 145, 150, 151, 152
        const row167Idx = findIdx(167);
        if (row167Idx !== -1 && row166Idx !== -1) {
            const quotePrice = getGross(s3Rows[row166Idx], 'right');
            const i144 = getGross(s3Rows[findIdx(144)], 'left');
            const i145 = getGross(s3Rows[findIdx(145)], 'left');
            const i150 = getGross(s3Rows[findIdx(150)], 'left');
            const i151 = getGross(s3Rows[findIdx(151)], 'left');
            const i152 = getGross(s3Rows[findIdx(152)], 'left');

            const result = quotePrice - i144 - i145 - i150 - i151 - i152;
            updateRowSide(row167Idx, 'right', {
                gross: result.toFixed(2),
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 20. ROW 168: GROSS MARGIN - depends on 161, 164, 166
        const row168Idx = findIdx(168);
        if (row168Idx !== -1 && row161Idx !== -1 && row164Idx !== -1 && row166Idx !== -1) {
            const overhead = getGross(s3Rows[row161Idx], 'right');
            const gain = getGross(s3Rows[row164Idx], 'right');
            const quotePrice = getGross(s3Rows[row166Idx], 'right');

            const newQty = quotePrice > 0 ? ((overhead + gain) / quotePrice).toFixed(4) : '';
            updateRowSide(row168Idx, 'right', {
                quantity: newQty,
                quantityReadonly: true,
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 21. ROW 169: BOND AMOUNT - depends on 166
        const row169Idx = findIdx(169);
        if (row169Idx !== -1 && row166Idx !== -1) {
            const totalQuote = getGross(s3Rows[row166Idx], 'right');
            const bondCalc = (totalQuote / 1000) * 1.5;
            updateRowSide(row169Idx, 'right', {
                quantity: (bondCalc < 100 ? 100 : bondCalc).toFixed(2),
                quantityReadonly: true,
                gross: ((totalQuote / 1000) * 12).toFixed(2),
                unitPriceReadonly: true,
                grossReadonly: true
            });
        }

        // 22. ROW 173: MATERIAL PER HEAD - depends on 155, 151, 144, 152, 145, 150
        const row173Idx = findIdx(173);
        if (row173Idx !== -1) {
            const i155 = getGross(s3Rows[findIdx(155)], 'left');
            const i144 = getGross(s3Rows[findIdx(144)], 'left');
            const i145 = getGross(s3Rows[findIdx(145)], 'left');
            const i150 = getGross(s3Rows[findIdx(150)], 'left');
            const i151 = getGross(s3Rows[findIdx(151)], 'left');
            const i152 = getGross(s3Rows[findIdx(152)], 'left');

            const materialCost = i155 - i151 - i144 - i152 - i145 - i150;
            
            updateRowSide(row173Idx, 'right', {
                quantity: materialCost.toFixed(2),
                quantityReadonly: true,
                unitPrice: headcountFromS1 > 0 ? headcountFromS1.toFixed(2) : '',
                unitPriceReadonly: true,
                gross: headcountFromS1 > 0 ? (materialCost / headcountFromS1).toFixed(2) : '',
                grossReadonly: true
            });
        }

        // 23. ROW 174: DIRECT COST PER HEAD - depends on 160
        const row174Idx = findIdx(174);
        if (row174Idx !== -1 && row160Idx !== -1) {
            const directCost = getGross(s3Rows[row160Idx], 'right');
            
            updateRowSide(row174Idx, 'right', {
                quantity: directCost > 0 ? directCost.toFixed(2) : '',
                quantityReadonly: true,
                unitPrice: headcountFromS1 > 0 ? headcountFromS1.toFixed(2) : '',
                unitPriceReadonly: true,
                gross: headcountFromS1 > 0 ? (directCost / headcountFromS1).toFixed(2) : '',
                grossReadonly: true
            });
        }

        // 24. ROW 175: BUILDING SQ. FOOTAGE
        const row175Idx = findIdx(175);
        if (row175Idx !== -1) {
            const qty = getQty(s3Rows[row175Idx], 'right');
            updateRowSide(row175Idx, 'right', {
                unitPrice: headcountFromS1 > 0 ? headcountFromS1.toFixed(2) : '',
                unitPriceReadonly: true,
                gross: (qty > 0 && headcountFromS1 > 0) ? (qty / headcountFromS1).toFixed(2) : '',
                grossReadonly: true
            });
        }

        // 25. ROW 176: SALES COST PER HEAD - depends on 166
        const row176Idx = findIdx(176);
        if (row176Idx !== -1 && row166Idx !== -1) {
            const quotePrice = getGross(s3Rows[row166Idx], 'right');
            const i144 = getGross(s3Rows[findIdx(144)], 'left');
            const i145 = getGross(s3Rows[findIdx(145)], 'left');
            const i150 = getGross(s3Rows[findIdx(150)], 'left');
            const i151 = getGross(s3Rows[findIdx(151)], 'left');
            const i152 = getGross(s3Rows[findIdx(152)], 'left');

            const salesCost = quotePrice - i144 - i145 - i150 - i151 - i152;
            
            updateRowSide(row176Idx, 'right', {
                quantity: salesCost.toFixed(2),
                quantityReadonly: true,
                unitPrice: headcountFromS1 > 0 ? headcountFromS1.toFixed(2) : '',
                unitPriceReadonly: true,
                gross: headcountFromS1 > 0 ? (salesCost / headcountFromS1).toFixed(2) : '',
                grossReadonly: true
            });
        }

        // 26. ROW 178: COST PER SQUARE FOOT - depends on 167, 175
        const row178Idx = findIdx(178);
        if (row178Idx !== -1 && row167Idx !== -1 && row175Idx !== -1) {
            const netPrice = getGross(s3Rows[row167Idx], 'right');
            const sqFootage = getQty(s3Rows[row175Idx], 'right');

            updateRowSide(row178Idx, 'right', {
                gross: sqFootage > 0 ? (netPrice / sqFootage).toFixed(2) : '',
                grossReadonly: true,
                quantityReadonly: true,
                unitPriceReadonly: true,
                quantityClass: 'readonly-cell',
                unitPriceClass: 'col-unit readonly-cell'
            });
        }

        if (sectionChanged) {
            this.section3Rows = s3Rows;
        }
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
                    quantity: parseFloat(row.left.quantityRaw || row.left.quantity) || 0,
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
                    quantity: parseFloat(row.right.quantityRaw || row.right.quantity) || 0,
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