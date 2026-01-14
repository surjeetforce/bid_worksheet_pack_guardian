import { LightningElement, api, track, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getSheet2Items from '@salesforce/apex/BidWorksheetUndergroundController.getSheet2Items';
import saveSheet from '@salesforce/apex/BidWorksheetUndergroundController.saveSheet';
import loadLatestSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestSheet';
import loadVersionById from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById';

export default class BidWorksheetUndergroundTwo extends LightningElement {
    @api recordId;
    @api sheetNumber = 2;
    
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

    // ========================================
    // ROW NUMBER CONSTANTS
    // ========================================
    static ROW_SHEET1_TOTAL = 71;
    static ROW_SALES_TAX = 73;
    static ROW_CARTAGE = 74;
    static ROW_MATERIAL_EQUIP_SUBTOTAL = 90;
    static ROW_FIELD_ENG_FAB_TOTAL_RIGHT = 91;
    static ROW_GRAND_TOTAL_MATERIAL = 92;
    static ROW_LABOR_FACTOR = 94;
    static ROW_SUBTOTAL_BEFORE_GAIN = 96;
    static ROW_OVERHEAD = 97;
    static ROW_SUBTOTAL_AFTER_OVERHEAD = 99;
    static ROW_GAIN = 100;
    static ROW_TOTAL_UNDERGROUND_PRICE = 102;
    static ROW_COMMENTS_LABEL = 104;
    static ROW_COMMENTS_INPUT = 105;
    static ROW_EXCLUSIONS_LABEL = 115;
    static ROW_TOTAL_LABOR_HOURS = 120;
    static ROW_LABOR_COST = 122;
    static ROW_ENGINEERING_COST = 123;
    static ROW_FABRICATION_COST = 124;
    static ROW_FIELD_ENG_FAB_TOTAL_LEFT = 126;
    
    // Row ranges
    static MATERIAL_ITEMS_START = 72;
    static MATERIAL_ITEMS_END = 89;
    static ADDITIONAL_RIGHT_ITEMS_START = 75;
    static ADDITIONAL_RIGHT_ITEMS_END = 88;
    static LABOR_HOURS_START = 96;
    static LABOR_HOURS_END = 119;

    @track tableRows = [];
    @track isLoading = true;
    @track isSaving = false;

    @track grandTotalMaterial = '0.00';
    @track salesTax = '0.00';
    @track cartage = '0.00';
    @track totalLaborHours = '0.00';
    @track laborCost = '0.00';
    @track engineeringCost = '0.00';
    @track fabricationCost = '0.00';
    @track fieldEngFabTotal = '0.00';
    @track assessments = '0.00';
    @track laborFactor = '0.00';
    @track materialEquipSubtotal = '0.00';
    @track subtotalBeforeGain = '0.00';
    @track gain = '0.00';
    @track subtotalAfterGain = '0.00';
    @track overhead = '0.00';
    @track totalUndergroundPrice = '0.00';

    @track gainPercent = 0;
    @track overheadPercent = 0;

    SALES_TAX_RATE = 0.0825;
    CARTAGE_RATE = 0.03;
    ASSESSMENTS_RATE = 0.67;
    LABOR_RATE = 37.09;
    ENGINEERING_RATE = 28;
    FABRICATION_RATE = 18;
    LABOR_FACTOR_DIVISOR = 8;

    @track revisionDate = '5/4/00';
    nextRowId = 0;

    calculationTimeout = null;

    get currentDate() {
        const today = new Date();
        return today.toLocaleDateString('en-US');
    }

    connectedCallback() {
        if (!this.recordId) {
            console.error('❌ No recordId provided');
            this.showToast('Error', 'Record ID required', 'error');
            this.isLoading = false;
            return;
        }
    }


    _sheet1Subtotal = 0;

    @api
    get sheet1Subtotal() {
        return this._sheet1Subtotal;
    }

    set sheet1Subtotal(value) {
        this._sheet1Subtotal = value;
    
        // Debounce the calculation to prevent excessive recalculations
        if (this.calculationTimeout) {
            clearTimeout(this.calculationTimeout);
        }
        
        this.calculationTimeout = setTimeout(() => {
            if (this.tableRows && this.tableRows.length > 0) {
                this.calculateTotals();
            }
        }, 300);
    }
    
    @wire(getSheet2Items)
    wiredItems({ error, data }) {
        if (data) {
            // ⭐ Set flag FIRST to prevent autosave during initialization
            this._isLoadingData = true;
            
            this.initializeDataFromMetadata(data);
            this.isLoading = false;
            this.calculateTotals();
            
            // Always ensure versionIdToLoad is set - if null/empty, set to 'draft'
            // This ensures the setter fires and loads data
            const versionToLoad = (this._versionIdToLoad && this._versionIdToLoad !== '') 
                ? this._versionIdToLoad 
                : 'draft';
            
            // Reset to force load and trigger setter
            this._lastLoadedVersionId = null;
            this.versionIdToLoad = versionToLoad;
            
            // Clear flag after initialization completes (loadSavedSheet sets its own flag, so this is a backup)
            // The flag will be cleared by loadSavedSheet's applyLoadedData, but we set a timeout as backup
            setTimeout(() => {
                if (this._isLoadingData) {
                    this._isLoadingData = false;
                }
            }, 1500); // Give enough time for loadSavedSheet to complete

        } else if (error) {
            console.error('❌ Error loading Sheet #2 metadata:', error);
            this.showToast('Error', 'Failed to load sheet configuration: ' + (error.body ? error.body.message : error.message), 'error');
            this.isLoading = false;
            this._isLoadingData = false; // Clear flag on error
        }
    }

    initializeDataFromMetadata(metadataItems) {
        this.tableRows = metadataItems.map((item, index) =>
            this.createRowFromMetadata(item, index)
        );
        this.nextRowId = this.tableRows.length;
    }

    createRowFromMetadata(data, id) {
        const rowId = id !== null ? id : this.nextRowId++;
        const excelRow = data.excelRow || null;

        const leftHasDescription = !!(data.left.description && data.left.description.trim());
        const rightHasDescription = !!(data.right.description && data.right.description.trim());

        // ========================================
        // CALCULATED ROWS - All fields readonly
        // ========================================
        const calculatedLeftRows = [
            BidWorksheetUndergroundTwo.ROW_SHEET1_TOTAL,
            BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL,
            BidWorksheetUndergroundTwo.ROW_LABOR_FACTOR,
            BidWorksheetUndergroundTwo.ROW_TOTAL_LABOR_HOURS,
            BidWorksheetUndergroundTwo.ROW_LABOR_COST,
            // ROW_ENGINEERING_COST and ROW_FABRICATION_COST removed - amount fields should be editable
            BidWorksheetUndergroundTwo.ROW_FIELD_ENG_FAB_TOTAL_LEFT
        ];

        const calculatedRightRows = [
            BidWorksheetUndergroundTwo.ROW_SHEET1_TOTAL,
            BidWorksheetUndergroundTwo.ROW_SALES_TAX,
            BidWorksheetUndergroundTwo.ROW_CARTAGE,
            BidWorksheetUndergroundTwo.ROW_MATERIAL_EQUIP_SUBTOTAL,
            BidWorksheetUndergroundTwo.ROW_FIELD_ENG_FAB_TOTAL_RIGHT,
            BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL,
            BidWorksheetUndergroundTwo.ROW_LABOR_FACTOR,
            BidWorksheetUndergroundTwo.ROW_SUBTOTAL_BEFORE_GAIN,
            BidWorksheetUndergroundTwo.ROW_OVERHEAD,
            BidWorksheetUndergroundTwo.ROW_SUBTOTAL_AFTER_OVERHEAD,
            BidWorksheetUndergroundTwo.ROW_GAIN,
            BidWorksheetUndergroundTwo.ROW_TOTAL_UNDERGROUND_PRICE,
            BidWorksheetUndergroundTwo.ROW_COMMENTS_LABEL,
            BidWorksheetUndergroundTwo.ROW_EXCLUSIONS_LABEL
        ];

        // ========================================
        // USER-EDITABLE UNIT PRICE
        // ========================================
        const editableUnitPriceRows = [
            BidWorksheetUndergroundTwo.ROW_SALES_TAX,
            BidWorksheetUndergroundTwo.ROW_CARTAGE,
            BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL,
            BidWorksheetUndergroundTwo.ROW_OVERHEAD,
            BidWorksheetUndergroundTwo.ROW_GAIN,
            94 // TRAVEL/SUBSISTANCE
        ];

        // ========================================
        // COMMENT ROWS
        // ========================================
        const isLeftCommentRow = data.left.isCommentRow || false;
        const isRightCommentRow = data.right.isCommentRow || false;

        // ========================================
        // DETERMINE READONLY STATE
        // ========================================
        const isLeftCalculated = excelRow && calculatedLeftRows.includes(excelRow);
        const isRightCalculated = excelRow && calculatedRightRows.includes(excelRow);
        const canEditRightUnitPrice = excelRow && editableUnitPriceRows.includes(excelRow);

        // LEFT SIDE
        const leftDescriptionReadonly = isLeftCommentRow || !!data.left.description;
        // Amount: readonly if calculated OR if no description yet
        const leftAmountReadonly = isLeftCommentRow || isLeftCalculated || !leftHasDescription;
        // Unit price: readonly if calculated OR if no description, BUT LABOR unit price should be editable
        const isLaborRow = excelRow === BidWorksheetUndergroundTwo.ROW_LABOR_COST;
        const leftUnitPriceReadonly = isLeftCommentRow || (isLeftCalculated && !isLaborRow) || !leftHasDescription;

        // RIGHT SIDE
        const rightDescriptionReadonly = isRightCommentRow || !!data.right.description;
        const rightAmountReadonly = isRightCommentRow || isRightCalculated || !rightHasDescription;
        const rightUnitPriceReadonly = isRightCommentRow
            ? true
            : canEditRightUnitPrice
                ? false
                : (isRightCalculated || !rightHasDescription);

        // Initialize overhead percentage from metadata
        if (excelRow === BidWorksheetUndergroundTwo.ROW_OVERHEAD) {
            const initialOverheadRate = parseFloat(data.right.unitPrice);
            if (!isNaN(initialOverheadRate)) {
                this.overheadPercent = initialOverheadRate * 100;
            }
        }

        // Initialize gain percentage from metadata
        if (excelRow === BidWorksheetUndergroundTwo.ROW_GAIN) {
            const initialGainRate = parseFloat(data.right.unitPrice);
            if (!isNaN(initialGainRate)) {
                this.gainPercent = initialGainRate * 100;
            }
        }

        // Initialize rates from metadata
        if (excelRow === BidWorksheetUndergroundTwo.ROW_SALES_TAX) {
            const taxRate = parseFloat(data.right.unitPrice);
            if (!isNaN(taxRate)) {
                this.SALES_TAX_RATE = taxRate;
            }
        }

        if (excelRow === BidWorksheetUndergroundTwo.ROW_CARTAGE) {
            const cartageRate = parseFloat(data.right.unitPrice);
            if (!isNaN(cartageRate)) {
                this.CARTAGE_RATE = cartageRate;
            }
        }

        if (excelRow === BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL) {
            const assessmentRate = parseFloat(data.right.unitPrice);
            if (!isNaN(assessmentRate)) {
                this.ASSESSMENTS_RATE = assessmentRate;
            }
        }

        return {
            id: rowId,
            excelRow: excelRow,
            left: {
                description: data.left.description || '',
                descriptionReadonly: leftDescriptionReadonly,
                amount: (data.left.amount !== undefined && data.left.amount !== null)
                    ? data.left.amount
                    : '',
                amountReadonly: leftAmountReadonly,
                unitPrice: (data.left.unitPrice !== undefined && data.left.unitPrice !== null)
                    ? data.left.unitPrice
                    : '',
                unitPriceReadonly: leftUnitPriceReadonly,
                gross: '',
                unitPriceFieldType: data.left.unitPriceFieldType || 'Currency',
                grossFieldType: data.left.grossFieldType || 'Currency',
                isPercentageUnitPrice: (data.left.unitPriceFieldType || 'Currency') === 'Percentage',
                isNumberUnitPrice: (data.left.unitPriceFieldType || 'Currency') === 'Number',
                isHoursGross: (data.left.grossFieldType || 'Currency') === 'Hours',
                isNumberGross: (data.left.grossFieldType || 'Currency') === 'Number',
                isTotalRow: data.left.isTotalRow || false,
                isCommentRow: isLeftCommentRow,
                isCalculated: isLeftCalculated,
                descriptionClass: (data.left.isIndent ? 'description-cell indent' : 'description-cell') + (leftDescriptionReadonly ? ' readonly-cell' : ''),
                amountClass: leftAmountReadonly ? 'readonly-cell' : '',
                unitPriceClass: leftUnitPriceReadonly ? 'col-unit readonly-cell' : 'col-unit'
            },
            right: {
                description: data.right.description || '',
                descriptionReadonly: rightDescriptionReadonly,
                amount: (data.right.amount !== undefined && data.right.amount !== null)
                    ? data.right.amount
                    : '',
                amountReadonly: rightAmountReadonly,
                unitPrice: (data.right.unitPrice !== undefined && data.right.unitPrice !== null)
                    ? data.right.unitPrice
                    : (excelRow === BidWorksheetUndergroundTwo.ROW_OVERHEAD ? '0.00' :
                        excelRow === BidWorksheetUndergroundTwo.ROW_SALES_TAX ? this.SALES_TAX_RATE.toFixed(2) :
                            excelRow === BidWorksheetUndergroundTwo.ROW_CARTAGE ? this.CARTAGE_RATE.toFixed(2) :
                                excelRow === BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL ? this.ASSESSMENTS_RATE.toFixed(2) :
                                    excelRow === BidWorksheetUndergroundTwo.ROW_GAIN ? '0.00' :
                                        ''),
                unitPriceReadonly: rightUnitPriceReadonly,
                gross: '',
                unitPriceFieldType: data.right.unitPriceFieldType || 'Currency',
                grossFieldType: data.right.grossFieldType || 'Currency',
                isPercentageUnitPrice: (data.right.unitPriceFieldType || 'Currency') === 'Percentage',
                isNumberUnitPrice: (data.right.unitPriceFieldType || 'Currency') === 'Number',
                isHoursGross: (data.right.grossFieldType || 'Currency') === 'Hours',
                isNumberGross: (data.right.grossFieldType || 'Currency') === 'Number',
                isTotalRow: false,
                isCommentRow: isRightCommentRow,
                isCalculated: isRightCalculated,
                descriptionClass: (data.right.isIndent ? 'description-cell indent' : 'description-cell') + (rightDescriptionReadonly ? ' readonly-cell' : ''),
                amountClass: rightAmountReadonly ? 'readonly-cell' : '',
                unitPriceClass: rightUnitPriceReadonly ? 'col-unit readonly-cell' : 'col-unit'
            }
        };
    }

    handleCellChange(event) {
        const rowId = parseInt(event.target.dataset.row);
        const col = event.target.dataset.col;
        const field = event.target.dataset.field;
        const value = event.target.value;

        const rowIndex = this.tableRows.findIndex(row => row.id === rowId);
        if (rowIndex !== -1) {
            const updatedRow = { ...this.tableRows[rowIndex] };
            updatedRow[col] = { ...updatedRow[col], [field]: value };

            const excelRow = updatedRow.excelRow;

            // Validate unitPrice changes for overhead and gain
            if (field === 'unitPrice' && col === 'right') {
                const numericValue = parseFloat(value);

                if (excelRow === BidWorksheetUndergroundTwo.ROW_OVERHEAD ||
                    excelRow === BidWorksheetUndergroundTwo.ROW_GAIN) {
                    if (numericValue < 0 || numericValue > 1) {
                        this.showToast('Warning', 'Percentage must be between 0 and 1 (e.g., 0.15 = 15%)', 'warning');
                        return;
                    }
                }

                // Capture rate changes
                if (excelRow === BidWorksheetUndergroundTwo.ROW_SALES_TAX) {
                    if (!isNaN(numericValue)) {
                        this.SALES_TAX_RATE = numericValue;
                    }
                } else if (excelRow === BidWorksheetUndergroundTwo.ROW_CARTAGE) {
                    if (!isNaN(numericValue)) {
                        this.CARTAGE_RATE = numericValue;
                    }
                } else if (excelRow === BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL) {
                    if (!isNaN(numericValue)) {
                        this.ASSESSMENTS_RATE = numericValue;
                    }
                } else if (excelRow === BidWorksheetUndergroundTwo.ROW_LABOR_FACTOR) {
                } else if (excelRow === BidWorksheetUndergroundTwo.ROW_OVERHEAD) {
                    if (!isNaN(numericValue)) {
                        this.overheadPercent = numericValue * 100;
                    }
                } else if (excelRow === BidWorksheetUndergroundTwo.ROW_GAIN) {
                    if (!isNaN(numericValue)) {
                        this.gainPercent = numericValue * 100;
                    }
                }
            }

            // Recalculate gross if amount or unitPrice changed
            if (field === 'amount' || field === 'unitPrice') {
                updatedRow[col].gross = this.calculateGross(
                    updatedRow[col].amount,
                    updatedRow[col].unitPrice
                );
            }

            // If description changed, recalculate readonly states for amount and unitPrice
            if (field === 'description') {
                const hasDescription = !!(value && value.trim());
                const isCalculated = excelRow && (
                    (col === 'left' && [
                        BidWorksheetUndergroundTwo.ROW_SHEET1_TOTAL,
                        BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL,
                        BidWorksheetUndergroundTwo.ROW_LABOR_FACTOR,
                        BidWorksheetUndergroundTwo.ROW_TOTAL_LABOR_HOURS,
                        BidWorksheetUndergroundTwo.ROW_LABOR_COST,
                        // ROW_ENGINEERING_COST and ROW_FABRICATION_COST removed - amount fields should be editable
                        BidWorksheetUndergroundTwo.ROW_FIELD_ENG_FAB_TOTAL_LEFT
                    ].includes(excelRow)) ||
                    (col === 'right' && [
                        BidWorksheetUndergroundTwo.ROW_SHEET1_TOTAL,
                        BidWorksheetUndergroundTwo.ROW_SALES_TAX,
                        BidWorksheetUndergroundTwo.ROW_CARTAGE,
                        BidWorksheetUndergroundTwo.ROW_MATERIAL_EQUIP_SUBTOTAL,
                        BidWorksheetUndergroundTwo.ROW_FIELD_ENG_FAB_TOTAL_RIGHT,
                        BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL,
                        BidWorksheetUndergroundTwo.ROW_LABOR_FACTOR,
                        BidWorksheetUndergroundTwo.ROW_SUBTOTAL_BEFORE_GAIN,
                        BidWorksheetUndergroundTwo.ROW_OVERHEAD,
                        BidWorksheetUndergroundTwo.ROW_SUBTOTAL_AFTER_OVERHEAD,
                        BidWorksheetUndergroundTwo.ROW_GAIN,
                        BidWorksheetUndergroundTwo.ROW_TOTAL_UNDERGROUND_PRICE,
                        BidWorksheetUndergroundTwo.ROW_COMMENTS_LABEL,
                        BidWorksheetUndergroundTwo.ROW_EXCLUSIONS_LABEL
                    ].includes(excelRow))
                );
                const isCommentRow = updatedRow[col].isCommentRow || false;
                
                // Update readonly states based on new description value
                updatedRow[col].amountReadonly = isCommentRow || isCalculated || !hasDescription;
                
                // For unit price, check if it's in the editable list (right side only)
                if (col === 'right') {
                    const editableUnitPriceRows = [
                        BidWorksheetUndergroundTwo.ROW_SALES_TAX,
                        BidWorksheetUndergroundTwo.ROW_CARTAGE,
                        BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL,
                        BidWorksheetUndergroundTwo.ROW_OVERHEAD,
                        BidWorksheetUndergroundTwo.ROW_GAIN,
                        94 // TRAVEL/SUBSISTANCE
                    ];
                    const canEditUnitPrice = excelRow && editableUnitPriceRows.includes(excelRow);
                    updatedRow[col].unitPriceReadonly = isCommentRow
                        ? true
                        : canEditUnitPrice
                            ? false
                            : (isCalculated || !hasDescription);
                } else {
                    updatedRow[col].unitPriceReadonly = isCommentRow || isCalculated || !hasDescription;
                }
                
                // Update class properties to reflect readonly state
                updatedRow[col].amountClass = updatedRow[col].amountReadonly ? 'readonly-cell' : '';
                updatedRow[col].unitPriceClass = updatedRow[col].unitPriceReadonly ? 'col-unit readonly-cell' : 'col-unit';
                // Description class already includes readonly-cell if needed, but update it
                const baseDescClass = updatedRow[col].descriptionClass ? updatedRow[col].descriptionClass.replace(' readonly-cell', '') : 'description-cell';
                updatedRow[col].descriptionClass = baseDescClass + (updatedRow[col].descriptionReadonly ? ' readonly-cell' : '');
            }

            this.tableRows = [
                ...this.tableRows.slice(0, rowIndex),
                updatedRow,
                ...this.tableRows.slice(rowIndex + 1)
            ];

            this.calculateTotals();
        }
    }

    updateAllCalculatedRowDisplays() {
        const findRow = (excelRowNum) => this.tableRows.find(r => r.excelRow === excelRowNum);

        const row71 = findRow(BidWorksheetUndergroundTwo.ROW_SHEET1_TOTAL);
        if (row71) {
            row71.left.gross = this.sheet1Subtotal || '0.00';
            row71.right.gross = this.grandTotalMaterial;
        }

        const row73 = findRow(BidWorksheetUndergroundTwo.ROW_SALES_TAX);
        if (row73) {
            row73.right.amount = this.grandTotalMaterial;
            row73.right.unitPrice = this.SALES_TAX_RATE.toFixed(4);
            row73.right.gross = this.salesTax;
        }

        const row74 = findRow(BidWorksheetUndergroundTwo.ROW_CARTAGE);
        if (row74) {
            row74.right.amount = this.grandTotalMaterial;
            row74.right.unitPrice = this.CARTAGE_RATE.toFixed(4);
            row74.right.gross = this.cartage;
        }

        const row90 = findRow(BidWorksheetUndergroundTwo.ROW_MATERIAL_EQUIP_SUBTOTAL);
        if (row90) {
            row90.right.gross = this.materialEquipSubtotal;
        }

        const row91 = findRow(BidWorksheetUndergroundTwo.ROW_FIELD_ENG_FAB_TOTAL_RIGHT);
        if (row91) {
            row91.right.gross = this.fieldEngFabTotal;
        }

        const row92 = findRow(BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL);
        if (row92) {
            row92.right.amount = this.fieldEngFabTotal;
            row92.right.unitPrice = this.ASSESSMENTS_RATE.toFixed(2);
            row92.right.gross = this.assessments;
            row92.left.gross = this.grandTotalMaterial;
        }

        const row94 = findRow(BidWorksheetUndergroundTwo.ROW_LABOR_FACTOR);
        if (row94) {
            const laborFactorAmount = (parseFloat(this.totalLaborHours) / this.LABOR_FACTOR_DIVISOR).toFixed(2);
            row94.right.amount = laborFactorAmount;
            row94.right.gross = this.laborFactor;
        }

        const row96 = findRow(BidWorksheetUndergroundTwo.ROW_SUBTOTAL_BEFORE_GAIN);
        if (row96) {
            row96.right.gross = this.subtotalBeforeGain;
        }

        const row97 = findRow(BidWorksheetUndergroundTwo.ROW_OVERHEAD);
        if (row97) {
            row97.right.amount = this.subtotalBeforeGain;
            row97.right.unitPrice = (this.overheadPercent / 100).toFixed(4);
            row97.right.gross = this.overhead;
        }

        const row99 = findRow(BidWorksheetUndergroundTwo.ROW_SUBTOTAL_AFTER_OVERHEAD);
        if (row99) {
            row99.right.gross = this.subtotalAfterGain;
        }

        const row100 = findRow(BidWorksheetUndergroundTwo.ROW_GAIN);
        if (row100) {
            row100.right.amount = this.subtotalAfterGain;
            row100.right.unitPrice = (this.gainPercent / 100).toFixed(4);
            row100.right.gross = this.gain;
        }

        const row102 = findRow(BidWorksheetUndergroundTwo.ROW_TOTAL_UNDERGROUND_PRICE);
        if (row102) {
            row102.right.gross = this.totalUndergroundPrice;
        }

        const row120 = findRow(BidWorksheetUndergroundTwo.ROW_TOTAL_LABOR_HOURS);
        if (row120) {
            row120.left.gross = this.totalLaborHours;
        }

        const row122 = findRow(BidWorksheetUndergroundTwo.ROW_LABOR_COST);
        if (row122) {
            row122.left.amount = this.totalLaborHours;
            row122.left.gross = this.laborCost;
        }

        const row123 = findRow(BidWorksheetUndergroundTwo.ROW_ENGINEERING_COST);
        if (row123) {
            row123.left.gross = this.engineeringCost;
        }

        const row124 = findRow(BidWorksheetUndergroundTwo.ROW_FABRICATION_COST);
        if (row124) {
            row124.left.gross = this.fabricationCost;
        }

        const row126 = findRow(BidWorksheetUndergroundTwo.ROW_FIELD_ENG_FAB_TOTAL_LEFT);
        if (row126) {
            row126.left.gross = this.fieldEngFabTotal;
        }

        this.tableRows = [...this.tableRows];
    }

    calculateTotals() {
        this.calculateGrandTotalMaterialCost();
        this.calculateSalesTaxAndCartage();
        this.calculateTotalLaborHours();
        this.calculateLaborCosts();
        this.calculateFieldEngFabTotal();
        this.calculateAssessments();
        this.calculateMaterialEquipSubtotal();
        this.calculateLaborFactor();
        this.calculateSubtotalBeforeGain();
        this.calculateOverhead();
        this.calculateSubtotalAfterGain();
        this.calculateGain();
        this.calculateFinalTotal();

        this.updateAllCalculatedRowDisplays();

        this.notifyParent();
        
        // Notify parent for auto-save (only if not loading data)
        if (!this._isLoadingData) {
            this.notifyParentForAutoSave();
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

    calculateGrandTotalMaterialCost() {
        const sheet1Total = parseFloat(this.sheet1Subtotal) || 0;
        let materialItemsTotal = 0;

        for (let excelRow = BidWorksheetUndergroundTwo.MATERIAL_ITEMS_START;
            excelRow <= BidWorksheetUndergroundTwo.MATERIAL_ITEMS_END;
            excelRow++) {
            const row = this.tableRows.find(r => r.excelRow === excelRow);
            if (row) {
                const gross = parseFloat(row.left.gross) || 0;
                materialItemsTotal += gross;
            }
        }

        this.grandTotalMaterial = (sheet1Total + materialItemsTotal).toFixed(2);
    }

    calculateSalesTaxAndCartage() {
        const grandTotal = parseFloat(this.grandTotalMaterial) || 0;
        this.salesTax = (grandTotal * this.SALES_TAX_RATE).toFixed(2);
        this.cartage = (grandTotal * this.CARTAGE_RATE).toFixed(2);
    }

    calculateTotalLaborHours() {
        let totalHours = 0;

        for (let excelRow = BidWorksheetUndergroundTwo.LABOR_HOURS_START;
            excelRow <= BidWorksheetUndergroundTwo.LABOR_HOURS_END;
            excelRow++) {
            const row = this.tableRows.find(r => r.excelRow === excelRow);
            if (row) {
                const hours = parseFloat(row.left.gross) || 0;
                totalHours += hours;
            }
        }

        this.totalLaborHours = totalHours.toFixed(2);
    }

    calculateLaborCosts() {
        const totalHours = parseFloat(this.totalLaborHours) || 0;
        
        // Get LABOR row and use actual unit price if available, otherwise use default rate
        const row122 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_LABOR_COST);
        const laborUnitPrice = parseFloat(row122?.left.unitPrice) || this.LABOR_RATE;
        this.laborCost = (totalHours * laborUnitPrice).toFixed(2);

        const row123 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_ENGINEERING_COST);
        const row124 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_FABRICATION_COST);

        const engHours = parseFloat(row123?.left.amount) || 0;
        const fabHours = parseFloat(row124?.left.amount) || 0;
        
        // Use the actual unit price from the row if available, otherwise use the default rate
        const engUnitPrice = parseFloat(row123?.left.unitPrice) || this.ENGINEERING_RATE;
        const fabUnitPrice = parseFloat(row124?.left.unitPrice) || this.FABRICATION_RATE;

        this.engineeringCost = (engHours * engUnitPrice).toFixed(2);
        this.fabricationCost = (fabHours * fabUnitPrice).toFixed(2);

    }

    calculateFieldEngFabTotal() {
        const labor = parseFloat(this.laborCost) || 0;
        const engineering = parseFloat(this.engineeringCost) || 0;
        const fabrication = parseFloat(this.fabricationCost) || 0;
        this.fieldEngFabTotal = (labor + engineering + fabrication).toFixed(2);
    }

    calculateAssessments() {
        const total = parseFloat(this.fieldEngFabTotal) || 0;
        this.assessments = (total * this.ASSESSMENTS_RATE).toFixed(2);
    }

    calculateMaterialEquipSubtotal() {
        const grandTotal = parseFloat(this.grandTotalMaterial) || 0;
        const salesTax = parseFloat(this.salesTax) || 0;
        const cartage = parseFloat(this.cartage) || 0;

        let additionalItems = 0;
        for (let excelRow = BidWorksheetUndergroundTwo.ADDITIONAL_RIGHT_ITEMS_START;
            excelRow <= BidWorksheetUndergroundTwo.ADDITIONAL_RIGHT_ITEMS_END;
            excelRow++) {
            const row = this.tableRows.find(r => r.excelRow === excelRow);
            if (row) {
                const gross = parseFloat(row.right.gross) || 0;
                additionalItems += gross;
            }
        }

        this.materialEquipSubtotal = (grandTotal + salesTax + cartage + additionalItems).toFixed(2);
    }

    calculateLaborFactor() {
        const totalHours = parseFloat(this.totalLaborHours) || 0;
        const factor = totalHours / this.LABOR_FACTOR_DIVISOR;

        const row94 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_LABOR_FACTOR);
        const userRate = parseFloat(row94?.right.unitPrice) || 0;

        this.laborFactor = (factor * userRate).toFixed(2);
    }

    calculateSubtotalBeforeGain() {
        const materialEquip = parseFloat(this.materialEquipSubtotal) || 0;
        const assessments = parseFloat(this.assessments) || 0;
        const laborFactor = parseFloat(this.laborFactor) || 0;
        const fieldEngFabTotal = parseFloat(this.fieldEngFabTotal) || 0;
        this.subtotalBeforeGain = (materialEquip + assessments + laborFactor + fieldEngFabTotal).toFixed(2);
    }

    calculateOverhead() {
        const subtotalBeforeGain = parseFloat(this.subtotalBeforeGain) || 0;
        this.overhead = (subtotalBeforeGain * (this.overheadPercent / 100)).toFixed(2);
    }

    calculateSubtotalAfterGain() {
        const beforeGain = parseFloat(this.subtotalBeforeGain) || 0;
        const overhead = parseFloat(this.overhead) || 0;
        this.subtotalAfterGain = (beforeGain + overhead).toFixed(2);
    }

    calculateGain() {
        const subtotal = parseFloat(this.subtotalAfterGain) || 0;
        this.gain = (subtotal * (this.gainPercent / 100)).toFixed(2);
    }

    calculateFinalTotal() {
        const subtotalAfterGain = parseFloat(this.subtotalAfterGain) || 0;
        const gain = parseFloat(this.gain) || 0;
        this.totalUndergroundPrice = (subtotalAfterGain + gain).toFixed(2);
    }

    calculateGross(amount, unitPrice) {
        const amountNum = parseFloat(amount) || 0;
        const priceNum = parseFloat(unitPrice) || 0;
        const gross = amountNum * priceNum;
        return gross > 0 ? gross.toFixed(2) : '';
    }

    notifyParent() {
        const event = new CustomEvent('sheetupdate', {
            detail: {
                sheetNumber: this.sheetNumber,
                totalPrice: this.totalUndergroundPrice
            }
        });
        this.dispatchEvent(event);
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

    collectFormData() {
        return {
            sheetNumber: this.sheetNumber,
            sheet1Subtotal: this.sheet1Subtotal,
            grandTotalMaterial: this.grandTotalMaterial,
            salesTax: this.salesTax,
            salesTaxRate: this.SALES_TAX_RATE,
            cartage: this.cartage,
            cartageRate: this.CARTAGE_RATE,
            totalLaborHours: this.totalLaborHours,
            laborCost: this.laborCost,
            engineeringCost: this.engineeringCost,
            fabricationCost: this.fabricationCost,
            fieldEngFabTotal: this.fieldEngFabTotal,
            assessments: this.assessments,
            assessmentsRate: this.ASSESSMENTS_RATE,
            laborFactor: this.laborFactor,
            materialEquipSubtotal: this.materialEquipSubtotal,
            subtotalBeforeGain: this.subtotalBeforeGain,
            gainPercent: this.gainPercent,
            gain: this.gain,
            subtotalAfterGain: this.subtotalAfterGain,
            overheadPercent: this.overheadPercent,
            overhead: this.overhead,
            totalUndergroundPrice: this.totalUndergroundPrice,
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

        if (!this.tableRows || this.tableRows.length === 0) {
            return;
        }

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

            this.applyLoadedData(savedState);

        } catch (error) {
            const errorMessage = error?.body?.message || error?.message || String(error);

            if (errorMessage.includes('not found') || errorMessage.includes('No ContentVersion') || errorMessage.includes('List has no rows')) {
                return;
            }
            this.logError('Load saved sheet failed', error);
            if (!errorMessage.includes('not found') && !errorMessage.includes('No ContentVersion')) {
                this.showToast('Error', 'Failed to load saved sheet: ' + errorMessage, 'error');
            }
        }
    }

    applyLoadedData(data) {
        if (!data) {
            return;
        }

        if (!this.tableRows || this.tableRows.length === 0) {
            return;
        }

        // Set flag to prevent autosave during data loading
        this._isLoadingData = true;

        try {
            const sheetData = data.sheet2 || data;

            const updateIfExists = (savedValue, currentValue) => {
                return (savedValue !== undefined && savedValue !== null) ? savedValue : currentValue;
            };

            this.sheetNumber = updateIfExists(sheetData.sheetNumber, this.sheetNumber);
            this.sheet1Subtotal = updateIfExists(sheetData.sheet1Subtotal, this.sheet1Subtotal);
            this.grandTotalMaterial = updateIfExists(sheetData.grandTotalMaterial, this.grandTotalMaterial);
            this.salesTax = updateIfExists(sheetData.salesTax, this.salesTax);
            this.SALES_TAX_RATE = updateIfExists(sheetData.salesTaxRate, this.SALES_TAX_RATE);
            this.cartage = updateIfExists(sheetData.cartage, this.cartage);
            this.CARTAGE_RATE = updateIfExists(sheetData.cartageRate, this.CARTAGE_RATE);
            this.totalLaborHours = updateIfExists(sheetData.totalLaborHours, this.totalLaborHours);
            this.laborCost = updateIfExists(sheetData.laborCost, this.laborCost);
            this.engineeringCost = updateIfExists(sheetData.engineeringCost, this.engineeringCost);
            this.fabricationCost = updateIfExists(sheetData.fabricationCost, this.fabricationCost);
            this.fieldEngFabTotal = updateIfExists(sheetData.fieldEngFabTotal, this.fieldEngFabTotal);
            this.assessments = updateIfExists(sheetData.assessments, this.assessments);
            this.ASSESSMENTS_RATE = updateIfExists(sheetData.assessmentsRate, this.ASSESSMENTS_RATE);
            this.laborFactor = updateIfExists(sheetData.laborFactor, this.laborFactor);
            this.materialEquipSubtotal = updateIfExists(sheetData.materialEquipSubtotal, this.materialEquipSubtotal);
            this.subtotalBeforeGain = updateIfExists(sheetData.subtotalBeforeGain, this.subtotalBeforeGain);
            this.gainPercent = updateIfExists(sheetData.gainPercent, 0);
            this.gain = updateIfExists(sheetData.gain, this.gain);
            this.subtotalAfterGain = updateIfExists(sheetData.subtotalAfterGain, this.subtotalAfterGain);
            this.overheadPercent = updateIfExists(sheetData.overheadPercent, 0);
            this.overhead = updateIfExists(sheetData.overhead, this.overhead);
            this.totalUndergroundPrice = updateIfExists(sheetData.totalUndergroundPrice, this.totalUndergroundPrice);

            if (sheetData.lineItems && Array.isArray(sheetData.lineItems) && sheetData.lineItems.length > 0) {
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

                // Update display rates
                const row73 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_SALES_TAX);
                if (row73) row73.right.unitPrice = this.SALES_TAX_RATE.toFixed(4);

                const row74 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_CARTAGE);
                if (row74) row74.right.unitPrice = this.CARTAGE_RATE.toFixed(4);

                const row92 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_GRAND_TOTAL_MATERIAL);
                if (row92) row92.right.unitPrice = this.ASSESSMENTS_RATE.toFixed(2);

                const row97 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_OVERHEAD);
                if (row97) row97.right.unitPrice = (this.overheadPercent / 100).toFixed(4);

                const row100 = this.tableRows.find(r => r.excelRow === BidWorksheetUndergroundTwo.ROW_GAIN);
                if (row100) row100.right.unitPrice = (this.gainPercent / 100).toFixed(4);

                this.tableRows = [...this.tableRows];
            }

            setTimeout(() => {
                this.calculateTotals();
                
                // Clear flag after a delay to allow DOM to settle and prevent autosave
                setTimeout(() => {
                    this._isLoadingData = false;
                }, 500);
            }, 50);

        } catch (err) {
            console.error('❌ [APPLY] Error applying loaded data:', err);
            this.logError('Apply loaded data failed', err);
            this.showToast('Error', 'Failed to apply loaded sheet data: ' + err.message, 'error');
            this._isLoadingData = false; // Clear flag on error
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