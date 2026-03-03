import { LightningElement, api, track } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import loadDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.loadDesignWorksheet';
import loadLatestDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestDesignWorksheet';
import loadVersionById_Design from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById_Design';

export default class DesignWorksheet extends LightningElement {
    @api recordId;
    @api isReadOnly = false;
    @track isLoading = true;

    // Version control properties
    _versionIdToLoad = null;
    _lastLoadedVersionId = null;
    _isLoadingData = false;
    _isUserEditing = false;
    _editingTimeout = null;

    get designContainerClass() {
        return this.isReadOnly ? 'slds-p-around_medium design-sheet readonly-sheet' : 'slds-p-around_medium design-sheet';
    }

    @api
    get currentJobName() {
        return this.formData ? this.formData.jobName : '';
    }

    @api
    get versionIdToLoad() {
        return this._versionIdToLoad;
    }

    set versionIdToLoad(value) {
        const oldValue = this._versionIdToLoad;
        this._versionIdToLoad = value;
        
        // Only reload if value changed and formData is initialized
        if (oldValue !== value && this.formData) {
            // Don't reload if user is actively editing
            if (!this._isUserEditing) {
                this.loadSavedData();
            }
        }
    }

    @track formData = {
        // Job Information
        jobName: '',
        jobAddress: '',
        description: '',

        // Service Territory & Floors
        serviceTerritory_SC: false,
        serviceTerritory_NC: false,
        numberOfFloors: '',
        penthouse: '',
        bidPlanDate: '',

        // Project Requirements (SINGLE CHECKBOX: checked = YES, unchecked = NO)
        residentialRates: false,
        localHire: false,
        apprenticePercent: false,
        textura: false,
        certifiedPayroll: false,
        bond: false,

        // OCIP (SINGLE CHECKBOX: checked = DEDUCT, unchecked = ADD LATER)
        ocipDeduct: false,
        ocupAmount: '',

        // Other Requirements
        marketRecovery: false,
        bimRequired: false,

        // Permit Fees (SINGLE CHECKBOX: checked = INCLUDED, unchecked = EXCLUDED)
        permitFeesIncluded: false,
        permitAmount: '',

        // Pre-Construction
        ammr: '',
        preApp: '',
        fpeRequired: '',
        ahj: '',

        // System Design
        hazardClassification: '',
        densityRequired: false,
        atticSprinklersRequired: false,

        headTypesAttic: '',
        headTypesCeiling: '',
        standpipeQty: '',

        tempSpRequired: false,

        // Fire Pump
        firePumpGpm: '',
        firePumpPsi: '',
        firePumpVoltage: '',
        firePumpTransferSwitch: '',

        // Materials (SINGLE CHECKBOX each)
        buyAmerican: false,
        steelPipe: false,
        importPipe: false,
        dynaflow: false,
        cpvc: false,

        // Head Details
        ceilingHeads: '',
        atticHeads: '',
        headTypeColorCeiling: '',
        headTypeColorAttic: '',

        // Metraflex Loops
        metraflexLoops: false,
        metraflexSize: '',
        metraflexQty: '',

        // Flexheads
        flexheads: false,
        flexheadsQty: '',

        // FDC
        fdcCount: '',

        // FDC Type 
        fdcType_FreeStanding: false,
        fdcType_2Way: false,
        fdcType_3Way: false,
        fdcType_4Way: false,
        fdcType_SP: false,
        fdcType_Flush: false,
        fdcType_CH: false,
        fdcType_PolBR: false,

        // Underground Scope (multiple checkboxes allowed)
        trenching: false,
        sawcut: false,
        import: false,
        export: false,
        pave: false,

        // Backflow (SINGLE CHECKBOX: checked = DDCV, unchecked = REDUCED PRESSURE)
        backflowDDCV: false,

        // Equipment Rental (SINGLE CHECKBOX: checked = YES)
        scissorLifts: false,
        scissorLiftsMonths: '',
        scissorLiftsSize: '',

        boomLifts: false,
        boomLiftsMonths: '',
        boomLiftsSize: '',

        forklift: false,
        forkliftMonths: '',
        forkliftSize: '',

        // Labor Hours
        designHours: '',
        fieldHours: '',
        fab: '',
        fm200: '',

        // Comments
        comments: '',

        // Labels
        label_serviceTerritory: 'SERVICE TERRITORY / FLOORS',
        label_bidPlanDate: 'BID PLAN DATE',
        label_residentialRates: 'RESIDENTIAL RATES',
        label_localHire: 'LOCAL HIRE',
        label_apprenticePercent: 'APPRENTICE %',
        label_textura: 'TEXTURA',
        label_certifiedPayroll: 'CERTIFIED PAYROLL',
        label_bond: 'BOND',
        label_ocip: 'OCIP DEDUCT OR ADD LATER',
        label_marketRecovery: 'MARKET RECOVERY',
        label_bimRequired: 'BIM REQUIRED',
        label_permitFees: 'PERMIT FEES',
        label_ammr: 'AMMR',
        label_preApp: 'PRE APP',
        label_fpeRequired: 'FPE REQUIRED',
        label_ahj: 'AHJ',
        label_hazardClassification: 'HAZARD CLASSIFICATION',
        label_densityRequired: 'DENSITY REQUIRED?',
        label_atticSprinklers: 'ATTIC SPRINKLERS REQ?',
        label_headTypes: 'HEAD TYPES',
        label_standpipeQty: 'STANDPIPE QTY AND HOSE VALVES',
        label_tempSpRequired: 'TEMP SP REQUIRED?',
        label_firePump: 'FIRE PUMP',
        label_buyAmerican: 'BUY AMERICAN?',
        label_steelPipe: 'STEEL PIPE',
        label_importPipe: 'IMPORT PIPE',
        label_dynaflow: 'DYNAFLOW, DYNATHREAD OK?',
        label_cpvc: 'CPVC',
        label_ceilingHeads: '# CEILING HEADS',
        label_atticHeads: '# ATTIC HEADS',
        label_headDetails: 'TYPE AND COLOR OF HEADS',
        label_metraflexLoops: 'METRAFLEX LOOPS',
        label_flexheads: 'FLEXHEADS',
        label_fdcCount: '# OF FDC',
        label_fdcType: 'TYPE FDC',
        label_undergroundScope: 'UNDERGROUND SCOPE',
        label_backflow: 'BACKFLOW',
        label_scissorLifts: 'SCISSOR LIFTS',
        label_boomLifts: 'BOOM LIFTS',
        label_forklift: 'FORKLIFT',
        label_designHours: 'DESIGN HOURS INC BIM',
        label_fieldHours: 'FIELD HOURS',
        label_fab: 'FAB',
        label_fm200: 'FM 200',
        label_comments: 'COMMENTS/OTHER'
    };

    get currentDate() {
        const today = new Date();
        return today.toLocaleDateString('en-US');
    }

    connectedCallback() {
        if (!this.recordId) {
            this.recordId = '006VF00000I9RJaYAN'; // Fallback for testing
        }
        this.loadSavedData();
    }

    async loadSavedData() {
        if (!this.recordId) {
            return;
        }

        // Don't load if user is actively editing
        if (this._isUserEditing) {
            return;
        }

        // Set loading flag to prevent autosave during load
        this._isLoadingData = true;

        try {
            let savedData;
            
            // If versionIdToLoad is set, load that specific version
            if (this.versionIdToLoad && this.versionIdToLoad !== 'draft') {
                const base64Data = await loadVersionById_Design({ versionId: this.versionIdToLoad });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                }
                this._lastLoadedVersionId = this.versionIdToLoad;
            } else {
                // Otherwise, load latest (autosave or most recent)
                const base64Data = await loadLatestDesignWorksheet({ opportunityId: this.recordId });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                } else {
                    // Fallback to old method for backward compatibility
                    savedData = await loadDesignWorksheet({ opportunityId: this.recordId });
                }
                this._lastLoadedVersionId = 'draft';
            }

            if (savedData) {
                const data = typeof savedData === 'string' ? JSON.parse(savedData) : savedData;

                // Load form data
                if (data.formData) {
                    this.formData = { ...this.formData, ...data.formData };
                }

            } else {
            }
        } catch (error) {
        } finally {
            this.isLoading = false;
            // Clear loading flag after a delay to allow DOM to settle
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

    handleInputChange(event) {
        // Set flag to indicate user is actively editing
        this._isUserEditing = true;
        if (this._editingTimeout) {
            clearTimeout(this._editingTimeout);
        }
        this._editingTimeout = setTimeout(() => {
            this._isUserEditing = false;
        }, 1000);

        const field = event.target.dataset.field;
        const value = event.target.value;
        
        // Notify parent for autosave (only if not loading data)
        if (!this._isLoadingData) {
            this.notifyParentForAutoSave();
        }
        this.formData[field] = value;
    }

    handleKeyDown(event) {

        const allowDecimal = event.target.dataset.isAllowedDecimal === 'true';
        const allowedKeys = [
            'Backspace',
            'Delete',
            'Tab',
            'ArrowLeft',
            'ArrowRight',
            'End'
        ];

        // Allow Ctrl / Cmd shortcuts (copy, paste, select all)
        if (event.ctrlKey || event.metaKey) {
            return;
        }

        // Allow special keys
        if (allowedKeys.includes(event.key)) {
            return;
        }

        // Allow numbers only (0–9)
        if (/^[0-9]$/.test(event.key)) {
            return;
        }

        // Allow decimal point (only once)
        if (allowDecimal && event.key === '.') {
            if (value.includes('.')) {
                event.preventDefault();
            }
            return;
        }

        // Block everything else
        event.preventDefault();

    }

    handleCheckboxChange(event) {
        // Set flag to indicate user is actively editing
        this._isUserEditing = true;
        if (this._editingTimeout) {
            clearTimeout(this._editingTimeout);
        }
        this._editingTimeout = setTimeout(() => {
            this._isUserEditing = false;
        }, 1000);

        const field = event.target.dataset.field;
        const checked = event.target.checked;
        this.formData[field] = checked;
        
        // Notify parent for autosave (only if not loading data)
        if (!this._isLoadingData) {
            this.notifyParentForAutoSave();
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


    get isTypeFdcCheckboxDisabled() {
        let fields = [
            'fdcType_FreeStanding',
            'fdcType_2Way',
            'fdcType_3Way',
            'fdcType_4Way',
            'fdcType_SP',
            'fdcType_Flush',
            'fdcType_CH',
            'fdcType_PolBR'
        ]

        let selectedfield = fields.find(field => this.formData[field] == true);

        let disabledMap = {};
        fields.forEach(field => {
            disabledMap[field] = selectedfield != null && selectedfield != field
        })
        return disabledMap;
    }

    @api
    async saveSheet() {
        const data = {
            formData: this.formData,
            savedDate: new Date().toISOString()
        };

        return data;
    }

    showToast(title, message, variant) {
        this.dispatchEvent(new ShowToastEvent({ title, message, variant }));
    }

    /**
     * Handle Download CSV button click
     */
    handleDownloadCSV() {
        if (!this.formData) {
            this.showToast('Info', 'No data to download', 'info');
            return;
        }

        const csvContent = this.convertToCSV();
        const dateStr = new Date().toISOString().split('T')[0];
        const filename = `Design_Worksheet_${dateStr}.csv`;
        
        this.downloadCSVFile(csvContent, filename);
    }

    /**
     * Convert Design Worksheet data to CSV string
     */
    convertToCSV() {
        const lines = [];
        
        // Basic Info
        lines.push('JOB START UP / DESIGN NARRATIVE');
        lines.push(`JOB NAME,${this.formatCSVValue(this.formData.jobName)}`);
        lines.push(`JOB ADDRESS,${this.formatCSVValue(this.formData.jobAddress)}`);
        lines.push(`DESCRIPTION,${this.formatCSVValue(this.formData.description)}`);
        lines.push('');

        // Form Fields
        lines.push('FIELD,VALUE');
        
        // Helper to add checkbox/text values
        const addField = (label, val) => {
            let displayVal = val;
            if (typeof val === 'boolean') {
                displayVal = val ? 'YES' : 'NO';
            }
            lines.push(`${this.formatCSVValue(label)},${this.formatCSVValue(displayVal)}`);
        };

        addField('SERVICE TERRITORY SC', this.formData.serviceTerritory_SC);
        addField('SERVICE TERRITORY NC', this.formData.serviceTerritory_NC);
        addField('# FLOORS', this.formData.numberOfFloors);
        addField('PENTHOUSE', this.formData.penthouse);
        addField('BID PLAN DATE', this.formData.bidPlanDate);
        addField('RESIDENTIAL RATES', this.formData.residentialRates);
        addField('LOCAL HIRE', this.formData.localHire);
        addField('APPRENTICE %', this.formData.apprenticePercent);
        addField('TEXTURA', this.formData.textura);
        addField('CERTIFIED PAYROLL', this.formData.certifiedPayroll);
        addField('BOND', this.formData.bond);
        addField('OCIP DEDUCT', this.formData.ocipDeduct);
        addField('OCIP AMOUNT', this.formData.ocupAmount);
        addField('MARKET RECOVERY', this.formData.marketRecovery);
        addField('BIM REQUIRED', this.formData.bimRequired);
        addField('PERMIT FEES INCLUDED', this.formData.permitFeesIncluded);
        addField('PERMIT AMOUNT', this.formData.permitAmount);
        addField('AMMR', this.formData.ammr);
        addField('PRE APP', this.formData.preApp);
        addField('FPE REQUIRED', this.formData.fpeRequired);
        addField('AHJ', this.formData.ahj);
        addField('HAZARD CLASSIFICATION', this.formData.hazardClassification);
        addField('DENSITY REQUIRED', this.formData.densityRequired);
        addField('ATTIC SPRINKLERS REQUIRED', this.formData.atticSprinklersRequired);
        addField('HEAD TYPES ATTIC', this.formData.headTypesAttic);
        addField('HEAD TYPES CEILING', this.formData.headTypesCeiling);
        addField('STANDPIPE QTY AND HOSE VALVES', this.formData.standpipeQty);
        addField('TEMP SP REQUIRED', this.formData.tempSpRequired);
        addField('FIRE PUMP GPM', this.formData.firePumpGpm);
        addField('FIRE PUMP PSI', this.formData.firePumpPsi);
        addField('FIRE PUMP VOLTAGE', this.formData.firePumpVoltage);
        addField('FIRE PUMP TRANSFER SWITCH', this.formData.firePumpTransferSwitch);
        addField('BUY AMERICAN', this.formData.buyAmerican);
        addField('STEEL PIPE', this.formData.steelPipe);
        addField('IMPORT PIPE', this.formData.importPipe);
        addField('DYNAFLOW OK', this.formData.dynaflow);
        addField('CPVC', this.formData.cpvc);
        addField('# CEILING HEADS', this.formData.ceilingHeads);
        addField('# ATTIC HEADS', this.formData.atticHeads);
        addField('TYPE/COLOR CEILING', this.formData.headTypeColorCeiling);
        addField('TYPE/COLOR ATTIC', this.formData.headTypeColorAttic);
        addField('METRAFLEX LOOPS', this.formData.metraflexLoops);
        addField('METRAFLEX SIZE', this.formData.metraflexSize);
        addField('METRAFLEX QTY', this.formData.metraflexQty);
        addField('FLEXHEADS', this.formData.flexheads);
        addField('FLEXHEADS QTY', this.formData.flexheadsQty);
        addField('# OF FDC', this.formData.fdcCount);
        addField('FDC TYPE FREE STANDING', this.formData.fdcType_FreeStanding);
        addField('FDC TYPE 2 WAY', this.formData.fdcType_2Way);
        addField('FDC TYPE 3 WAY', this.formData.fdcType_3Way);
        addField('FDC TYPE 4 WAY', this.formData.fdcType_4Way);
        addField('FDC TYPE SP', this.formData.fdcType_SP);
        addField('FDC TYPE FLUSH', this.formData.fdcType_Flush);
        addField('FDC TYPE CH', this.formData.fdcType_CH);
        addField('FDC TYPE POL BR', this.formData.fdcType_PolBR);
        addField('TRENCHING', this.formData.trenching);
        addField('SAWCUT', this.formData.sawcut);
        addField('IMPORT', this.formData.import);
        addField('EXPORT', this.formData.export);
        addField('PAVE', this.formData.pave);
        addField('BACKFLOW DDCV', this.formData.backflowDDCV);
        addField('SCISSOR LIFTS', this.formData.scissorLifts);
        addField('SCISSOR LIFTS MONTHS', this.formData.scissorLiftsMonths);
        addField('SCISSOR LIFTS SIZE', this.formData.scissorLiftsSize);
        addField('BOOM LIFTS', this.formData.boomLifts);
        addField('BOOM LIFTS MONTHS', this.formData.boomLiftsMonths);
        addField('BOOM LIFTS SIZE', this.formData.boomLiftsSize);
        addField('FORKLIFT', this.formData.forklift);
        addField('FORKLIFT MONTHS', this.formData.forkliftMonths);
        addField('FORKLIFT SIZE', this.formData.forkliftSize);
        addField('DESIGN HOURS', this.formData.designHours);
        addField('FIELD HOURS', this.formData.fieldHours);
        addField('FAB', this.formData.fab);
        addField('FM 200', this.formData.fm200);
        addField('COMMENTS', this.formData.comments);

        return lines.join('\n');
    }

    /**
     * Format a single value for CSV
     */
    formatCSVValue(val) {
        if (val === null || val === undefined) return '';
        let stringVal = String(val);
        stringVal = stringVal.replace(/"/g, '""');
        if (stringVal.includes(',') || stringVal.includes('\n') || stringVal.includes('"')) {
            stringVal = `"${stringVal}"`;
        }
        return stringVal;
    }

    /**
     * Trigger browser download of CSV file
     */
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