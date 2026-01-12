import { LightningElement, api, track } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import loadDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.loadDesignWorksheet';
import loadLatestDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestDesignWorksheet';
import loadVersionById_Design from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById_Design';

export default class DesignWorksheet extends LightningElement {
    @api recordId;
    @track isLoading = true;

    // Version control properties
    _versionIdToLoad = null;
    _lastLoadedVersionId = null;
    _isLoadingData = false;
    _isUserEditing = false;
    _editingTimeout = null;

    @api
    get versionIdToLoad() {
        return this._versionIdToLoad;
    }

    set versionIdToLoad(value) {
        const oldValue = this._versionIdToLoad;
        this._versionIdToLoad = value;
        
        // Only reload if value changed and formData is initialized
        if (oldValue !== value && this.formData) {
            console.log(`📍 [Design] versionIdToLoad changed from ${oldValue} to ${value}`);
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
        firePumpGpm: false,
        firePumpPsi: false,
        firePumpVoltage: false,
        firePumpTransferSwitch: false,

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
        comments: ''
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
            console.log('❌ [LOAD Design] No recordId, skipping load');
            return;
        }

        // Don't load if user is actively editing
        if (this._isUserEditing) {
            console.log('📍 [LOAD Design] User is editing, skipping load to prevent data loss');
            return;
        }

        // Set loading flag to prevent autosave during load
        this._isLoadingData = true;

        try {
            let savedData;
            
            // If versionIdToLoad is set, load that specific version
            if (this.versionIdToLoad && this.versionIdToLoad !== 'draft') {
                console.log('🔍 [LOAD Design] Loading specific version:', this.versionIdToLoad);
                const base64Data = await loadVersionById_Design({ versionId: this.versionIdToLoad });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                }
                this._lastLoadedVersionId = this.versionIdToLoad;
            } else {
                // Otherwise, load latest (autosave or most recent)
                console.log('🔍 [LOAD Design] Loading latest (draft)');
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

                console.log('✅ [LOAD Design] Design Worksheet data loaded');
            } else {
                console.log('⚠️ [LOAD Design] No saved Design Worksheet data found');
            }
        } catch (error) {
            console.log('⚠️ [LOAD Design] No saved data found or error:', error);
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
            console.error('Decode data failed:', err);
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
        console.log(`Field updated: ${field} = ${value}`);
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
        console.log(`Checkbox updated: ${field} = ${checked ? 'YES/TRUE' : 'NO/FALSE'}`);
        
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
        console.log('selectedfield :- ', selectedfield);

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

        console.log('Saving Design Worksheet:', data);
        return data;
    }

    showToast(title, message, variant) {
        this.dispatchEvent(new ShowToastEvent({ title, message, variant }));
    }
}