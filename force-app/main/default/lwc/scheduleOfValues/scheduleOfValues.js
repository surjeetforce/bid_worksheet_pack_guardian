import { LightningElement, track, api, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getSOVItems from '@salesforce/apex/BidWorksheetUndergroundController.getSOVItems';
import loadSOVSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadSOVSheet';
import loadLatestSOVSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestSOVSheet';
import loadVersionById_SOV from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById_SOV';
import loadDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.loadDesignWorksheet';
import loadLatestDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestDesignWorksheet';

const FLOOR_COLUMNS = [
    { key: 'total', label: 'TOTAL', isTotal: true },
    { key: 'ground', label: 'GROUND / 1ST FLOOR' },
    { key: 'second', label: '2ND FLOOR' },
    { key: 'third', label: '3RD FLOOR' },
    { key: 'fourth', label: '4TH FLOOR' },
    { key: 'fifth', label: '5TH FLOOR' },
    { key: 'sixth', label: '6TH FLOOR' },
    { key: 'seventh', label: '7TH FLOOR' },
    { key: 'eighth', label: '8TH FLOOR' },
    { key: 'ninth', label: '9TH FLOOR' },
    { key: 'tenth', label: '10TH FLOOR' }
];

export default class ScheduleOfValues extends LightningElement {
    @api recordId;

    currencyFormatter;
    floorColumns = FLOOR_COLUMNS;

    @track isLoading = true;
    @track jobOverview = {
        jobName: '',
        numberOfBuildings: 1
    };

    @track summaryRows = [];
    @track metadataSummaryRows = [];
    @track metadataBuildingRows = [];
    @track buildings = [];

    // Version control properties
    _versionIdToLoad = null;
    _lastLoadedVersionId = null;
    _isLoadingData = false;
    _isUserEditing = false;
    _editingTimeout = null;
    _isInitializing = true; // Flag to prevent autosave during initialization
    _isRecalculating = false; // Flag to prevent infinite calculation loops
    _suppressAutoSave = false; // Flag to suppress autosave during programmatic updates

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
        const rowsReady = this.summaryRows.length > 0 || this.buildings.length > 0;
        
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
            // Don't reload if user is actively editing (but allow during initial load)
            // During initial load (_isLoadingData is true), we should load even if _isUserEditing is true
            if (!this._isUserEditing || this._isLoadingData) {
                this.loadSavedData();
            } else {
            }
        } else {
        }
    }

    constructor() {
        super();
        this.currencyFormatter = new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD',
            minimumFractionDigits: 2
        });
    }

    connectedCallback() {
        if (!this.recordId) {
            this.recordId = '006VF00000I9RJaYAN'; // Fallback for testing
        }

        // Load metadata first, then try to load saved data
        this.loadMetadata();
    }

    async loadMetadata() {
        // Metadata will be loaded via @wire
        // After metadata loads, we'll load saved data
    }

    @wire(getSOVItems)
    wiredSOVItems({ error, data }) {
        if (data) {
            // Set loading flag first to prevent autosave during initialization
            this._isLoadingData = true;

            this.metadataSummaryRows = data.summaryRows || [];
            this.metadataBuildingRows = data.buildingRows || [];

            // Initialize with metadata
            this.summaryRows = this.createSummaryRows(this.metadataSummaryRows);
            this.buildings = this.createBuildings(this.jobOverview.numberOfBuildings);

            this.refreshBuildingCalculations();

            // Now try to load saved data
            setTimeout(() => {
                // Always ensure versionIdToLoad is set - if null/empty, set to 'draft'
                // This ensures the setter fires and loads data
                const versionToLoad = (this._versionIdToLoad && this._versionIdToLoad !== '') 
                    ? this._versionIdToLoad 
                    : 'draft';
                
                // Reset to force load and trigger setter
                this._lastLoadedVersionId = null;
                this.versionIdToLoad = versionToLoad;
                
                // Clear flags after initialization completes
                setTimeout(() => {
                    this._isLoadingData = false;
                    this._isInitializing = false;
                }, 1500);
            }, 100);
        } else if (error) {
            console.error('Error loading SOV metadata:', error);
            this.showToast('Error', 'Failed to load SOV metadata', 'error');
            this.isLoading = false;
            this._isLoadingData = false;
        }
    }

    async loadSavedData() {
        if (!this.recordId) {
            return;
        }

        // Don't load if rows haven't been initialized from metadata yet
        if (this.summaryRows.length === 0 && this.buildings.length === 0) {
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
                const base64Data = await loadVersionById_SOV({ versionId: this.versionIdToLoad });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                }
                this._lastLoadedVersionId = this.versionIdToLoad;
            } else {
                // Otherwise, load latest (autosave or most recent)
                const base64Data = await loadLatestSOVSheet({ opportunityId: this.recordId });
                if (base64Data) {
                    savedData = this.decodeData(base64Data);
                } else {
                    // Fallback to old method for backward compatibility
                    savedData = await loadSOVSheet({ opportunityId: this.recordId });
                }
                this._lastLoadedVersionId = 'draft';
            }

            if (savedData) {
                const data = typeof savedData === 'string' ? JSON.parse(savedData) : savedData;
                // Apply saved data
                if (data.jobOverview) {
                    this.jobOverview.numberOfBuildings = data.jobOverview.numberOfBuildings;
                }

                if (data.summaryRows && Array.isArray(data.summaryRows) && data.summaryRows.length > 0) {
                    // Ensure summaryRows have the correct structure with displayValue
                    this.summaryRows = data.summaryRows.map(row => ({
                        ...row,
                        displayValue: row.displayValue || this.formatCurrency(row.value || 0)
                    }));
                } else {
                }

                if (data.buildings && Array.isArray(data.buildings) && data.buildings.length > 0) {
                    this.buildings = data.buildings;
                } else {
                }

                // Only refresh calculations if we have buildings data
                // This ensures summaryRows are preserved if they were loaded
                if (this.buildings && this.buildings.length > 0) {
                    this.refreshBuildingCalculations();
                } else {
                    // If no buildings, just recalculate summary totals from existing summaryRows
                    this.recalculateSummaryTotals();
                    this.notifyParent();
                }
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

    // Populate SOV Job Name from Design Worksheet file when SOV does not have a value.
    @api
    async populateJobNameFromDesign(manualJobName) {
        if (manualJobName) {
            this.jobOverview = {
                ...this.jobOverview,
                jobName: manualJobName
            };
            return;
        }

        if (!this.recordId) {
            return;
        }
        try {
            this.isLoading = true;
            let savedData;

            const base64Data = await loadLatestDesignWorksheet({ opportunityId: this.recordId });
            if (base64Data) {
                savedData = this.decodeData(base64Data);
            } else {
                savedData = await loadDesignWorksheet({ opportunityId: this.recordId });
            }

            if (!savedData) {
                return;
            }

            const data = typeof savedData === 'string' ? JSON.parse(savedData) : savedData;
            const designJobName =
                data &&
                data.formData &&
                typeof data.formData.jobName === 'string' &&
                data.formData.jobName.trim() !== ''
                    ? data.formData.jobName
                    : null;
            
            if (!designJobName) {
                return;
            }
            this.jobOverview = {
                ...this.jobOverview,
                jobName: designJobName
            };
        } catch (error) {
            console.error('Error populating job name from Design Worksheet', error);
        } finally {
            this.isLoading = false;
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

    get currentDate() {
        const today = new Date();
        return today.toLocaleDateString('en-US');
    }

    get totalSummaryDisplay() {
        const total = this.summaryRows.reduce((sum, row) => sum + (row.value || 0), 0);
        return this.formatCurrency(total);
    }

    get totalFromAllBuildingsDisplay() {
        const total = this.buildings.reduce((sum, building) => sum + (building.total || 0), 0);
        return this.formatCurrency(total);
    }

    get visibleBuildings() {
        const count = this.jobOverview.numberOfBuildings || 1;
        return this.buildings.slice(0, count);
    }

    createSummaryRows(metadataRows) {
        if (!metadataRows) metadataRows = [];
        return metadataRows.map(item => ({
            id: item.id,
            label: item.label,
            mirrorsBuilding: item.mirrorsBuilding ?? true,
            value: item.defaultValue || 0,
            displayValue: this.formatCurrency(item.defaultValue || 0)
        }));
    }

    createBuildings(desiredCount = 1) {
        const count = Math.max(1, parseInt(desiredCount, 10) || 1);

        // Use building-specific metadata rows if present; otherwise fall back to summaryRows
        const source = (this.metadataBuildingRows && this.metadataBuildingRows.length)
            ? this.metadataBuildingRows
            : this.metadataSummaryRows;

        const buildings = [];
        for (let i = 1; i <= count; i++) {
            const letter = this.columnNameFromIndex(i);
            const label = `${letter} | ${i}`;
            buildings.push({
                id: `building-${i - 1}`,
                label,
                categoryRows: this.createCategoryRows(source),
                total: 0,
                totalDisplay: this.formatCurrency(0)
            });
        }

        return buildings;
    }

    createCategoryRows(source) {
        const floorDefaults = this.createFloorDefaults();
        return (source || []).map(category => ({
            id: category.id,
            label: category.label,
            values: { ...floorDefaults },
            total: 0,
            totalDisplay: this.formatCurrency(0)
        }));
    }

    createFloorDefaults() {
        const defaults = {};
        this.floorColumns
            .filter(column => !column.isTotal)
            .forEach(column => {
                defaults[column.key] = 0;
            });
        return defaults;
    }

    formatCurrency(value) {
        return this.currencyFormatter.format(value || 0);
    }

    refreshBuildingCalculations() {
        // Prevent infinite loops - don't recalculate if already in a calculation
        if (this._isRecalculating) {
            return;
        }
        this._isRecalculating = true;

        // Only set editing flag if not loading data (user-initiated changes)
        // Don't set flag during initialization/data loading
        if (!this._isLoadingData && !this._isInitializing) {
            this._isUserEditing = true;
            if (this._editingTimeout) {
                clearTimeout(this._editingTimeout);
            }
            this._editingTimeout = setTimeout(() => {
                this._isUserEditing = false;
            }, 1000);
        }

        const newBuildings = this.buildings.map(building => {
            const floorTotals = {};
            this.floorColumns
                .filter(column => !column.isTotal)
                .forEach(column => (floorTotals[column.key] = 0));

            const categoryRows = building.categoryRows.map(row => {
                const total = Object.values(row.values).reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
                const totalDisplay = this.formatCurrency(total);
                const valueMap = { ...row.values };
                const cells = this.floorColumns.map(column => ({
                    key: column.key,
                    isTotal: column.isTotal,
                    value: column.isTotal ? totalDisplay : valueMap[column.key] || 0,
                    displayValue: column.isTotal ? totalDisplay : null
                }));

                this.floorColumns
                    .filter(column => !column.isTotal)
                    .forEach(column => {
                        const numericValue = parseFloat(valueMap[column.key]) || 0;
                        floorTotals[column.key] += numericValue;
                    });

                return {
                    ...row,
                    total,
                    totalDisplay,
                    cells
                };
            });

            const buildingTotal = categoryRows.reduce((sum, row) => sum + row.total, 0);
            const floorSummaryCells = this.floorColumns.map(column => ({
                key: column.key,
                displayValue: column.isTotal
                    ? this.formatCurrency(buildingTotal)
                    : this.formatCurrency(floorTotals[column.key] || 0)
            }));

            return {
                ...building,
                categoryRows,
                total: buildingTotal,
                totalDisplay: this.formatCurrency(buildingTotal),
                floorSummaryCells
            };
        });

        // Only update if buildings actually changed to prevent reactive loops
        const buildingsChanged = JSON.stringify(newBuildings) !== JSON.stringify(this.buildings);
        if (buildingsChanged) {
            this.buildings = newBuildings;
        }

        this._isRecalculating = false;

        // Always recalculate summary totals when buildings change
        this.recalculateSummaryTotals();
        this.notifyParent();
        
        // Only suppress autosave if this is NOT a user-initiated change
        // User-initiated changes should trigger autosave (handlers will call notifyParentForAutoSave separately)
        if (this._isLoadingData || this._isInitializing) {
            // During loading/initialization, suppress autosave
            this._suppressAutoSave = true;
            setTimeout(() => {
                this._suppressAutoSave = false;
            }, 100);
        }
        // For user-initiated changes, don't suppress - let the handler trigger autosave
    }

    recalculateSummaryTotals() {
        // ONLY these specific rows should auto-calculate from building totals
        // Based on Excel: C5, C10, C11, C14
        const AUTO_CALC_LABELS = [
            '3D BIM (LUMP)',
            'PUMP MATERIAL',
            'PUMP ROUGH IN',
            'STANDPIPE AND FHV (BY FLOOR / BY BUILDING)'
        ];

        // Build totals maps from all buildings
        const totalsById = {};
        const totalsByLabel = {};

        (this.buildings || []).forEach(building => {
            (building.categoryRows || []).forEach(cat => {
                const id = cat.id;
                const label = (cat.label || '').toString();
                const total = parseFloat(cat.total || 0) || 0;

                if (id) totalsById[id] = (totalsById[id] || 0) + total;
                if (label) totalsByLabel[label] = (totalsByLabel[label] || 0) + total;
            });
        });

        // Update summary rows - ONLY auto-calculate the 4 specific rows
        // Make sure we have summaryRows to work with
        if (!this.summaryRows || this.summaryRows.length === 0) {
            return;
        }

        const newSummaryRows = this.summaryRows.map(row => {
            // Check if this row should auto-calculate
            const shouldAutoCalc = AUTO_CALC_LABELS.includes(row.label);

            if (!shouldAutoCalc) {
                // Keep existing manual value for all other rows
                return row;
            }

            // Auto-calculate from building totals for the 4 rows only
            let matchingTotal = 0;
            if (row.id && totalsById.hasOwnProperty(row.id)) {
                matchingTotal = totalsById[row.id];
            } else if (row.label && totalsByLabel.hasOwnProperty(row.label)) {
                matchingTotal = totalsByLabel[row.label];
            } else {
                matchingTotal = row.value || 0;
            }

            return {
                ...row,
                value: matchingTotal,
                displayValue: this.formatCurrency(matchingTotal)
            };
        });

        // Only update if values actually changed to prevent reactive loops
        const hasChanged = newSummaryRows.some((newRow, index) => {
            const oldRow = this.summaryRows[index];
            return !oldRow || newRow.value !== oldRow.value;
        });

        if (hasChanged) {
            // Only suppress autosave if this is a programmatic update (not user-initiated)
            if (this._isLoadingData || this._isInitializing) {
                this._suppressAutoSave = true;
                this.summaryRows = newSummaryRows;
                setTimeout(() => {
                    this._suppressAutoSave = false;
                }, 50);
            } else {
                // User-initiated change - allow autosave
                this.summaryRows = newSummaryRows;
            }
        }
    }

    handleSummaryValueChange(event) {
        // Set flag to indicate user is actively editing
        this._isUserEditing = true;
        if (this._editingTimeout) {
            clearTimeout(this._editingTimeout);
        }
        this._editingTimeout = setTimeout(() => {
            this._isUserEditing = false;
        }, 1000);

        if (event.target.type == 'number') {
            this.sanitizeNonNegative(event);
        }

        const rowId = event.target.dataset.rowId;
        const newValue = parseFloat(event.target.value) || 0;
        this.summaryRows = this.summaryRows.map(row =>
            row.id === rowId ? { ...row, value: newValue, displayValue: this.formatCurrency(newValue) } : row
        );

        this.notifyParent();
        
        // Notify parent for autosave (only if not loading data and not during initialization)
        if (!this._isLoadingData && !this._isInitializing) {
            this.notifyParentForAutoSave();
        }
    }

    handleBuildingValueChange(event) {
        // Set flag to indicate user is actively editing
        this._isUserEditing = true;
        if (this._editingTimeout) {
            clearTimeout(this._editingTimeout);
        }
        this._editingTimeout = setTimeout(() => {
            this._isUserEditing = false;
        }, 1000);

        if (event.target.type == 'number') {
            this.sanitizeNonNegative(event);
        }

        const { buildingId, rowId, floorKey } = event.target.dataset;
        const newValue = parseFloat(event.target.value) || 0;

        this.buildings = this.buildings.map(building => {
            if (building.id !== buildingId) {
                return building;
            }

            const categoryRows = building.categoryRows.map(row => {
                if (row.id !== rowId) {
                    return row;
                }

                const updatedValues = {
                    ...row.values,
                    [floorKey]: newValue
                };

                return { ...row, values: updatedValues };
            });

            return { ...building, categoryRows };
        });

        this.refreshBuildingCalculations();
        
        // Notify parent for autosave (only if not loading data and not during initialization)
        if (!this._isLoadingData && !this._isInitializing) {
            this.notifyParentForAutoSave();
        }
    }

    handleJobInfoChange(event) {

        if (event.target.type == 'number') {
            this.sanitizeNonNegative(event);
        }

        const field = event.target.dataset.field;
        if (!field) {
            return;
        }

        let value = event.target.value;

        if (field === 'numberOfBuildings') {
            // Allow empty string temporarily - validation will happen on blur
            // Only update buildings array if we have a valid number
            const numericValue = value === '' ? null : parseFloat(value);
            
            if (numericValue !== null && !isNaN(numericValue)) {
                const maxAllowed = 200;
                const validatedValue = Math.max(1, Math.min(maxAllowed, Math.floor(numericValue)));
                
                // Update the stored value
                this.jobOverview = { ...this.jobOverview, [field]: validatedValue };

                const currentBuildings = this.buildings || [];
                const currentCount = currentBuildings.length;

                // if new count is larger, add new buildings
                if (validatedValue > currentCount) {
                    const newBuildings = this.createBuildings(validatedValue);
                    this.buildings = [...currentBuildings, ...newBuildings.slice(currentCount)];
                }
                // if new count is smaller, slice off extra buildings
                else if (validatedValue < currentCount) {
                    this.buildings = currentBuildings.slice(0, validatedValue);
                }

                this.refreshBuildingCalculations();
            } else {
                // Empty or invalid - store as empty string but don't modify buildings array
                this.jobOverview = { ...this.jobOverview, [field]: '' };
            }
            
            // Notify parent for autosave (only if not loading data and not during initialization)
            if (!this._isLoadingData && !this._isInitializing) {
                this.notifyParentForAutoSave();
            }
            return;
        }

        // generic update for other jobOverview fields
        this.jobOverview = {
            ...this.jobOverview,
            [field]: value
        };

        // Set flag to indicate user is actively editing
        this._isUserEditing = true;
        if (this._editingTimeout) {
            clearTimeout(this._editingTimeout);
        }
        this._editingTimeout = setTimeout(() => {
            this._isUserEditing = false;
        }, 1000);

        this.notifyParent();
        
        // Notify parent for autosave (only if not loading data and not during initialization)
        if (!this._isLoadingData && !this._isInitializing) {
            this.notifyParentForAutoSave();
        }
    }

    handleJobInfoBlur(event) {
        const field = event.target.dataset.field;
        
        if (field === 'numberOfBuildings') {
            const value = event.target.value;
            const numericValue = value === '' ? null : parseFloat(value);
            const maxAllowed = 200;
            
            // Enforce minimum of 1 (default to 1 if empty or invalid)
            const validatedValue = numericValue !== null && !isNaN(numericValue) 
                ? Math.max(1, Math.min(maxAllowed, Math.floor(numericValue)))
                : 1;
            
            // Update the stored value
            this.jobOverview = { ...this.jobOverview, [field]: validatedValue };
            
            // Update input field to reflect the validated value
            event.target.value = validatedValue;

            const currentBuildings = this.buildings || [];
            const currentCount = currentBuildings.length;

            // if new count is larger, add new buildings
            if (validatedValue > currentCount) {
                const newBuildings = this.createBuildings(validatedValue);
                this.buildings = [...currentBuildings, ...newBuildings.slice(currentCount)];
            }
            // if new count is smaller, slice off extra buildings
            else if (validatedValue < currentCount) {
                this.buildings = currentBuildings.slice(0, validatedValue);
            }

            this.refreshBuildingCalculations();
            
            // Notify parent for autosave (only if not loading data and not during initialization)
            if (!this._isLoadingData && !this._isInitializing) {
                this.notifyParentForAutoSave();
            }
        }
    }

    notifyParent() {
        // Dispatch event to parent with total
        const total = this.summaryRows.reduce((sum, row) => sum + (row.value || 0), 0);
        this.dispatchEvent(new CustomEvent('sovupdate', {
            detail: { totalSOV: total.toFixed(2) }
        }));
    }

    /**
     * Notify parent component of cell change for autosave
     */
    notifyParentForAutoSave() {
        // Don't trigger autosave if we're suppressing it (during programmatic updates)
        if (this._suppressAutoSave || this._isLoadingData || this._isInitializing) {
            return;
        }
        
        const event = new CustomEvent('cellchange', {
            bubbles: true,
            composed: true
        });
        this.dispatchEvent(event);
    }

    /**
     * @api method called by parent to save data
     * Returns the data structure to be saved
     */
    @api
    async saveSheet() {
        // Enforce minimum of 1 for numberOfBuildings before saving
        if (!this.jobOverview.numberOfBuildings || this.jobOverview.numberOfBuildings < 1) {
            this.jobOverview = { ...this.jobOverview, numberOfBuildings: 1 };
        }
        
        this.refreshBuildingCalculations();

        return {
            jobOverview: this.jobOverview,
            summaryRows: this.summaryRows,
            buildings: this.buildings
        };
    }

    columnNameFromIndex(index) {
        let name = '';
        let i = index;
        while (i > 0) {
            const rem = (i - 1) % 26;
            name = String.fromCharCode(65 + rem) + name; // 65 = 'A'
            i = Math.floor((i - 1) / 26);
        }
        return name;
    }

    showToast(title, message, variant = 'info') {
        this.dispatchEvent(new ShowToastEvent({ title, message, variant }));
    }

    handleNumberKeyDown(event) {
        if (event.key === '-' || event.key === 'Minus') {
            event.preventDefault();
        }
    }

    sanitizeNonNegative(event) {
        let value = event.target.value;

        if (value === '') return;

        value = parseFloat(value);
        if (isNaN(value) || value < 0) {
            event.target.value = 0;
        }
    }

}