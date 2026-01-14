import { LightningElement, api, track } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import saveSheet from '@salesforce/apex/BidWorksheetUndergroundController.saveSheet';
import saveEstimateSheet from '@salesforce/apex/BidWorksheetUndergroundController.saveEstimateSheet';
import saveDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.saveDesignWorksheet';
import saveSOVSheet from '@salesforce/apex/BidWorksheetUndergroundController.saveSOVSheet';
import updateOpportunityFields from '@salesforce/apex/BidWorksheetUndergroundController.updateOpportunityFields';
import getNextVersionNumber from '@salesforce/apex/BidWorksheetUndergroundController.getNextVersionNumber';
import getVersionHistory from '@salesforce/apex/BidWorksheetUndergroundController.getVersionHistory';
import loadVersionById from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById';
import autoSaveSheet from '@salesforce/apex/BidWorksheetUndergroundController.autoSaveSheet';
// Import new wrapper methods for other worksheets
import getNextVersionNumber_Estimate from '@salesforce/apex/BidWorksheetUndergroundController.getNextVersionNumber_Estimate';
import getVersionHistory_Estimate from '@salesforce/apex/BidWorksheetUndergroundController.getVersionHistory_Estimate';
import loadVersionById_Estimate from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById_Estimate';
import loadLatestEstimateSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestEstimateSheet';
import autoSaveEstimateSheet from '@salesforce/apex/BidWorksheetUndergroundController.autoSaveEstimateSheet';
import getNextVersionNumber_SOV from '@salesforce/apex/BidWorksheetUndergroundController.getNextVersionNumber_SOV';
import getVersionHistory_SOV from '@salesforce/apex/BidWorksheetUndergroundController.getVersionHistory_SOV';
import loadVersionById_SOV from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById_SOV';
import loadLatestSOVSheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestSOVSheet';
import autoSaveSOVSheet from '@salesforce/apex/BidWorksheetUndergroundController.autoSaveSOVSheet';
import getNextVersionNumber_Design from '@salesforce/apex/BidWorksheetUndergroundController.getNextVersionNumber_Design';
import getVersionHistory_Design from '@salesforce/apex/BidWorksheetUndergroundController.getVersionHistory_Design';
import loadVersionById_Design from '@salesforce/apex/BidWorksheetUndergroundController.loadVersionById_Design';
import loadLatestDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.loadLatestDesignWorksheet';
import autoSaveDesignWorksheet from '@salesforce/apex/BidWorksheetUndergroundController.autoSaveDesignWorksheet';

export default class BidWorksheetUndergroundParent extends LightningElement {
    @api recordId;
    @track activeSections = ['sheet1']; // Sheet #1 open by default
    @track activeTab = 'underground'; // Track active tab
    @track sheet1Subtotal = 0;
    @track sheet2Subtotal = 0;
    @track estimateGrandTotal = 0;
    @track sovTotal = 0;
    @track isSaving = false;
    
    // Configuration object for all worksheets
    @track worksheetConfig = {
        'underground': {
            sheetTitle: 'UndergroundBid',
            autoSaveTitle: 'UndergroundBid_AutoSave',
            versionList: [],
            selectedVersionId: '',
            nextVersionNumber: 1,
            autoSaveTimeout: null,
            autoSaveStatus: '',
            getVersionHistoryMethod: getVersionHistory,
            getNextVersionMethod: getNextVersionNumber,
            autoSaveMethod: autoSaveSheet,
            loadLatestMethod: null, // Uses loadLatestSheet (already exists)
            loadVersionMethod: loadVersionById,
            childComponents: ['c-bid-worksheet-underground', 'c-bid-worksheet-underground-two']
        },
        'estimates': {
            sheetTitle: 'BidWorksheet_Estimate',
            autoSaveTitle: 'BidWorksheet_Estimate_AutoSave',
            versionList: [],
            selectedVersionId: '',
            nextVersionNumber: 1,
            autoSaveTimeout: null,
            autoSaveStatus: '',
            getVersionHistoryMethod: getVersionHistory_Estimate,
            getNextVersionMethod: getNextVersionNumber_Estimate,
            autoSaveMethod: autoSaveEstimateSheet,
            loadLatestMethod: loadLatestEstimateSheet,
            loadVersionMethod: loadVersionById_Estimate,
            childComponents: ['c-bid-worksheet-estimate']
        },
        'schedule': {
            sheetTitle: 'BidWorksheet_SOV',
            autoSaveTitle: 'BidWorksheet_SOV_AutoSave',
            versionList: [],
            selectedVersionId: '',
            nextVersionNumber: 1,
            autoSaveTimeout: null,
            autoSaveStatus: '',
            getVersionHistoryMethod: getVersionHistory_SOV,
            getNextVersionMethod: getNextVersionNumber_SOV,
            autoSaveMethod: autoSaveSOVSheet,
            loadLatestMethod: loadLatestSOVSheet,
            loadVersionMethod: loadVersionById_SOV,
            childComponents: ['c-schedule-of-values']
        },
        'design': {
            sheetTitle: 'DesignWorksheet',
            autoSaveTitle: 'DesignWorksheet_AutoSave',
            versionList: [],
            selectedVersionId: '',
            nextVersionNumber: 1,
            autoSaveTimeout: null,
            autoSaveStatus: '',
            getVersionHistoryMethod: getVersionHistory_Design,
            getNextVersionMethod: getNextVersionNumber_Design,
            autoSaveMethod: autoSaveDesignWorksheet,
            loadLatestMethod: loadLatestDesignWorksheet,
            loadVersionMethod: loadVersionById_Design,
            childComponents: ['c-design-worksheet']
        }
    };
    
    // Legacy version control state (kept for backward compatibility during refactor)
    @track versionList = [];
    @track selectedVersionId = '';
    @track nextVersionNumber = 1;
    @track isLoadingVersion = false; // Loading state for version data
    
    // Auto-save state (legacy - will be replaced by config)
    autoSaveTimeout = null;
    @track autoSaveStatus = ''; // 'saving', 'saved', ''
    _isInitializing = true; // Flag to prevent autosave during component initialization

    connectedCallback() {
        // Set initialization flag - will be cleared after components are ready
        this._isInitializing = true;
        
        // Immediately set selectedVersionId to 'draft' for the active tab so child components get the right value
        const config = this.getCurrentConfig();
        if (config && !config.selectedVersionId) {
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [this.activeTab]: {
                    ...config,
                    selectedVersionId: 'draft'
                }
            };
        }
        
        if (this.recordId && this.activeTab === 'underground') {
            this.loadVersionData();
        }
        
        // Clear initialization flag after a delay to allow all components to initialize
        setTimeout(() => {
            this._isInitializing = false;
        }, 3000); // Give enough time for both sheets to initialize and load data
    }

    handleTabChange(event) {
        const selectedTab = event.target.value;
        this.activeTab = selectedTab;
        
        // Load version data for the new tab
        if (this.recordId) {
            // Set initialization flag when switching tabs
            this._isInitializing = true;
            this.loadVersionDataForTab(selectedTab);
            // Clear initialization flag after components load
            setTimeout(() => {
                this._isInitializing = false;
            }, 3000);
        }
    }

    // ========================================
    // GENERIC HELPER METHODS (Config-Driven)
    // ========================================

    /**
     * Get configuration for the current active tab
     */
    getCurrentConfig() {
        return this.worksheetConfig[this.activeTab] || null;
    }

    /**
     * Get configuration for a specific tab
     */
    getConfigForTab(tabName) {
        return this.worksheetConfig[tabName] || null;
    }

    /**
     * Load version data for a specific tab
     */
    async loadVersionDataForTab(tabName) {
        const config = this.getConfigForTab(tabName);
        if (!config || !this.recordId) {
            return;
        }

        try {

            if (tabName === 'schedule') {
                
                // Get current job name from design worksheet component if it exists
                const designComponent = this.template.querySelector('c-design-worksheet');
                const currentJobName = designComponent ? designComponent.currentJobName : null;

                setTimeout(() => {
                    const sovComponent = this.template.querySelector('c-schedule-of-values');
                    if (sovComponent) {
                        try {
                            sovComponent.populateJobNameFromDesign(currentJobName);
                        } catch (err) {
                            console.error('Error calling refreshJobNameFromDesign on SOV component', err);
                        }
                    }
                }, 0);
            }

            // Load next version number first
            await this.loadNextVersionNumberForTab(tabName);
            // Then load version list
            await this.loadVersionListForTab(tabName);
        } catch (error) {
            console.error(`Error loading version data for ${tabName}:`, error);
        }
    }

    /**
     * Load next version number for a specific tab
     */
    async loadNextVersionNumberForTab(tabName) {
        const config = this.getConfigForTab(tabName);
        if (!config || !this.recordId) return;
        
        try {
            const nextVersion = await config.getNextVersionMethod({ opportunityId: this.recordId });
            // Update reactively
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [tabName]: {
                    ...config,
                    nextVersionNumber: nextVersion
                }
            };
        } catch (error) {
            console.error(`Error loading next version number for ${tabName}:`, error);
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [tabName]: {
                    ...config,
                    nextVersionNumber: 1
                }
            };
        }
    }

    /**
     * Load version list for a specific tab
     */
    async loadVersionListForTab(tabName) {
        const config = this.getConfigForTab(tabName);
        if (!config || !this.recordId) return;
        
        try {
            // Ensure nextVersionNumber is loaded
            if (!config.nextVersionNumber) {
                await this.loadNextVersionNumberForTab(tabName);
                // Get updated config after loading version number
                config = this.getConfigForTab(tabName);
            }
            
            const versions = await config.getVersionHistoryMethod({ opportunityId: this.recordId });
            
            // Map saved versions
            const savedVersions = versions.map(v => ({
                label: `Version ${v.versionNumber} - ${this.formatDate(v.createdDate)} - ${v.createdBy}`,
                value: v.versionId,
                versionNumber: v.versionNumber,
                isDraft: false
            }));
            
            // Add draft option at the top
            const draftOption = {
                label: `Draft - Version ${config.nextVersionNumber}`,
                value: 'draft',
                versionNumber: config.nextVersionNumber,
                isDraft: true
            };
            
            // Draft always appears first, then saved versions
            // Update config reactively
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [tabName]: {
                    ...config,
                    versionList: [draftOption, ...savedVersions]
                }
            };
            
            // Set draft as default selection if no version is currently selected
            const currentConfig = this.getConfigForTab(tabName);
            if (!currentConfig.selectedVersionId || currentConfig.selectedVersionId === 'draft') {
                // Update selectedVersionId reactively
                this.worksheetConfig = {
                    ...this.worksheetConfig,
                    [tabName]: {
                        ...currentConfig,
                        selectedVersionId: 'draft'
                    }
                };
                // Set versionIdToLoad on children so they load latest (draft) data
                this.updateChildrenVersionIdForTab(tabName);
            }
        } catch (error) {
            console.error(`Error loading version list for ${tabName}:`, error);
            // Even on error, show draft option - update reactively
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [tabName]: {
                    ...config,
                    versionList: [{
                        label: `Draft - Version ${config.nextVersionNumber || 1}`,
                        value: 'draft',
                        versionNumber: config.nextVersionNumber || 1,
                        isDraft: true
                    }],
                    selectedVersionId: 'draft'
                }
            };
        }
    }

    /**
     * Update versionIdToLoad on child components for a specific tab
     */
    updateChildrenVersionIdForTab(tabName) {
        const config = this.getConfigForTab(tabName);
        if (!config) return;

        setTimeout(() => {
            
            config.childComponents.forEach(selector => {
                const component = this.template.querySelector(selector);
                if (component) {
                    component.versionIdToLoad = config.selectedVersionId;
                } else {
                }
            });
        }, 200);
    }

    /**
     * Legacy loadVersionData (kept for backward compatibility, now uses config)
     */
    async loadVersionData() {
        await this.loadVersionDataForTab(this.activeTab);
    }

    /**
     * Legacy loadVersionList (kept for backward compatibility, now uses config)
     */
    async loadVersionList() {
        await this.loadVersionListForTab(this.activeTab);
    }

    // Removed loadSelectedVersion() - version loading is now handled automatically by child components
    // via versionIdToLoad property and renderedCallback()

    /**
     * Legacy loadNextVersionNumber (kept for backward compatibility, now uses config)
     */
    async loadNextVersionNumber() {
        await this.loadNextVersionNumberForTab(this.activeTab);
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
     * Legacy updateChildrenVersionId (kept for backward compatibility, now uses config)
     */
    updateChildrenVersionId() {
        this.updateChildrenVersionIdForTab(this.activeTab);
    }

    /**
     * Generic version change handler for any tab
     */
    async handleVersionChangeGeneric(event, tabName) {
        const config = this.getConfigForTab(tabName);
        if (!config) return;

        const newVersionId = event.detail.value;
        
        if (!newVersionId) return;
        
        // If switching away from draft and there are unsaved changes, save first
        if (config.selectedVersionId === 'draft' && newVersionId !== 'draft') {
            // Wait for any in-progress autosave to complete
            if (config.autoSaveStatus === 'saving') {
                // Wait up to 5 seconds for autosave to complete
                let waitCount = 0;
                while (config.autoSaveStatus === 'saving' && waitCount < 25) {
                    await new Promise(resolve => setTimeout(resolve, 200));
                    waitCount++;
                }
            }
            
            // Check if there's a pending autosave (user made changes but debounce hasn't fired yet)
            if (config.autoSaveTimeout) {
                
                // Clear the pending timeout
                clearTimeout(config.autoSaveTimeout);
                config.autoSaveTimeout = null;
                
                // Save immediately
                await this.performAutoSaveGeneric(tabName);
                
            }
        }
        
        // Update selectedVersionId reactively
        this.worksheetConfig = {
            ...this.worksheetConfig,
            [tabName]: {
                ...config,
                selectedVersionId: newVersionId
            }
        };
        
        // Pass version ID to child components - they'll auto-reload
        this.updateChildrenVersionIdForTab(tabName);
        
        // Show loading state for non-draft versions
        if (newVersionId !== 'draft') {
            this.isLoadingVersion = true;
            // Hide loading after a delay (children handle their own loading)
            setTimeout(() => {
                this.isLoadingVersion = false;
            }, 2000);
        }
    }

    /**
     * Legacy version change handler (kept for backward compatibility)
     */
    async handleVersionChange(event) {
        await this.handleVersionChangeGeneric(event, this.activeTab);
    }

    decodeData(base64Data) {
        try {
            return decodeURIComponent(escape(atob(base64Data)));
        } catch (err) {
            console.error('Decode data failed:', err);
            throw err;
        }
    }

    get showSaveUndergroundButton() {
        return this.activeTab === 'underground';
    }

    get showSaveEstimatesButton() {
        return this.activeTab === 'estimates';
    }

    get showSaveDesignButton() {
        return this.activeTab === 'design';
    }

    get showSaveSOVButton() {
        return this.activeTab === 'schedule';
    }

    get isAutoSaving() {
        return this.currentAutoSaveStatus === 'saving';
    }

    get isAutoSaved() {
        return this.currentAutoSaveStatus === 'saved';
    }

    // ========================================
    // DYNAMIC GETTERS (Config-Driven)
    // ========================================

    /**
     * Get current version list from config
     */
    get currentVersionList() {
        const config = this.getCurrentConfig();
        if (!config) return [];
        return config.versionList.length > 0 ? config.versionList : [{
            label: `Draft - Version ${config.nextVersionNumber || 1}`,
            value: 'draft',
            versionNumber: config.nextVersionNumber || 1,
            isDraft: true
        }];
    }

    /**
     * Get current selected version ID from config
     */
    get currentSelectedVersionId() {
        const config = this.getCurrentConfig();
        if (!config) return 'draft';
        return config.selectedVersionId || 'draft';
    }

    /**
     * Get current autosave status from config
     */
    get currentAutoSaveStatus() {
        const config = this.getCurrentConfig();
        if (!config) return '';
        return config.autoSaveStatus || '';
    }

    /**
     * Check if version control should be shown for current tab
     */
    get showVersionControl() {
        // Show version control for all tabs that have it configured
        return ['underground', 'estimates', 'schedule', 'design'].includes(this.activeTab);
    }

    /**
     * Check if version list is disabled
     */
    get versionListDisabled() {
        const config = this.getCurrentConfig();
        if (!config) return true;
        // Version list should never be empty since draft is always present
        return config.versionList.length === 0;
    }

    /**
     * Legacy getters (kept for backward compatibility)
     */
    get legacyVersionList() {
        return this.currentVersionList;
    }

    get legacySelectedVersionId() {
        return this.currentSelectedVersionId;
    }

    get legacyAutoSaveStatus() {
        return this.currentAutoSaveStatus;
    }

    get undergroundGrandTotal() {
        const s2 = parseFloat(this.sheet2Subtotal) || 0;
        return (s2).toFixed(2);
    }

    handleSheet1Update(event) {
        const { subtotal } = event.detail;
        this.sheet1Subtotal = subtotal;

    }

    handleSheet2Update(event) {
        const { totalPrice } = event.detail;
        this.sheet2Subtotal = totalPrice;

    }

    handleEstimateUpdate(event) {
        const { grandTotal } = event.detail;
        this.estimateGrandTotal = grandTotal;

    }

    handleSOVUpdate(event) {
        const { totalSOV } = event.detail;
        this.sovTotal = totalSOV;

    }

    async handleSaveUnderground() {
        this.isSaving = true;

        try {

            if (!this.activeSections.includes('sheet1')) {
                this.activeSections = [...this.activeSections, 'sheet1'];
            }
            if (!this.activeSections.includes('sheet2')) {
                this.activeSections = [...this.activeSections, 'sheet2'];
            }

            await new Promise(resolve => setTimeout(resolve, 100));

            let sheet1Component = this.template.querySelector('c-bid-worksheet-underground');
            let sheet2Component = this.template.querySelector('c-bid-worksheet-underground-two');

            if (!sheet1Component || !sheet2Component) {
                throw new Error('Sheet components not found. Please ensure both sections are expanded.');
            }

            const sheet1Data = await sheet1Component.saveSheet();
            const sheet2Data = await sheet2Component.saveSheet();

            const payload = {
                worksheetType: 'Underground',
                version: '1.0',
                savedDate: new Date().toISOString(),
                opportunityId: this.recordId,
                sheet1: sheet1Data,
                sheet2: sheet2Data,
                summary: {
                    sheet1Subtotal: this.sheet1Subtotal,
                    sheet2Total: this.sheet2Subtotal,
                    grandTotal: this.undergroundGrandTotal
                }
            };

            const base64Payload = this.encodeData(payload);
            if (!this.recordId) {
                throw new Error('Opportunity ID is required');
            }
            const targetId = this.recordId;

            await saveSheet({
                opportunityId: targetId,
                base64Data: base64Payload
            });


            // Refresh version data after save
            await this.loadVersionDataForTab('underground');
            
            // Set draft as selected after save (user continues working on new draft)
            const config = this.getConfigForTab('underground');
            // Get the version number that was just saved (nextVersionNumber - 1, since nextVersionNumber is now the next draft)
            const savedVersionNumber = config ? (config.nextVersionNumber - 1) : 0;
            if (config) {
                config.selectedVersionId = 'draft';
            }
            // Don't auto-load - user is already working on the draft

            const fieldData = await this.extractOpportunityFieldData('underground');

            if (Object.keys(fieldData).length > 0) {
                await updateOpportunityFields({
                    opportunityId: targetId,
                    fieldDataJson: JSON.stringify(fieldData)
                });

                this.showToast('Success', 'Underground Worksheet saved successfully!', 'success');
            } else {
                this.showToast('Success', 'Underground Worksheet saved successfully!', 'success');
            }

        } catch (error) {
            this.logError('Error saving Underground estimate', error);
            const message = error?.body?.message || error?.message || 'Unknown error';
            this.showToast('Error', 'Failed to save: ' + message, 'error');
        } finally {
            this.isSaving = false;
        }
    }

    async handleSaveDesign() {
        this.isSaving = true;

        try {

            await new Promise(resolve => setTimeout(resolve, 100));

            let designComponent = this.template.querySelector('c-design-worksheet');

            if (!designComponent) {
                throw new Error('Design Worksheet component not found');
            }


            // ⭐ UPDATED: Call saveSheet instead of saveWorksheet
            if (typeof designComponent.saveSheet !== 'function') {
                throw new Error('Design saveSheet method not available');
            }

            const designData = await designComponent.saveSheet();

            // Validate designData
            if (!designData) {
                throw new Error('Design data is null or undefined');
            }

            if (!designData.formData) {
                throw new Error('Design formData is missing. Please ensure the form is properly initialized.');
            }

            const payload = {
                worksheetType: 'Design',
                version: '1.0',
                savedDate: new Date().toISOString(),
                opportunityId: this.recordId,
                formData: designData.formData
            };

            const base64Payload = this.encodeData(payload);
            if (!this.recordId) {
                throw new Error('Opportunity ID is required');
            }
            const targetId = this.recordId;

            await saveDesignWorksheet({
                opportunityId: targetId,
                base64Data: base64Payload
            });


            // Refresh version data after save
            await this.loadVersionDataForTab('design');
            
            // Set draft as selected after save
            const designConfig = this.getConfigForTab('design');
            if (designConfig) {
                designConfig.selectedVersionId = 'draft';
            }

            // ⭐ Extract and update Opportunity fields from Design
            let fieldData = {};
            try {
                fieldData = await this.extractOpportunityFieldData('design', designData, null);
            } catch (extractError) {
                console.error('❌ Error extracting field data:', extractError);
                // Continue with empty fieldData - don't fail the entire save
                fieldData = {};
            }

            // Validate fieldData is an object
            if (!fieldData || typeof fieldData !== 'object' || Array.isArray(fieldData)) {
                this.showToast('Success', 'Design Worksheet saved successfully!', 'success');
            } else {
                const fieldKeys = Object.keys(fieldData);
                if (fieldKeys.length > 0) {
                    try {
                        const fieldDataJson = JSON.stringify(fieldData);
                        if (!fieldDataJson || fieldDataJson === '{}') {
                            throw new Error('Field data JSON is empty or invalid');
                        }
                        await updateOpportunityFields({
                            opportunityId: targetId,
                            fieldDataJson: fieldDataJson
                        });

                        this.showToast('Success', 'Design Worksheet saved successfully!', 'success');
                    } catch (updateError) {
                        console.error('❌ Error updating Opportunity fields:', updateError);
                        // Still show success for saving the worksheet file
                        this.showToast('Warning',
                            'Design Worksheet saved, but failed to update Opportunity fields: ' + (updateError?.body?.message || updateError?.message || 'Unknown error'),
                            'warning');
                    }
                } else {
                    this.showToast('Success', 'Design Worksheet saved successfully!', 'success');
                }
            }

        } catch (error) {
            this.logError('Error saving Design Worksheet', error);
            const message = error?.body?.message || error?.message || 'Unknown error';
            this.showToast('Error', 'Failed to save: ' + message, 'error');
        } finally {
            this.isSaving = false;
        }
    }

    async handleSaveEstimate() {
        this.isSaving = true;

        try {

            await new Promise(resolve => setTimeout(resolve, 100));

            let estimateComponent = this.template.querySelector('c-bid-worksheet-estimate');

            if (!estimateComponent) {
                const estimateComponents = this.template.querySelectorAll('c-bid-worksheet-estimate');
                if (estimateComponents && estimateComponents.length > 0) {
                    estimateComponent = estimateComponents[0];
                }
            }

            if (!estimateComponent) {
                throw new Error('Estimate component not found. Please ensure you are on the Estimates tab.');
            }


            if (typeof estimateComponent.saveSheet !== 'function') {
                throw new Error('Estimate saveSheet method not available');
            }

            const estimateData = await estimateComponent.saveSheet();

            const payload = {
                worksheetType: 'Estimate',
                version: '1.0',
                savedDate: new Date().toISOString(),
                opportunityId: this.recordId,
                section1: estimateData.section1,
                section2: estimateData.section2,
                section3: estimateData.section3,
                summary: {
                    section1Subtotal: estimateData.section1?.reduce((sum, item) => sum + (parseFloat(item.gross) || 0), 0).toFixed(2) || '0.00',
                    section2Subtotal: estimateData.section2?.reduce((sum, item) => sum + (parseFloat(item.gross) || 0), 0).toFixed(2) || '0.00',
                    section3Subtotal: estimateData.section3?.reduce((sum, item) => sum + (parseFloat(item.gross) || 0), 0).toFixed(2) || '0.00',
                    grandTotal: this.estimateGrandTotal
                }
            };

            const base64Payload = this.encodeData(payload);
            if (!this.recordId) {
                throw new Error('Opportunity ID is required');
            }
            const targetId = this.recordId;

            await saveEstimateSheet({
                opportunityId: targetId,
                base64Data: base64Payload
            });


            // Refresh version data after save
            await this.loadVersionDataForTab('estimates');
            
            // Set draft as selected after save
            const estimateConfig = this.getConfigForTab('estimates');
            if (estimateConfig) {
                estimateConfig.selectedVersionId = 'draft';
            }

            // ⭐ Extract and update Opportunity fields from Estimate
            const fieldData = await this.extractOpportunityFieldData('estimate', estimateData);


            if (Object.keys(fieldData).length > 0) {
                await updateOpportunityFields({
                    opportunityId: targetId,
                    fieldDataJson: JSON.stringify(fieldData)
                });

                this.showToast('Success', 'Estimate Worksheet saved successfully!', 'success');
            } else {
                this.showToast('Success', 'Estimate Worksheet saved successfully!', 'success');
            }

        } catch (error) {
            this.logError('Error saving Estimate worksheet', error);
            const message = error?.body?.message || error?.message || 'Unknown error';
            this.showToast('Error', 'Failed to save Estimate worksheet: ' + message, 'error');
        } finally {
            this.isSaving = false;
        }
    }

    async handleSaveSOV() {
        this.isSaving = true;

        try {

            await new Promise(resolve => setTimeout(resolve, 100));

            let sovComponent = this.template.querySelector('c-schedule-of-values');

            if (!sovComponent) {
                throw new Error('Schedule of Values component not found. Please ensure you are on the SOV tab.');
            }


            if (typeof sovComponent.saveSheet !== 'function') {
                throw new Error('SOV saveSheet method not available');
            }

            const sovData = await sovComponent.saveSheet();

            // Validate sovData
            if (!sovData) {
                throw new Error('SOV data is null or undefined');
            }

            const payload = {
                worksheetType: 'ScheduleOfValues',
                version: '1.0',
                savedDate: new Date().toISOString(),
                opportunityId: this.recordId,
                jobOverview: sovData.jobOverview || {},
                summaryRows: sovData.summaryRows || [],
                buildings: sovData.buildings || [],
                summary: {
                    totalSOV: this.sovTotal
                }
            };

            const base64Payload = this.encodeData(payload);
            if (!this.recordId) {
                throw new Error('Opportunity ID is required');
            }
            const targetId = this.recordId;

            await saveSOVSheet({
                opportunityId: targetId,
                base64Data: base64Payload
            });


            // Refresh version data after save
            await this.loadVersionDataForTab('schedule');
            
            // Set draft as selected after save
            const sovConfig = this.getConfigForTab('schedule');
            if (sovConfig) {
                sovConfig.selectedVersionId = 'draft';
            }

            // ⭐ Extract and update Opportunity fields from SOV
            let fieldData = {};
            try {
                fieldData = await this.extractOpportunityFieldData('sov', null, {
                    summaryRows: sovData.summaryRows || [],
                    summary: payload.summary
                });
            } catch (extractError) {
                console.error('❌ Error extracting field data:', extractError);
                // Continue with empty fieldData - don't fail the entire save
                fieldData = {};
            }

            // Validate fieldData is an object
            if (!fieldData || typeof fieldData !== 'object' || Array.isArray(fieldData)) {
                this.showToast('Success', 'Schedule of Values saved successfully!', 'success');
            } else {
                const fieldKeys = Object.keys(fieldData);
                if (fieldKeys.length > 0) {
                    try {
                        const fieldDataJson = JSON.stringify(fieldData);
                        if (!fieldDataJson || fieldDataJson === '{}') {
                            throw new Error('Field data JSON is empty or invalid');
                        }
                        await updateOpportunityFields({
                            opportunityId: targetId,
                            fieldDataJson: fieldDataJson
                        });

                        this.showToast('Success', 'Schedule of Values saved successfully!', 'success');
                    } catch (updateError) {
                        console.error('❌ Error updating Opportunity fields:', updateError);
                        // Still show success for saving the worksheet file
                        this.showToast('Warning',
                            'Schedule of Values saved, but failed to update Opportunity fields: ' + (updateError?.body?.message || updateError?.message || 'Unknown error'),
                            'warning');
                    }
                } else {
                    this.showToast('Success', 'Schedule of Values saved successfully!', 'success');
                }
            }

        } catch (error) {
            this.logError('Error saving Schedule of Values', error);
            const message = error?.body?.message || error?.message || 'Unknown error';
            this.showToast('Error', 'Failed to save Schedule of Values: ' + message, 'error');
        } finally {
            this.isSaving = false;
        }
    }

    showToast(title, message, variant) {
        const event = new ShowToastEvent({
            title: title,
            message: message,
            variant: variant
        });
        this.dispatchEvent(event);
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

    logError(context, error) {
        const safeMsg = error?.body?.message || error?.message || String(error);
        console.error(`❌ ${context}:`, safeMsg, error);
    }

    /**
     * Generic cell change handler (for native HTML inputs using oninput)
     */
    handleCellChangeGeneric(event) {
        
        // Don't trigger autosave during initialization or version loading
        if (this._isInitializing || this.isLoadingVersion) {
            return;
        }
        
        const config = this.getCurrentConfig();
        if (!config) {
            return;
        }
        
        
        // Debounce auto-save - clear existing timeout
        if (config.autoSaveTimeout) {
            clearTimeout(config.autoSaveTimeout);
        }

        // Set new timeout for 2 seconds
        const timeoutId = setTimeout(() => {
            this.performAutoSaveGeneric(this.activeTab);
        }, 2000);
        
        // Update config with new timeout (for reactivity)
        this.worksheetConfig = {
            ...this.worksheetConfig,
            [this.activeTab]: {
                ...config,
                autoSaveTimeout: timeoutId
            }
        };
        
    }

    /**
     * Cell change handler for lightning-input components (Design worksheet uses onchange)
     */
    handleCellChangeLightning(event) {
        // Same logic as generic handler
        this.handleCellChangeGeneric(event);
    }

    /**
     * Legacy cell change handler (kept for backward compatibility)
     */
    handleCellChange(event) {
        this.handleCellChangeGeneric(event);
    }

    /**
     * Generic autosave performer for any tab
     */
    async performAutoSaveGeneric(tabName) {
        
        const config = this.getConfigForTab(tabName);
        if (!config || !this.recordId) {
            return;
        }

        try {
            // Force reactivity by reassigning the config object
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [tabName]: {
                    ...config,
                    autoSaveStatus: 'saving'
                }
            };
            
            // Get child components for this tab
            const components = config.childComponents.map(selector => 
                this.template.querySelector(selector)
            ).filter(c => c !== null);

            if (components.length === 0) {
                this.worksheetConfig = {
                    ...this.worksheetConfig,
                    [tabName]: {
                        ...this.worksheetConfig[tabName],
                        autoSaveStatus: ''
                    }
                };
                return;
            }

            await new Promise(resolve => setTimeout(resolve, 100));

            // Collect data from all child components
            let payload = {};
            
            if (tabName === 'underground') {
                // Underground has two sheets
                const sheet1Component = components[0];
                const sheet2Component = components[1];
                if (!sheet1Component || !sheet2Component) {
                    this.worksheetConfig = {
                        ...this.worksheetConfig,
                        [tabName]: {
                            ...this.worksheetConfig[tabName],
                            autoSaveStatus: ''
                        }
                    };
                    return;
                }
                const sheet1Data = await sheet1Component.saveSheet();
                const sheet2Data = await sheet2Component.saveSheet();
                payload = {
                    worksheetType: 'Underground',
                    version: '1.0',
                    savedDate: new Date().toISOString(),
                    opportunityId: this.recordId,
                    sheet1: sheet1Data,
                    sheet2: sheet2Data,
                    summary: {
                        sheet1Subtotal: this.sheet1Subtotal,
                        sheet2Total: this.sheet2Subtotal,
                        grandTotal: this.undergroundGrandTotal
                    }
                };
            } else if (tabName === 'estimates') {
                const estimateComponent = components[0];
                if (!estimateComponent) {
                    this.worksheetConfig = {
                        ...this.worksheetConfig,
                        [tabName]: {
                            ...this.worksheetConfig[tabName],
                            autoSaveStatus: ''
                        }
                    };
                    return;
                }
                const estimateData = await estimateComponent.saveSheet();
                payload = {
                    worksheetType: 'Estimate',
                    version: '1.0',
                    savedDate: new Date().toISOString(),
                    opportunityId: this.recordId,
                    section1: estimateData.section1,
                    section2: estimateData.section2,
                    section3: estimateData.section3,
                    summary: {
                        section1Subtotal: estimateData.section1?.reduce((sum, item) => sum + (parseFloat(item.gross) || 0), 0).toFixed(2) || '0.00',
                        section2Subtotal: estimateData.section2?.reduce((sum, item) => sum + (parseFloat(item.gross) || 0), 0).toFixed(2) || '0.00',
                        section3Subtotal: estimateData.section3?.reduce((sum, item) => sum + (parseFloat(item.gross) || 0), 0).toFixed(2) || '0.00',
                        grandTotal: this.estimateGrandTotal
                    }
                };
            } else if (tabName === 'schedule') {
                const sovComponent = components[0];
                if (!sovComponent) {
                    this.worksheetConfig = {
                        ...this.worksheetConfig,
                        [tabName]: {
                            ...this.worksheetConfig[tabName],
                            autoSaveStatus: ''
                        }
                    };
                    return;
                }
                const sovData = await sovComponent.saveSheet();
                payload = {
                    worksheetType: 'ScheduleOfValues',
                    version: '1.0',
                    savedDate: new Date().toISOString(),
                    opportunityId: this.recordId,
                    jobOverview: sovData.jobOverview || {},
                    summaryRows: sovData.summaryRows || [],
                    buildings: sovData.buildings || [],
                    summary: {
                        totalSOV: this.sovTotal
                    }
                };
            } else if (tabName === 'design') {
                const designComponent = components[0];
                if (!designComponent) {
                    this.worksheetConfig = {
                        ...this.worksheetConfig,
                        [tabName]: {
                            ...this.worksheetConfig[tabName],
                            autoSaveStatus: ''
                        }
                    };
                    return;
                }
                const designData = await designComponent.saveSheet();
                payload = {
                    worksheetType: 'Design',
                    version: '1.0',
                    savedDate: new Date().toISOString(),
                    opportunityId: this.recordId,
                    formData: designData.formData
                };
            }

            const base64Payload = this.encodeData(payload);

            await config.autoSaveMethod({
                opportunityId: this.recordId,
                base64Data: base64Payload
            });

            // Update status to 'saved' with reactivity
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [tabName]: {
                    ...this.worksheetConfig[tabName],
                    autoSaveStatus: 'saved'
                }
            };


            // Clear status after 2 seconds
            setTimeout(() => {
                this.worksheetConfig = {
                    ...this.worksheetConfig,
                    [tabName]: {
                        ...this.worksheetConfig[tabName],
                        autoSaveStatus: ''
                    }
                };
            }, 2000);

        } catch (error) {
            console.error(`Error during auto-save for ${tabName}:`, error);
            this.worksheetConfig = {
                ...this.worksheetConfig,
                [tabName]: {
                    ...this.worksheetConfig[tabName],
                    autoSaveStatus: ''
                }
            };
            // Don't show error toast for auto-save failures - just log
        }
    }

    /**
     * Legacy autosave method (kept for backward compatibility)
     */
    async performAutoSave() {
        await this.performAutoSaveGeneric(this.activeTab);
    }

    /**
     * Extract Opportunity field data from child components
     * @param {string} componentType - 'underground' or 'estimate'
     * @param {object} estimateData - Optional estimate data (section1, section2, section3)
     */
    async extractOpportunityFieldData(componentType, estimateData = null, sovData = null) {
        const fieldData = {};

        if (componentType === 'underground') {
            // ========================================
            // UNDERGROUND WORKSHEET - SHEET 2 ONLY
            // Only 2 fields needed
            // ========================================

            const sheet2Component = this.template.querySelector('c-bid-worksheet-underground-two');

            if (!sheet2Component) {
                throw new Error('Underground Sheet 2 component not found');
            }

            const sheet2Data = await sheet2Component.saveSheet();
            const tableRows = sheet2Data.lineItems || [];


            // Field #1: Pump Direct Cost Total (Row 83, Column R = right.gross)
            const pumpRow = tableRows.find(r => r.excelRow === 83);
            if (pumpRow && pumpRow.right && pumpRow.right.gross) {
                fieldData.Pump_Direct_Cost_Total__c = parseFloat(pumpRow.right.gross) || 0;
            } else {
            }

            // Field #2: FHV Standpipe Total Direct Cost (Row 84, Column R = right.gross)
            const fhvRow = tableRows.find(r => r.excelRow === 84);
            if (fhvRow && fhvRow.right && fhvRow.right.gross) {
                fieldData.FHV_Standpipe_Total_Direct_Cost__c = parseFloat(fhvRow.right.gross) || 0;
            } else {
            }


            return fieldData;

        } else if (componentType === 'estimate') {
            // ========================================
            // ESTIMATE WORKSHEET - SECTION 3
            // 37 fields total
            // ========================================

            if (!estimateData || !estimateData.section3) {
                return fieldData;
            }

            const section3Items = estimateData.section3;

            // Helper function to find item by excel row and column
            const findItem = (excelRow, column) => {
                return section3Items.find(item =>
                    item.excelRow === excelRow &&
                    item.column === column
                );
            };

            // Helper function to get decimal value
            const getDecimal = (item, field) => {
                if (!item) return null;
                const value = item[field];
                if (value === null || value === undefined || value === '') return null;
                return parseFloat(value) || 0;
            };

            // Row 155 - Grand Total Material Cost (Left, Gross) - Column I
            const i155 = findItem(155, 'Left');
            if (i155) {
                fieldData.GRAND_TOTAL_MATERIAL_COST__c = getDecimal(i155, 'gross');
            }

            // Row 157 - Head Count (Left, Size field contains headcount) - Column C
            const c157 = findItem(157, 'Left');
            if (c157 && c157.size) {
                fieldData.Head_Count__c = parseFloat(c157.size) || 0;
            }

            // Row 184 - Labor (FM+7th Period)
            const e184 = findItem(184, 'Left');
            if (e184) {
                fieldData.LABOR_FM_7TH_PERIOD_QUANTITY__c = getDecimal(e184, 'quantity');
                fieldData.LABOR_FM_7TH_PERIOD_UNIT__c = getDecimal(e184, 'unitPrice');
                fieldData.LABOR_FM_7TH_PERIOD_GROSS__c = getDecimal(e184, 'gross');
            }

            // Row 185 - Engineering Half Hour Head
            const e185 = findItem(185, 'Left');
            if (e185) {
                fieldData.ENGINEERING_HALF_HOUR_HEAD_QUANTITY__c = getDecimal(e185, 'quantity');
                fieldData.ENGINEERING_HALF_HOUR_HEAD_UNIT__c = getDecimal(e185, 'unitPrice');
                fieldData.ENGINEERING_HALF_HOUR_HEAD_GROSS__c = getDecimal(e185, 'gross');
            }

            // Row 186 - BIM
            const e186 = findItem(186, 'Left');
            if (e186) {
                fieldData.BIM_QUANTITY__c = getDecimal(e186, 'quantity');
                fieldData.BIM_UNIT__c = getDecimal(e186, 'unitPrice');
                fieldData.BIM_GROSS__c = getDecimal(e186, 'gross');
            }

            // Row 187 - Fabrication Quarter Hour Per
            const e187 = findItem(187, 'Left');
            if (e187) {
                fieldData.FABRICATION_QUARTER_HOUR_PER_QUANTITY__c = getDecimal(e187, 'quantity');
                fieldData.FABRICATION_QUARTER_HOUR_PER_UNIT__c = getDecimal(e187, 'unitPrice');
                fieldData.FABRICATION_QUARTER_HOUR_PER_GROSS__c = getDecimal(e187, 'gross');
            }

            // Row 160 - Total Direct Cost (Right, Gross) - Column R
            const r160 = findItem(160, 'Right');
            if (r160) {
                fieldData.TOTAL_DIRECT_COST_GROSS__c = getDecimal(r160, 'gross');
            }

            // Row 161 - % Overhead (Right)
            const n161 = findItem(161, 'Right');
            if (n161) {
                fieldData.OVERHEAD_QUANTITY__c = getDecimal(n161, 'quantity');
                fieldData.OVERHEAD_UNIT__c = getDecimal(n161, 'unitPrice');
                fieldData.OVERHEAD_GROSS__c = getDecimal(n161, 'gross');
            }

            // Row 163 - Subtotal (Right, Gross) - Column R
            const r163 = findItem(163, 'Right');
            if (r163) {
                fieldData.SUBTOTAL_GROSS__c = getDecimal(r163, 'gross');
            }

            // Row 164 - % Gain (Right)
            const n164 = findItem(164, 'Right');
            if (n164) {
                fieldData.GAIN_QUANTITY__c = getDecimal(n164, 'quantity');
                fieldData.GAIN_UNIT__c = getDecimal(n164, 'unitPrice');
                fieldData.GAIN_GROSS__c = getDecimal(n164, 'gross');
            }

            // Row 166 - Total Quote Price (Right, Gross) - Column R
            const r166 = findItem(166, 'Right');
            if (r166) {
                fieldData.TOTAL_QUOTE_PRICE__c = getDecimal(r166, 'gross');
            }

            const r167 = findItem(167, 'Right');
            if (r167) {
                fieldData.PRICE_MINUS_SP_PUMP_BF_MFLEX_DON_T_C__c = getDecimal(r167, 'gross');
            }

            // Row 168 - Gross Margin (Right, Quantity) - Column N
            const n168 = findItem(168, 'Right');
            if (n168) {
                fieldData.GROSS_MARGIN__c = getDecimal(n168, 'quantity');
            }

            // Row 169 - Bond Amount (Right, Quantity) - Column N
            const n169 = findItem(169, 'Right');
            if (n169) {
                fieldData.BOND_AMOUNT__c = getDecimal(n169, 'quantity');
            }

            // Row 173 - Material Per Head (Right)
            const n173 = findItem(173, 'Right');
            if (n173) {
                fieldData.MATERIAL_PER_HEAD_QUANTITY__c = getDecimal(n173, 'quantity');
                fieldData.MATERIAL_PER_HEAD_UNIT__c = getDecimal(n173, 'unitPrice');
                fieldData.MATERIAL_PER_HEAD_GROSS__c = getDecimal(n173, 'gross');
            }

            // Row 174 - Direct Cost Per Head (Right)
            const n174 = findItem(174, 'Right');
            if (n174) {
                fieldData.DIRECT_COST_PER_HEAD_QUANTITY__c = getDecimal(n174, 'quantity');
                fieldData.DIRECT_COST_PER_HEAD_UNIT__c = getDecimal(n174, 'unitPrice');
                fieldData.DIRECT_COST_PER_HEAD_GROSS__c = getDecimal(n174, 'gross');
            }

            // Row 175 - Building Sq. Footage (Right)
            const p175 = findItem(175, 'Right');
            if (p175) {
                fieldData.BUILDING_SQ_FOOTAGE_UNIT__c = getDecimal(p175, 'unitPrice');
                fieldData.BUILDING_SQ_FOOTAGE_GROSS__c = getDecimal(p175, 'gross');
            }

            // Row 176 - Sales Cost Per Head (Right)
            const n176 = findItem(176, 'Right');
            if (n176) {
                fieldData.SALES_COST_PER_HEAD_QUANTITY__c = getDecimal(n176, 'quantity');
                fieldData.SALES_COST_PER_HEAD_UNIT__c = getDecimal(n176, 'unitPrice');
                fieldData.SALES_COST_PER_HEAD_GROSS__c = getDecimal(n176, 'gross');
            }

            // Row 178 - Cost Per Square Foot (Right, Gross) - Column R
            const r178 = findItem(178, 'Right');
            if (r178) {
                fieldData.COST_PER_SQUARE_FOOT__c = getDecimal(r178, 'gross');
            }

            return fieldData;

        } else if (componentType === 'sov') {
            // ========================================
            // SCHEDULE OF VALUES - SUMMARY ROWS
            // 12 fields from summary table
            // ========================================

            if (!sovData) {
                return fieldData;
            }

            if (!sovData.summaryRows || !Array.isArray(sovData.summaryRows)) {
                return fieldData;
            }

            const summaryRows = sovData.summaryRows;

            // Helper to find row by label (case-insensitive, flexible matching)
            const findRowByLabel = (searchLabel) => {
                return summaryRows.find(row => {
                    const label = (row.label || '').toLowerCase().trim();
                    const search = searchLabel.toLowerCase().trim();
                    return label.includes(search) || search.includes(label);
                });
            };

            // Map each summary row to its Opportunity field
            // Row 1: DESIGN (LUMP)
            const designRow = findRowByLabel('design');
            if (designRow) {
                fieldData.Design_Total__c = parseFloat(designRow.value) || 0;
            }

            // Row 2: 3D BIM (LUMP)
            const bimRow = findRowByLabel('3d bim');
            if (bimRow) {
                fieldData.X3D_BIM_Total__c = parseFloat(bimRow.value) || 0;
            }

            // Row 3: MATERIAL / FAB
            const materialRow = findRowByLabel('material');
            if (materialRow) {
                fieldData.Material_FAB_Total__c = parseFloat(materialRow.value) || 0;
            }

            // Row 4: ROUGH IN
            const roughInRow = findRowByLabel('rough in');
            if (roughInRow) {
                fieldData.Rough_In_Total__c = parseFloat(roughInRow.value) || 0;
            }

            // Row 5: DROP CUT
            const dropCutRow = findRowByLabel('drop cut');
            if (dropCutRow) {
                fieldData.Drop_Cut_Total__c = parseFloat(dropCutRow.value) || 0;
            }

            // Row 6: TRIM
            const trimRow = findRowByLabel('trim');
            if (trimRow) {
                fieldData.Trim_Total__c = parseFloat(trimRow.value) || 0;
            }

            // Row 7: PUMP MATERIAL
            const pumpMaterialRow = findRowByLabel('pump material');
            if (pumpMaterialRow) {
                fieldData.Pump_Material_Total__c = parseFloat(pumpMaterialRow.value) || 0;
            }

            // Row 8: PUMP ROUGH IN
            const pumpRoughInRow = findRowByLabel('pump rough in');
            if (pumpRoughInRow) {
                fieldData.Pump_Rough_In_Total__c = parseFloat(pumpRoughInRow.value) || 0;
            }

            // Row 9: Permit (summary) or BACK FLOW DEVICE (building)
            // We'll use "permit" from summary rows, but also check for backflow
            const permitRow = findRowByLabel('permit');
            if (permitRow) {
                fieldData.Backflow_Device_Total__c = parseFloat(permitRow.value) || 0; // per request: use Permit row
            }

            // Row 10: Scissor (summary) used for Underground Work total
            const scissorRow = findRowByLabel('scissor');
            if (scissorRow) {
                fieldData.Underground_Work_Total__c = parseFloat(scissorRow.value) || 0; // per request: use Scissor row
            }

            // Row 11: STANDPIPE AND FHV
            const standpipeRow = findRowByLabel('standpipe');
            if (standpipeRow) {
                fieldData.Standpipe_and_VHF_Total__c = parseFloat(standpipeRow.value) || 0;
            }

            // Building Total (from totalSOV)
            if (sovData.summary && sovData.summary.totalSOV) {
                fieldData.Building_Total__c = parseFloat(sovData.summary.totalSOV) || 0;
            }

            return fieldData;


        } else if (componentType === 'design') {
            // ========================================
            // DESIGN WORKSHEET - FORM DATA
            // Extract ALL 56 fields with corrected data types
            // ========================================


            if (!estimateData) {
                return fieldData;
            }

            if (!estimateData.formData) {
                return fieldData;
            }

            const form = estimateData.formData;

            // ========================================
            // SAFE HELPER FUNCTIONS
            // ========================================

            // Add string field only if it has a non-empty value
            const addStringField = (fieldName, value, maxLength = 255) => {
                if (value !== undefined && value !== null) {
                    const strValue = String(value).trim();
                    if (strValue !== '') {
                        fieldData[fieldName] = strValue.substring(0, maxLength);
                        return true;
                    }
                }
                return false;
            };

            // Add number field only if it has a valid numeric value
            const addNumberField = (fieldName, value) => {
                if (value !== undefined && value !== null && value !== '') {
                    const numValue = parseFloat(value);
                    if (!isNaN(numValue)) {
                        fieldData[fieldName] = numValue;
                        return true;
                    }
                }
                return false;
            };

            // Add picklist field (Yes/No) only if boolean value is explicitly set
            const addPicklistField = (fieldName, booleanValue) => {
                if (booleanValue !== undefined && booleanValue !== null) {
                    fieldData[fieldName] = booleanValue ? 'Yes' : 'No';
                    return true;
                }
                return false;
            };

            // Add date field only if it has a value
            const addDateField = (fieldName, dateValue) => {
                if (dateValue !== undefined && dateValue !== null && dateValue !== '') {
                    fieldData[fieldName] = dateValue;
                    return true;
                }
                return false;
            };

            // ========================================
            // EXTRACT FIELDS WITH SAFE NULL HANDLING
            // ========================================

            // 1. Design Description (STRING 120)
            addStringField('Design_Description__c', form.description, 120);

            // 2. # Floors (NUMBER 18,0)
            addNumberField('Floors__c', form.numberOfFloors);

            // 3. Penthouse (STRING 120) - "Yes" or empty string
            addStringField('Penthouse__c', form.penthouse, 120);

            // 4. Bid Plan Date (DATE)
            addDateField('Bid_Plan_Date__c', form.bidPlanDate);

            // 5-14. Yes/No Picklist Fields
            addPicklistField('Residential_Rates__c', form.residentialRates);
            addPicklistField('Local_Hire__c', form.localHire);
            addPicklistField('Apprentice__c', form.apprenticePercent);
            addPicklistField('Textura__c', form.textura);
            addPicklistField('Certified_Payroll__c', form.certifiedPayroll);
            addPicklistField('OCIP_Deduct_or_Add_Later__c', form.ocipDeduct);

            // 11. OCIP Amount (CURRENCY 16,2)
            addNumberField('OCIP_Deduct_or_Add_Later_Amount__c', form.ocupAmount);

            // 12-14. More Yes/No Picklists
            addPicklistField('Market_Recovery__c', form.marketRecovery);
            addPicklistField('BIM_Required__c', form.bimRequired);
            addPicklistField('Permit_Fees__c', form.permitFeesIncluded);

            // 15. Permit Fees Amount (CURRENCY 16,2)
            addNumberField('Permit_Fees_Amount__c', form.permitAmount);

            // 16-19. Text Fields
            addStringField('AMMR__c', form.ammr, 120);
            addStringField('Pre_App__c', form.preApp, 120);
            addStringField('FPE_Required__c', form.fpeRequired, 120);
            addStringField('AHJ_Account__c', form.ahj, 1300);

            // 20. Hazard Classification (STRING 120)
            addStringField('Hazard_Classification__c', form.hazardClassification, 120);

            // 21-22. System Design Picklists
            addPicklistField('Hazard_Classification_Density_Required__c', form.densityRequired);
            addPicklistField('Attic_Sprinklers_Req__c', form.atticSprinklersRequired);

            // 23-25. Head Types
            addStringField('Attic_Head_Type__c', form.headTypesAttic, 120);
            addStringField('Ceiling_Head_Type__c', form.headTypesCeiling, 120);
            addStringField('Sandpipe_Qty_and_Hose_Valves__c', form.standpipeQty, 120);

            // 26. Temp SP Required
            addPicklistField('Temp_SP_Required__c', form.tempSpRequired);

            // 27. Fire Pump (STRING 120) - Combined field
            let firePumpParts = [];
            if (form.firePumpGpm === true) firePumpParts.push('GPM');
            if (form.firePumpPsi === true) firePumpParts.push('PSI');
            if (form.firePumpVoltage === true) firePumpParts.push('Voltage');
            if (form.firePumpTransferSwitch === true) firePumpParts.push('Transfer Switch');
            if (firePumpParts.length > 0) {
                addStringField('Fire_Pump__c', firePumpParts.join(', '), 120);
            }

            // 28-32. Material Picklists
            addPicklistField('Buy_American__c', form.buyAmerican);
            addPicklistField('Steel_Pipe__c', form.steelPipe);
            addPicklistField('Import_Pipe__c', form.importPipe);
            addPicklistField('Dynaflow_Dynathread_OK__c', form.dynaflow);
            addPicklistField('CPVC__c', form.cpvc);

            // 33-36. Head Counts and Colors
            addNumberField('Ceiling_Heads__c', form.ceilingHeads);
            addStringField('Color_of_Ceiling_Heads__c', form.headTypeColorCeiling, 120);
            addNumberField('Attic_Heads__c', form.atticHeads);
            addStringField('Color_of_Attic_Heads__c', form.headTypeColorAttic, 120);

            // 37. # of FDC (STRING 120)
            addStringField('of_FDC__c', form.fdcCount, 120);

            // 38. Type FDC (PICKLIST 255 - SINGLE SELECT)
            // Priority order: Free Standing > 2 Way > 3 Way > 4 Way > SP > Flush > CH > POL BR
            let fdcType = null;
            if (form.fdcType_FreeStanding) fdcType = 'Free Standing';
            else if (form.fdcType_2Way) fdcType = '2 Way';
            else if (form.fdcType_3Way) fdcType = '3 Way';
            else if (form.fdcType_4Way) fdcType = '4 Way';
            else if (form.fdcType_SP) fdcType = 'SP';
            else if (form.fdcType_Flush) fdcType = 'Flush';
            else if (form.fdcType_CH) fdcType = 'CH';
            else if (form.fdcType_PolBR) fdcType = 'POL BR';

            if (fdcType) {
                fieldData.Type_FDC__c = fdcType;
            }

            // 39-43. Metraflex and Flexheads
            addPicklistField('Metraflex_Loops__c', form.metraflexLoops);
            addNumberField('Metraflex_Loops_Qty__c', form.metraflexQty);
            addStringField('Metraflex_Loops_Size__c', form.metraflexSize, 120);
            addPicklistField('Flexheads__c', form.flexheads);
            addNumberField('Flexheads_Qty__c', form.flexheadsQty);

            // 44. Underground Scope (MULTIPICKLIST 4099)
            let undergroundScope = [];
            if (form.trenching) undergroundScope.push('Trenching');
            if (form.sawcut) undergroundScope.push('Sawcut');
            if (form.import) undergroundScope.push('Import');
            if (form.export) undergroundScope.push('Export');
            if (form.pave) undergroundScope.push('Pave');
            if (undergroundScope.length > 0) {
                fieldData.Underground_Scope__c = undergroundScope.join(';').substring(0, 4099);
            }

            // 45. Backflow (PICKLIST 255)
            if (form.backflowDDCV !== undefined && form.backflowDDCV !== null) {
                fieldData.Backflow__c = form.backflowDDCV ? 'DDCV' : 'Reduced Pressure';
            }

            // 46-54. Equipment (Scissor Lifts, Boom Lifts, Forklift)
            addPicklistField('Scissor_Lifts__c', form.scissorLifts);
            addNumberField('Scissor_Lifts_Months__c', form.scissorLiftsMonths);
            addStringField('Scissor_Lifts_Size__c', form.scissorLiftsSize, 120);

            addPicklistField('Boom_Lifts__c', form.boomLifts);
            addNumberField('Boom_Lifts_Months__c', form.boomLiftsMonths);
            addStringField('Boom_Lifts_Size__c', form.boomLiftsSize, 120);

            addPicklistField('Forklift__c', form.forklift);
            addNumberField('Forklift_Months__c', form.forkliftMonths);
            addStringField('Forklift_Size__c', form.forkliftSize, 120);

            // 55-56. Hours
            addNumberField('Design_Hours_Inc_BIM__c', form.designHours);
            addNumberField('Field_Hours__c', form.fieldHours);
            addNumberField('FAB__c', form.fab);
            addNumberField('FM_200__c', form.fm200);
            addStringField('Comments__c', form.comments, 120);
            return fieldData;
        }
        return fieldData;
    }
}