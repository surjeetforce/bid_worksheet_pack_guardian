({
    syncRecordIdFromPageReference: function (component) {
        var pageRef = component.get('v.pageReference');
        var state = pageRef && pageRef.state ? pageRef.state : null;

        // When opened as /lightning/n/<TabApiName>?c__recordId=<id>, the id comes through as state.c__recordId.
        // Keep a fallback for state.recordId in case other navigation methods are used.
        var recordId = state ? (state.c__recordId || state.recordId) : null;

        if (recordId) {
            component.set('v.recordId', recordId);
        }
    }
});