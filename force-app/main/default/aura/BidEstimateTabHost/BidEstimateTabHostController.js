({
    doInit: function (component, event, helper) {
        helper.syncRecordIdFromPageReference(component);
    },

    handlePageReferenceChange: function (component, event, helper) {
        helper.syncRecordIdFromPageReference(component);
    }
});