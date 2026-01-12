({
    doInit: function (component, event, helper) {
        var recordId = component.get('v.recordId');
        var tabApiName = component.get('v.tabApiName');
        var navService = component.find('navService');

        var pageReference = {
            type: 'standard__navItemPage',
            attributes: {
                apiName: tabApiName
            },
            state: {
                c__recordId: recordId
            }
        };

        navService.generateUrl(pageReference).then(
            $A.getCallback(function (url) {
                window.open(url, '_blank');
                $A.get('e.force:closeQuickAction').fire();
            }),
            $A.getCallback(function (error) {
                // eslint-disable-next-line no-console
                console.error('Failed to open ITM tab:', error);
                $A.get('e.force:closeQuickAction').fire();
            })
        );
    }
});