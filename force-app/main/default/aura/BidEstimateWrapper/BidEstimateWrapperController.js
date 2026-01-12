({
    doInit: function (component, event, helper) {
        var recordId = component.get('v.recordId');
        var navService = component.find('navService');

        var pageReference = {
            type: 'standard__navItemPage',
            attributes: {
                apiName: 'Estimates_Calculator'
            },
            state: {
                c__recordId: recordId
            }
        };

        navService.generateUrl(pageReference).then(
            $A.getCallback(function (url) {
                // Open the custom Lightning tab in a new browser tab
                window.open(url, '_blank');
                $A.get('e.force:closeQuickAction').fire();
            }),
            $A.getCallback(function (error) {
                // Fallback: close the quick action even if URL generation fails
                // eslint-disable-next-line no-console
                console.error('Failed to open Estimates_Calculator tab:', error);
                $A.get('e.force:closeQuickAction').fire();
            })
        );
    }
});