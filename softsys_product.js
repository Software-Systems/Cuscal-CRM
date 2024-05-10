function platformTypeLoad(executionContext) {
    formContext = executionContext.getFormContext();
    var platformType = formContext.getAttribute("softsys_productplatform").getValue();
    if (platformType == null) {
        formContext.getAttribute("softsys_productplatformsubtype").setValue(null);
        formContext.getControl('softsys_productplatformsubtype').setDisabled(true);
    }
    else { 
        formContext.getControl('softsys_productplatformsubtype').setDisabled(false);
    }
}


function platformTypeChange(executionContext) {
    formContext = executionContext.getFormContext();
    var platformType = formContext.getAttribute("softsys_productplatform").getValue();
    if (platformType == null) {
        formContext.getAttribute("softsys_productplatformsubtype").setValue(null);
        formContext.getControl('softsys_productplatformsubtype').setDisabled(true);
    }
    else {
        formContext.getAttribute("softsys_productplatformsubtype").setValue(null);
        formContext.getControl('softsys_productplatformsubtype').setDisabled(false);
    }
}

function platformSubTypeChange(executionContext) {
    formContext = executionContext.getFormContext();
    var platformType = formContext.getAttribute("softsys_productplatform").getValue();
    var platformSubType = formContext.getAttribute("softsys_productplatformsubtype").getValue();
    if (platformType != null && platformSubType != null) {
        //Get related platform type
        Xrm.WebApi.retrieveRecord('softsys_productplatformsubtype', platformSubType[0].id.replace('{', '').replace('}', ''), "?$select=_softsys_productplatform_value,softsys_name").then(
            function success(result) {
                if (result != null && result["_softsys_productplatform_value"] != null) {
                    var relatedPlatformType = result["_softsys_productplatform_value"];
                    if (relatedPlatformType.toLowerCase() != platformType[0].id.replace('{', '').replace('}', '').toLowerCase()) {
                        formContext.getAttribute("softsys_productplatformsubtype").setValue(null);
                        Xrm.Navigation.openAlertDialog("Invalid subtype specified, please update the product platform subtype to a valid combination");
                    }
                }
                else {
                    formContext.getAttribute("softsys_productplatformsubtype").setValue(null);
                    Xrm.Navigation.openAlertDialog("Invalid subtype specified, please update the product platform subtype to a valid combination");
                }
            },
            function (error) {
            }
        );
    }
}