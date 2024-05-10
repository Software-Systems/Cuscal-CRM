// JavaScript source code
var formContext;

function form_onload(executionContext) {
    formContext = executionContext.getFormContext();
    if (formContext.ui.getFormType() == 1) {
        formContext.getAttribute('bacp_startdatetime').setValue(new Date());
        if (formContext.getAttribute('bacp_notificationtemplate').getValue() != null) {
            Template_OnChange();
        }
    }
    else {
        if (!formContext.getAttribute('softsys_z_firstload').getValue()) {
            Xrm.Utility.showProgressIndicator();
            var data = {};
            data.softsys_z_firstload = true;
            Xrm.WebApi.updateRecord(formContext.data.entity.getEntityName(), formContext.data.entity.getId().replace('{', '').replace('}', ''), data).then(
                function success(result) {
                    // perform operations on record update
                    var entityFormOptions = {
                        entityId: formContext.data.entity.getId().replace('{', '').replace('}', ''),
                        entityName: formContext.data.entity.getEntityName()
                    };
                    var formParameters = {};
                    Xrm.Utility.closeProgressIndicator();

                    Xrm.Navigation.openForm(entityFormOptions, formParameters).then(
                        function (result) {
                            // perform operations if record is saved in the quick create form

                        },
                        function (error) {
                            console.log(error.message);
                            // handle error conditions
                        }
                    );
                },
                function (error) {
                    console.log(error.message);
                    // handle error conditions
                }
            );
        }
    }
    formContext.getAttribute('bacp_notificationtemplate').addOnChange(Template_OnChange);
    formContext.getAttribute('bacp_rollupquery').addOnChange(Query_OnChange);
    formContext.getAttribute('softsys_recipientcriteriatemplate').addOnChange(QueryTemplate_Onchange);

    lockMList();
    lockRecipientCriteria();
    formContext.getAttribute('softsys_recipientcriteriatemplate').addOnChange(lockMList);
    formContext.getAttribute('softsys_recipientcriteriatemplate').addOnChange(lockRecipientCriteria);
    formContext.getAttribute('softsys_marketinglist').addOnChange(lockMList);
    formContext.getAttribute('softsys_marketinglist').addOnChange(lockRecipientCriteria);
    //COPY IMPACTED SERVICES FROM PARENT OF CLONED RECORD
    if (formContext.getAttribute('softsys_clonereference').getValue() != null && formContext.getControl("Products").getGrid().getTotalRecordCount() < 1) {
        var cloneRefId = formContext.getAttribute('softsys_clonereference').getValue()[0].id.replace('{', '').replace('}', '');
        var products = new Array();
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/products?fetchXml=%3Cfetch%20version%3D%221.0%22%20output-format%3D%22xml-platform%22%20mapping%3D%22logical%22%20distinct%3D%22true%22%3E%3Centity%20name%3D%22product%22%3E%3Cattribute%20name%3D%22name%22%20%2F%3E%3Cattribute%20name%3D%22productid%22%20%2F%3E%3Corder%20attribute%3D%22productnumber%22%20descending%3D%22false%22%20%2F%3E%3Clink-entity%20name%3D%22softsys_bacp_casemanagementnotification_product%22%20from%3D%22productid%22%20to%3D%22productid%22%20visible%3D%22false%22%20intersect%3D%22true%22%3E%3Clink-entity%20name%3D%22bacp_casemanagementnotification%22%20from%3D%22bacp_casemanagementnotificationid%22%20to%3D%22bacp_casemanagementnotificationid%22%20alias%3D%22aa%22%3E%3Cfilter%20type%3D%22and%22%3E%3Ccondition%20attribute%3D%22bacp_casemanagementnotificationid%22%20operator%3D%22eq%22%20value%3D%22" + cloneRefId + "%22%20%2F%3E%3C%2Ffilter%3E%3C%2Flink-entity%3E%3C%2Flink-entity%3E%3C%2Fentity%3E%3C%2Ffetch%3E", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    var relatedEnts = new Array();
                    for (var i = 0; i < results.value.length; i++) {
                        relatedEnts.push({ entityType: "product", id: results.value[i].productid })
                    }
                    var manyToManyAssociateRequest = {
                        getMetadata: () => ({
                            boundParameter: null,
                            parameterTypes: {},
                            operationType: 2,
                            operationName: "Associate"
                        }),


                        relationship: "softsys_bacp_casemanagementnotification_product",


                        target: {
                            entityType: "bacp_casemanagementnotification",
                            id: formContext.data.entity.getId().replace('{', '').replace('}', '')
                        },

                        relatedEntities: relatedEnts
                    }


                    Xrm.WebApi.online.execute(manyToManyAssociateRequest).then(
                        (success) => {
                            console.log("Success", success);
                            formContext.data.refresh(false).then(
                                function success() {
                                    formContext.getControl('Products').refresh();
                                },
                                function (error) {
                                    console.log(error.message);
                                    // handle error conditions
                                }
                            );
                        },
                        (error) => {
                            console.log("Error", error);
                        }
                    )
                }
            }
        };
        req.send();
    }
}

function form_onsave(executionContext) {
    lockMList();
    lockRecipientCriteria();
    Emailtemplatetext_OnChange();
}

function Emailtemplatetext_OnChange() {
    if (formContext.getAttribute('bacp_emailtext').getValue() != null && formContext.getAttribute('bacp_emailtext').getIsDirty()) {
        var emailText = stripHtml(formContext.getAttribute('bacp_emailtext').getValue());
        if (emailText != null && emailText.length > 0) {
            formContext.getAttribute('bacp_smstext').setValue(emailText.trim());
        }
    }
}

function lockMList(executionContext) {
    if (formContext.getAttribute('softsys_marketinglist').getValue() == null && formContext.getAttribute('softsys_recipientcriteriatemplate').getValue() != null) {
        formContext.getControl('softsys_marketinglist').setDisabled(true);
    }
    else {
        formContext.getControl('softsys_marketinglist').setDisabled(false);
    }
}

function lockRecipientCriteria() {
    if (formContext.getAttribute('softsys_marketinglist').getValue() != null && formContext.getAttribute('softsys_recipientcriteriatemplate').getValue() == null) {
        formContext.getControl('softsys_recipientcriteriatemplate').setDisabled(true);
    }
    else {
        formContext.getControl('softsys_recipientcriteriatemplate').setDisabled(false);
    }
}

function ProcessFillers() {
    try {
        if (formContext.ui.getFormType() != 2) return;

        var cons = formContext.ui.tabs.get('general').sections.get('ContentSection').controls;

        if ((formContext.getAttribute('bacp_emailtext').getValue() || formContext.getAttribute('bacp_smstext').getValue())) {
            var email = formContext.getAttribute('bacp_emailtext').getValue();
            var emailText = stripHtml(email);
            if (emailText != null && emailText.length > 0) {
                formContext.getAttribute('bacp_smstext').setValue(emailText.trim());
            }
            var sms = formContext.getAttribute('bacp_smstext').getValue();

            cons.forEach(con => {
                var token = '{' + con.getLabel() + '}';
                //process email text
                email = mergeFillers(token, con, email);
                //process sms text
                sms = mergeFillersSMS(token, con, sms);
            });

            //search for associated products
            var products = new Array();
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/products?fetchXml=%3Cfetch%20version%3D%221.0%22%20output-format%3D%22xml-platform%22%20mapping%3D%22logical%22%20distinct%3D%22true%22%3E%3Centity%20name%3D%22product%22%3E%3Cattribute%20name%3D%22name%22%20%2F%3E%3Corder%20attribute%3D%22productnumber%22%20descending%3D%22false%22%20%2F%3E%3Clink-entity%20name%3D%22softsys_bacp_casemanagementnotification_product%22%20from%3D%22productid%22%20to%3D%22productid%22%20visible%3D%22false%22%20intersect%3D%22true%22%3E%3Clink-entity%20name%3D%22bacp_casemanagementnotification%22%20from%3D%22bacp_casemanagementnotificationid%22%20to%3D%22bacp_casemanagementnotificationid%22%20alias%3D%22aa%22%3E%3Cfilter%20type%3D%22and%22%3E%3Ccondition%20attribute%3D%22bacp_casemanagementnotificationid%22%20operator%3D%22eq%22%20value%3D%22" + formContext.data.entity.getId().replace('{', '').replace('}', '') + "%22%20%2F%3E%3C%2Ffilter%3E%3C%2Flink-entity%3E%3C%2Flink-entity%3E%3C%2Fentity%3E%3C%2Ffetch%3E", false);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var results = JSON.parse(this.response);
                        for (var i = 0; i < results.value.length; i++) {
                            products.push(results.value[i].name);
                        }
                    }
                }
            };
            req.send();

            //merge product name
            if (products.length > 0) {
                var html = '<ul type="disc">';
                for (let index = 0; index < products.length; index++) {
                    const element = products[index];
                    html += '<li><span style="color: #4d4e53;"><span style="font-family: Verdana,sans-serif;"><span style="font-size: 9.0pt;">';
                    html += element;
                    html += '</span></span></span></li>';
                }
                html += '</ul>';
                if (email) email = email.replace('{Impacted Services}', html);
                if (sms) sms = sms.replace('{Impacted Services}', products.join(" ,"));
            }
            else {
                if (email) email = email.replace('{Impacted Services}', '');
                if (sms) sms = sms.replace('{Impacted Services}', '');
            }

            if (email) formContext.getAttribute('softsys_draftemail').setValue(email);
            if (sms) {
                if (sms.indexOf("You have received") > 0) {
                    sms = sms.substring(0, sms.indexOf("You have received"));
                }
                formContext.getAttribute('softsys_draftsms').setValue(sms.trim());
            }

            //refreshEmailEditor();

        }
    }
    catch (error) {

    }

    function mergeFillers(token, con, text) {
        if (text.indexOf(token) > -1) {
            if (con.getAttribute().getValue() != null) {
                if (con.getAttribute().getFormat() == "datetime") {
                    var datetime = con.getAttribute().getValue().toLocaleString('en-AU');
                    text = text.replace(token, datetime);
                }
                else if (con.getAttribute().getAttributeType() == 'multiselectoptionset') {
                    var optionsSelected = con.getAttribute().getText();
                    if (optionsSelected.length > 0) {
                        /*
                        var html = '<ul type="disc">';
                        for (let index = 0; index < optionsSelected.length; index++) {
                            const element = optionsSelected[index];
                            html += '<li><span style="color: black;"><span style="font-family: Verdana,sans-serif;"><span style="font-size: 9.0pt;">';
                            html += element;
                            html += '</span></span></span></li>';
                        }
                        html += '</ul>';
                        */
                        text = text.replace(token, optionsSelected.join(", "));
                    }
                }
                else {
                    text = text.replace(token, con.getAttribute().getValue());
                }
            } else {
                text = text.replace(token, '');
            }
        }
        return text;
    }

    function mergeFillersSMS(token, con, text) {
        if (text.indexOf(token) > -1) {
            if (con.getAttribute().getValue() != null) {
                if (con.getAttribute().getFormat() == "datetime") {
                    var datetime = con.getAttribute().getValue().toLocaleString('en-AU');
                    text = text.replace(token, datetime);
                }
                else if (con.getAttribute().getAttributeType() == 'multiselectoptionset') {
                    var optionsSelected = con.getAttribute().getText();
                    text = text.replace(token, optionsSelected.join(' ,'));

                }
                else {
                    text = text.replace(token, con.getAttribute().getValue());
                }
            } else {
                text = text.replace(token, '');
            }
        }
        return text;
    }
}

//Confirm Filler switch

function filler_confirm() {
    try {
        if (formContext.getAttribute('bacp_confirmfillers').getValue()) {
            var cons = formContext.ui.tabs.get('general').sections.get('ContentSection').controls;
            cons.forEach(con => {
                con.setDisabled(true);
            });
        }
    } catch (error) {

    }

}

//Notification Template on change

function Template_OnChange() {
    if (formContext.getAttribute('bacp_notificationtemplate').getValue()) {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/bacp_notificationtemplates(" + formContext.getAttribute("bacp_notificationtemplate").getValue()[0].id.replace("{", "").replace("}", "") + ")", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.response);
                    if (result["bacp_emailtemplate"] && result["bacp_emailtemplate"].length > 0) {
                        formContext.getAttribute('bacp_emailtext').setValue(result["bacp_emailtemplate"]);
                        //refreshEmailEditor();
                    }
                    if (result["bacp_smstemplate"] && result["bacp_smstemplate"].length > 0) {
                        formContext.getAttribute('bacp_smstext').setValue(result["bacp_smstemplate"].trim());

                    }
                    formContext.getAttribute('bacp_fromrecordowner').setValue(result["bacp_fromrecordowner"]);
                    formContext.getAttribute('bacp_name').setValue(result["bacp_subject"]);
                    if (result["_bacp_fromqueue_value"]) {
                        formContext.getAttribute('bacp_fromqueue').setValue([{
                            id: result["_bacp_fromqueue_value"],
                            name: result["_bacp_fromqueue_value@OData.Community.Display.V1.FormattedValue"],
                            entityType: result["_bacp_fromqueue_value@Microsoft.Dynamics.CRM.lookuplogicalname"]
                        }]);
                    }
                    if (result["_bacp_fromuser_value"]) {
                        formContext.getAttribute('bacp_fromuser').setValue([{
                            id: result["_bacp_fromuser_value"],
                            name: result["_bacp_fromuser_value@OData.Community.Display.V1.FormattedValue"],
                            entityType: "systemuser"
                        }]);
                    }
                    if (result["_bacp_rollupquery_value"]) {
                        formContext.getAttribute('softsys_recipientcriteriatemplate').setValue([{
                            id: result["_bacp_rollupquery_value"],
                            name: result["_bacp_rollupquery_value@OData.Community.Display.V1.FormattedValue"],
                            entityType: result["_bacp_rollupquery_value@Microsoft.Dynamics.CRM.lookuplogicalname"]
                        }]);
                    }

                    //test contact
                    formContext.getAttribute('emailaddress').setValue(result['softsys_testemailaddress']);
                    formContext.getAttribute('bacp_testsmsno').setValue(result['softsys_testsms']);

                    //New fillers
                    if (result["softsys_highlevelproductcategory"]) {
                        var hlpc = [];
                        var hlpc1 = [];
                        hlpc = result["softsys_highlevelproductcategory"].split(',');
                        for (let index = 0; index < hlpc.length; index++) {
                            hlpc1.push(parseInt(hlpc[index]));
                        }
                        formContext.getAttribute('softsys_highlevelproductcategory').setValue(hlpc1);
                    }
                    QueryTemplate_Onchange();

                    if (formContext.ui.getFormType() == 2) {
                        var req1 = new XMLHttpRequest();
                        req1.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/products?fetchXml=%3Cfetch%20version%3D%221.0%22%20output-format%3D%22xml-platform%22%20mapping%3D%22logical%22%20distinct%3D%22true%22%3E%3Centity%20name%3D%22product%22%3E%3Cattribute%20name%3D%22name%22%20%2F%3E%3Cattribute%20name%3D%22productid%22%20%2F%3E%3Clink-entity%20name%3D%22softsys_bacp_notificationtemplate_product%22%20from%3D%22productid%22%20to%3D%22productid%22%20visible%3D%22false%22%20intersect%3D%22true%22%3E%3Clink-entity%20name%3D%22bacp_notificationtemplate%22%20from%3D%22bacp_notificationtemplateid%22%20to%3D%22bacp_notificationtemplateid%22%20alias%3D%22aa%22%3E%3Cfilter%20type%3D%22and%22%3E%3Ccondition%20attribute%3D%22bacp_notificationtemplateid%22%20operator%3D%22eq%22%20value%3D%22%7B" + result["bacp_notificationtemplateid"] + "%7D%22%20%2F%3E%3C%2Ffilter%3E%3C%2Flink-entity%3E%3C%2Flink-entity%3E%3C%2Fentity%3E%3C%2Ffetch%3E", true);
                        req1.setRequestHeader("OData-MaxVersion", "4.0");
                        req1.setRequestHeader("OData-Version", "4.0");
                        req1.setRequestHeader("Accept", "application/json");
                        req1.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                        req1.onreadystatechange = function () {
                            if (this.readyState === 4) {
                                req1.onreadystatechange = null;
                                if (this.status === 200) {
                                    var results = JSON.parse(this.response);
                                    var relatedEnts = new Array();
                                    for (var i = 0; i < results.value.length; i++) {
                                        relatedEnts.push({ entityType: "product", id: results.value[i].productid })
                                    }
                                    if (relatedEnts.length > 0) {
                                        var manyToManyAssociateRequest = {
                                            getMetadata: () => ({
                                                boundParameter: null,
                                                parameterTypes: {},
                                                operationType: 2,
                                                operationName: "Associate"
                                            }),


                                            relationship: "softsys_bacp_casemanagementnotification_product",


                                            target: {
                                                entityType: "bacp_casemanagementnotification",
                                                id: formContext.data.entity.getId().replace('{', '').replace('}', '')
                                            },

                                            relatedEntities: relatedEnts
                                        }


                                        Xrm.WebApi.online.execute(manyToManyAssociateRequest).then(
                                            (success) => {
                                                console.log("Success", success);
                                                //refreshEmailEditor();
                                            },
                                            (error) => {
                                                console.log("Error", error);
                                            }
                                        );
                                    }
                                    else {
                                        //refreshEmailEditor();
                                    }
                                }
                            }
                        };
                        req1.send();
                    }
                    else {
                        //refreshEmailEditor();
                    }
                }
            }
        };
        req.send();
    }
}

//Rollup Query On change
function Query_OnChange() {
    if (formContext.getAttribute("bacp_rollupquery").getValue()) {
        //retreive contact xml from rollup query
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/goalrollupqueries(" + formContext.getAttribute("bacp_rollupquery").getValue()[0].id.replace("{", "").replace("}", "") + ")?$select=fetchxml", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.response);
                    formContext.getAttribute('bacp_contactsxml').setValue(result["fetchxml"]);

                }
            }
        };
        req.send();
    } else {
        formContext.getAttribute('bacp_contactsxml').setValue(null);
    }
}

function QueryTemplate_Onchange() {
    if (formContext.getAttribute('softsys_recipientcriteriatemplate').getValue() != null) {
        Xrm.WebApi.retrieveRecord('goalrollupquery', formContext.getAttribute("softsys_recipientcriteriatemplate").getValue()[0].id.replace("{", "").replace("}", "")).then(
            function success(result) {
                // perform operations on record retrieval
                var data = {};
                data["fetchxml"] = result["fetchxml"];
                var dateTime = new Date();

                var name = dateTime.getFullYear() + "-" + (dateTime.getMonth() + 1) + "-" + dateTime.getDate() + " " + (dateTime.getHours() < 10 ? '0' : '') + dateTime.getHours() + ":" + (dateTime.getMinutes() < 10 ? '0' : '') + dateTime.getMinutes() + " " + result["name"];
                data["name"] = name;
                data["queryentitytype"] = result["queryentitytype"];
                Xrm.WebApi.createRecord('goalrollupquery', data).then(
                    function success(result) {
                        // perform operations on record creation

                        formContext.getAttribute('bacp_rollupquery').setValue([{
                            id: result.id,
                            name: name,
                            entityType: "goalrollupquery"
                        }]);
                        Query_OnChange();

                        formContext.data.entity.save().then(
                            function success(result) {
                                if (!formContext.getAttribute('softsys_z_firstload').getValue()) {
                                    Xrm.Utility.showProgressIndicator();
                                    var data = {};
                                    data.softsys_z_firstload = true;
                                    Xrm.WebApi.updateRecord(formContext.data.entity.getEntityName(), formContext.data.entity.getId().replace('{', '').replace('}', ''), data).then(
                                        function success(result) {
                                            // perform operations on record update
                                            var entityFormOptions = {
                                                entityId: formContext.data.entity.getId().replace('{', '').replace('}', ''),
                                                entityName: formContext.data.entity.getEntityName()
                                            };
                                            var formParameters = {}
                                            Xrm.Utility.closeProgressIndicator();
                                            Xrm.Navigation.openForm(entityFormOptions, formParameters).then(
                                                function (result) {
                                                    // perform operations if record is saved in the quick create form

                                                },
                                                function (error) {
                                                    console.log(error.message);
                                                    // handle error conditions
                                                }
                                            );
                                        },
                                        function (error) {
                                            console.log(error.message);
                                            // handle error conditions
                                        }
                                    );
                                }
                            },
                            function error(error) {

                            }
                        );
                    },
                    function (error) {
                        console.log(error.message);
                        // handle error conditions
                    }
                );
            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }
}

//Generate Contacts
function GenerateContacts() {
    if (formContext.getControl("RecipientsGrid").getGrid().getTotalRecordCount() > 0) {
        var confirmStrings = {
            cancelButtonLabel: 'Cancel',
            confirmButtonLabel: 'Confirm',
            subtitle: 'Click confirm to proceed',
            text: 'This will add the new contacts on to the existing list displayed below.',
            title: 'Confirmation Dialog'
        };
        var confirmOptions = { height: 200, width: 450 };

        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
            function (success) {
                if (success.confirmed) {
                    Query_OnChange();
                    formContext.getAttribute('bacp_z_button_contacts').setValue(true);
                    formContext.data.entity.save();
                    formContext.getControl('RecipientsGrid').setFocus();
                    formContext.getControl('RecipientsGrid').refresh();

                } else {
                    console.log('Dialog closed using Cancel button or X.');
                }
            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }
    else {
        Query_OnChange();
        formContext.getAttribute('bacp_z_button_contacts').setValue(true);
        formContext.data.entity.save();
        formContext.getControl('RecipientsGrid').setFocus();
        formContext.getControl('RecipientsGrid').refresh();
    }
}

function Disassociate() {
    var req = new XMLHttpRequest();
    req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/bacp_bacp_casemanagementnotification_contactset?$select=contactid&$filter=bacp_casemanagementnotificationid eq " + Xrm.Page.data.entity.getId().replace("{", "").replace("}", ""), true);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 200) {
                var results = JSON.parse(this.response);
                for (var i = 0; i < results.value.length; i++) {
                    var contactid = results.value[i]["contactid"];
                    SDK.WEBAPI.diassociateEntities(
                        "bacp_casemanagementnotification",
                        Xrm.Page.data.entity.getId().replace("{", "").replace("}", ""),
                        contactid.replace("{", "").replace("}", ""),
                        "bacp_bacp_casemanagementnotification_contact"
                    );
                }
            } else {

                var alertStrings = {
                    confirmButtonLabel: 'Ok',
                    text: this.statusText
                };
                var alertOptions = {
                    height: 120,
                    width: 260
                };

                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                    function success() {
                        // perform operations on alert dialog close

                    },
                    function (error) {
                        console.log(error.message);
                        // handle error conditions
                    }
                );
            }
        }
    };
    req.send();
}

function Associate() {
    if (formContext.getAttribute("bacp_contactsxml").getValue()) {
        var contacts = SDK.WEBAPI.executeFetch(formContext.getAttribute("bacp_contactsxml").getValue(), true, 'contact');
        if (contacts.length > 0) {
            for (var i = 0; i < contacts.length; i++) {
                SDK.WEBAPI.associateEntities(
                    "bacp_casemanagementnotification",
                    Xrm.Page.data.entity.getId().replace("{", "").replace("}", ""),
                    "contact",
                    contacts[i].contactid.replace("{", "").replace("}", ""),
                    "bacp_bacp_casemanagementnotification_contact"
                );
            }
        }
    }
}

//Generate Contacts

//Email template
var templateId;
var templateViewId = "4D5BE2E9-6828-482B-99AA-2387AFED7B37";

function EmailTemplate() {
    try {

        var lookupOptions = {};
        lookupOptions.allowMultiSelect = false;
        lookupOptions.entityTypes = ["template"];
        lookupOptions.defaultEntityType = "template";
        lookupOptions.defaultViewId = templateViewId;

        Xrm.Utility.lookupObjects(lookupOptions).then(
            function succes(result) {
                templateId = result[0].id;
                var templateContent = InstTemplate();
                formContext.getAttribute('bacp_emailtext').setValue(templateContent.value[0].description);
                //refreshEmailEditor();
                formContext.getAttribute('bacp_name').setValue(templateContent.value[0].subject);
                var oldsrc = formContext.getControl('WebResource_Email').getSrc();
                formContext.getControl('WebResource_Email').setSrc('about:blank');
                formContext.getControl('WebResource_Email').setSrc(oldsrc);
                formContext.ui.tabs.get('content').setFocus();
            }
        )
    } catch (e) {
        var alertStrings = {
            confirmButtonLabel: 'Ok',
            text: "Email Template error"
        };
        var alertOptions = {
            height: 120,
            width: 260
        };

        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success() {
                // perform operations on alert dialog close

            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }
}

//Instantiate email template and generate subject & body
function InstTemplate() {
    try {
        var results;
        var parameters = {};
        parameters.TemplateId = templateId;
        parameters.ObjectType = "bacp_casemanagementnotification";
        parameters.ObjectId = Xrm.Page.data.entity.getId().replace("{", "").replace("}", "");

        var req = new XMLHttpRequest();
        req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/InstantiateTemplate", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    results = JSON.parse(this.response);
                } else {
                    var alertStrings = {
                        confirmButtonLabel: 'Ok',
                        text: "Inst. Template error"
                    };
                    var alertOptions = {
                        height: 120,
                        width: 260
                    };

                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                        function success() {
                            // perform operations on alert dialog close

                        },
                        function (error) {
                            console.log(error.message);
                            // handle error conditions
                        }
                    );
                }
            }
        };
        req.send(JSON.stringify(parameters));
        return results;
    } catch (e) {
        Xrm.Utility.alertDialog(e);
        Xrm.Utility.closeProgressIndicator();
    }
}
//Email template

//SMS Template
function SMSTemplate() {
    try {

        var lookupOptions = {};
        lookupOptions.allowMultiSelect = false;
        lookupOptions.entityTypes = ["template"];
        lookupOptions.defaultEntityType = "template";
        lookupOptions.defaultViewId = templateViewId;

        Xrm.Utility.lookupObjects(lookupOptions).then(
            function succes(result) {
                templateId = result[0].id;
                var templateContent = InstTemplate();
                var text = templateContent.value[0].subject + "\n\n" + templateContent.value[0].description;
                formContext.getAttribute('bacp_smstext').setValue(text.trim());
                formContext.getControl('bacp_smstext').setFocus();
            }
        )
    } catch (e) {
        var alertStrings = {
            confirmButtonLabel: 'Ok',
            text: "SMS Template error"
        };
        var alertOptions = {
            height: 120,
            width: 260
        };

        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success() {
                // perform operations on alert dialog close

            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }
}
//SMS Template

//Refresh Orgs
function RefreshOrgs() {
    try {
        formContext.getAttribute('bacp_refreshorgsaffected').setValue(true);
        formContext.data.entity.save();
        formContext.getControl('AffectedOrgs').setFocus();
    } catch (error) {

    }
}
//Refresh Orgs

//Send Email
function SendEmail() {
    if (formContext.getAttribute('bacp_z_button_sendemail').getValue()) {
        var alertStrings = { confirmButtonLabel: 'Ok', text: 'Emails already sent out for this record. Please check status under Emails Tab.' };
        var alertOptions = { height: 180, width: 450 };

        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success() {
                formContext.ui.tabs.get('tab_3').setDisplayState('expanded');
                formContext.ui.tabs.get('tab_3').setFocus();
            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    } else {
        formContext.getControl('RecipientsGrid').setFocus();
        formContext.getControl("RecipientsGrid").refresh();
        setTimeout(() => {
            if (formContext.getControl("RecipientsGrid").getGrid().getTotalRecordCount() > 0) {
                ProcessFillers();
                var confirmStrings = {
                    cancelButtonLabel: 'Cancel',
                    confirmButtonLabel: 'Confirm',
                    subtitle: 'Click confirm to proceed',
                    text: 'This will send email to ' + formContext.getControl("RecipientsGrid").getGrid().getTotalRecordCount() + ' contacts. Are you sure?',
                    title: 'Confirmation Dialog'
                };
                var confirmOptions = { height: 200, width: 450 };

                Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                    function (success) {
                        if (success.confirmed) {
                            formContext.getAttribute('bacp_z_button_sendemail').setValue(true);
                            formContext.getAttribute('softsys_sent').setValue(new Date());
                            formContext.data.entity.save();
                            formContext.ui.tabs.get('tab_3').setDisplayState('expanded');
                            formContext.ui.tabs.get('tab_3').setFocus();
                        }

                        //refreshEmailEditor();
                    },
                    function (error) {
                    }
                );
            }
            else {
                var alertStrings = { confirmButtonLabel: 'Ok', text: 'No contacts in the list', title: 'Alert' };
                var alertOptions = { height: 120, width: 260 };

                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                    function success() {
                        // perform operations on alert dialog close

                    },
                    function (error) {
                        console.log(error.message);
                        // handle error conditions
                    }
                );
            }
        }, 1000);

    }
}
//Send Email

//Send SMS
function SendSMS() {
    if (formContext.getAttribute('bacp_z_button_sendsms').getValue()) {
        var alertStrings = { confirmButtonLabel: 'Ok', text: 'SMS already sent out for this record. Please check status under SMS Tab.', title: 'Alert' };
        var alertOptions = { height: 180, width: 450 };

        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success() {
                // perform operations on alert dialog close

            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }
    else if (formContext.getAttribute('bacp_smstext').getValue() == null) {

        var alertStrings = { confirmButtonLabel: 'Ok', text: 'There is no content in SMS to send', title: 'Alert' };
        var alertOptions = { height: 180, width: 450 };

        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success() {
                // perform operations on alert dialog close

            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }
    else {
        formContext.getControl('RecipientsGrid').setFocus();
        formContext.getControl("RecipientsGrid").refresh();
        setTimeout(() => {
            if (formContext.getControl("RecipientsGrid").getGrid().getTotalRecordCount() > 0) {
                ProcessFillers();
                if (formContext.getAttribute('softsys_draftsms').getValue() != null && formContext.getAttribute('softsys_draftsms').getValue().length > 1600) {
                    var alertStrings = { confirmButtonLabel: 'Ok', text: 'SMS length exceeds 1600 characters. Please edit and try again.', title: 'Alert' };
                    var alertOptions = { height: 120, width: 260 };

                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                        function success() {
                            // perform operations on alert dialog close

                        },
                        function (error) {
                            console.log(error.message);
                            // handle error conditions
                        }
                    );
                    return;
                }
                var confirmStrings = {
                    cancelButtonLabel: 'Cancel',
                    confirmButtonLabel: 'Confirm',
                    subtitle: 'Click confirm to proceed',
                    text: 'This will send sms to ' + formContext.getControl("RecipientsGrid").getGrid().getTotalRecordCount() + ' contacts. Are you sure?',
                    title: 'Confirmation Dialog'
                };
                var confirmOptions = { height: 200, width: 450 };

                Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                    function (success) {
                        if (success.confirmed) {
                            formContext.getAttribute('bacp_z_button_sendsms').setValue(true);
                            formContext.getAttribute('softsys_sent').setValue(new Date());
                            formContext.data.entity.save();
                            formContext.ui.tabs.get('tab_4').setDisplayState('expanded');
                            formContext.ui.tabs.get('tab_4').setFocus();
                        }
                        //refreshEmailEditor();
                    },
                    function (error) {
                    }
                );
            }
            else {
                var alertStrings = { confirmButtonLabel: 'Ok', text: 'No contacts in the list', title: 'Alert' };
                var alertOptions = { height: 120, width: 260 };

                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                    function success() {
                        // perform operations on alert dialog close

                    },
                    function (error) {
                        console.log(error.message);
                        // handle error conditions
                    }
                );
            }
        }, 1000);

    }
}
//Send SMS

function SendTestMessage() {
    ProcessFillers();
    formContext.data.refresh(true).then(
        function success() {
            if (formContext.getAttribute('emailaddress').getValue() == null) {
                var alertStrings = { confirmButtonLabel: 'Ok', text: 'Please provide a test email address.', title: 'Alert' };
                var alertOptions = { height: 120, width: 260 };

                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                    function success() {
                        // perform operations on alert dialog close

                    },
                    function (error) {
                        console.log(error.message);
                        // handle error conditions
                    }
                );
            } else {
                SendTestEMail();
                formContext.ui.tabs.get('tab_3').setDisplayState('expanded');
                formContext.ui.tabs.get('tab_4').setDisplayState('expanded');
                formContext.ui.tabs.get('tab_3').setFocus();
                formContext.data.refresh(false).then(
                    function success() {
                        // perform operations on refresh

                    },
                    function (error) {
                        console.log(error.message);
                        // handle error conditions
                    }
                );
            }
            if (formContext.getAttribute('bacp_testsmsno').getValue() != null) {
                SendTestSMS();
            }

        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}

//Send Test Email
function SendTestEMail() {

    var entity = {};
    entity.bacp_z_testaction = 395280000;

    var req = new XMLHttpRequest();
    req.open("PATCH", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/bacp_casemanagementnotifications(" + formContext.data.entity.getId().replace("{", "").replace("}", "") + ")", false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
            } else {
                Xrm.Utility.alertDialog(this.statusText);
            }
        }
    };
    req.send(JSON.stringify(entity));
}
//Send Test Email

//Send Test SMS
function SendTestSMS() {

    if (formContext.getAttribute('softsys_draftsms').getValue() != null && formContext.getAttribute('softsys_draftsms').getValue().length > 1600) {
        var alertStrings = { confirmButtonLabel: 'Ok', text: 'SMS length exceeds 1600 characters. Please edit and try again.', title: 'Alert' };
        var alertOptions = { height: 120, width: 260 };

        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success() {
                // perform operations on alert dialog close

            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
        return;
    }
    var parameters = {};
    parameters.EntityId = formContext.data.entity.getId().replace("{", "").replace("}", "");

    var req = new XMLHttpRequest();
    req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/workflows(287D1211-F6C1-49A6-9740-ACEDD0B1D295)/Microsoft.Dynamics.CRM.ExecuteWorkflow", false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
            } else {
                Xrm.Utility.alertDialog(this.statusText);
            }
        }
    };
    req.send(JSON.stringify(parameters));
}
//Send Test SMS

//Clone

function clone() {
    var parameters = {};
    parameters.EntityId = formContext.data.entity.getId().replace("{", "").replace("}", "");

    var req = new XMLHttpRequest();
    req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/workflows(5B3B4F09-9035-4576-B8FF-BD849BEA0981)/Microsoft.Dynamics.CRM.ExecuteWorkflow", false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
                var req1 = new XMLHttpRequest();
                req1.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/bacp_casemanagementnotifications?$select=bacp_casemanagementnotificationid&$filter=_softsys_clonereference_value eq " + formContext.data.entity.getId().replace("{", "").replace("}", "") + "&$orderby=createdon desc", true);
                req1.setRequestHeader("OData-MaxVersion", "4.0");
                req1.setRequestHeader("OData-Version", "4.0");
                req1.setRequestHeader("Accept", "application/json");
                req1.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req1.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                req1.onreadystatechange = function () {
                    if (this.readyState === 4) {
                        req1.onreadystatechange = null;
                        if (this.status === 200) {
                            var results = JSON.parse(this.response);
                            if (results.value.length > 0) {
                                var bacp_casemanagementnotificationid = results.value[0]["bacp_casemanagementnotificationid"];
                                var entityFormOptions = {
                                    entityId: bacp_casemanagementnotificationid,
                                    entityName: 'bacp_casemanagementnotification'
                                };
                                var formParameters = {}

                                Xrm.Navigation.openForm(entityFormOptions, formParameters).then(
                                    function (result) {
                                        // perform operations if record is saved in the quick create form

                                    },
                                    function (error) {
                                        console.log(error.message);
                                        // handle error conditions
                                    }
                                );
                            }
                        }
                    }
                };
                req1.send();
            }
        }
    };
    req.send(JSON.stringify(parameters));
}

//Clone

//Clear
function clearList() {
    Xrm.Utility.showProgressIndicator();
    var data = {};
    data.softsys_z_button_clear = true;
    Xrm.WebApi.updateRecord('bacp_casemanagementnotification', formContext.data.entity.getId().replace('{', '').replace('}', ''), data).then(
        function success(result) {
            Xrm.Utility.closeProgressIndicator();
            formContext.getControl('RecipientsGrid').setFocus();
            formContext.getControl('RecipientsGrid').refresh();
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            console.log(error.message);
            // handle error conditions
        }
    );
}

//Clear

//refresh email editor
function refreshEmailEditor() {
    formContext.data.refresh(true).then(
        function success() {
            // perform operations on refresh
            debugger;
            var webResourceControl = Xrm.Page.getControl("WebResource_Email");
            var iframeSrc = Xrm.Page.getControl("WebResource_Email").getSrc();
            webResourceControl.setSrc("about:blank");
            setTimeout(() => {
                webResourceControl.setSrc(iframeSrc);
            }, 1000);
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}

/**
 * Returns the text from a HTML string
 * 
 * @param {html} String The html string
 */
function stripHtml(html) {
    // Create a new div element
    var temporalDivElement = document.createElement("div");
    // Set the HTML content with the providen
    temporalDivElement.innerHTML = html;
    // Retrieve the text property of the element (cross-browser support)
    var text = temporalDivElement.textContent || temporalDivElement.innerText || "";

    var newText = text.replace(/Notification\nInvestigation\nRestored\nResolved\n/, '');
    newText = newText.replace(/\n\n/g, '\n');
    newText = newText.replace(/\n\n/g, '\n');
    newText = newText.replace(/\n\n/g, '\n');
   
    return newText;
}
