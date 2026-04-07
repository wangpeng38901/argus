<!--#INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2B_BulkIncomingResult.asp
' Description  : E2B Bulk Incoming Result
'******************************************************************************
' Revision History
' Date		   Author		Description
' 15MAY2006 	Umar       	Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title><%=GetTranslationData("E2B_BULK_INCOMING_RESULTS")%></title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<%
Dim E2BViewType, lError, sError

E2BViewType = GetRequest("E2BViewType")
If (Len(E2BViewType)= 0) Then E2BViewType = 1	'CIOMS
%>
<body onload="fn_init()">
    <form name="Frm_BulkImport" method="post">
        <%Call BuildHiddenControlDirect("Initial", "") %>
        <%Call BuildHiddenControlDirect("Initial_notes", "") %>
        <%Call BuildHiddenControlDirect("Followup", "") %>
        <%Call BuildHiddenControlDirect("Followup_notes", "") %>
        <%Call BuildHiddenControlDirect("Nullification", "") %>
        <%Call BuildHiddenControlDirect("Nullification_notes", "") %>
        <%Call BuildHiddenControlDirect("E2b_Type", "") %>
        <%Call BuildHiddenControlDirect("E2BViewType", E2BViewType) %>
        <!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
    </form>
    <table width="100%">
        <tr>
            <td width="60%" align="left">
                <%
                    Dim cfE2BBulkImportTabs
                    Set cfE2BBulkImportTabs = BuildTabs ("CF_BULK_IMPORT", 3, CONST_TABS_TOP, 1)
                    cfE2BBulkImportTabs.AddTab GetTranslationData("INITIAL"), "fn_view_display(1)"
                    cfE2BBulkImportTabs.AddTab GetTranslationData("E2B_RPT_TYPE_FOLLOW_UP"), "fn_view_display(2)"
                    cfE2BBulkImportTabs.AddTab GetTranslationData("E2B_RPT_TYPE_NULLIFICATION"), "fn_view_display(3)"
                    cfE2BBulkImportTabs.SetTabWidth "150px"
                    cfE2BBulkImportTabs.Render
                %>
            </td>
            <td width="40%" align="right">
                <button name="Btn_Print" class="button button-large" onclick="fn_print()"><%=GetTranslationData("PRINT")%></button>
                <button name="Btn_Close" class="button button-large" onclick="window.close()"><%=GetTranslationData("CLOSE")%></button>
            </td>
        </tr>
    </table>
    <div id="dInitialReport" style="position: absolute; width: 870px; height: 610px; z-index: 1; left: 20px;">
        <iframe id="initial_report" name="initial_report" width="100%" height="100%" framespacing="0" frameborder="0"></iframe>
    </div>
    <div id="dFollowupReport" style="position: absolute; width: 870px; height: 610px; z-index: 1; left: 20px;">
        <iframe id="followup_report" name="followup_report" width="100%" height="100%" framespacing="0" frameborder="0"></iframe>
    </div>
    <div id="dNullificationReport" style="position: absolute; width: 870px; height: 610px; z-index: 1; left: 20px;">
        <iframe id="nullification_report" name="nullification_report" width="100%" height="100%" framespacing="0" frameborder="0"></iframe>
    </div>
    <div id="dLoading" style="position: absolute; width: 812px; height: 600px; z-index: 1; left: 20px;">
        <span id='Span1'>
            <br>
            <table width='100%'>
                <tr>
                    <td align='middle'>
                        <img name='status' src='/img/Common/loader.gif'>
                    </td>
                </tr>
            </table>
        </span>
    </div>
</body>
</html>

<script type="text/javascript">
    var objParent = window.dialogArguments;
    var pInput = new Array();
    var bAcceptButtonPress = true;
    var esm_report_type;
    var stitle_BlkIncome = '<%=GetTranslationData("INCOME_E2B")%>';
    pInput[0] = "<%=GetUserPreferences("E2bImport_pinput0", lError, sError) %>";
    pInput[1] = "<%=GetUserPreferences("E2bImport_pinput1", lError, sError) %>";
    pInput[2] = "<%=GetUserPreferences("E2bImport_pinput2", lError, sError) %>";

    var initial_notes = "";
    var followup_notes = "", nullification_notes = "";

    async function fn_init() {
        if (pInput[0].length < 1) {
            f_Tabs_HideTab("CF_BULK_IMPORT", 1);
            dInitialReport.style.display = "none";
        }
        if (pInput[1].length < 1) {
            f_Tabs_HideTab("CF_BULK_IMPORT", 2);
            dFollowupReport.style.display = "none";
        }
        if (pInput[2].length < 1) {
            f_Tabs_HideTab("CF_BULK_IMPORT", 3);
            dNullificationReport.style.display = "none";
        }
        await fn_DisplayInitial();
        await fn_DisplayFollowup();
        await fn_DisplayNullification();
        if (pInput[0].length <= 0 && pInput[1].length <= 0 && pInput[2].length <= 0) {
            window.close();
            return;
        }

        dLoading.style.display = "block";
        fn_import_initial();
        fn_import_followup();
        fn_import_nullification();

        if (pInput[0].length > 0)
            fn_view_display(1);
        else if (pInput[1].length > 0)
            fn_view_display(2);
        else if (pInput[2].length > 0)
            fn_view_display(3);
    }

    async function fn_LockedReport() {
        var xmlDoc = this.req.responseXML;
        var sErrStr;
        var sLocked_User;
        var sesm_report_ids = "";

        sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
        if (sErrStr.length > 0) {
            await MessageBoxRes("GENERAL_ERROR", stitle_BlkIncome, sErrStr);
            return;
        }

        var asLocked_User = xmlDoc.getElementsByTagName("LOCKED_USER");
        var asEsmReportIDs = xmlDoc.getElementsByTagName("MODIFY_REPORT_IDS");
        if (asLocked_User) {
            sLocked_User = GetTextContentFromXML(asLocked_User[0]);
            if (asEsmReportIDs)
                sesm_report_ids = GetTextContentFromXML(asEsmReportIDs[0]);
            if (sLocked_User.length > 0)
                await MessageBoxRes("GENERAL_INFORMATION", stitle_BlkIncome, sLocked_User);
            if (bAcceptButtonPress)
                await fn_AcceptE2BMultiCase(sesm_report_ids);
            else {
                fn_RejectSingleE2BCase();
            }
        }
        else {
            if (bAcceptButtonPress)
                fn_AcceptE2BSingleCase(sesm_report_ids);
            else {
                fn_RejectSingleE2BCase();
            }
        }
        return;
    }

    async function fn_UnLockedReport() {
        var xmlDoc = this.req.responseXML;
        var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
        if (sErrStr.length > 0) {
            await MessageBoxRes("GENERAL_ERROR", stitle_BlkIncome, sErrStr);
            return;
        }
        return;
    }
    async function fn_DisplayInitial() {
        esm_report_type = 1;
        if (pInput[0].length > 0) {
            await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "bulk=1&esm_report_id=" + pInput[0]);
        }
    }
    async function fn_DisplayFollowup() {
        esm_report_type = 3;
        if (pInput[1].length > 0) {
            await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "bulk=1&esm_report_id=" + pInput[1]);
        }
    }
    async function fn_DisplayNullification() {
        esm_report_type = 4;
        if (pInput[2].length > 0) {
            await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "bulk=1&esm_report_id=" + pInput[2]);
        }
    }

    async function fn_AcceptE2BMultiCase(esm_report_ids) {
        var strURL = "";

        if (esm_report_type == 1) {
            pInput[0] = esm_report_ids;
            if (esm_report_ids.length > 0)
                strURL = "/E2B/Actions/E2B_AcceptE2BCase.asp?Bulk=1";
        }
        else if (esm_report_type == 3 || esm_report_type == 5 || esm_report_type == 6) {
            pInput[1] = esm_report_ids;
            if (esm_report_ids.length > 0)
                strURL = "/E2B/Actions/E2B_AcceptFollowupE2BCase.asp?Bulk=1";
        }
        else if (esm_report_type == 4) {
            pInput[2] = esm_report_ids;
            if (esm_report_ids.length > 0)
                strURL = "/E2B/Actions/E2B_AcceptNullificationE2BCase.asp?Bulk=1&esm_report_id=" + esm_report_ids;
        }
        if (strURL.length > 0) {
            var sDialogStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };
            if (esm_report_type == 4)
                sDialogStyle = { dialogHeight: "640", dialogWidth: "480", resizable: false, scrollable: false };
            notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);

            if ((!notes) && (notes !== sSessionTimeOutDialogReturn)) {
                if (esm_report_type == 1)
                    pInput[0] = "";
                else if (esm_report_type == 3 || esm_report_type == 5 || esm_report_type == 6)
                    pInput[1] = "";
                else if (esm_report_type == 4)
                    pInput[2] = "";
                await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + esm_report_ids);
            }
            else {
                if (notes !== sSessionTimeOutDialogReturn) {
                    if (esm_report_type == 1) {
                        fn_getElementByName('Initial').value = esm_report_ids;
                        fn_getElementByName('Initial_notes').value = notes;
                    }
                    else if (esm_report_type == 3 || esm_report_type == 5 || esm_report_type == 6) {
                        fn_getElementByName('Followup').value = esm_report_ids;
                        fn_getElementByName('Followup_notes').value = notes;
                    }
                    else if (esm_report_type == 4) {
                        fn_getElementByName('Nullification').value = esm_report_ids;
                        fn_getElementByName('Nullification_notes').value = notes;
                    }
                }
            }
        }
    }

    function fn_import_initial() {
        if (fn_getElementByName('Initial_notes').value) {
            fn_getElementByName('E2b_Type').value = 1;
            document.Frm_BulkImport.action = "/E2B/Incoming/BulkAcceptE2b.asp?<%=GetRequestKeyValue()%>";
            document.Frm_BulkImport.target = "initial_report";
            fn_ValidateSubmitForm(document.Frm_BulkImport);
            dInitialReport.style.display = "block";
            fn_view_display(1);
        }
    }

    function fn_import_followup() {
        if (fn_getElementByName('Followup_notes').value) {
            fn_getElementByName('E2b_Type').value = 3;
            document.Frm_BulkImport.action = "/E2B/Incoming/BulkAcceptE2b.asp?<%=GetRequestKeyValue()%>";
            document.Frm_BulkImport.target = "followup_report";
            fn_ValidateSubmitForm(document.Frm_BulkImport);
            dFollowupReport.style.display = "block";
            fn_view_display(2);
        }
    }

    function fn_import_nullification() {
        if (fn_getElementByName('Nullification_notes').value) {
            fn_getElementByName('E2b_Type').value = 4;
            document.Frm_BulkImport.action = "/E2B/Incoming/BulkAcceptE2b.asp?<%=GetRequestKeyValue()%>";
            document.Frm_BulkImport.target = "nullification_report";
            fn_ValidateSubmitForm(document.Frm_BulkImport);
            dNullificationReport.style.display = "block";
            fn_view_display(3);
        }
    }

    function fn_print() {
        if (fn_getElementByName('Initial_notes').value) {
           document.getElementById("initial_report").contentWindow.print();
        }
        if (fn_getElementByName('Followup_notes').value) {
             document.getElementById("followup_report").contentWindow.print();
        }
        if (fn_getElementByName('Nullification_notes').value) {
            document.getElementById("nullification_report").contentWindow.print();
        }
    }

    var was_at = 0, onsection = 0;
    function fn_view_display(id) {
        if (id != onsection) {
            was_at = onsection;
            set_off_force(was_at);
            f_Tabs_SetTab("CF_BULK_IMPORT", id);
            onsection = id;
        }
        dInitialReport.style.display = (id == 1) ? "block" : "none";
        dFollowupReport.style.display = (id == 2) ? "block" : "none";
        dNullificationReport.style.display = (id == 3) ? "block" : "none";
    }

    function set_off_force(num) {
        if (num > 0)
            f_Tabs_ShowTab("CF_BULK_IMPORT", num);
    }
</script>

<!--#INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
