<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : GetLoadE2bStagingWarning.asp
' Description  : This page is called form E2bIncomingReports.asp 
'******************************************************************************
' Revision History
' Date		Author		Description
' 15jun2006 UMAR     Original
' 11MAR2008 UMAR     Passing Justification Id to save in case followup
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title><%=GetTranslationData("E2B_STAGING_WARNING") %></title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <style type="text/css">
        TABLE.inner-table {
            padding: unset !important;
        }
    </style>
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
Dim oMessage, oOutMsg, lUserID, lErrNo, check_warning, oRecList, lNumRec, oRec, oTable, l_case_num, notes
Dim l_case_id, l_user_id, l_esm_report_id, l_esm_seq_num, l_esm_transmit, l_view, filePathName, fileName
Dim AcceptedCase, OriginalCase, Sender, MessageNumber, l_load_e2b_report, l_status, lWarning, DisplayAcceptedCase
Dim sError, s_pgError, custom_import, lcase_id, lError, load_case_num, notes_id
Dim sSQL
Dim lSaveError, sSaveError
l_esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
check_warning = GetLong(GetRequest("check_warning"), 1)
l_case_num = GetRequest("case_num")
notes = GetRequest("notes")
notes_id = GetRequest("notes_id")
lUserID = GetRequest("user_id")
custom_import = GetLong(GetRequest("custom_import"), 1)
lcase_id = GetLong(GetRequest("case_id"), -1)

If custom_import = 0 Then
    'Validate Auth Token
    Call ValidateAuthToken()
End If

DisplayAcceptedCase = "N/A"
'if it is an intial report and case id has been sent then it should be accepted as a follow up
'otherwise it should always be treated as initial
if lcase_id > 0 then
    Call SetParameter("P_REPORT_ID", l_esm_report_id, PARAM_NUMBER)
     sSQL = "UPDATE SAFETYREPORT set e2b_type_accept_as = 3 WHERE e2b_type = 1 and report_id = :P_REPORT_ID"
    Call UpdateSQL(sSQL,lError,sError)
else 
    Call SetParameter("P_REPORT_ID", l_esm_report_id, PARAM_NUMBER)
    sSQL = "UPDATE SAFETYREPORT set e2b_type_accept_as = NULL WHERE e2b_type = 1 and report_id = :P_REPORT_ID"
    Call UpdateSQL(sSQL,lError,sError)
end if

if IsNull(lUserID) or len(lUserID) <= 0 then
	lUserID = oArgusUser.GetUserId()
end if
if IsNull(l_case_num) or len(l_case_num) <= 0 then
	l_case_num = ""
end if
if IsNull(notes) or len(notes) <= 0 then
	notes = ""
end if
		
'Create Message
Call CreateMessage (oMessage, 300100261) 'MID_db_app_load_e2b_report
Call SetXMLValueDirect (oMessage, "CSM_CASE_NUM", l_case_num)
Call SetXMLValueDirect (oMessage, "GN_UI_JUSTIFICATION", notes)
Call SetXMLValueDirect (oMessage, "LM_JUSTIFICATIONS_JUSTIFICATION_ID", notes_id)
Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", l_esm_report_id)
Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
Call SetXMLValueDirect (oMessage, "RPT_E2B_CHECK_WARNINGS", check_warning)
Call SetXMLValueDirect (oMessage, "RPT_E2B_VIEW_STAGING_WARNING", custom_import)
Call SetXMLValueDirect (oMessage, "CSM_CASE_ID", lcase_id)
Set oOutMsg = ServiceRequest3(oArgusSvr, oMessage, lErrNo, sError,"Argus.CaseFormApp")
s_pgError = ""
If (lErrNo <> 0) Then
	s_pgError = "Error No. " & lErrNo & " : " & GetXMLValueDirect (oOutMsg, "GN_ERROR_STRING")
else
	Set oRecList = oOutMsg.selectnodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")
	lNumRec = oRecList.length
	For Each oRec in oRecList
		l_case_id = GetXMLValueDirect(oRec, "CSM_CASE_ID")
		l_user_id = GetXMLValueDirect(oRec, "CFG_USERS_USER_ID")
		l_esm_report_id = GetXMLValueDirect(oRec, "RPT_E2B_REPORT_ID")
		l_esm_seq_num = GetXMLValueDirect(oRec, "RPT_E2B_SEQ_NUM")
		l_esm_transmit = GetXMLValueDirect(oRec, "RPT_E2B_TRANSMIT_WARNING")
		l_load_e2b_report = GetXMLValueDirect(oRec, "RPT_E2B_LOAD_E2B_REPORT")
		load_case_num = GetXMLValueDirect(oRec, "CSM_CASE_NUM")
		if CCur(l_load_e2b_report) > 0 then
			l_status = "101"
		else
			l_status = "102"
		end if
	Next
end if

if len(s_pgError) = 0 then
	Call CreateMessage (oMessage, 403000003) 'MID_rpt_app_prt_E2B_warnings
	Set oTable = AddXMLNode (oMessage, "TABLE_RPT_E2B")
	Set oRec = AddXMLNode (oTable, "RPT_E2B")
	Call SetXMLValueDirect (oRec, "CSM_CASE_ID",l_case_id)
	Call SetXMLValueDirect (oRec, "CFG_USERS_USER_ID",l_user_id)
	Call SetXMLValueDirect (oRec, "RPT_E2B_REPORT_ID",l_esm_report_id)
	Call SetXMLValueDirect (oRec, "RPT_E2B_SEQ_NUM",l_esm_seq_num)
	Call SetXMLValueDirect (oRec, "RPT_E2B_TRANSMIT_WARNING",l_esm_transmit)
	l_view = 1
	Call SetXMLValueDirect (oRec, "RPT_E2B_VIEW_TYPE",l_view)
	Call SetXMLValueDirect (oRec, "CSM_CASE_NUM",load_case_num)
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lErrNo, sError)
	s_pgError = ""
	lWarning = 1
	If (lErrNo <> 0) Then
		s_pgError = "An Error Occurred. Error No. " & lErrNo & " : " & GetXMLValueDirect (oOutMsg, "GN_ERROR_STRING")
	else
		lWarning = GetXMLValueDirect (oOutMsg, "RPT_E2B_STATUS_ID")
		Set oRecList = oOutMsg.selectnodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B_E2B")
		lNumRec = oRecList.length
		for each oRec in oRecList
			AcceptedCase = GetXMLValueDirect(oRec, "CSM_CASE_NUM")
			if len(AcceptedCase) = 0 then
				DisplayAcceptedCase = "N/A"
			else
				DisplayAcceptedCase = AcceptedCase
			end if
			OriginalCase = GetXMLValueDirect(oRec, "RPT_E2B_COMPANYNUMB")
			Sender = GetXMLValueDirect(oRec, "RPT_E2B_AGENCY_NAME")
			MessageNumber = GetXMLValueDirect(oRec, "RPT_E2B_MESSAGE_ID")
		next
	end if
end if

if check_warning = "1" then
  DisplayAcceptedCase = "N/A"
end if	

'check for error EID_DBAPP_ESM_E2B_LOAD_FAIL
if CCur(l_load_e2b_report) =  305100038 then 'if case is not created then 
    AcceptedCase ="" ' it will stop from showing case successful update message from E2BIncomingReport.asp
ElseIf (l_status = 102 and Len(s_pgError) = 0 and check_warning = 0) Then
	SaveFollowupDiffData()
end if
%>

<script type="text/javascript">
    var objParent = window.dialogArguments;
    objParent.hideLoading();
    async function fn_PrintList() {
        var strURL;
        var strStyle = { dialogHeight: "480", dialogWidth: "720", resizable: false, scrollable: false };
        strURL = "/E2B/Misc/GetLoadE2bWarningPrint.asp?view_warning=0&esm_report_id=" + document.all.E2bReportID.value + "&case_id=" + document.all.case_id.value + "&esm_seq_num=" + document.all.esm_seq_num.value + "&esm_transmit=" + document.all.esm_transmit.value + "&user_id=" + document.all.user_id.value;
        await fn_OpenModalDialog(strURL, window, strStyle);
    }
</script>

<%Call BuildHiddenControlDirect("status", l_status) %>
<%if len(s_pgError) > 0 then%>
<body>
    <form name="frm_check_warning">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <table style="width: 100%; height: 100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td colspan="3" align="center" class="Label">
                    <div id="error">
                        <%=Fn_Sanitize(s_pgError) %>
                    </div>
                </td>
            </tr>
            <tr>
                <td width="47%" align="right">&nbsp;
                </td>
                <td width="4%" align="center">&nbsp;
                </td>
                <td width="49%" align="left">&nbsp;
                </td>
            </tr>
            <tr>
                <td width="47%" align="right">&nbsp;
                </td>
                <td width="4%" align="center">
                    <%BuildButtonDirect("btn_OK", "OK", 3).Style("width:50px").OnClick("setWindowReturnValue('');window.close(); ").Render() %>
                </td>
                <td width="49%" align="left">&nbsp;
                </td>
            </tr>
        </table>
    </form>
</body>
<%elseif lWarning = 0 then%>

<script type="text/javascript">
	<% If Len(sSaveError) > 0 Then %>
	setWindowReturnValue(<%=lWarning%> + "~" + document.all.status.value + "~" + "<%=AcceptedCase%>" + "~" + <%=JavaScriptSanitize(sSaveError) %>);
	<% else %>
	setWindowReturnValue(<%=lWarning %> + "~" + document.all.status.value + "~" + "<%=AcceptedCase%>");
	<% End If %>
        fn_CheckErrors();
</script>

<%elseif lWarning = 1 then%>
<body>
    <form name="frm_check_warning" method="post" action="/E2B/Misc/GetLoadE2bWarning.asp" class="margin_0_override">
        <%Call BuildHiddenControlDirect("E2bReportID", l_esm_report_id) %>
        <%Call BuildHiddenControlDirect("case_id", l_case_id) %>
        <%Call BuildHiddenControlDirect("esm_seq_num", l_esm_seq_num) %>
        <%Call BuildHiddenControlDirect("esm_transmit", l_esm_transmit) %>
        <%Call BuildHiddenControlDirect("user_id", l_user_id) %>
        <!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <table style="width: 100%; height: 100%" class="table border-blue inner-table">
            <tr style="height: 70px">
                <td>
                    <table class="inner-table" style="width: 100%;">
                        <tr>
                            <td valign="top" width="100%">
                                <table class="table border-blue inner-table" width="100%" border="0" cellspacing="2"
                                    cellpadding="2">
                                    <col width="19%" />
                                    <col width="31%" />
                                    <col width="17%" />
                                    <col width="33%" />
                                    <tr>
                                        <td class="Padding-All" align="left">
                                            <%BuildLocalLabel("ACCPTD_CASE_NUM").Render()%>
                                        </td>
                                        <td>
                                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "AcceptedCase", DisplayAcceptedCase, true, 1, "").Style("width:100%").Render() %>
                                        </td>
                                        <td style="float: left;">
                                            <%BuildLocalLabel("ORG_CASE_NUM").Render()%>
                                        </td>
                                        <td>
                                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "OriginalCase", OriginalCase, true, 1, "").Style("width:100%").Render() %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="float: left;">
                                            <%BuildLocalLabel("WL_BP_SENDER_NAME").Render()%>
                                        </td>
                                        <td>
                                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "Sender", Sender, true, 1, "").Style("width:100%").Render() %>
                                        </td>
                                        <td style="float: left;">
                                            <%BuildLocalLabel("E2B_MESSAGE_NUMBER").Render() %>
                                        </td>
                                        <td>
                                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "MessageNumber", MessageNumber, true, 1, "").Style("width:100%").Render() %>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr height="48px">
                <td style="padding-left: 0">
                    <div id="Divheader" class="table-scroll scroll_hide" style="width: 100%; overflow-y: scroll; height: 100%; border-bottom-width: 0px; background: #b3c7e2;">
                        <table class="table" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <col width="10%" />
                            <col width="30%" />
                            <col width="30%" />
                            <col width="30%" />
                            <tr class="tblheader-lightblue">
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLocalLabel("S_NUM").Render() %>
                                </td>
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLocalLabel("E2B_ELEM").Render() %>
                                </td>
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLocalLabel("E2B_VALUE").Render() %>
                                </td>
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLocalLabel("VALUE_SELECTED").Render() %>
                                </td>
                            </tr>
                            <tr class="tblheader-lightblue">
                                <td class="alc-header" colspan="4" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLocalLabel("ERR_WARNING_MSG").Render() %>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td width="100%" style="padding-left: 0">
                    <div id="DivResult" class="table-scroll" style="width: 100%; height: 100%; overflow-y: scroll">
                        <table class="table" border="0" cellspacing="0" cellpadding="0" border="1" style="width: 100%;">
                            <col width="10%" />
                            <col width="30%" />
                            <col width="30%" />
                            <col width="30%" />
                            <%	for each oRec in oRecList  %>
                            <tr valign="top">
                                <td class="alc-header" style="padding-left: 5px;" width="10%">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(GetXMLValueDirect(oRec,"RPT_E2B_ERROR_CODE")) %>"
                                        size="5" id="text1" readonly name="text1">
                                </td>
                                <td class="alc-header" style="padding-left: 5px;" width="30%">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(GetXMLValueDirect(oRec,"RPT_E2B_WARNING1")) %>"
                                        size="35" id="text2" readonly name="text2">
                                </td>
                                <td class="alc-header" style="padding-left: 5px;" width="30%">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(GetXMLValueDirect(oRec,"RPT_E2B_WARNING2")) %>"
                                        size="35" id="text4" readonly name="text4">
                                </td>
                                <td class="alc-header" style="padding-left: 5px;" width="30%">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(GetXMLValueDirect(oRec,"RPT_E2B_WARNING3")) %>"
                                        size="30" id="text5" readonly name="text5">
                                </td>
                            </tr>
                            <tr valign="top">
                                <td colspan="4">
                                    <%BuildControlDirect(CTL_TEXTAREA, "text6", GetXMLValueDirect(oRec,"RPT_E2B_WARNING4"), true, 5, "").Style("width:100%;").Render() %>
                            </tr>
                            <% Next %>
                        </table>
                    </div>
                </td>
            </tr>
            <tr class="tblheader-gray" height="25px">
                <td align="center" valign="middle">
                    <%BuildButton("btnPrint", "PRINT", 5).Style("width:60px")_
                                .OnClick("fn_PrintList();").Render()%>
                &nbsp;
                <%BuildButton("btnClose", "BTN_CLOSE", 6).Style("width:60px")_
                                .OnClick("fn_CheckErrors();").Render() 
                %>
                </td>
            </tr>
        </table>
    </form>
</body>
<%End If%>

<script type="text/javascript">
async function fn_CheckErrors()
{
    await loadArgusMessage( "/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + document.all.E2bReportID.value);
}

function fn_UnLockedReport()
{
    var xmlDoc = this.req.responseXML;
    var sErrStr;
    var sLocked_User;
    sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
    if (sErrStr.length > 0)
    {
      MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("INCOME_E2B")%>',sErrStr); 
    }
	<% If Len(sSaveError) > 0 Then %>
	setWindowReturnValue( <%=lWarning %> + "~" + document.all.status.value + "~" + "<%=AcceptedCase%>" + "~" + <%=JavaScriptSanitize(sSaveError) %>);
	<% else %>
	setWindowReturnValue( <%=lWarning %> + "~" + document.all.status.value + "~" + "<%=AcceptedCase%>");
	<% End If %>
        window.close();
}

</script>

</html>
<script runat="SERVER" language="VBSCRIPT">
	Sub SaveFollowupDiffData()
		Dim lE2bType, esm_initial_report_id
		Dim lASP, lGmtOffSet, sDocId, sAccept_date, lIsJReport

		lE2bType = GetLong(GetRequest("E2bType"), 0)

		if (lE2bType = 1) or (lE2bType = 3) then	
			lGmtOffSet = GetXMLValueDirect (oSession, "GMT")
			sDocId = GetString(GetRequest("DocId"), "")
			lcase_id = GetLong(GetRequest("case_id"), -1)		
			esm_initial_report_id = GetLong(GetRequest("esm_initial_report_id"), 0)
			sAccept_date = GetRequest("accept_date")    
			lIsJReport = GetRequest("IsJReport")

			if IsNullOrEmpty(sAccept_date) then
				sAccept_date = "-99"
			end if

			lSaveError = 0
			sSaveError = ""
			lASP = 1

			if (lE2bType = 1) then
				Call SetParameter("P_REPORT_ID", l_esm_report_id, PARAM_NUMBER)
				sSQL = "select case_xref from safetyreport where report_id = :P_REPORT_ID"
				lcase_id = ExecuteSQLReturnStr (sSQL, lSaveError, sSaveError)
			end if

			if (lcase_id > 0) then
				Call SetParameter("P_CASE_ID", lcase_id, PARAM_NUMBER)
				sSQL = "select case_num from case_master where case_id = :P_CASE_ID"
				l_case_num = ExecuteSQLReturnStr (sSQL, lSaveError, sSaveError)
			else
				l_case_num = ""
			end if
			'Create Message
			Call CreateMessage (oMessage, 300400064) 'MID_db_app_E2b_save_followup_report
			Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", l_esm_report_id)
			Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
			Call SetXMLValueDirect (oMessage, "GN_THIS_IS_ASP", lASP)
			Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
			Call SetXMLValueDirect (oMessage, "GN_REPORT_IDENTIFIER", sDocId)
			Call SetXMLValueDirect (oMessage, "CSM_CASE_ID", lcase_id)
			Call SetXMLValueDirect (oMessage, "CSM_CASE_NUM", l_case_num)
			Call SetXMLValueDirect (oMessage, "RPT_E2B_E2B_TYPE", lE2bType)
			Call SetXMLValueDirect (oMessage, "CSAC_DATE_DONE", sAccept_date)
			Call SetXMLValueDirect (oMessage, "GN_NUMBER2", lIsJReport)
			Call SetXMLValueDirect (oMessage, "GN_UI_JUSTIFICATION", notes)
	
			Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lSaveError, sSaveError)
			if lcase_id < 1 then
				sSaveError = sSaveError +  " " + GetTranslationData("E2BINCOME_FUDIFF_ERROR")
			end if
			if esm_initial_report_id > 0 then
				Call SetParameter("P_REPORT_ID", esm_initial_report_id, PARAM_NUMBER)
				Call UpdateSQL("update SAFETYREPORT set status = 8 where report_id = :P_REPORT_ID", lError, sError)
			end if
		End If
	End Sub
</script>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
