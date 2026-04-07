<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : GetLoadE2bWarning.asp
' Description  : This page is called form worklist/worklistbulktransmit.asp 
'******************************************************************************
' Revision History
' Date		Author		Description
' 15jun2006 Prakash     Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title>E2B Warnings</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
Dim oMessage, oOutMsg, lUserID, lErrNo, check_warning, oRecList, lNumRec, oRec, oTable, l_case_num, notes
Dim l_case_id, l_user_id, l_esm_report_id, l_esm_seq_num, l_esm_transmit, l_view, filePathName, fileName
Dim AcceptedCase, OriginalCase, Sender, MessageNumber, l_load_e2b_report, l_status, lWarning, DisplayAcceptedCase
Dim sError, s_pgError, load_case_num
l_esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
check_warning = GetLong(GetRequest("check_warning"), 1)
l_case_num = GetRequest("case_num")
notes = GetRequest("notes")
lUserID = GetRequest("user_id")

DisplayAcceptedCase = "N/A"

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
	Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", l_esm_report_id)
	Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_CHECK_WARNINGS", check_warning)
	Set oOutMsg = ServiceRequest3(oArgusSvr, oMessage, lErrNo, sError,"Argus.CaseFormApp")
	s_pgError = ""
	If (lErrNo <> 0) Then
		s_pgError = "An Error Occurred. Error No. " & lErrNo & " : " & GetXMLValueDirect (oOutMsg, "GN_ERROR_STRING")
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
%>

<script type="text/javascript">
    var objParent = window.dialogArguments;
    objParent.hideLoading();
    async function fn_PrintList() {
        var strURL;
        strURL = "/E2B/Misc/GetLoadE2bWarningPrint.asp?view_warning=0&esm_report_id=" + document.all.E2bReportID.value + "&case_id=" + document.all.case_id.value + "&esm_seq_num=" + document.all.esm_seq_num.value + "&esm_transmit=" + document.all.esm_transmit.value + "&user_id=" + document.all.user_id.value;
        var sDialogStyle = { dialogHeight: "480", dialogWidth: "720", resizable: false, scrollable: false };
        await fn_OpenModalDialog(strURL, window,sDialogStyle );
    }
</script>

<%Call BuildHiddenControlDirect("status", l_status) %>
<%if len(s_pgError) > 0 then%>
<body>
    <form name="frm_check_warning">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <table style="width: 100%; height: 100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td colspan="3" align="center" class="label">
                    <div id="error">
                        <%=Fn_Sanitize(s_pgError) %>
                    </div>
                </td>
            </tr>
            <tr>
                <td width="47%" align="right">
                    <form name="form1">
                    </form>
                </td>
                <td width="4%" align="center">&nbsp;
                </td>
                <td width="49%" align="left">
                    <form name="form2">
                    </form>
                </td>
            </tr>
            <tr>
                <td width="47%" align="right">&nbsp;
                </td>
                <td width="4%" align="center">
                    <%BuildButtonDirect("btn_OK", "OK", 3).Style("width:50px").OnClick("window.close();").Render() %>
                </td>
                <td width="49%" align="left">&nbsp;
                </td>
            </tr>
        </table>
    </form>
</body>
<%elseif lWarning = 0 then%>

<script type="text/javascript">
	setWindowReturnValue(<%=lWarning%> + "~" + document.all.status.value + "~" + "<%=AcceptedCase%>");
	window.close();
</script>

<%elseif lWarning = 1 then%>
<body>
    <form name="frm_check_warning" method="post" action="/E2B/Misc/GetLoadE2bWarning.asp">
        <%Call BuildHiddenControlDirect("E2bReportID", l_esm_report_id) %>
        <%Call BuildHiddenControlDirect("case_id", l_case_id) %>
        <%Call BuildHiddenControlDirect("esm_seq_num", l_esm_seq_num) %>
        <%Call BuildHiddenControlDirect("esm_transmit", l_esm_transmit) %>
        <%Call BuildHiddenControlDirect("user_id", l_user_id) %>
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
                                            <%BuildLabelDirect("Accepted Case Number").Render()%>
                                        </td>
                                        <td>
                                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "AcceptedCase", DisplayAcceptedCase, true, 1, "").Style("width:100%").Render() %>
                                        </td>
                                        <td style="float: left;">
                                            <%BuildLabelDirect("Original Case Number").Render()%>
                                        </td>
                                        <td>
                                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "OriginalCase", OriginalCase, true, 1, "").Style("width:100%").Render() %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="float: left;">
                                            <%BuildLabelDirect("Sender Name").Render()%>
                                        </td>
                                        <td>
                                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "Sender", Sender, true, 1, "").Style("width:100%").Render() %>
                                        </td>
                                        <td style="float: left;">
                                            <%BuildLabelDirect("Message Number").Render() %>
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
            <tr height="50px">
                <td>
                    <div id="Divheader" class="table-scroll table-scroll-hide" style="width: 100%; border-bottom-width: 0px;">
                        <table class="table" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <col width="10%" />
                            <col width="30%" />
                            <col width="30%" />
                            <col width="30%" />
                            <tr class="tblheader-lightblue">
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLabelDirect("S.No.").Render() %>
                                </td>
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLabelDirect("E2B Element").Render() %>
                                </td>
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLabelDirect("E2B Value").Render() %>
                                </td>
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLabelDirect("Value Selected").Render() %>
                                </td>
                            </tr>
                            <tr class="tblheader-lightblue">
                                <td class="alc-header" colspan="4" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <%BuildLabelDirect("Error / Warning Message").Render() %>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td width="100%">
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
                    <%BuildButtonDirect("btnPrint", "Print", 5).Style("width:60px")_
                                .OnClick("fn_PrintList();").Render()%>
                &nbsp;
                <%BuildButtonDirect("btnClose", "Close", 6).Style("width:60px")_
                                .OnClick("window.close();").Render() 
                %>
                </td>
            </tr>
        </table>
    </form>
</body>
<%End If%>

<script type="text/javascript">
	setWindowReturnValue(<%=lWarning%> + "~" + document.all.status.value + "~" + "<%=AcceptedCase%>");
</script>

</html>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
