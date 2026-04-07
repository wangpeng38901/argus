<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<html>
<head>
    <title>Report Generation</title>
</head>
<%
	Dim oMessage, bSuccess, oOutMsg, sDocId, lUserID
	Dim e2b_lError, e2b_sError, e2b_sReturn
	Dim l_esm_report_id, l_case_id, lGmtOffSet

	l_esm_report_id = GetLong(GetRequest("esm_report_id"), -1)
	l_case_id = GetLong(GetRequest("case_id"), 0)
	lUserID = oArgusUser.GetUserId()
	lGmtOffSet = GetXMLValueDirect (oSession, "GMT")

	e2b_sError = ""
	e2b_lError = 0
   
	Call CreateMessage (oMessage, 400300070)
	Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", l_esm_report_id)
	Call SetXMLValueDirect (oMessage, "CSM_CID", l_case_id)
    Call SetXMLValueDirect (oMessage, "RPT_E2B_PREVIOUS_FOLLOWUP", GetLong(GetRequest("prev_id"), -1))
    Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
    Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
    Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)
		
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, e2b_lError, e2b_sError)	
	If e2b_lError = 0 Then
        sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
	End If
%>
<body onload="fn_Load()">
</body>
</html>

<script type="text/javascript">
async function fn_Load()
{
    <%If e2b_lError <> 0 Then%>
    await MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("ERROR_OCCURED")%>', <%=JavaScriptClean(e2b_sError) %>);
    <%Else %>
    setWindowReturnValue("<%=sDocId %>");
    <%End If %>
    window.close();
}
</script>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
