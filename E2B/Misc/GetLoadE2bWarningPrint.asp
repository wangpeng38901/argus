<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : GetLoadE2bWarningPrint.asp
' Description  : 
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
    <title>Get Load ICSR Warnings</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
Dim oMessage, oOutMsg, check_warning, oRecList, lNumRec, oRec, oTable
Dim l_case_id, l_user_id, l_esm_report_id, l_esm_seq_num, l_esm_transmit, l_view, sDocId
Dim AcceptedCase, OriginalCase, Sender, MessageNumber
Dim bSuccess, lErrNo, sError, s_pgError

bSuccess= true
l_esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
l_view = GetLong(GetRequest("view_warning"), 0)
l_case_id = GetLong(GetRequest("case_id"), 0)
l_esm_seq_num = GetLong(GetRequest("esm_seq_num"), 0)
l_esm_transmit = GetLong(GetRequest("esm_transmit"), 0)
l_user_id = GetLong(GetRequest("user_id"), 0)
bSuccess = True
if l_user_id <= 0 then
	l_user_id = oArgusUser.GetUserId()
end if
	
Call CreateMessage (oMessage, 403000003) 'MID_rpt_app_prt_E2B_warnings
Set oTable = AddXMLNode (oMessage, "TABLE_RPT_E2B")
Set oRec = AddXMLNode (oTable, "RPT_E2B")
Call SetXMLValueDirect (oRec, "CSM_CASE_ID",l_case_id)
Call SetXMLValueDirect (oRec, "CFG_USERS_USER_ID",l_user_id)
Call SetXMLValueDirect (oRec, "RPT_E2B_REPORT_ID",l_esm_report_id)
Call SetXMLValueDirect (oRec, "RPT_E2B_SEQ_NUM",l_esm_seq_num)
Call SetXMLValueDirect (oRec, "RPT_E2B_TRANSMIT_WARNING",l_esm_transmit)
Call SetXMLValueDirect (oRec, "RPT_E2B_VIEW_TYPE",l_view)
Call SetXMLValueDirect (oRec, "GN_GEN_SAVE_REPORT",1)
Call SetXMLValueDirect (oRec, "GN_DISPLAY_LANGUAGE", glDisplayLang)
Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lErrNo,sError)
s_pgError = ""
If (lErrNo <> 0) Then
	s_pgError = "An Error Occurred. Error No. " & lErrNo & " : " & sError
	bSuccess = false
Else
	sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
End If
If (bSuccess) Then
%>
<script type="text/javascript">
	fn_ViewDocument("<%=sDocId %>", "");
	window.close();
</script>

<% Elseif len(s_pgError) > 0 Then %>
<body bgcolor="white" leftmargin="5" class="PgBody">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
                </form>
            </td>
            <td width="4%" align="center">&nbsp;
            </td>
            <td width="49%" align="left">
                <form name="form2">
                	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->                	
                </form>
            </td>
        </tr>
        <tr>
            <td width="47%" align="right">&nbsp;
            </td>
            <td width="4%" align="center">
                <%BuildButtonDirect("btnPrint", "OK", 15).Style("width:50px")_
				.OnClick("window.close();").Render()%>
            </td>
            <td width="49%" align="left">&nbsp;
            </td>
        </tr>
    </table>
</body>
<% Else %>
<body bgcolor="white" leftmargin="5" class="PgBody">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td colspan="3" align="center" class="label">No data was found for this report
            </td>
        </tr>
        <tr>
            <td width="47%" align="right">
                <form name="form1">
                	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
                </form>
            </td>
            <td width="4%" align="center">&nbsp;
            </td>
            <td width="49%" align="left">
                <form name="form2">
                	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
                </form>
            </td>
        </tr>
        <tr class="tblheader-gray" height="25px">
            <td align="CENTER" valign="middle" colspan="3">
                <%BuildButtonDirect("btnPrint", "OK", 15).Style("width:50px")_
				.OnClick("window.close();").Render()%>
            </td>
        </tr>
    </table>
</body>
<% End If %>
</html>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
