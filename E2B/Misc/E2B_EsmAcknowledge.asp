<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2B_ESMAcknowledge.asp
' Description  : Display the ESM Acknowledgement Information
'******************************************************************************
' Revision History
' Date		    Author		Description
' 02JUN2006 	Umar       	Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title><%=GetTranslationData("ACK_INFO") %></title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<!-- Declaration of Page Scope variables Ends -->
<!-- Assign Values to Page Scope variables Starts -->
<!-- Assign Values to Page Scope variables Ends -->
<!-- Page Processing Starts -->
<%
	Dim sMsgTrackNum, sRptTrackNum, sDateInit, sICSRNum, sCaseNum, sesm_report_id, sreport_ack_code
	sMsgTrackNum = Request.QueryString("MsgTrackNum")
	sRptTrackNum = Request.QueryString("RptTrackNum")
	sDateInit = Request.QueryString("DateInit")
	sICSRNum = Request.QueryString("ICSRNum")
	sCaseNum = Request.QueryString("CaseNum")
	sesm_report_id = Request.QueryString("esm_report_id")
	sreport_ack_code = Request.QueryString("report_ack_code")
	if len(sreport_ack_code) <= 0 or IsNull(sreport_ack_code) Then
		sreport_ack_code = ""
	end if
	sDateInit = FormatDateTime(DateAdd("h",-8,sDateInit))
%>
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body onload="fn_init();">
    <form name="EsmAck" method="post" action="">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <%Call BuildHiddenControlDirect("parsingerror", "") %>
        <%Call BuildHiddenControlDirect("errormessage", "") %>
        <table width="100%" height="100%" class="table border" cellspacing="0" cellpadding="0">
            <tr class="tblheader-lightblue" height="25px">
                <td class="section-header-middle">
                    <%BuildLocalLabel("ACK_INFO").SetStyleSheet("label label-section").Render() %>
                </td>
            </tr>
            <tr height="100%">
                <td class="padding-all">
                    <table width="100%" height="100%" class="table" cellspacing="0" cellpadding="0" border="0">
                        <col width="40%" />
                        <tr height="25px">
                            <td>
                                <%BuildLocalLabel("ACK_MSG_TRACKING_NUM").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "textfield", sMsgTrackNum, true, 1, "").Style("width:100%").Render() %>
                            </td>
                        </tr>
                        <tr height="25px">
                            <td>
                                <%BuildLocalLabel("ACK_RPT_TRACKING_NUM").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "textfield2", sRptTrackNum, true, 2, "").Style("width:100%").Render() %>
                            </td>
                        </tr>
                        <tr height="25px">
                            <td>
                                <%BuildLocalLabel("ACK_DATE_INIT").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "textfield3", sDateInit, true, 3, "").Style("width:100%").Render() %>
                            </td>
                        </tr>
                        <tr height="25px">
                            <td>
                                <%BuildLocalLabel("ICSR_MSG_NUM").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "textfield4", sICSRNum, true, 4, "").Style("width:100%").Render() %>
                            </td>
                        </tr>
                        <tr height="25px">
                            <td colspan="2" valign="bottom">
                                <%BuildLocalLabel("ERROR_RPTED_RCV").Render() %>
                            </td>
                        </tr>
                        <tr height="100%">
                            <td colspan="2" width="100%">
                                <%BuildControlDirect(CTL_TEXTAREA, "ackError", "", true, 5, "").Style("width:100%;height:99%").SetMaxLength(32767).Render() %>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr class="tblheader-gray" height="25px">
                <td align="center" valign="middle">
                    <% BuildButton("btn_OK", "OK", 6).Style("width:50px").OnClick("fn_OK();").Render() %>
                    <% BuildButton("btn_Print", "PRINT", 7).Style("width:60px").OnClick("fn_Print();").Render() %>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
<!-- Page Display Ends -->
<!-- Javascript Functions Starts -->

<script type="text/javascript">
var oParentWindow = window.dialogArguments;
var sError, report_ack_code;
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2B_ESMAcknowledge.asp
' Description  : Close the Dialogue
'******************************************************************************
' Revision History
' Date		    Author		Description
' 02JUN2006 	Umar       	Original
'******************************************************************************
%>
function fn_OK()
{
	window.close();
}
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2B_ESMAcknowledge.asp
' Description  : Print the ACK in PDF
'******************************************************************************
' Revision History
' Date		    Author		Description
' 02JUN2006 	Umar       	Original
'******************************************************************************
%>
async function fn_Print()
{
	var strURL, sMessage, sReport, sDate, sICSR, sCaseNum, sesm_report_id, sParam;
	
	sMessage = <%=JavaScriptSanitize(sMsgTrackNum) %>;
	sReport = <%=JavaScriptSanitize(sRptTrackNum) %>;
	sDate = <%=JavaScriptSanitize(sDateInit) %>;
	sICSR = <%=JavaScriptSanitize(sICSRNum) %>;
	sCaseNum = <%=JavaScriptSanitize(sCaseNum) %>;
	sesm_report_id = <%=JavaScriptSanitize(sesm_report_id) %>;
	
	strURL = "/E2B/Misc/E2B_EsmAcknowledgePrint.asp";
	sParam = "msg=" + sMessage + "&rpt=" + sReport + "&date=" + sDate + "&icsr=" + sICSR + "&CaseNum=" + fn_URLEncode(sCaseNum) + "&esm_report_id=" + sesm_report_id + "&DSPLYLNG=" + glDisplayLang;
	await loadArgusMessage(strURL, fn_CallBackDone,sParam);
}

async function fn_CallBackDone()
{
    var xmlDoc = this.req.responseXML;
    var sError = xmlDoc.getElementsByTagName("ERROR_STRING");
    var asDocId = xmlDoc.getElementsByTagName("GN_REPORT_IDENTIFIER");
  
    if (asDocId && (asDocId.length > 0))
        fn_ViewDocument(GetTextContentFromXML(asDocId[0]), "");
    else
        await MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("ESM_ACK")%>', GetTextContentFromXML(sError[0]));
}
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2B_ESMAcknowledge.asp
' Description  : Initialize the Page Content
'******************************************************************************
' Revision History
' Date		    Author		Description
' 02JUN2006 	Umar       	Original
'******************************************************************************
%>
function fn_init()
{
	report_ack_code = <%=JavaScriptSanitize(sreport_ack_code) %>;
	sError = "";
	if (oParentWindow.reportdetails.parsingerror.value.length > 0)
	{
		sError = oParentWindow.reportdetails.parsingerror.value;
		document.EsmAck.parsingerror.value = sError.substring(0, 32767);
	}
	if (oParentWindow.reportdetails.errormessage.value.length > 0)
	{
		document.EsmAck.errormessage.value = oParentWindow.reportdetails.errormessage.value;
		
		if (sError != "")
		{
		    sError = sError + '\n' + oParentWindow.reportdetails.errormessage.value;
		}
		else
		{
		    sError = oParentWindow.reportdetails.errormessage.value;
		}
	}
	document.EsmAck.ackError.value = sError.substring(0, 32767);
}
</script>

<!-- Javascript Functions Ends -->
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
