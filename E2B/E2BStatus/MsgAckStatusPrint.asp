<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : E2bTransmitStatusPrint.asp
' Description  : Print the Data for E2B Report that is Transmitted from Argus in PDF
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
    <title>Acknowledgement Status Print List</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
Dim oMessage, bSuccess, sError, oOutMsg, sDocId
Dim AckType, MessageNumb, lErrNo, lGmtOffSet

bSuccess = True
sError = ""
AckType = GetRequest("AckType")
MessageNumb = GetLong(GetRequest("MessageNumb"), 0)
lGmtOffSet = GetFloat(GetRequest("GmtOffSet"), 0)

Call CreateMessage (oMessage, 300400004)
Call SetXMLValueDirect (oMessage, "RPT_E2B_ACK_TYPE", AckType)
Call SetXMLValueDirect (oMessage, "RPT_E2B_LMESSAGENUMB", MessageNumb)
Call SetXMLValueDirect (oMessage, "GN_PRINT", 1)
Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)
Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lErrNo, sError)

If lErrNo <> 0 Then
	bSuccess = False
Else
	sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
End If
%>
<body onload="fn_init();">
</body>
</html>

<script type="text/javascript">
    async function fn_init()
	{
		<%If bSuccess Then %>
		fn_ViewDocument("<%=sDocId%>", "window");
		<%Else %>
            await MessageBoxResEx("EL_PRINT_ERROR",<%=glDisplayLang %>);
		<%End If %>
		window.close();
	}
</script>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
