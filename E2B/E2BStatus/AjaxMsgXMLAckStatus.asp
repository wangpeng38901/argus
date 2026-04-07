<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author		: UMAR
' Page			: AjaxMsgXMLAckStatus.asp
' Description	: Dsiplay ACK XML.
'
'******************************************************************************
' Revision History
' Date		    Author		Description
' 13JULY2006	UMAR         Original 
'******************************************************************************
%>

<%
	Dim oOutMessage, oMessageIn, MessageNumb
	Dim sDocId
	Dim lError, sError
	
	lError = 0
	sError = ""
	MessageNumb = GetRequest("MessageNumb")
	
	Call CreateMessage (oMessageIn, 300400066)
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_LMESSAGENUMB", MessageNumb)
	Call SetXMLValueDirect (oMessageIn, "GN_GEN_SAVE_REPORT", 1)
	
	Set oOutMessage = ServiceRequest(oArgusSvr, oMessageIn, lError, sError)
	
	If lError = 0 Then
		sDocId = GetXMLValueDirect (oOutMessage, "GN_REPORT_IDENTIFIER")
		Response.Write "<MESSAGE><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER></MESSAGE>"
	Else
		Response.Write ConstructAjaxErrorMessage(lError, sError)
	End If
%>


