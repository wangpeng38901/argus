<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Eric Popejoy
' Page         : Ajax_E2BPendingReportIDs.asp
' Description  : Move the report IDs into user preferences so the dialog
'                can read them from the report server
'******************************************************************************
' Revision History
' Date		   Author		Description
' 26SEP2007    Eric P       Original
'******************************************************************************
%>

<%
	Dim p0, p1, p2, lError, sError
	
	p0 = cfCmn_FindRegEx(GetRequest("pinput0"), "0-9,")
	p1 = cfCmn_FindRegEx(GetRequest("pinput1"), "0-9,")
	p2 = cfCmn_FindRegEx(GetRequest("pinput2"), "0-9,")
	
	Call SetUserPreferences("E2bImport_pinput0", p0, lError, sError)
	If lError <> 0 Then
		Response.Write ConstructAjaxErrorMessage(lError, sError)
	Else
		Call SetUserPreferences("E2bImport_pinput1", p1, lError, sError)
		If lError <> 0 Then
			Response.Write ConstructAjaxErrorMessage(lError, sError)
		Else
			Call SetUserPreferences("E2bImport_pinput2", p2, lError, sError)
			If lError <> 0 Then
				Response.Write ConstructAjaxErrorMessage(lError, sError)
			End If
		End If
	End If

%>
