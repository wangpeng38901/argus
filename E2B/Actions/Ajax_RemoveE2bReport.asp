<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : DeleteReports.asp
' Description  : DeleteReport
'******************************************************************************
' Revision History
' Date		Author		Description
' 27MAY2006 Prakash     Original
'******************************************************************************
%>

<!-- Declaration of Page Scope variables Starts -->
<%
	Dim lReportId, esm_deleted
	Dim lErrNo, sError
	
	lErrNo = 0
	sError = ""
	
	esm_deleted = GetLong(Request.Querystring("deleted"), 0)
	lReportId = GetLong(Request.Querystring("esm_report_id"), -1)

	If lReportId > 0 Then
		Call SetParameter("P_ESM_REPORT_ID", lReportId, PARAM_NUMBER)
		If esm_deleted = 1 Then
			Call UpdateSQL("declare var number; begin var := ESM_PKG.F_CLEAN_CHILD_TABLES(:P_ESM_REPORT_ID); end;", lErrNo, sError)
		Else
			Call UpdateSQL("UPDATE SAFETYREPORT SET E2B = NULL WHERE STATUS = 9 AND REPORT_ID = :P_ESM_REPORT_ID", lErrNo, sError)
			Call UpdateSQL("UPDATE MHLWADMINITEMSICSR SET E2B = NULL WHERE STATUS = 9 AND REPORT_ID = :P_ESM_REPORT_ID", lErrNo, sError)
		End If
	End If
	Response.Write ConstructAjaxErrorMessage(lErrNo, sError)
%>
