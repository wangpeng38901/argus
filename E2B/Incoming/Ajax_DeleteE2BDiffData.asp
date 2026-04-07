<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Ankur Gupta
' Page         : Ajax_PrintE2BDiffReport.asp
' Description  : Ajax call for Printing the Difference Report PDF
'******************************************************************************
' Revision History
' Date		Author		Description
' 10SEP2008 Ankur G     Original
'******************************************************************************

	Dim e2b_lError, e2b_sError, e2b_sReturn
	Dim l_esm_report_id, l_user_id, strSQL

	l_esm_report_id = GetRequest("esm_report_id")
	l_user_id = oArgusUser.GetUserId()
	e2b_sError = ""
	e2b_lError = 0
	
	Call SetParameter("USER_NUM", l_user_id, PARAM_NUMBER)
	Call SetParameter("P_REPORT_ID", l_esm_report_id, PARAM_NUMBER)
	strSQL = "begin ESM_IMP_DIFF_REPORT.p_clear_inter_tables_all (:USER_NUM, :P_REPORT_ID); end;"
		
	Call UpdateSQL(strSQL, e2b_lError, e2b_sError)
	e2b_sReturn = ConstructAjaxErrorMessage(e2b_lError, e2b_sError)
	Response.Write e2b_sReturn
%>
