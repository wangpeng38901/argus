<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : AjaxSaveFollowUpDiffData.asp
' Description  : Save Followup Difference Data
'******************************************************************************
' Revision History
' Date		Author		Description
' 27MAY2006 Umar     Original
'******************************************************************************
%>

<!-- Declaration of Page Scope variables Starts -->
<%
	Dim oMessage, bSuccess, oOutMsg, lUserID, lError, sError, notes
	Dim e2b_lError, e2b_sError, e2b_sReturn, sql, lE2bType, esm_initial_report_id
	Dim l_esm_report_id, lASP, sSQL, lGmtOffSet, sDocId, case_id, case_num, sAccept_date, lIsJReport
	
	lGmtOffSet = GetXMLValueDirect (oSession, "GMT")
	l_esm_report_id = GetLong(GetRequest("esm_report_id"), 0)	
	sDocId = GetString(GetRequest("DocId"), "")
	case_id = GetLong(GetRequest("case_id"), -1)	
	lE2bType = GetLong(GetRequest("E2bType"), 0)	
	lUserID = oArgusUser.GetUserId()
	esm_initial_report_id = GetLong(GetRequest("esm_initial_report_id"), 0)
	sAccept_date = GetRequest("accept_date")    
	lIsJReport = GetRequest("IsJReport")
	notes = GetRequest("notes")

	if IsNullOrEmpty(sAccept_date) then
		sAccept_date = "-99"
	end if
	e2b_sError = ""
	e2b_lError = 0
	lASP = 1
	if (lE2bType = 1) then
		Call SetParameter("P_REPORT_ID", l_esm_report_id, PARAM_NUMBER)
		sql = "select case_xref from safetyreport where report_id = :P_REPORT_ID"
		case_id = ExecuteSQLReturnStr (sql, e2b_lError, e2b_sError)
	end if
	if (case_id > 0) then
		Call SetParameter("P_CASE_ID", case_id, PARAM_NUMBER)
		sql = "select case_num from case_master where case_id = :P_CASE_ID"
		case_num = ExecuteSQLReturnStr (sql, e2b_lError, e2b_sError)
	else
		case_num = ""
	end if

	'Create Message
	Call CreateMessage (oMessage, 300400064) 'MID_db_app_E2b_save_followup_report
	Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", l_esm_report_id)
	Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
	Call SetXMLValueDirect (oMessage, "GN_THIS_IS_ASP", lASP)
	Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
	Call SetXMLValueDirect (oMessage, "GN_REPORT_IDENTIFIER", sDocId)
	Call SetXMLValueDirect (oMessage, "CSM_CASE_ID", case_id)
	Call SetXMLValueDirect (oMessage, "CSM_CASE_NUM", case_num)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_E2B_TYPE", lE2bType)
	Call SetXMLValueDirect (oMessage, "CSAC_DATE_DONE", sAccept_date)
	Call SetXMLValueDirect (oMessage, "GN_NUMBER2", lIsJReport)
	Call SetXMLValueDirect (oMessage, "GN_UI_JUSTIFICATION", notes)
	
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, e2b_lError, e2b_sError)
	if case_id < 1 then
		e2b_sError = e2b_sError +  " " + GetTranslationData("E2BINCOME_FUDIFF_ERROR")
	end if
	if len(esm_initial_report_id) > 0 then
		Call SetParameter("P_REPORT_ID", esm_initial_report_id, PARAM_NUMBER)
		Call UpdateSQL("update SAFETYREPORT set status = 8 where report_id = :P_REPORT_ID", lError,sError)
	end if
	e2b_sReturn = ConstructAjaxErrorMessage(e2b_lError, e2b_sError)
	Response.Write e2b_sReturn
%>
