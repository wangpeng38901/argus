<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : AjaxE2bStatusChange.asp
' Description  : Change the Status for E2B Report
'******************************************************************************
' Revision History
' Date		Author		Description
' 27MAY2006 Umar     Original
'******************************************************************************
%>

<!-- Declaration of Page Scope variables Starts -->
<%
Dim oMessage, oOutMsg, esm_report_id, notes, e2b_lError, esm_status, e2b_sError, e2b_sReturn, esm_initial_report_id

esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
notes = GetRequest("notes")
esm_status = GetLong(GetRequest("esm_status"), 0)
esm_initial_report_id = GetLong(GetRequest("esm_initial_report_id"), 0)
	  
e2b_sError = ""
e2b_lError = 0
if esm_initial_report_id > 0 then
	Call SetParameter("P_REPORT_ID", esm_initial_report_id, PARAM_NUMBER)
	CALL UpdateSQL("update SAFETYREPORT set status = 8, e2b_type_accept_as = NULL WHERE report_id = :P_REPORT_ID", e2b_lError,e2b_sError)
end if
Call CreateMessage (oMessage, 300100262)
Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", esm_report_id)
Call SetXMLValueDirect (oMessage, "GN_UI_JUSTIFICATION", notes)
Call SetXMLValueDirect (oMessage, "GN_STATUS_NUMBER", esm_status)
Call SetXMLValueDirect (oMessage, "CSAC_DATE_DONE", fn_date_from_iso_gmt(TodayNowGMT(), 8, true))
	
Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, e2b_lError, e2b_sError)
e2b_sReturn = ConstructAjaxErrorMessage(e2b_lError, e2b_sError)
Response.Write e2b_sReturn
%>
