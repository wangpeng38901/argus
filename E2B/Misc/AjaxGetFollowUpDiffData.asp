<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : AjaxGetFollowUpDiffData.asp
' Description  : Followup Difference Data
'******************************************************************************
' Revision History
' Date		Author		Description
' 27MAY2006 Umar     Original
'******************************************************************************
%>


<!-- Declaration of Page Scope variables Starts -->
<%
	Dim oMessage, oOutMsg
	Dim e2b_lError, e2b_sError, e2b_sReturn
	Dim l_esm_report_id, l_report_id, lUserID, lIsJReport

	l_esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
	l_report_id = GetLong(GetRequest("current_report_id"), 0)
	lIsJReport  = GetLong(GetRequest("IsJReport"), 0)
	lUserID = oArgusUser.GetUserId()

	e2b_sError = ""
	e2b_lError = 0
   
	Call CreateMessage (oMessage, 400300069)
	Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", l_esm_report_id)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID", l_report_id)
	Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
	Call SetXMLValueDirect (oMessage, "GN_NUMBER1", lIsJReport)
		
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, e2b_lError, e2b_sError)
	
	if (e2b_lError <> 0) and (IsNullOrEmpty(e2b_sError)) then
		e2b_sError = "System Error. Unable to run the followup difference report."
	end if
	e2b_sReturn = ConstructAjaxErrorMessage(e2b_lError, e2b_sError)
   Response.Write e2b_sReturn
%>
