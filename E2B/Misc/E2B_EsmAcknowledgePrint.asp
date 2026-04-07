<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2b_ESMAcknowledgePrint.asp
' Description  : Call from E2b_ESMAcknowledge on Print button
'******************************************************************************
' Revision History
' Date		    Author		  Description
' 15MAY2006 	Umar       	Original
'******************************************************************************
%>
<%
	Dim sMessage, sReport, sDate, sICSR, sError1, sError2, sCaseNum, oOutMsg
	Dim lErrNo, s_pgError, bSuccess, sDocId, oMessage, sEsm_report_id
	Dim E2BAck_lError, E2BAck_sError
	
	E2BAck_lError = 0
	E2BAck_sError = ""
	sMessage = GetRequest("msg")
	sReport = GetRequest("rpt")
	sDate = GetRequest("date")
	sICSR = GetRequest("icsr")
	sCaseNum = GetRequest("CaseNum")
	sEsm_report_id = GetRequest("esm_report_id")
	bSuccess = true
	Call CreateMessage (oMessage, 300400014)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_LMESSAGENUMB", sMessage)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_RMESSAGENUMB", sICSR)
	Call SetXMLDate (oMessage, "RPT_E2B_MESSAGE_DATE", sDate)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_LOCALREPORTNUMB", sReport)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID", sEsm_report_id)
	Call SetXMLValueDirect (oMessage, "CSM_CASE_NUM", sCaseNum)
	Call SetXMLValueDirect (oMessage, "GN_UI_KANJI_FLAG", glDisplayLang)
	Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)

	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, E2BAck_lError, E2BAck_sError)
	sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
	if E2BAck_lError <> 0 then
		sError1 = ConstructAjaxErrorMessage(E2BAck_lError, E2BAck_sError)
	else
		sError1 = "<MESSAGE><ERROR_NUM>0</ERROR_NUM><ERROR_STRING></ERROR_STRING><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER></MESSAGE>"
	end if
	Response.Write sError1
%>
