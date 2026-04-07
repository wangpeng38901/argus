<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2bCheck.asp
' Description  : E2B Check
'******************************************************************************
' Revision History
' Date		Author		Description
' 27MAY2006 Umar     Original
'******************************************************************************
%>
<%
	Dim oMessage, oOutMsg
	Dim lCaseID, lUserID, sLengthCheckWarning, sDocId
	Dim e2b_lError, e2b_sError
	Dim e2b_sReturn

	e2b_lError = 0
	e2b_sError = ""
	sDocId = ""
	
	If Not (HasMenuAccess(171) AND HasMenuAccess(153)) Then 
		e2b_lError = "1"
		e2b_sError = "Not authorized to perform this operation."
	Else   
		lCaseID = GetLong(GetRequest("case_id"), -1)
		lUserID = oArgusUser.GetUserId()
		Call CreateMessage (oMessage, 300400061) 'MID_db_app_E2B_length_check_warning
		Call SetXMLValueDirect (oMessage, "RPT_E2B_COMMON_ID", lCaseID)
		Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
		Call SetXMLValueDirect (oMessage, "RPT_E2B_CHECK_E2B", "1")
		Call SetXMLValueDirect (oMessage, "RPT_E2B_IS_LM_DATA", 0)
		Call SetXMLValueDirect (oMessage, "GN_LANG_PREF", glDisplayLang)
		Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)
		
		Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, e2b_lError, e2b_sError) 
		If e2b_lError = 0 Then
			sLengthCheckWarning = GetXMLValueDirect (oOutMsg, "E2B_LENGTH_CHECK_WARNING")
			If sLengthCheckWarning <> "0" Then
				sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
			End If
		Else
			If IsNullOrEmpty(e2b_sError) Then 
				e2b_sError = "System Error. Unable to run the E2B check."
			End If
		End If
	End If
	e2b_sReturn = "<MESSAGE><ERROR_NUM>" & e2b_lError & "</ERROR_NUM><ERROR_STRING>" & ConvertXMLSpecialChars(e2b_sError) & "</ERROR_STRING><E2B_LENGTH_CHECK_WARNING>" & sLengthCheckWarning & "</E2B_LENGTH_CHECK_WARNING><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER></MESSAGE>"
	Response.Write e2b_sReturn
%>
