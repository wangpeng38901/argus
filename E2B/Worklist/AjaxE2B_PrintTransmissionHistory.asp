<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author		: UMAR
' Page			: AjaxPrintTransmissionHistory.asp
' Description	: Print the E2B Transmission History.
'
'******************************************************************************
' Revision History
' Date		    Author		Description
' 13JULY2006	UMAR         Original 
'******************************************************************************
%>
<%
    Dim oOutMessage, oMessageIn, lRetVal
    Dim sDocId, sAjaxResponse
    Dim CaseNum, messageId, agency, messageNum, screen, reportId, lGmtOffSet, lError, sError
    
    lError = 0
    sError = ""
    CaseNum = GetRequest("casenum")
    agency = GetRequest("agency")
    messageId = GetLong(GetRequest("messageId"), 0)
    screen = GetLong(GetRequest("screen"), 0)
    reportId = GetLong(GetRequest("reportId"), 0)
    messageNum = GetRequest("messageNum")
    lGmtOffSet = GetRequest("GmtOffSet")
    
    Call CreateMessage (oMessageIn, 300400062)
    Call SetXMLValueDirect (oMessageIn, "RPT_E2B_AGENCY_NAME", Agency)
    Call SetXMLValueDirect (oMessageIn, "CSM_CASE_NUM",CaseNum)
    Call SetXMLValueDirect (oMessageIn, "RPT_E2B_MESSAGE_ID",messageId)
    Call SetXMLValueDirect (oMessageIn, "RPT_E2B_LMESSAGENUMB",messageNum)
    Call SetXMLValueDirect (oMessageIn, "RPT_E2B_REG_REPORT_ID",reportId)
    Call SetXMLValueDirect (oMessageIn, "RPT_E2B_STAGE_ID",screen)
    Call SetXMLValueDirect (oMessageIn, "GN_PRINT",11)
    Call SetXMLValueDirect (oMessageIn, "GN_RPT_GMT_OFFSET", lGmtOffSet)    
    Call SetXMLValueDirect (oMessageIn, "GN_UI_KANJI_FLAG", glDisplayLang)
    Call SetXMLValueDirect (oMessageIn, "GN_GEN_SAVE_REPORT", 1)
    
    Set oOutMessage = ServiceRequest(oArgusSvr, oMessageIn, lError, sError)
    
    If lError = 0 Then
        sDocId = GetXMLValueDirect (oOutMessage, "GN_REPORT_IDENTIFIER")
        sAjaxResponse = "<MESSAGE><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER></MESSAGE>"
    Else
        sAjaxResponse = ConstructAjaxErrorMessage(lError, sError)
    End If
    Response.Write sAjaxResponse
%>


