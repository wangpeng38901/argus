<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
	Dim ReportID, sDocId, lGmtOffSet, lIsJReport, sCaseNum
	Dim VDR_lErrNo, VDR_sError, oMessage, oOutMsg, e2b_sReturn
	sDocId = ""
	VDR_lErrNo = 0

	lGmtOffSet = GetXMLValueDirect (oSession, "GMT")
	ReportID = GetLong(GetRequest("report_id"), -1)
	sCaseNum = GetString(GetRequest("case_num"), "")
	lIsJReport = GetLong(GetRequest("IsJReport"), -1)

	Call CreateMessage (oMessage, 400200037)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID",ReportID)		
	Call SetXMLValueDirect (oMessage, "RPT_E2B_SAFETYREPORTID",sCaseNum)
	Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
	Call SetXMLValueDirect (oMessage, "GN_DISPLAY_LANGUAGE", lIsJReport)
	Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)
	
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage,  VDR_lErrNo, VDR_sError)    	    
	If VDR_lErrNo = 0 then
		sDocId =  GetXMLValueDirect(oOutMsg, "GN_REPORT_IDENTIFIER")
		e2b_sReturn = "<MESSAGE><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER></MESSAGE>"
	Else
		if IsNullOrEmpty(VDR_sError) then VDR_sError = "System Error. Unable to run the followup difference report."
		e2b_sReturn = ConstructAjaxErrorMessage(VDR_lErrNo, VDR_sError)
	End If
	
	Response.Write e2b_sReturn
%>


