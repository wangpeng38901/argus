<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<!DOCTYPE html>
<html>
<head>
	<!-- Page Title -->
	<title>Argus Web ICSR Viewer</title>
	<!-- Include Stylesheet here -->
	<link rel="stylesheet" href="/css/Relsys.css" />
	<!-- Client Library Includes Starts -->
	<!-- Client Library Includes Ends -->
	<style type="text/css">
		.collapse {
			position: absolute;
			visibility: hidden;
			display: none;
		}

		.expand {
			position: relative;
			visibility: visible;
			display: block;
		}
	</style>
</head>
<%
	Call BuildHiddenControlDirect("BlindingType", -1) 
Dim CaseID, ReportID, ProductID, LicenseID, CountryID, AgencyID, Draft, TimeFrame, E2BViewType, sDocId, Aware_Date, Aware_Date_J
Dim ErrorNum, Error, sCaseNumb, sCompanyNumb, sDTDVersion, IncomingE2b, E2bReportID, sSQL, StudyOver
Dim StateID, GenerateDate, UserID, EsmReportID, xmlDoc, Nodelist, ClickViewReport, E2BLicensed, Protected
Dim sTitle, e2b_lError, e2b_sError, e2b_type, ReGenerate, toolbardraft
Dim sValidationErrorReportDocId, ProdSeqNum
Dim bPMDAAuthority, AuthorityId, lAssocCount, lReportFormId, lApplyNewFW, lRpt_category_id
Dim oReport, oRecordList, oMsg
Dim sResults, oReportList, oMessage
Dim PMDAFormId
Dim PMDAFormDesc
Dim btnExportStyle
Dim PMDAJReportReceived
Dim sProfileRelease
Dim sHL7Profile
Dim iBlindValue
Dim lBlindedReport
Dim lBlindPMDAPprRep ,lJ10ProfileBlinded, lJ10LicBlinded,iProtectValue,bBlind,Blind,lInitialPMDA
	
PMDAJReportReceived = false
bPMDAAuthority=false
E2BLicensed = 1
CaseID = query_Crypt.Decrypt(GetRequest("CaseID"), GetEncryptKey())
ReportID = GetLong(query_Crypt.Decrypt(GetRequest("ReportID"), GetEncryptKey()), -1)
ProductID = GetRequest("ProductID")
ProdSeqNum = GetRequest("ProdSeqNum")
LicenseID = GetRequest("LicenseID")
CountryID = GetRequest("CountryID")
AgencyID = GetRequest("AgencyID")
Draft = GetLong(GetRequest("Draft"), 1)
TimeFrame = GetLong(GetRequest("TimeFrame"), 0)
E2BViewType = GetLong(GetRequest("E2BViewType"), -1)
IncomingE2b = GetLong(GetRequest("IncomingE2b"), 0)
E2bReportID = GetLong(query_Crypt.Decrypt(GetRequest("E2bReportID"), GetEncryptKey()), -1)
ClickViewReport = GetLong(GetRequest("ClickViewReport"), 0)
StudyOver = GetLong(GetRequest("StudyOver"), 1)
Protected = GetLong(GetRequest("Protected"), 0)
toolbardraft = GetLong(GetRequest("toolbardraft"), 0)
AuthorityId = GetLong(GetRequest("AuthorityId"), 0)
lApplyNewFW = GetLong(GetRequest("ApplyNewFW"), 0)
lReportFormId = GetLong(GetRequest("ReportFormId"), 0)
lBlindPMDAPprRep = GetLong(GetRequest("BlindedReport"), 0)
lRpt_category_id = GetLong(GetRequest("Rpt_category_id"), 0) ' ONLY To handle PMDA R3 BIP Reports, from draft view and medicial review screen
UserID = GetRequest("UserID")
sDocId = GetRequest("DocID")
Aware_Date = GetRequest("AwareDate")
Aware_Date_J = GetRequest("AwareDate_J")

lBlindedReport =0
lJ10ProfileBlinded = 0
lJ10LicBlinded = 0
iProtectValue=0
bBlind = 0
Blind = 0
lInitialPMDA=0

	'This call is mainly for the request from case forms
	If (AgencyID = "-1") then
		AgencyID = GetE2BAgencyId(E2bReportID)
	End If

	If (lReportFormId = 0 AND ReportID <> "-1" ) Then
		lReportFormId = GetReportFormId(ReportID)
	End If
	EsmReportID = 0
	StateID = 0
	UserID = -1
	GenerateDate = ""

	If ReportID > 0 Then
		Call SetParameter("P_REPORT_ID", ReportID, PARAM_NUMBER)
		sSQL = "Select esm_report_id, state_id, user_id, date_generated from cmn_reg_reports where reg_report_id = :P_REPORT_ID"
		sResults = "8410067, 8410019, 8410007, 8440010"
		Set oMessage = ExecuteSQL (sSQL, sResults, e2b_lError, e2b_sError)
		If e2b_lError = 0 Then
			Set oReportList = oMessage.selectNodes ("/MESSAGE/TABLE_CMN_REG_REPORTS/CMN_REG_REPORTS")
			For Each oReport in oReportList
				EsmReportID = GetLong(GetXMLValueDirect(oReport, "CMN_REG_REPORTS_ESM_REPORT_ID"), 0)
				StateID = GetLong(GetXMLValueDirect(oReport, "CMN_REG_REPORTS_STATE_ID"), 0)
				UserID = GetLong(GetXMLValueDirect(oReport, "CMN_REG_REPORTS_USER_ID"), -1)
				GenerateDate = GetXMLValueDirect(oReport, "CMN_REG_REPORTS_DATE_GENERATED")
				Exit For
			Next
		End If
	End If

   
	IF EsmReportID>0 OR E2bReportID>0 THEN
		IF EsmReportID>0 THEN
			Call SetParameter("P_REPORT_ID", EsmReportID, PARAM_NUMBER)
		   ELSE
			Call SetParameter("P_REPORT_ID", E2bReportID, PARAM_NUMBER)
		END IF
		sSQL = "select a.authority_id,Profile_Release, NVL(hl7_profile,0), report_form, NVL(apply_new_fw,0) apply_new_fw from cfg_profile a, safetyreport s where a.profile = s.profile and s.report_id = :P_REPORT_ID"
	ELSEIF AgencyID>0 THEN
		Call SetParameter("P_AGENCY_ID", AgencyID, PARAM_NUMBER)
		sSQL = "select a.authority_id,Profile_Release, NVL(hl7_profile,0), report_form, NVL(apply_new_fw,0) apply_new_fw from cfg_profile a, cfg_receiver b where b.message_profile = a.profile and b.agency_id = :P_AGENCY_ID"
	END IF
	
	sResults = "300150037, 300150038, 300150039, 13810002, 300150040"

	Set oMessage = ExecuteSQL (sSQL, sResults, e2b_lError, e2b_sError)
	AuthorityId = GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_unknown_300100000/unknown_300100000/GN_STRING1")
	sProfileRelease =GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_unknown_300100000/unknown_300100000/GN_STRING2")
	sHL7Profile = GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_unknown_300100000/unknown_300100000/GN_STRING3")
	lApplyNewFW = GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_unknown_300100000/unknown_300100000/GN_STRING4")

	If (lReportFormId <= 0) Then
		lReportFormId = GetLong(GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_unknown_300100000/unknown_300100000/LM_REPORT_FORMS_REPORT_FORM_ID"), 27)
	End If
			  
	If AuthorityId = "4" Then
		bPMDAAuthority = True
	Else
		bPMDAAuthority = False
	End If
	If UserID = -1 Then 
		UserID = GetLong(oArgusUser.GetUserId(), -1)
	End If
	   
	'End If
	ErrorNum = 0
	If (bPMDAAuthority) Then
		Call SetParameter("P_CASE_ID", CaseID, PARAM_NUMBER)
		sSQL = "select (CASE WHEN lm_studies.unblind_ok=1 AND code_broken>0 THEN 1 when study_type_id = 3 then 1 ELSE 0 END) " &_
					"  from case_study, lm_studies, case_master, lm_report_type, lm_study_cohorts lsc  " &_
					   "where lm_studies.study_key = case_study.study_key " &_
					   "and case_master.case_id = :P_CASE_ID and case_master.case_id = case_study.case_id " &_
					   "and case_master.rpt_type_id = lm_report_type.rpt_type_id " &_
					   "and case_study.cohort_id = lsc.cohort_id " &_
					   "and lm_report_type.incl_trial=1"
		iBlindValue = ExecuteSQLReturnStr(sSQL, e2b_lError, e2b_sError)

		Call SetParameter("P_REPORT_ID", ReportID, PARAM_NUMBER)
		sSQL = "SELECT count(1) FROM cfg_e2b ce,  cfg_profile cp,  cfg_receiver cr,  lm_regulatory_contact lrc,  cmn_reg_reports cmr " &_
				"WHERE ce.BLIND_PMDA_AE_PAPER_RPT = 1 AND ce.profile = cp.profile AND cp.profile  IN (cr.message_profile,cr.message_profile2) " &_
				"AND ce.dtd_element = 'MHLWADMICSRREPORTTIMESEVENT' AND cr.agency_id = lrc.agency_id AND lrc.agency_id = cmr.agency_id "&_
				"AND cmr.reg_report_id = :P_REPORT_ID AND cp.deleted is null AND ce.deleted is null AND cr.deleted is null " &_
				"AND cmr.deleted IS NULL AND lrc.deleted IS NULL "

		lJ10ProfileBlinded = ExecuteSQLReturnStr(sSQL, e2b_lError, e2b_sError)   

		sSQL = "SELECT NVL(blinded_field,0) blinded_field FROM lm_license ll,   cmn_reg_reports cmr " &_
				"WHERE ll.license_id   = cmr.license_id AND cmr.reg_report_id = :P_REPORT_ID and cmr.deleted is null and ll.deleted is null "

		lJ10LicBlinded = ExecuteSQLReturnStr(sSQL, e2b_lError, e2b_sError)  
		iProtectValue = GetLong(GetXMLValueDirect(oSession, "CFG_USERS_PROTECT_FROM_UNBLIND"), 0)        
		If (IncomingE2b = 1) Then 
			Call FetchPMDAReportFormForIncomingE2BReport(E2bReportID, sHL7Profile)
		Else
			Call FetchPMDAReportFormForE2BReport(ReportID, CaseID, LicenseID, ProdSeqNum) 'it internally sets PMDAFormId, PMDAFormDesc
		End If        

		If (PMDAFormId <= 0) Then
			ErrorNum = 402000713 'EID_PMDA_RPT_CATEGORY_NOT_DEFINED
			Error = "PMDA_NO_CATEGORY_EXISTS"
		End If
		If ((PMDAFormId = 81 OR PMDAFormId = 84 ) AND (E2BViewType = 6))Then
			If  iProtectValue = 0 Then
				If iBlindValue = "1" AND lJ10ProfileBlinded = "1" AND lJ10LicBlinded ="1" and lBlindedReport <> "1" Then
					bBlind = 1
				ElseIf iBlindValue = "1" AND lJ10ProfileBlinded = "1" AND lJ10LicBlinded ="1" and lBlindedReport = "1" Then
						bBlind = 2	'2 means blind the report
						Blind = 1
						lBlindPMDAPprRep=1
				ElseIf iBlindValue = "0" Then
					bBLind = 1
					Blind = 1
					lBlindPMDAPprRep=1
				End If
			Else
				If iBlindValue = "1" AND lJ10ProfileBlinded = "1" AND lJ10LicBlinded ="1" Then
					bBLind = 2
					Blind = 2
					lBlindPMDAPprRep=2
				End If
			End If
		End If
	End If

	If(sHL7Profile = "1" and(E2BViewType = 1 or E2BViewType = 2)) and IncomingE2b = 0 Then
		E2BViewType=0
	End If
	StateID = 0     'same reason as above
	
	If (ErrorNum = 0) Then ' if error PMDA_NO_CATEGORY_EXISTS error not faced so far
		e2b_type = GetRequest("e2b_type")
		ReGenerate = GetLong(GetRequest("ReGenerate"), 0)
		' AS 7.1 Bug 18363817 - FWDPRT: E2B GENERATED DATE IS UPDATED WHEN REPORT STATUS IS APPROVED 
		If ReGenerate = -1 then ReGenerate = 0 End If
		If Draft = 1 Then 
			sTitle = GetTranslationData("ICSR_VIEWER") & " (" & GetTranslationData("DRAFT") & ")"
		Else
			sTitle = GetTranslationData("ICSR_VIEWER")
		End If
		e2b_lError = 0
		e2b_sError = ""

		'Find if the Report is Draft or Final
	
		If E2BViewType < 0 Then
			If (bPMDAAuthority and sHL7Profile = "1") Then 
				E2BViewType = GetValueFromCMN_PROFILEOnKey ("PMDA_E2BR3_VIEW_PRINT")            
			ElseIf (bPMDAAuthority and sHL7Profile <> "1") Then 'PMDA Forms are considered only when both PMDA I and J Files have been receieved
				E2BViewType = GetValueFromCMN_PROFILEOnKey ("PMDA_E2B_VIEW_PRINT")
				If (Not PMDAJReportReceived And (E2BViewType = "3" Or E2BViewType = "5")) Then
					'if J report has not been reveived we cannot show J SGML(E2BViewType = 3), J Decoded View (E2BViewType = 5)
					E2BViewType = 0 'I-SGML would be shown
				End If
			ElseIf (sHL7Profile = "1" and lReportFormId = 27) Then 'For E2BR3 Non-PMDA
				E2BViewType = GetValueFromCMN_PROFILEOnKey ("E2BR3_VIEW_PRINT")                
			Else
				If lReportFormId = 44 Then
					E2BViewType = GetValueFromCMN_PROFILEOnKey ("EVAERS_VIEW_PRINT")
				ElseIf lReportFormId = 45 Then
					E2BViewType = GetValueFromCMN_PROFILEOnKey ("EMDR_VIEW_PRINT")
				Else
					E2BViewType = GetValueFromCMN_PROFILEOnKey ("E2B_VIEW_PRINT")
				End If    
			End If
		End If
		lInitialPMDA=E2BViewType

		If E2BViewType = 6 or E2BViewType = 11 Then ' if e2b viewer is requested to be seen for PMDA Form (E2BViewType = 6 or PMDA_E2B_VIEW_PRINT = 6)
			E2BViewType = PMDAFormId ' then it is actually the PMDA Form derived from the reporting category of the report
		End If
	
		If GetXMLValueDirect (oSession, "GN_AUTHORIZED_INTERCHANGE") = "0" Then
			E2BViewType = 0
			E2BLicensed = 0
		End If

        If Len(sDocID) > 0 Then
            ErrorNum = "0"
        ElseIf StudyOver = 0 and Protected > 0 then 'report was unnecessarily getting generated in this case, hence this check has been done here
	        'error raised is "E2BVIEW_STUDY_RESTRICT_PREVIEW"
            ErrorNum = "-999999" 'some arbitrary no. to raise "E2BVIEW_STUDY_RESTRICT_PREVIEW" below
	        sDocId = ""
        Else
            If ((PMDAFormId = 81 OR PMDAFormId = 84)  AND (lInitialPMDA = 6))Then
                If (iProtectValue = 0 and iBlindValue = "1" AND lJ10ProfileBlinded = "1" AND lJ10LicBlinded ="1" )Then 'only J10 blinding 2,0
                    Response.Write "<script type=""text/javascript"" language=""javascript"">"
					Response.Write "async function e2bViewMessage(){ var result = await MessageBoxResEx('EXP_REP_UNBINDED_INFO'," & glDisplayLang & ");if (result == 1){ fn_getElementByName(""BlindingType"").value=2;}else {fn_getElementByName(""BlindingType"").value=0;}} e2bViewMessage();" 
					Response.Write "</SCRIPT>"
				ElseIf (bBLind =1) Then                                                                                 'all blinding 1,0
					Response.Write "<script type=""text/javascript"" language=""javascript"">"
					Response.Write "async function e2bViewMessage(){ var result = await MessageBoxResEx('EXP_REP_UNBINDED_INFO'," & glDisplayLang & ");if (result == 1){fn_getElementByName(""BlindingType"").value=1;}else {fn_getElementByName(""BlindingType"").value=0;}} e2bViewMessage();"
					Response.Write "</SCRIPT>"
				Else
					If (IncomingE2b = 0 and ClickViewReport = 0) Then
						Call ExecSQLE2B()
					Else
						Call ExecE2BViewer()
					End If

					If IncomingE2b = 0 and EsmReportID > 0 and E2bReportID <= 0 Then
						E2bReportID = EsmReportID
					End If

					If IncomingE2b = 0 and EsmReportID <= 0 and E2bReportID > 0 Then
						EsmReportID = E2bReportID
					End If
				End If                    
			Else
				If (IncomingE2b = 0 and ClickViewReport = 0) Then
					Call ExecSQLE2B()
				Else
					Call ExecE2BViewer()
				End If

				If IncomingE2b = 0 and EsmReportID > 0 and E2bReportID <= 0 Then
					E2bReportID = EsmReportID
				End If

				If IncomingE2b = 0 and EsmReportID <= 0 and E2bReportID > 0 Then
					EsmReportID = E2bReportID
				End If
			End If
		End If
	End If

btnExportStyle = "width:60px"
If (glDisplayLang = cfCMN_LANG_JP) Then
	btnExportStyle = "width:80px"
End If
%>
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body class="no-margin" onload="initForm();">
	<form action="/E2B/E2bViewer/E2bViewer.asp?DSPLYLNG=<%=glDisplayLang%>&<%=GetRequestKeyValue()%>" method="post" name="fmHeader">
		<!-- Outer Box Starts -->
		<%Call BuildHiddenControlDirect("CaseID", query_Crypt.Encrypt(CaseID, GetEncryptKey())) %>
		<%Call BuildHiddenControlDirect("ReportID", query_Crypt.Encrypt(ReportID, GetEncryptKey())) %>
		<%Call BuildHiddenControlDirect("ProductID", ProductID) %>
		<%Call BuildHiddenControlDirect("ProdSeqNum", ProdSeqNum) %>
		<%Call BuildHiddenControlDirect("LicenseID", LicenseID) %>
		<%Call BuildHiddenControlDirect("CountryID", CountryID) %>
		<%Call BuildHiddenControlDirect("AgencyID", AgencyID) %>
		<%Call BuildHiddenControlDirect("Draft", Draft) %>
		<%Call BuildHiddenControlDirect("TimeFrame", TimeFrame) %>
		<%Call BuildHiddenControlDirect("IncomingE2b", IncomingE2b) %>
		<%Call BuildHiddenControlDirect("E2bReportID", query_Crypt.Encrypt(E2bReportID, GetEncryptKey())) %>
		<%Call BuildHiddenControlDirect("StateID", StateID) %>
		<%Call BuildHiddenControlDirect("UserID", UserID) %>
		<%Call BuildHiddenControlDirect("GenerateDate", GenerateDate) %>
		<%Call BuildHiddenControlDirect("EsmReportID", EsmReportID) %>
		<%Call BuildHiddenControlDirect("ClickViewReport", "0") %>
		<%Call BuildHiddenControlDirect("ExportPath", "") %>
		<%Call BuildHiddenControlDirect("e2b_type", e2b_type) %>
		<%Call BuildHiddenControlDirect("toolbardraft", toolbardraft) %>
		<%Call BuildHiddenControlDirect("AuthorityId", AuthorityId) %>
		<%Call BuildHiddenControlDirect("ReportFormId", lReportFormId) %>
		<%Call BuildHiddenControlDirect("BlindedStudy", iBlindValue) %>
		<%Call BuildHiddenControlDirect("BlindedReport", lBlindedReport) %>
		<%Call BuildHiddenControlDirect("J10ProfileBlinded", lJ10ProfileBlinded) %>
		<%Call BuildHiddenControlDirect("J10LicBlinded", lJ10LicBlinded) %>
		<%Call BuildHiddenControlDirect("RPT_CATEGORY_ID", lRpt_category_id) %>
        <%Call BuildHiddenControlDirect("Aware_Date", Aware_Date) %>
        <%Call BuildHiddenControlDirect("Aware_Date_J", Aware_Date_J) %>

		<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
                                    
		<table class="table" cellspacing="0" cellpadding="0" style="width: 100%; height: 100%;">
			<col width="30%" />
			<col width="30%" />
			<col width="40%" />
			<tr style="height: 25px">
				<td class="section-header-middle">
					<%BuildLabelDirect(sTitle).SetStyleSheet("label label-section").Render() %>
				</td>
			</tr>
			<%If GetLong(GetXMLValueDirect (oSession, "GN_AUTHORIZED_INTERCHANGE"), 0) > 0 Then%>
			<tr style="height: 80px;" valign="top">
				<td valign="top" style="width: 100%">
					<table cellpadding="2" cellspacing="2" style="width: 100%">
						<tr>
							<td>
								<% BuildLocalLabel("REP_TYPE").Render()%>
							</td>
							<td>
								<% BuildLocalLabel("SENDER_CASE_NUM").Render()%>
							</td>
							<td>
								<% BuildLocalLabel("VIEW_FORMAT").Render()%>
							</td>
						</tr>
						<tr valign="top">
							<td>
								<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "MsgHeader", e2b_type, true, 2, "").Style("width:80%").Render()%>
							</td>
							<td>
								<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "CaseNumb", sCaseNumb, true, 3, "").Style("width:80%").Render()%>
							</td>
							<td>
								<select id="E2BViewType" name="E2BViewType" class="ddlist" onchange="fn_ViewReport();" style="width: 90%;" tabindex="1">
									<%If bPMDAAuthority and sHL7Profile <> "1" Then %>
									<option value="0" <%if E2BViewType = 0 then Response.Write(" Selected ")%>>I-SGML</option>
									<%If PMDAJReportReceived Then 'if only I report has been received and imported then no J report and PMDA Form is shown%>
									<option value="3" <%if E2BViewType = 3 then Response.Write(" Selected ")%>>J-SGML</option>
									<%End If%>
									<option value="4" <%if E2BViewType = 4 then Response.Write(" Selected ")%>>I-Decoded View</option>
									<%If PMDAJReportReceived Then %>
									<option value="5" <%if E2BViewType = 5 then Response.Write(" Selected ")%>>J-Decoded View</option>
									<%End If%>
									<%If PMDAFormId > 0 Then %>
									<option value='<%="" & PMDAFormId & "" %>' <%if CCur(E2BViewType) = CCur(PMDAFormId) then Response.Write(" Selected ")%> title=<%= "'" & Fn_Sanitize(PMDAFormDesc) & "'" %>><%=Fn_Sanitize(PMDAFormDesc)%></option>
									<%End If%>
									<%Else %>
									<option value="0" <%if E2BViewType = 0 then Response.Write(" Selected ")%>>XML</option>
									<option value="4" <%if E2BViewType = 4 then Response.Write(" Selected ")%>><%=GetTranslationData("DECODED_VIEW") %></option>

									<!--Show conditional HL7 view if profile is for E2BR3-->
									<%If sHL7Profile = "1" And IncomingE2b = 0 Then %>
									<option value="7" <%if E2BViewType = 7 then Response.Write(" Selected ")%>>HL7 View</option>
									<%Else %>
									<%If sHL7Profile = "1" And lApplyNewFW = "1" Then %>
									<option value="7" <%if E2BViewType = 7 then Response.Write(" Selected ")%>>HL7 View</option>
									<%Else %>
									<option value="1" <%if E2BViewType = 1 then Response.Write(" Selected ")%>>CIOMS</option>
									<option value="2" <%if E2BViewType = 2 then Response.Write(" Selected ")%>>MedWatch</option>
									<%End If%>
									<%End If%>

									<%If PMDAFormId > 0 Then %>
									<option value='<%="" & PMDAFormId & "" %>' <%if CCur(E2BViewType) = CCur(PMDAFormId) then Response.Write(" Selected ")%> title=<%= "'" & Fn_Sanitize(PMDAFormDesc) & "'" %>><%=Fn_Sanitize(PMDAFormDesc)%></option>
                                    <%End If%>                                    
                                    <%End If%>
                                </select>
                            </td>
                        </tr>
                        <tr valign="top">
                            <td>
                                <% BuildLocalLabel("REP_ID_NUM").Render()%>
                            </td>
                            <td>
                                <% BuildLocalLabel("DTD_VER").Render()%>
                            </td>
                            <td>&nbsp;
                            </td>
                        </tr>
                        <tr valign="top">
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "CompanyNumb", sCompanyNumb, true, 4, "").Style("width:80%").Render()%>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "DTDVersion", sDTDVersion, true, 5, "").Style("width:80%").Render()%>
                            </td>
                            <td>&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <%Else %>
            <%Call BuildHiddenControlDirect("E2BViewType", E2BViewType) %>
            <%End If%>
           <tr>
                <td>
                    <% If ((E2BViewType = 0) or (E2BViewType = 3) or (E2BViewType = 7)) Then %>
                        <%If Not IsNullOrEmpty(sDocId) Then %>
                            <iframe name="DisplayReport" id="DisplayReport" style="height: 100%; width: 100%;" src="" border="0"></iframe>
                        <%End If %>
                    <%ElseIf E2BViewType <> 4 And E2BViewType <> 5 Then%>
                        <%If Not IsNullOrEmpty(sDocId) Then %>
                            <iframe name="DisplayReport" id="DisplayReport" style="height: 100%; width: 100%;" src="/ArgusNet/Common/DocViewer.aspx?DocId=<%=Server.URLEncode(sDocId)%>&<%=GetRequestKeyValue()%>" border="0"></iframe>
                        <%End If %>
                    <%Else%>
                    <iframe name="DisplayReport" id="DisplayReport" style="height: 100%; width: 100%" src="/E2B/E2bViewer/E2bViewerSGMLDecode.asp?SGMLDecode=1&DocId=<%=Server.URLEncode(sDocId)%>&casenumb=<%=Server.URLEncode(sCompanyNumb)%>&EsmReportID=<%=EsmReportID%>&HL7Profile=<%=sHL7Profile%>&sDTDVersion=<%=sDTDVersion%>&E2BViewType=<%=E2BViewType%>&DSPLYLNG=<%=glDisplayLang%>&sProfileRelease=<%=sProfileRelease%>&<%=GetRequestKeyValue()%>&IncomingE2b=<%=IncomingE2b%>&sValidationErrorReportDocId=<%=sValidationErrorReportDocId%>&AuthorityId=<%=AuthorityId%>"
                        border="0"></iframe>
                    <%End If%>
                </td>
            </tr>
			<tr class="tblheader-gray" style="height: 25px">
				<td align="center" valign="middle">
					<%If Draft = 0 and (E2BViewType = 0 or E2BViewType = 3) Then%>
					<%BuildButton("btnExport", "EXPORT", 7).Style(btnExportStyle)_
					.OnClick("fn_Export();").Render()%>
					<%End If %>
					<%BuildButton("btnPrint", "PRINT", 8).Style("width:60px")_
					.OnClick("fn_PrintReport();").Render()%>
					<%BuildButton("btnClose", "BTN_CLOSE", 9).Style("width:60px")_
					.OnClick("fn_close();").Render() 
					%>
				</td>
			</tr>
		</table>
	</form>
	<!-- Outer Box Ends -->
</body>
</html>

<script type="text/javascript">
    var isFormSubmitted=0;
    var xmlSrc = "/ArgusNet/Common/DocViewer.aspx?DocId=<%=Server.URLEncode(sDocId)%>&<%=GetRequestKeyValue()%>";
	var isFormSubmitted=0;
	async function fn_ViewReport() {
		//Hide the iframe
		var lInitialPMDA=<%=JavaScriptSanitize(lInitialPMDA) %>;
		var obj = document.getElementById("DisplayReport");
		var ddl = document.getElementById("E2BViewType");
		var lBlindedStudy= fn_getElementByName("BlindedStudy").value;
		var PMDAFormId = ddl.options[ddl.selectedIndex].value;
		var lJ10ProfileBlinded = fn_getElementByName("J10ProfileBlinded").value;
		var lJ10LicBlinded = fn_getElementByName("J10LicBlinded").value;
		if(lInitialPMDA == 6)
		{
			fn_getElementByName("BlindedReport").value=fn_getElementByName("BlindingType").value;
			showLoading();
			document.fmHeader.ClickViewReport.value = 0;
			isFormSubmitted=1;
			fn_ValidateSubmitForm(document.fmHeader);
		}
		else
		{
			if((PMDAFormId == "81" || PMDAFormId == "84") && (lBlindedStudy === "0"))
			{
				var result = await MessageBoxRes("EXP_REP_UNBINDED_INFO");
				if (result == 1)
					fn_getElementByName("BlindedReport").value=1;
				else
					fn_getElementByName("BlindedReport").value=0;
			}
			else if((PMDAFormId == "81" || PMDAFormId == "84") && lJ10ProfileBlinded == "1" && lJ10LicBlinded == "1")
			{
				var result = await MessageBoxRes("EXP_REP_UNBINDED_INFO");
				if (result == 1)
					fn_getElementByName("BlindedReport").value=2;
				else
					fn_getElementByName("BlindedReport").value=0;
			}

			obj.style.display = "none";
			showLoading();
			document.fmHeader.ClickViewReport.value = 1;
			isFormSubmitted=1;
			fn_ValidateSubmitForm(document.fmHeader);
		}
	}

	function fn_PrintReport() {
		var dataTab,dataTabtd,TabData, iframe;
		document.DisplayReport.focus();
		//only for SGML decode View
		if((document.fmHeader.E2BViewType.value == 4) || (document.fmHeader.E2BViewType.value == 5))
		{
			dataTab = document.getElementById("DisplayReport").contentWindow.TabDisplay;
			dataTabtd = document.getElementById("DisplayReport").contentWindow.TabPrint;

			dataTab.style.position = 'absolute';
			dataTab.style.visibility = 'hidden';
			dataTab.style.display = 'none';

			dataTabtd.style.position = 'relative';
			dataTabtd.style.visibility = 'visible';
            dataTabtd.style.display = 'table-cell';
		}
		try {
            iframe = document.getElementById('DisplayReport');
			iframe.contentWindow.document.execCommand('print', false, null);
		}
		catch(e) {
			document.getElementById("DisplayReport").contentWindow.print();
		}

		//only for SGML decode View
		if((document.fmHeader.E2BViewType.value == 4) || (document.fmHeader.E2BViewType.value == 5))
        {
			dataTab.style.position = 'relative';
			dataTab.style.visibility = 'visible';
            dataTab.style.display = 'table';
			dataTabtd.style.position = 'absolute';
			dataTabtd.style.visibility = 'hidden';
			dataTabtd.style.display = 'none';
		}
	}

	function fn_Export()
	{
		fn_ViewDocument("<%=sDocId%>", "");
	}

	function fn_close()
	{
		window.close();
	}

	async function fn_RemoveReport()
    {
		var reg_report_id;
		var Draft;
		var E2bIncoming;
		var state_id = <%=JavaScriptSanitize(StateID) %>;
		var timeframe = <%=JavaScriptSanitize(TimeFrame) %>;
		var user_id = <%=JavaScriptSanitize(UserID)%>;
		var date_generated = <%=JavaScriptSanitize(GenerateDate) %>;
		var esm_report_id = <%=JavaScriptSanitize(EsmReportID) %>;
		reg_report_id = <%=JavaScriptSanitize(ReportID) %>;
		Draft = <%=JavaScriptSanitize(Draft) %>;
		E2bIncoming = <%=JavaScriptSanitize(IncomingE2b) %>;
		if (Draft == 1 && E2bIncoming == 0  && isFormSubmitted==0)
        {
			var strURL = "/E2B/Actions/Ajax_RemoveE2bReport.asp?deleted=1&esm_report_id=" + esm_report_id + "&reg_report_id=" + reg_report_id;
            SendArgusMessage(strURL);
			if ((state_id == 2) && (date_generated.length == 0) && (timeframe == 0) && (user_id == 0))
			{
                strURL = "/Reports/ExpeditedReports/Ajax_removereport.asp?reg_report_id=" + reg_report_id;
                SendArgusMessage(strURL);	
			}
		}
	}

	async function fn_CallbackRemove()
	{
		var xmlDoc = this.req.responseXML; 
		var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
		if (sErrStr.length > 0)
		{
			await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("E2B_REPORTS")%>', sErrStr);
			return;
		}
	}

	function fn_ShowErrorPdf()
	{
		fn_ViewDocument("<%=sValidationErrorReportDocId%>", "");
		window.close();
	}
    var xmlResponse = "";
    var xmlRet = 1;
    async function loadXMLDoc(filename) {
  		xmlResponse = "";
    	xmlRet = 1;
        var xhttp = new XMLHttpRequest();
        var error;
        try {
            xhttp.onreadystatechange = function () {
                if (this.readyState == 4) {
                    if (this.status == 200) {
                        xmlResponse = xhttp.response;
                        xmlRet = 0;
                        return;
                    } else {
                        xmlRet = -1;
                        return xmlRet;
                    }
                }
            }
            xhttp.open("GET", filename, false);
            xhttp.send("");
        }
        catch (error) {
            xmlRet = -1;
            return xmlRet;
        }
    }
  
    async function initForm()
    {
        var strMsg;  
        var drpdwn = document.getElementById('E2BViewType');
        var report = document.getElementById('DisplayReport');
        var returnValue = 0;
        if ((drpdwn.options[drpdwn.selectedIndex].value == '0' || drpdwn.options[drpdwn.selectedIndex].value == '3' || drpdwn.options[drpdwn.selectedIndex].value == '7') && report != null) {
            await fn_ViewXMLReport();
			return;
        }
        if (document.fmHeader.ClickViewReport.value == 1)
        {
            var obj = document.getElementById("DisplayReport");
            obj.style.display = "block";
        }
        var ErrorNumber = "<%=ErrorNum%>";	

		if((fn_getElementByName("BlindingType").value >=0) && (ErrorNumber == "0"))
		{
		    await fn_ViewReport();
			return;
		}
		hideLoading();
		
		if (ErrorNumber == "402000713") // EID_PMDA_RPT_CATEGORY_NOT_DEFINED
		{                      
			await MessageBoxRes(<%=JavaScriptClean(Error) %>);
			fn_close();
		}
		else if (ErrorNumber == "402000714") // EID_PMDA_VALIDATION_ERROR_FOUND
		{            
			fn_close();
			fn_ShowErrorPdf();
		}
		else if (ErrorNumber == "402000640") // EID_RPT_CONFORMANCE_RULES_FAILURE
		{            
			fn_close();
			fn_ShowErrorPdf();
		}
		else if (ErrorNumber == "-999999")
		{
			await MessageBoxRes("E2BVIEW_STUDY_RESTRICT_PREVIEW", "[[E2B_VIEW_ERROR]]");
			fn_close();
		}
		else if (ErrorNumber == "402000718") //handled specifically for E2B Viewer and PMDA ICSR
		{
			await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("E2B_VIEW_ERROR")%>', <%=JavaScriptClean(Error) %>);		
			fn_close();
		}
		else if (ErrorNumber != "0")
		{
			var iCharsPerLine = 48;
			var ErrorText = <%=JavaScriptClean(Error)%>;
			ErrorText = ErrorText.replace(new RegExp(/,/g), ", ");

			var temp = "";
			var temp2 = "";
			var aRow = new Array();
			aRow[0] = "";
			if (ErrorText.length > iCharsPerLine)
			{
				temp = ErrorText.split(" ");
				for (var i=0,j=0;i<temp.length;++i)
				{
					if (aRow[j].length + temp[i].length > iCharsPerLine)
					{
						j++;
						aRow[j] = ""
					}
					aRow[j] = aRow[j] + " " + temp[i];
				}
				for (var i=0;i<=j;++i)
					temp2 = temp2 + aRow[i] + "\n";
			}
			else
				temp2 = ErrorText;

			strMsg = "";
			if (ErrorNumber != 402000630)                                         
			{ 
				strMsg = '<%=GetAgencyError("RPT_ERROR_NO")%>' + ErrorNumber + "\n";
				strMsg = strMsg + '<%=GetAgencyError("DESC") %>'
			}
			strMsg = strMsg + temp2;

			await MessageBoxRes("GENERAL_WARNING", '<%=GetTranslationData("E2B_VIEW_ERROR")%>', strMsg); 
			fn_close();
		}        
		// 19453481,19621934 : Do not update icon or generate date.  Keep it like paper reports for now.    
	}

    async function fn_ViewXMLReport() {
        //this is xml report
        var xmlUrl = xmlSrc;
        var btnPrint = document.getElementById('btnPrint');
        var windowWidth = screen.width / 2;
        var windowHeight = screen.availHeight;
        var windowDim = 'height=' + windowHeight + ',width=' + windowWidth;
        btnPrint.disabled = false;
        var iframe = document.getElementById("DisplayReport");
        var pre = document.createElement("pre");
        pre.style.wordBreak= 'break-word';
        pre.style.whiteSpace= 'break-spaces';

        // using HXMLTTPRequest
        returnValue = await loadXMLDoc(xmlUrl);
		if (xmlRet == 0) {
            pre.textContent = xmlResponse;
            var idoc = iframe.contentDocument || iframe.contentWindow.document; // IE compat
            let div = document.createElement('div');
            div.id = 'container';
            idoc.body.appendChild(div);
            var container = idoc.getElementById("container");
            container.appendChild(pre);
        }
        else
             await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("E2B_REPORTS")%>', "");
    }
</script>

<script type="text/vbscript" language="vbscript" runat="server">
Function GetCaseNumFromCompanyNum(sCaseNum, sCompanyNum)
    Dim lPos, sReturnCaseNum
    sReturnCaseNum = sCaseNum
	If lReportFormId <> 45 Then
        If (Right(sCompanyNum, Len(sCaseNum)) <> sCaseNum) OR (Len(sCaseNum) = 0) Then
		    lPos = InStrRev(sCompanyNumb, "-", -1, 1)
		    If lPos <> 0 Then
			    sReturnCaseNum = Mid(sCompanyNumb, lPos+1)
            End If
	    End If
	End If
    GetCaseNumFromCompanyNum = sReturnCaseNum
End Function

Function ExecSQLE2B()
	Dim oMessage, oOutMsg, sGenerateError, sSQL
	Call CreateMessage (oMessage, 400200015)
	Call SetXMLValueDirect (oMessage, "CSM_CASE_ID", CaseID)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_REG_REPORT_ID", ReportID)
	Call SetXMLValueDirect (oMessage, "LM_PRODUCT_PRODUCT_ID", ProductID)
	Call SetXMLValueDirect (oMessage, "LM_LICENSE_LICENSE_ID", LicenseID)
	Call SetXMLValueDirect (oMessage, "LM_MANUFACTURER_MANUFACTURER_ID", -1)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_COUNTRY_ID", CountryID)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AGENCY_ID", AgencyID)
	Call SetXMLValueDirect (oMessage, "GN_RPT_PRT_DRAFT", Draft)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_TIMEFRAME", TimeFrame)
	Call SetXMLValueDirect (oMessage, "GN_GUI_WORKLIST_VIEW_ALL", -1)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_VIEW_TYPE",E2BViewType)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID",-1)
	Call SetXMLValueDirect (oMessage, "GN_THIS_IS_ASP",1)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_USER_ID", UserID)
	Call SetXMLValueDirect (oMessage, "GN_RPT_REGENERATE", ReGenerate)
	Call SetXMLValueDirect (oMessage, "GN_AUTHORIZED_INTERCHANGE", E2BLicensed)    
	Call SetXMLValueDirect (oMessage, "RPT_E2B_TOOLBAR_DRAFT",toolbardraft)
	Call SetXMLValueDirect (oMessage, "GN_UI_KANJI_FLAG", glDisplayLang)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_PROD_SEQ_NUM", ProdSeqNum)' this is needed in case report id is blank i.e. non scheduled draft report like from toolbar         
	'GN_GUI_LM_GENERAL_TEXT is used by PMDA ICSR Validation Report to display user name on Validation Report
	Call SetXMLValueDirect (oMessage, "GN_GUI_LM_GENERAL_TEXT", GetXMLValueDirect(oSession, "CFG_USERS_USER_FULLNAME"))
	Call SetXMLValueDirect (oMessage, "LM_REPORT_FORMS_REPORT_FORM_ID", lReportFormId)
	Call SetXMLValueDirect (oMessage, "RPT_PDF_BLINDED", lBlindPMDAPprRep)
	Call SetXMLValueDirect (oMessage, "CFG_PROFILE_hl7_profile",sHL7Profile)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_RPT_CATEGORY_ID",lRpt_category_id)
	Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)						
    If toolbardraft = 1 Then
        If bPMDAAuthority = True then
            Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AWARE_DATE", Aware_Date_J)
        Else
            Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AWARE_DATE", Aware_Date)	
        End If
    End If

	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, e2b_lError, e2b_sError) 
	sGenerateError = ""
	ErrorNum = e2b_lError
	Error = e2b_sError
	If (ErrorNum <> 0) Then
		If IsNullOrEmpty(Error) Then
			If ErrorNum = 402000618 Then
			  Error = GetAgencyError("ESM_INVALID_FILEVIEW")
			ElseIf ErrorNum = 402000616 Then
			  Error = GetAgencyError("ESM_DLL_NO_REGISTER")
			ElseIf ErrorNum = 402000610 Then
			  Error = GetAgencyError("ESM_MAND_MISS")
			ElseIf ErrorNum = 402000609 Then
			  Error = GetAgencyError("ESM_GENERAL_ERROR")
			ElseIf ErrorNum = 402000611 Then
			  Error = GetAgencyError("ESM_OPT_MISS")
			ElseIf ErrorNum = 402000612 Then
			  Error = GetAgencyError("ESM_UNSPEC_ERR")
			ElseIf ErrorNum = 402000617 Then
			  Error = GetAgencyError("ESM_INVALID_DTDVER")
			ElseIf ErrorNum = 402000608 Then
			  Error = GetAgencyError("ESM_PCK_FAILED")
			ElseIf ErrorNum = 402000602 Then
			  Error = GetAgencyError("ESM_NO_RPT_DATA")
			End If
		End If
		
		'We have to show Validation Error Pdf in case of EID_PMDA_VALIDATION_ERROR_FOUND
		If  ErrorNum = 402000714 or ErrorNum = 402000640 then
			Error = ""
			sValidationErrorReportDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
		End If
				   
		If StateID = 2 and GenerateDate = "" and TimeFrame = 0 and UserID = 0 and EsmReportID <= 0 Then
			Error = Error
		Else
			If not IsNullOrEmpty(ReportID) and (ErrorNum <> 402000630) and (ErrorNum <> 402000718) Then
				'for EID_RPT_E2B_GENERATION_FAILED(402000718) no need for the following logic
				Call SetParameter("P_REPORT_ID", ReportID, PARAM_NUMBER)
				sSQL="select generation_error from cmn_reg_reports where reg_report_id =:P_REPORT_ID"
				sGenerateError = ExecuteSQLReturnStr (sSQL, e2b_lError, e2b_sError)
				If Len(sGenerateError) > 0 Then
					Error = Error & " " & GetAgencyError("ESM_SEE_RPT_DETAIL")
				End If
			  End If
		End If
		EsmReportID = 0
		ExecSQLE2B = false
		
		Exit Function
	End If
	e2b_type = GetXMLValueDirect(oOutMsg, "RPT_E2B_REPORT_TYPE")
	if IsNullOrEmpty(e2b_type) then e2b_type = ""
	sDocId = GetXMLValueDirect(oOutMsg, "GN_REPORT_IDENTIFIER")
	'Get Case Numb, CompanyNumb and DTD Version
	sCaseNumb = GetXMLValueDirect(oOutMsg, "CSM_CASE_NUM")
	sCompanyNumb = GetXMLValueDirect(oOutMsg, "RPT_E2B_COMPANYNUMB")
	sDTDVersion = GetXMLValueDirect(oOutMsg, "RPT_E2B_DTD_VERSION")
	If E2BViewType = 3 or E2BViewType=5 Then
		sDTDVersion = GetXMLValueDirect(oOutMsg, "RPT_E2B_MHLW_DTD_VERSION")
	End If
	EsmReportID = GetLong(GetXMLValueDirect(oOutMsg, "RPT_E2B_REPORT_ID"), 0)
    sCaseNumb = GetCaseNumFromCompanyNum(sCaseNumb, sCompanyNumb)
	ExecSQLE2B = true
End Function

Function GetAgencyError(sToken)
	If (bPMDAAuthority And glDisplayLang = cfCMN_LANG_JP) Or (Not bPMDAAuthority And glDisplayLang <> cfCMN_LANG_JP) Then
		GetAgencyError = GetTranslationData(sToken)
	Else
		GetAgencyError = GetTranslationDataAlt(sToken)
	End If
End Function

Function ExecE2BViewer()
	Dim oMessage, oOutMsg

	Call CreateMessage (oMessage, 400200016)
	If E2bReportID > 0 then
		Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID",E2bReportID)
	ElseIf EsmReportID > 0 then
		Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID",EsmReportID)
	End If
	Call SetXMLValueDirect (oMessage, "RPT_E2B_VIEW_TYPE",E2BViewType)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_E2B_TYPE",IncomingE2b)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_REG_REPORT_ID",ReportID)
	Call SetXMLValueDirect (oMessage, "GN_THIS_IS_ASP",1)
	Call SetXMLValueDirect (oMessage, "GN_UI_USER_DATABASE", oArgusUser.GetDbName())
	Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_TYPE",e2b_type)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_TOOLBAR_DRAFT",toolbardraft)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_USER_ID", UserID)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_PROD_SEQ_NUM", ProdSeqNum) ' this is needed in case report id is blank i.e. non scheduled draft report like from toolbar         
	Call SetXMLValueDirect (oMessage, "GN_GUI_LM_GENERAL_TEXT", GetXMLValueDirect(oSession, "CFG_USERS_USER_FULLNAME"))
	Call SetXMLValueDirect (oMessage, "GN_UI_KANJI_FLAG", glDisplayLang)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_STATE_ID", StateID)
	Call SetXMLValueDirect (oMessage, "LM_REPORT_FORMS_REPORT_FORM_ID", lReportFormId)
	Call SetXMLValueDirect (oMessage, "RPT_PDF_BLINDED", lBlindPMDAPprRep)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_LICENSE_ID", LicenseID)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AGENCY_ID", AgencyID)
	Call SetXMLValueDirect (oMessage, "CFG_PROFILE_hl7_profile",sHL7Profile)
	Call SetXMLValueDirect (oMessage, "GN_RPT_PRT_DRAFT",Draft)
	Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_RPT_CATEGORY_ID",lRpt_category_id)
	Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)
    If toolbardraft = 1 Then
        If bPMDAAuthority = True then
            Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AWARE_DATE", Aware_Date_J)
        Else
            Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AWARE_DATE", Aware_Date)	
        End If
    End If

	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, e2b_lError, e2b_sError)
	ErrorNum = e2b_lError
	Error = e2b_sError
	If (ErrorNum <> 0) Then
		EsmReportID = 0
		Error = "ERROR_GEN_FILE"
		ExecE2BViewer = false
		Exit Function
	End If
	e2b_type = GetXMLValueDirect(oOutMsg, "RPT_E2B_REPORT_TYPE")
	If IsNullOrEmpty(e2b_type) Then e2b_type = ""

	sDocId = GetXMLValueDirect(oOutMsg, "GN_REPORT_IDENTIFIER")    
	If IsNullOrEmpty(sDocId) Then
		ErrorNum = "-1"
		Error = "ERROR_GEN_FILE"
		ExecE2BViewer = false
		Exit Function
	End If
	'Get Case Numb, CompanyNumb and DTD Version
	sCaseNumb = GetXMLValueDirect(oOutMsg, "CSM_CASE_NUM")
	sCompanyNumb = GetXMLValueDirect(oOutMsg, "RPT_E2B_COMPANYNUMB")
	sDTDVersion = GetXMLValueDirect(oOutMsg, "RPT_E2B_DTD_VERSION")
	If E2BViewType = 3 or E2BViewType = 5 Then
		sDTDVersion = GetXMLValueDirect(oOutMsg, "RPT_E2B_MHLW_DTD_VERSION")
	End If
	EsmReportID = GetLong(GetXMLValueDirect(oOutMsg, "RPT_E2B_REPORT_ID"), 0)
    sCaseNumb = GetCaseNumFromCompanyNum(sCaseNumb, sCompanyNumb)
	ExecE2BViewer = true
End Function

'This function fetches PMDA Form id and Desc for displaying it in View Types DDL on screen
'there is only one report form fetched here based on reporting category id derived from the parameters sent to the function
Sub FetchPMDAReportFormForE2BReport(ReportID, CaseID, LicenseID, ProdSeqNum)
	Dim sSql
	Dim sFormId
	Dim Message
	Dim RecordList
	Dim Record
	PMDAJReportReceived = true  'in case of exported E2B, it is assumed that J is always received
	
	If IsNullOrEmpty(ReportID) Then 
		ReportID = 0 
	End If    

	Call SetParameter("P_CASE_ID", CaseID, PARAM_NUMBER)
	Call SetParameter("P_REPORT_ID", ReportID, PARAM_NUMBER)    
	Call SetParameter("P_LICENSE_ID", LicenseID, PARAM_NUMBER)
	Call SetParameter("PROD_SEQ_NUM", ProdSeqNum, PARAM_NUMBER)
	
	sSql = "SELECT report_form_id, form_desc FROM ("
	sSql = sSql & "SELECT lrf.report_form_id, lrf.form_desc "
	sSql = sSql & "FROM a$cmn_reg_reports crr, lm_rpt_category lrc, lm_report_forms lrf "
	sSql = sSql & "WHERE  crr.reg_report_id = :P_REPORT_ID AND crr.license_id = :P_LICENSE_ID AND crr.prod_seq_num = :PROD_SEQ_NUM "
	sSql = sSql & "AND crr.rpt_category_id = lrc.rpt_category_id AND lrc.report_form_id = lrf.report_form_id ) "
	sSql = sSql & "WHERE rownum = 1"

	Set Message = ExecuteSQL (sSQL, "13810002,13850003",e2b_lError,e2b_sError)
	Set RecordList = Message.selectNodes ("/MESSAGE/TABLE_LM_REPORT_FORMS/LM_REPORT_FORMS")    

	For Each Record in RecordList
		PMDAFormId = GetXMLValueDirect(Record, "LM_REPORT_FORMS_REPORT_FORM_ID")
		PMDAFormDesc = GetXMLValueDirect(Record, "LM_REPORT_FORMS_FORM_DESC")            
	Next

	If IsNullOrEmpty(PMDAFormId) And ReportID > 0 Then
	'in caes of nullification report scheduled, it is possible that above parameters like product id does not exist
	'in such a case look for the originally e2b report form    
		Call SetParameter("P_REPORT_ID", ReportID, PARAM_NUMBER)
		sSql = "SELECT lrf.report_form_id, lrf.form_desc "
		sSql = sSql & "FROM   a$cmn_reg_reports crr, mhlwadminitemsicsr m, mhlwadmicsrcasenum mr, lm_rpt_category lrc, lm_report_forms lrf "
		sSql = sSql & "WHERE  crr.reg_report_id = :P_REPORT_ID AND crr.nullification = 1 "    
		sSql = sSql & "AND crr.prev_esm_rpt_id = m.report_id AND m.mhlw_report_id = mr.mhlw_report_id "
		sSql = sSql & "AND (mr.mhlwadmicsrcasenumclass = lrc.rpt_category_id  OR GSS_UTIL.GET_RPT_CATEGORY_ID(mr.mhlwadmicsrcasenumclassr3) = lrc.rpt_category_id) "
		sSql = sSql & "AND lrc.report_form_id = lrf.report_form_id"
		
		Set Message = ExecuteSQL (sSQL, "13810002,13850003",e2b_lError,e2b_sError)
		Set RecordList = Message.selectNodes ("/MESSAGE/TABLE_LM_REPORT_FORMS/LM_REPORT_FORMS")    

		For Each Record in RecordList
			PMDAFormId = GetXMLValueDirect(Record, "LM_REPORT_FORMS_REPORT_FORM_ID")
			PMDAFormDesc = GetXMLValueDirect(Record, "LM_REPORT_FORMS_FORM_DESC")            
		Next
	End If

	If IsNullOrEmpty(PMDAFormId) And ReportID <= 0 Then
		'in caes of reports generated from Draft Tool Bar or Medical Review screen
		'in such a case look for case_pmda license table 
		sSql = "SELECT report_form_id, form_desc FROM ("
		sSql = sSql & "SELECT lrf.report_form_id, lrf.form_desc "
		sSql = sSql & "FROM case_pmda_license cpl, lm_rpt_category lrc, lm_report_forms lrf "
		sSql = sSql & "WHERE  cpl.case_id = :P_CASE_ID AND cpl.license_id = :P_LICENSE_ID AND cpl.prod_seq_num = :PROD_SEQ_NUM "
		sSql = sSql & "AND cpl.rpt_category_id = lrc.rpt_category_id AND lrc.report_form_id = lrf.report_form_id)"
		
		Set Message = ExecuteSQL (sSQL, "13810002,13850003",e2b_lError,e2b_sError)
		Set RecordList = Message.selectNodes ("/MESSAGE/TABLE_LM_REPORT_FORMS/LM_REPORT_FORMS")    

		For Each Record in RecordList
			PMDAFormId = GetXMLValueDirect(Record, "LM_REPORT_FORMS_REPORT_FORM_ID")
			PMDAFormDesc = GetXMLValueDirect(Record, "LM_REPORT_FORMS_FORM_DESC")            
		Next
	End If

	If IsNullOrEmpty(PMDAFormId) Then 
		PMDAFormId = 0 
		PMDAFormDesc = ""
	End If
End Sub

Function FetchPMDAReportFormForIncomingE2BReport(ReportID, sHL7Profile)
	Dim sSql
	Dim Message
	Dim RecordList
	Dim Record    
	Dim SafetyReportId
	Dim MHLWReportId
	
	If IsNullOrEmpty(ReportID) Then 
		ReportID = 0 
	End If    
	SafetyReportId = ""
	MHLWReportId = ""
	
	Call SetParameter("P_REPORT_ID", ReportID, PARAM_NUMBER)
	If (sHL7Profile = 1) Then
		sSQL = "SELECT lrc.REPORT_FORM_ID REPORT_FORM_ID, lrc.DESCRIPTION_J DESCRIPTION_J " 
		sSQL = sSQL & " FROM LM_RPT_CATEGORY lrc, ESM_IMP_REPORT_DET EIRD "     
		sSQL = sSQL & " WHERE EIRD.REPORT_ID = :P_REPORT_ID AND EIRD.RPT_CATEGORY_ID = lrc.RPT_CATEGORY_ID  AND ROWNUM = 1 "
		
		Set Message = ExecuteSQL (sSQL, "88110007, 88150301",e2b_lError,e2b_sError) 
		Set RecordList = Message.selectNodes ("/MESSAGE/TABLE_LM_RPT_CATEGORY/LM_RPT_CATEGORY") 
		   
		If(RecordList.length = 0) Then
			sSQL = "SELECT lrc.REPORT_FORM_ID REPORT_FORM_ID, lrc.DESCRIPTION_J DESCRIPTION_J " 
			sSQL = sSQL & " FROM LM_RPT_CATEGORY lrc, ESM_RPT_STAGE ERS "     
			sSQL = sSQL & " WHERE ERS.REPORT_ID = :P_REPORT_ID AND ERS.DATA = lrc.E2B_R3 AND TABLE_NAME ='MHLWADMICSRCASENUM'  "
			sSQL = sSQL & " AND DTD_ELEMENT = 'MHLWADMICSRCASENUMCLASSR3' AND ROWNUM = 1  "

			Set Message = ExecuteSQL (sSQL, "88110007, 88150301",e2b_lError,e2b_sError) 
			Set RecordList = Message.selectNodes ("/MESSAGE/TABLE_LM_RPT_CATEGORY/LM_RPT_CATEGORY")     
		End If
			
		
		For Each Record in RecordList
			PMDAFormId          = GetXMLValueDirect(Record, "LM_RPT_CATEGORY_REPORT_FORM_ID")
			PMDAFormDesc        = GetXMLValueDirect(Record, "LM_RPT_CATEGORY_DESCRIPTION_J")
		Next 
	Else
		sSQL = "SELECT sr.REPORT_ID safetyreportid, mai.REPORT_ID mhlwreportid, NVL(lrc.REPORT_FORM_ID, lrc1.REPORT_FORM_ID) REPORT_FORM_ID, NVL(lrc.DESCRIPTION_J, lrc1.DESCRIPTION_J) DESCRIPTION_J " 
		sSQL = sSQL & " FROM SAFETYREPORT sr, LM_RPT_CATEGORY lrc, MHLWADMICSRCASENUM mcn, MHLWADMINITEMSICSR mai, LM_RPT_CATEGORY lrc1 "     
		sSQL = sSQL & " WHERE sr.REPORT_ID = :P_REPORT_ID AND  sr.REPORT_ID  = mai.REPORT_ID (+) AND mai.MHLW_REPORT_ID = mcn.MHLW_REPORT_ID (+) AND mcn.MHLWADMICSRCASENUMCLASS = lrc.E2B_CODE (+) AND lrc1.Rpt_Category_Id = 1 AND ROWNUM = 1"
		'The Report Category is always assumed to be "A" in case it is not found.
	
		Set Message = ExecuteSQL (sSQL, "25110036, 300110042, 88110007, 88150301",e2b_lError,e2b_sError) 'GN_number1 has been used for mhlwreportid
		Set RecordList = Message.selectNodes ("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")        

		For Each Record in RecordList
			PMDAFormId          = GetXMLValueDirect(Record, "LM_RPT_CATEGORY_REPORT_FORM_ID")
			PMDAFormDesc        = GetXMLValueDirect(Record, "LM_RPT_CATEGORY_DESCRIPTION_J")
			SafetyReportId      = GetXMLValueDirect(Record, "RPT_E2B_REPORT_ID")
			MHLWReportId        = GetXMLValueDirect(Record, "GN_NUMBER1")
		Next    

		If (SafetyReportId > 0) Then
			If (MHLWReportId > 0) Then
				'if the record is not at all stored in MHLWAdminItemsICSR table it will be assumed that J report is not received
				PMDAJReportReceived = true  'it could be the case that only PMDA I report has been received and imported
				'if J report is not received then PMDA will be assumed to be "A" in consistent with E2BReceive.exe
			End If
		End If	
	End If
End Function

Function GetE2BAgencyId(E2BReportID)
	Dim lE2BAgencyId
	Call SetParameter("P_E2B_REPORT_ID", E2BReportID, PARAM_NUMBER)
	sSQL = "select agency_id from safetyreport where report_id = :P_E2B_REPORT_ID"
	lE2BAgencyId = ExecuteSQLReturnStr (sSQL, e2b_lError, e2b_sError)
	GetE2BAgencyId = lE2BAgencyId
End Function

Function GetReportFormId(ReportId)
	Dim lReportFormId
	Call SetParameter("P_REG_REPORT_ID", ReportId, PARAM_NUMBER)
	sSQL = "select report_form_id from cmn_reg_reports where reg_report_id = :P_REG_REPORT_ID"
	lReportFormId = ExecuteSQLReturnStr (sSQL, e2b_lError, e2b_sError)
	GetReportFormId = lReportFormId
End Function

</script>
<script type="text/javascript">
    addListener(this, "unload", async function(){await fn_RemoveReport();});
</script>
<!-- Server-Side Functions Ends -->
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
