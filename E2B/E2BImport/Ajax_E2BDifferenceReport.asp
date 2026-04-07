<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
	Dim ReportID, Draft, ReGenerate, sSQL, CaseID 
	Dim StateID ,ReportFormID, ReportType
	Dim oOutMsg, Pos, sGenerateError, tmpRptId
	Dim iBlindValue, sUserID, iProtectValue ,LockDate 
	Dim RptStateID, ProductID, LicenseID,CountryID 
	Dim AgencyID, TimeFrame 
	Dim iCase_id, strSQL
	Dim strResult, VDR_lErrNo, VDR_sError
	Dim lCountry_Id,sAgency_ID,lProduct_ID 
	Dim lLicense_Id, lProd_SeqNum, lCplSeqNum, dAward_Date
	
	Dim dReceivedDate,dFollowupDate, lForm_Id
	Dim oMessage, VDR_iBlindValue, VDR_iProtectValue ,sSQLDevice
	Dim protect, iAgencyID, e2b_sReturn
	Dim CurrentfollowupRptID,PreviousfollowupRptID,GeneratedRptID
	Dim sDocId, lUserID, l_call_pdf, l_e2b_type
	Dim lGmtOffSet, sSuccessMessage, lIsJReport, Message, RecordList, Record,oTable, sCaseNum
	Dim selectedIsApplyNewFw
	VDR_lErrNo = 0
	lUserID = oArgusUser.GetUserId()
	lGmtOffSet = GetXMLValueDirect (oSession, "GMT")

	ReportID = GetLong(GetRequest("report_id"), -1)
	l_call_pdf = GetLong(GetRequest("call_pdf"), 0)
	l_e2b_type = GetLong(GetRequest("e2b_type"), 0)
	
	'logic has been changed since earlier PreviousfollowupRptID and iCase_id were fetched through sql 
	'this was leading to a scenario where View difference was showing difference between followup and true initial case
	'even if a different case was selected in duplicate search screen.
	PreviousfollowupRptID = GetLong(GetRequest("initialreport"), 0)
	iCase_id = GetLong(GetRequest("case_id"), 0)    
	CurrentfollowupRptID = ReportID
	lIsJReport = GetLong(GetRequest("IsJReport"), 0)
	selectedIsApplyNewFw = GetLong(Request("ApplyNewFw"), 0)
	lCplSeqNum = 0
	sCaseNum = GetRequest("case_num")
	
	' if it is not an initial report then only find previous E2B and run case draft E2B
	If (l_e2b_type <> 1) then
		If  (PreviousfollowupRptID <= 0 and iCase_id <= 0) then
			Call SetParameter("P_REPORT_ID", ReportID, PARAM_NUMBER)
			sSQL= "select report_id from (  "
			sSQL= sSQL & "select max(report_id) report_id from safetyreport S where S.status = 102 and s.reportacknowledgment.reportacknowledgmentcode = '01' "
			sSQL= sSQL & "and upper(s.companynumb) = (select upper(ss.companynumb) from safetyreport ss where ss.report_id =:P_REPORT_ID and s.sender_agency=ss.sender_agency) "
			sSQL= sSQL & "UNION "
			sSQL= sSQL & "select max(report_id) from safetyreport S where S.status = 102 and s.reportacknowledgment.reportacknowledgmentcode = '01' "
			sSQL= sSQL & "and upper(s.authoritynumb) = (select upper(ss.authoritynumb) from safetyreport ss where ss.report_id =:P_REPORT_ID and s.sender_agency=ss.sender_agency) "
			sSQL= sSQL & ") B where report_id is not null and rownum = 1 "

			PreviousfollowupRptID = ExecuteSQLReturnStr (sSQL, VDR_lErrNo, VDR_sError)    

			Call SetParameter("P_PREV_REPORT_ID", PreviousfollowupRptID, PARAM_NUMBER)
			sSQL="select case_xref from safetyreport where report_id=:P_PREV_REPORT_ID"
			iCase_id = ExecuteSQLReturnStr (sSQL, VDR_lErrNo, VDR_sError)            
		ElseIf (iCase_id > 0  and PreviousfollowupRptID <= 0) then        '
			Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
			sSQL="select max(report_id) from safetyreport s where status = 102 and s.reportacknowledgment.reportacknowledgmentcode = '01' and s.case_xref = :P_CASE_ID"
			PreviousfollowupRptID = ExecuteSQLReturnStr (sSQL, VDR_lErrNo, VDR_sError)
			
			If trim(PreviousfollowupRptID) = ""  then
				PreviousfollowupRptID = 0
			End If			
		End If

		lForm_Id = "27"
		sSQLDevice = ""
		If lForm_Id = "2" or lForm_Id = "6" or lForm_Id = "7" or lForm_Id = "42" Then
			sSQLDevice = " + decode(views_available, 2, 0, 3, 0, 6, 0, 7, 0, 50)"
		End If
		Draft = 1
		ReGenerate = 0
		
		Call SetParameter("P_REPORT_ID", GetRequest("report_id"), PARAM_NUMBER)
		sSQL="select agency_id from safetyreport where report_id=:P_REPORT_ID"
		iAgencyID = ExecuteSQLReturnStr (sSQL, VDR_lErrNo, VDR_sError)
		
		If (lIsJReport = 1) Then
			Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
			Call SetParameter("P_ESM_REPORT_ID", ReportID, PARAM_NUMBER)
			Call SetParameter("P_APPLY_NEW_FW", selectedIsApplyNewFw, PARAM_NUMBER)
			sSQL = "SELECT SF_GET_CASE_PMDA_LICENSE_SEQ(:P_CASE_ID,:P_ESM_REPORT_ID, :P_APPLY_NEW_FW) FROM DUAL"
			lCplSeqNum = GetLong (ExecuteSQLReturnStr(sSQL, VDR_lErrNo, VDR_sError),-1)

			If (lCplSeqNum > 0) Then  
				Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
				Call SetParameter("P_AGENCY_ID", iAgencyID, PARAM_NUMBER)
				Call SetParameter("P_CPL_SEQ_NUM", lCplSeqNum, PARAM_NUMBER)
			
				strSQL = strSQL & " SELECT distinct -100 + cp.sort_id"& sSQLDevice &" style,"
				strSQL = strSQL & " :P_AGENCY_ID agency_id, cm.country_id, nvl(cp.product_id, cp.pat_exposure) product_id,"
				strSQL = strSQL & " cpl.license_id, cp.seq_num FROM "
				strSQL = strSQL & " case_master cm, case_product cp, case_pmda_license cpl WHERE "
				strSQL = strSQL & " cm.case_id = :P_CASE_ID AND cp.case_id = cm.case_id AND "
				strSQL = strSQL & " cpl.prod_seq_num = cp.seq_num AND cpl.case_id = cp.case_id AND cpl.seq_num = :P_CPL_SEQ_NUM AND "
				strSQL = strSQL & " NVL(cp.product_id, cp.pat_exposure) > 0 And ROWNUM =1"
				
				strResult = "13750519,13710001,2110002,3510004,11810002,3510011"
			Else
				VDR_lErrNo = 1
				VDR_sError = GetTranslationData("NO_PROD_LICENCE_FOUND_MSG")
			End If
		Else
			Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
			strSQL = "SELECT distinct -100 + case_product.sort_id"& sSQLDevice &" style, " &_
							"agency_id, case_master.country_id, nvl(case_product.product_id, pat_exposure),"
			strSQL = strSQL & "lm_lic_products.license_id, case_product.seq_num, lm_license.AWARD_DATE FROM "
			strSQL = strSQL & "lm_regulatory_contact, case_master, case_product, lm_lic_products, lm_license, cmn_profile WHERE "
			strSQL = strSQL & "case_master.case_id = :P_CASE_ID AND case_product.case_id (+) = case_master.case_id AND "
			strSQL = strSQL & "case_product.drug_type (+) = 1 AND lm_lic_products.PRODUCT_ID (+) = nvl(case_product.product_id, pat_exposure) AND "
			strSQL = strSQL & " NVL (case_product.product_id, pat_exposure) > 0 and "
			strSQL = strSQL & " lm_license.LICENSE_ID (+) = lm_lic_products.LICENSE_ID AND lm_license.COUNTRY_ID = case_master.COUNTRY_ID and "
			strSQL = strSQL & "agency_id = value and upper(section) = 'CASE' and upper(key) = 'DRAFT_AGENCY' union " &_
							  "select distinct 000 + case_product.sort_id"& sSQLDevice &" style, "
			strSQL = strSQL & "agency_id, case_master.country_id, nvl(case_product.product_id, pat_exposure),lm_lic_products.license_id, case_product.seq_num, "
			strSQL = strSQL & "lm_license.AWARD_DATE from lm_regulatory_contact, case_master, case_product, lm_lic_products, lm_license "
			strSQL = strSQL & "where case_master.case_id = :P_CASE_ID AND lm_regulatory_contact.COUNTRY (+) = case_master.COUNTRY_ID AND "
			strSQL = strSQL & "case_product.case_id (+) = case_master.case_id AND case_product.drug_type (+) = 1 AND "
			strSQL = strSQL & "lm_lic_products.PRODUCT_ID (+) = nvl(case_product.product_id, pat_exposure) AND " 
			strSQL = strSQL & " NVL (case_product.product_id, pat_exposure) > 0 and "
			strSQL = strSQL & "lm_license.LICENSE_ID (+) = lm_lic_products.LICENSE_ID AND lm_license.COUNTRY_ID = case_master.COUNTRY_ID "
			strSQL = strSQL & "and agency_id is not null union " &_
							  "select distinct 100 + case_product.sort_id"& sSQLDevice &" style, agency_id, case_master.country_id, "
			strSQL = strSQL & "nvl(case_product.product_id, pat_exposure), lm_lic_products.license_id, case_product.seq_num, lm_license.AWARD_DATE "
			strSQL = strSQL & "from lm_regulatory_contact, case_master, case_product, lm_lic_products, lm_license "
			strSQL = strSQL & "where case_master.case_id = :P_CASE_ID AND lm_regulatory_contact.COUNTRY (+) = case_master.COUNTRY_ID AND "
			strSQL = strSQL & "case_product.case_id (+) = case_master.case_id AND case_product.drug_type (+) = 1 AND "
			strSQL = strSQL & "lm_lic_products.PRODUCT_ID (+) = nvl(case_product.product_id, pat_exposure) AND"
			strSQL = strSQL & " NVL (case_product.product_id, pat_exposure) > 0 and "
			strSQL = strSQL & " lm_license.license_id = lm_lic_products.license_id and agency_id is not null union "
			strSQL = strSQL & " SELECT " 
			strSQL = strSQL & "  distinct 200 + case_product.sort_id"& sSQLDevice &" style, "
			strSQL = strSQL & "  nvl(lm_regulatory_contact.agency_id,"
			strSQL = strSQL & "  lrc.agency_id), "
			strSQL = strSQL & "  case_master.country_id, "
			strSQL = strSQL & "  nvl(case_product.product_id, "
			strSQL = strSQL & "  pat_exposure), "
			strSQL = strSQL & "  lm_lic_products.license_id, "
			strSQL = strSQL & "  case_product.seq_num, "
			strSQL = strSQL & "  lm_license.AWARD_DATE "
			strSQL = strSQL & "FROM  "
			strSQL = strSQL & "  lm_regulatory_contact, (select * from lm_regulatory_contact WHERE rownum=1) lrc, " &_
							  "  case_master, case_product, lm_lic_products, lm_license " &_
							  "WHERE case_master.case_id = :P_CASE_ID AND "
			strSQL = strSQL & "  lm_regulatory_contact.COUNTRY (+) = lm_license.COUNTRY_ID AND "
			strSQL = strSQL & "  case_product.case_id (+) = case_master.case_id AND "
			strSQL = strSQL & "  case_product.drug_type (+) = 1 AND "
			strSQL = strSQL & "  lm_lic_products.PRODUCT_ID (+) = nvl(case_product.product_id, pat_exposure) AND "
			strSQL = strSQL & " NVL (case_product.product_id, pat_exposure) > 0 and "
			strSQL = strSQL & "  lm_license.LICENSE_ID (+) = lm_lic_products.LICENSE_ID order by style, award_date"
			
			strResult = "13750519,13710001,2110002,3510004,11810002,3510011,12140008"
		End If
		
		If VDR_lErrNo = 0 then
			set oMessage = ExecuteSQL(strSQL, strResult, VDR_lErrNo, VDR_sError)
			Set oTable = oMessage.selectNodes("/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT")
			'oMessage.Save "c:\temp\ViewNewDraftProductsResult.xml"
			
			If VDR_lErrNo = "0" And oTable.length < 1 And lIsJReport = 1 Then
				VDR_lErrNo = 1
				VDR_sError = GetTranslationData("NO_PROD_LICENCE_FOUND_MSG")
			End If            
			
			If VDR_lErrNo = 0 then
				lCountry_Id = GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT/CSM_COUNTRY_ID")
				sAgency_ID = GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT_AGENCY_ID")
				lProduct_ID = GetLong(GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT/CSPD_PRODUCT_ID"), 0)
				lLicense_Id = GetLong(GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT/LM_LIC_COUNTRIES_LICENSE_ID"), 0)
				lProd_SeqNum = GetLong(GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT/CSPD_SEQ_NUM"), 0)
				dAward_Date = GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT/LM_LICENSE_AWARD_DATE")
				If lProduct_ID < 1 Then
					lProduct_ID = 0
				End If
				If lProd_SeqNum < 1 Then 
					lProd_SeqNum = 0
				End If
				
				If lProduct_ID <> 0 then
					if len(protect) > 0 then
						'get the aware date
						Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
						strSQL = "Select init_rept_date, followup_date from case_master where case_id = :P_CASE_ID"
						strResult = "2140009, 2140008"
						set oMessage = ExecuteSQL(strSQL, strResult, VDR_lErrNo, VDR_sError)
						dReceivedDate = fn_date_from_iso(GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_CASE_MASTER/CASE_MASTER/CSM_INIT_REPT_DATE"), 8, false)
						dFollowupDate = fn_date_from_iso(GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_CASE_MASTER/CASE_MASTER/CSM_FOLLOWUP_DATE"), 8, false)
						if len(dFollowupDate) > 0 then
							dReceivedDate = dFollowupDate
						end if
					end if
					Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
					sSQL = "select csuo.unblind_ok from case_study_unblind_ok csuo, case_master, lm_report_type " &_
						   "where case_master.case_id =  csuo.case_id " &_
						   "and case_master.case_id = :P_CASE_ID " &_
						   "and case_master.rpt_type_id = lm_report_type.rpt_type_id " &_
						   "and lm_report_type.incl_trial=1"
					VDR_iBlindValue = GetLong(ExecuteSQLReturnStr(sSQL, VDR_lErrNo, VDR_sError), 1)
					VDR_iProtectValue = GetLong(GetXMLValueDirect (oSession, "CFG_USERS_PROTECT_FROM_UNBLIND"), 0)
				End If
				
				//*************************************************************************************
				//            Code Block from E2B Viewer
				//  Create Message for generating XML.
				//*************************************************************************************            
				Call CreateMessage (oMessage, 400200015)
				Call SetXMLValueDirect (oMessage, "CSM_CASE_ID", iCase_id)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_REG_REPORT_ID", tmpRptId)
				Call SetXMLValueDirect (oMessage, "LM_PRODUCT_PRODUCT_ID", lProduct_ID)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_PROD_SEQ_NUM", lProd_SeqNum)
				Call SetXMLValueDirect (oMessage, "LM_LICENSE_LICENSE_ID", lLicense_Id)
				Call SetXMLValueDirect (oMessage, "LM_MANUFACTURER_MANUFACTURER_ID", -1)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_COUNTRY_ID", lCountry_Id)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AGENCY_ID", iAgencyID)
				Call SetXMLValueDirect (oMessage, "GN_RPT_PRT_DRAFT", Draft)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_TIMEFRAME", TimeFrame)
				Call SetXMLValueDirect (oMessage, "GN_GUI_WORKLIST_VIEW_ALL", -1)
				Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID",-1)		
				Call SetXMLValueDirect (oMessage, "GN_THIS_IS_ASP",1)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_USER_ID", lUserID)
				Call SetXMLValueDirect (oMessage, "GN_UI_KANJI_FLAG", glDisplayLang)
				Call SetXMLValueDirect (oMessage, "RPT_E2B_R3_IMPORT", 1) 'Supplied to identify that e2b generation is from View Difference
				Call SetXMLValueDirect (oMessage, "LM_REPORT_FORMS_REPORT_FORM_ID", 27)
				Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)
			
				Set oOutMsg = ServiceRequest(oArgusSvr, oMessage,  VDR_lErrNo, VDR_sError) 

				'We have to show Validation Error Pdf in case of EID_PMDA_VALIDATION_ERROR_FOUND
				If  VDR_lErrNo = 402000714 or VDR_lErrNo = 402000640 then
					Error = ""
					sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
				End If

				If VDR_lErrNo = 0 then
					sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
					If Not IsNullOrEmpty(sDocId) Then
						GeneratedRptID = GetXMLValueDirect(oOutMsg, "/MESSAGE/RPT_E2B_REPORT_ID")
						if trim(GeneratedRptID) = "" or trim(PreviousfollowupRptID) = "" or trim(GeneratedRptID) = "" or trim(iCase_id) = ""  then
							VDR_lErrNo = 1
							VDR_sError = GetTranslationData("NO_DATA_IN_REPORT")
						else
							if (PreviousfollowupRptID <= 0) then
								'this will be a case when there was no safetyreport for a case
								'in such a case latest generated e2b is considered to be the last imported e2b 
								PreviousfollowupRptID = GeneratedRptID
							end if
						end if
					Else
						VDR_lErrNo = 1
						VDR_sError = GetTranslationData("NO_PDF_FILE")
					End If
				Else
					if IsNullOrEmpty(VDR_sError) then VDR_sError = "System Error. Unable to run the followup difference report."
				End If
			End If
		End If    
	End if 

	if VDR_lErrNo = 0 then
		'Code block to generate the difference report PDF
		Call CreateMessage (oMessage, 400300069)
		Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", GeneratedRptID)
		Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
		Call SetXMLValueDirect (oMessage, "RPT_E2B_PREVIOUS_FOLLOWUP", PreviousfollowupRptID)
		Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID", CurrentfollowupRptID)
		Call SetXMLValueDirect (oMessage, "CSM_CID", iCase_id)
		Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AGENCY_ID", iAgencyID)
		Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
		Call SetXMLValueDirect (oMessage, "GN_NUMBER1", lIsJReport)
		
		Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, VDR_lErrNo, VDR_sError)
		if VDR_lErrNo <> 0 then
			VDR_lErrNo = 1
			VDR_sError = VDR_sError
			sDocId = ""
		end if
		if (l_call_pdf = 1) and (VDR_lErrNo <> 1) then
			Call CreateMessage (oMessage, 400300070)
			Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", CurrentfollowupRptID)
			Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
			Call SetXMLValueDirect (oMessage, "RPT_E2B_PREVIOUS_FOLLOWUP", PreviousfollowupRptID)
			Call SetXMLValueDirect (oMessage, "CSM_CID", iCase_id)
			Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
			Call SetXMLValueDirect (oMessage, "GN_NUMBER1", lIsJReport)
			Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)
			
			Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, VDR_lErrNo, VDR_sError)
			if VDR_lErrNo = 0 then
				sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
			else
				 VDR_lErrNo = 1
				 VDR_sError = VDR_sError
				 sDocId = ""
			end if
		end if
	end if
			  
	if(VDR_lErrNo <> 0) then
		e2b_sReturn = ConstructAjaxErrorMessageWithDoc(VDR_lErrNo, VDR_sError, sDocId)
	else
		e2b_sReturn = "<MESSAGE><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER>"
		e2b_sReturn = e2b_sReturn & "<USER_ID>" & lUserID & "</USER_ID>"
		e2b_sReturn = e2b_sReturn & "<PREVIOUS_REPORT_ID>" & PreviousfollowupRptID & "</PREVIOUS_REPORT_ID>"
		e2b_sReturn = e2b_sReturn & "<CURRENT_REPORT_ID>" & CurrentfollowupRptID & "</CURRENT_REPORT_ID>"
		e2b_sReturn = e2b_sReturn & "<CASE_ID>" & iCase_id & "</CASE_ID></MESSAGE>"
	end if    
	Response.Write e2b_sReturn
%>


