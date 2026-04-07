<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
	Dim esm_report_id, e2b_type, lError, sError, sReturn, iCase_id
	Dim oMessage, AgencyID, Draft, E2BViewType, UserID, dAward_Date
	Dim oOutMsg, EsmInitialRptID, lIndex
	Dim strSQL, lCountry_Id, lProduct_ID, lLicense_Id, lProd_SeqNum, sCaseNum, lCplSeqNum
	Dim TimeFrame, sSQL, strResult, accept_init_fup, lIsJReport,Message, RecordList, Record, oTable
	Dim selectedIsApplyNewFw
	Dim sDocId

	esm_report_id = GetRequest("esm_report_id")
	e2b_type = GetRequest("e2b_type")
	iCase_id = GetRequest("case_id")
	accept_init_fup = GetLong(GetRequest("accept_init_fup"), 0)
	Draft = 1
	E2BViewType = 0
	TimeFrame = 5
	UserID = oArgusUser.GetUserId()
	lIsJReport = GetLong(GetRequest("IsJReport"), 0)
	sCaseNum = GetRequest("case_num")
	lCplSeqNum = 0
	lError = 0
	sDocId = ""
	selectedIsApplyNewFw = GetLong(Request("ApplyNewFw"), 0)
  
	Call SetParameter("P_REPORT_ID", esm_report_id, PARAM_NUMBER)
	sSQL = "select agency_id from SAFETYREPORT where report_id = :P_REPORT_ID"
	AgencyID = ExecuteSQLReturnStr(sSQL,lError, sError)
	IF lError = "0" THEN
		If (lIsJReport = 1) Then
			Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
			Call SetParameter("P_ESM_REPORT_ID", esm_report_id, PARAM_NUMBER)
			Call SetParameter("P_APPLY_NEW_FW", selectedIsApplyNewFw, PARAM_NUMBER)
			sSQL = "SELECT SF_GET_CASE_PMDA_LICENSE_SEQ(:P_CASE_ID,:P_ESM_REPORT_ID, :P_APPLY_NEW_FW) FROM DUAL"
			lCplSeqNum = GetLong(ExecuteSQLReturnStr(sSQL, lError, sError),-1)
   
			If (lCplSeqNum > 0) Then
				Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
				Call SetParameter("P_AGENCY_ID", AgencyID, PARAM_NUMBER)
				Call SetParameter("P_CPL_SEQ_NUM", lCplSeqNum, PARAM_NUMBER)
				
				strSQL = strSQL & "SELECT distinct -100 + cp.sort_id style, "
				strSQL = strSQL & " :P_AGENCY_ID agency_id, cm.country_id, nvl(cp.product_id, cp.pat_exposure) product_id,"
				strSQL = strSQL & "cpl.license_id, cp.seq_num FROM "
				strSQL = strSQL & "case_master cm, case_product cp, case_pmda_license cpl WHERE "
				strSQL = strSQL & "cm.case_id = :P_CASE_ID AND cp.case_id = cm.case_id AND "
				strSQL = strSQL & "cpl.prod_seq_num = cp.seq_num AND cpl.case_id = cp.case_id AND cpl.seq_num = :P_CPL_SEQ_NUM and "
				strSQL = strSQL & "NVL(cp.product_id, cp.pat_exposure) > 0 and rownum=1"
				
				strResult = "13750519,13710001,2110002,3510004,11810002,3510011"
			Else
				lError = 1
				sError = GetTranslationData("NO_PROD_LICENCE_FOUND_MSG")
			End If       
		Else
			Call SetParameter("P_CASE_ID", iCase_id, PARAM_NUMBER)
			strSQL = "SELECT distinct -100 + case_product.sort_id style, " &_
					"agency_id, case_master.country_id, nvl(case_product.product_id, pat_exposure),"
			strSQL = strSQL & "lm_lic_products.license_id, case_product.seq_num, lm_license.AWARD_DATE FROM "
			strSQL = strSQL & "lm_regulatory_contact, case_master, case_product, lm_lic_products, lm_license, cmn_profile WHERE "
			strSQL = strSQL & "case_master.case_id = :P_CASE_ID AND case_product.case_id (+) = case_master.case_id AND "
			strSQL = strSQL & "case_product.drug_type (+) = 1 AND lm_lic_products.PRODUCT_ID (+) = nvl(case_product.product_id, pat_exposure) AND "
			strSQL = strSQL & " NVL (case_product.product_id, pat_exposure) > 0 and "
			strSQL = strSQL & " lm_license.LICENSE_ID (+) = lm_lic_products.LICENSE_ID AND lm_license.COUNTRY_ID = case_master.COUNTRY_ID and "
			strSQL = strSQL & "agency_id = value and upper(section) = 'CASE' and upper(key) = 'DRAFT_AGENCY' union " &_
							  "select distinct 000 + case_product.sort_id style, "
			strSQL = strSQL & "agency_id, case_master.country_id, nvl(case_product.product_id, pat_exposure),lm_lic_products.license_id, case_product.seq_num, "
			strSQL = strSQL & "lm_license.AWARD_DATE from lm_regulatory_contact, case_master, case_product, lm_lic_products, lm_license "
			strSQL = strSQL & "where case_master.case_id = :P_CASE_ID AND lm_regulatory_contact.COUNTRY (+) = case_master.COUNTRY_ID AND "
			strSQL = strSQL & "case_product.case_id (+) = case_master.case_id AND case_product.drug_type (+) = 1 AND "
			strSQL = strSQL & "lm_lic_products.PRODUCT_ID (+) = nvl(case_product.product_id, pat_exposure) AND " 
			strSQL = strSQL & " NVL (case_product.product_id, pat_exposure) > 0 and "
			strSQL = strSQL & "lm_license.LICENSE_ID (+) = lm_lic_products.LICENSE_ID AND lm_license.COUNTRY_ID = case_master.COUNTRY_ID "
			strSQL = strSQL & "and agency_id is not null union " &_
							  "select distinct 100 + case_product.sort_id style, agency_id, case_master.country_id, "
			strSQL = strSQL & "nvl(case_product.product_id, pat_exposure), lm_lic_products.license_id, case_product.seq_num, lm_license.AWARD_DATE "
			strSQL = strSQL & "from lm_regulatory_contact, case_master, case_product, lm_lic_products, lm_license "
			strSQL = strSQL & "where case_master.case_id = :P_CASE_ID AND lm_regulatory_contact.COUNTRY (+) = case_master.COUNTRY_ID AND "
			strSQL = strSQL & "case_product.case_id (+) = case_master.case_id AND case_product.drug_type (+) = 1 AND "
			strSQL = strSQL & "lm_lic_products.PRODUCT_ID (+) = nvl(case_product.product_id, pat_exposure) AND"
			strSQL = strSQL & " NVL (case_product.product_id, pat_exposure) > 0 and "
			strSQL = strSQL & " lm_license.license_id = lm_lic_products.license_id and agency_id is not null union "
			strSQL = strSQL & " SELECT " 
			strSQL = strSQL & "  distinct 200 + case_product.sort_id style, "
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
		
		If lError = 0 Then
			set oMessage = ExecuteSQL(strSQL, strResult, lError, sError)
			Set oTable = oMessage.selectNodes("/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT")
			
			IF lError = "0" And oTable.length > 0 THEN
				lCountry_Id = GetXMLValueDirect(oMessage, "/MESSAGE/TABLE_LM_REGULATORY_CONTACT/LM_REGULATORY_CONTACT/CSM_COUNTRY_ID")
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
				If Len(lLicense_Id)=0 or lLicense_Id = "-1" Then 
					lLicense_Id = "NULL"
				End If

				Call CreateMessage (oMessage, 400200015)
				Call SetXMLValueDirect (oMessage, "CSM_CASE_ID", iCase_id)
				Call SetXMLValueDirect (oMessage, "LM_PRODUCT_PRODUCT_ID", lProduct_ID)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_PROD_SEQ_NUM", lProd_SeqNum)
				Call SetXMLValueDirect (oMessage, "LM_LICENSE_LICENSE_ID", lLicense_ID)
				Call SetXMLValueDirect (oMessage, "LM_MANUFACTURER_MANUFACTURER_ID", -1)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_COUNTRY_ID", lCountry_Id)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_AGENCY_ID", AgencyID)
				Call SetXMLValueDirect (oMessage, "GN_RPT_PRT_DRAFT", Draft)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_TIMEFRAME", TimeFrame)
				Call SetXMLValueDirect (oMessage, "GN_GUI_WORKLIST_VIEW_ALL", -1)
				Call SetXMLValueDirect (oMessage, "RPT_E2B_VIEW_TYPE",E2BViewType)
				Call SetXMLValueDirect (oMessage, "GN_THIS_IS_ASP",1)
				Call SetXMLValueDirect (oMessage, "CMN_REG_REPORTS_USER_ID", UserID)                
				Call SetXMLValueDirect (oMessage, "GN_UI_KANJI_FLAG", glDisplayLang)
				Call SetXMLValueDirect (oMessage, "RPT_E2B_R3_IMPORT", 1) 'Supplied to identify that e2b generation is from Accept FUP
				Call SetXMLValueDirect (oMessage, "LM_REPORT_FORMS_REPORT_FORM_ID", 27)
				Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)

				Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lError, sError)

				'We have to show Validation Error Pdf in case of EID_PMDA_VALIDATION_ERROR_FOUND
				If  lError = 402000714 or lError = 402000640 then
					sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
				End If

				IF lError = "0" THEN
					EsmInitialRptID = GetXMLValueDirect(oOutMsg, "RPT_E2B_REPORT_ID")
					Call SetParameter("E2B_TYPE", e2b_type, PARAM_NUMBER)
					Call SetParameter("P_REPORT_ID", esm_report_id, PARAM_NUMBER)
					sSQL = "UPDATE SAFETYREPORT set e2b_type_accept_as = :E2B_TYPE WHERE report_id = :P_REPORT_ID"
					Call UpdateSQL(sSQL,lError,sError)
					IF lError = "0" THEN
						Call SetParameter("P_REPORT_ID", EsmInitialRptID, PARAM_NUMBER)
                        sSQL = "UPDATE (SELECT srpt.e2b_type_accept_as e2b_type_accept_as, srpt.STATUS STATUS, srpt.reportacknowledgment.reportacknowledgmentcode reportacknowledgmentcode  "
                        sSQL = sSQL + " from SAFETYREPORT srpt WHERE srpt.report_id = :P_REPORT_ID) sr "                    
                        sSQL = sSQL + " SET sr.e2b_type_accept_as = 8, sr.STATUS = 102, sr.reportacknowledgmentcode = '01'"
						Call UpdateSQL(sSQL, lError, sError)
					END IF
				END IF
			Else
				lError = 1
				sError = GetTranslationData("NO_PROD_LICENCE_FOUND_MSG")
			END IF
		END IF
	END IF
	sReturn = "<MESSAGE><ERROR_NUM>" & lError & "</ERROR_NUM><ERROR_STRING>" & ConvertXMLSpecialChars(sError) & "</ERROR_STRING><ESM_INITIAL_REPORT>" & EsmInitialRptID & "</ESM_INITIAL_REPORT><CLICK_INIT_FUP>" & accept_init_fup & "</CLICK_INIT_FUP><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER></MESSAGE>"
	Response.Write sReturn
%>
