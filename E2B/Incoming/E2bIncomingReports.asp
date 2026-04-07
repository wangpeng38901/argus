<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<!-- #INCLUDE VIRTUAL="/Controls/Utilities/CSearchEngine_inc.asp" -->
<%
'******************************************************************************
' Author       : Pankaj Gautam
' Page         : E2bIncomingReports.asp
' Description  : Incoming E2B Reports Dialogue 
'******************************************************************************
' Revision History
' Date		Author		Description
' 11jul2006 Prakash    Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
	<!-- Page Title -->
	<title><%=GetTranslationData("DUPLICATE_SEARCH")%></title>
	<!-- Include Stylesheet here -->
	<link rel="stylesheet" href="/css/Relsys.css" />
	<!-- Client Library Includes Starts -->

	<script type="text/javascript" src="/js/Menu/JMenu.js"></script>

	<!-- Client Library Includes Ends -->
	<base target="_self" />
	<style type="text/css">
		td {
			float: unset !important;
		}

		TABLE.inner-table {
			outline: none;
		}
	</style>
</head>
<!-- Declaration of Page Scope variables Starts -->
<%   
Dim oOutMsg, Error, ErrorNum, oRecList, lNumRec, search, oOutCaseMsg, oRecCaseList
Dim rb_report_status, E2BViewType, notes, esm_report_id, esm_status, oRec
Dim oCountryList, oReportTypeList, oGenderList, oAgeUnitsList
Dim bPrevSort, bCurrentSort, bSortOrder, sArrow, lNextSortOrder, h_product_name
Dim h_report_type, h_receipt_date, h_generic_name, h_study_id, h_center_id
Dim h_sal, h_first_name, h_last_name, h_suffix
Dim h_country_of_incidence, h_state, h_postal_code
Dim h_patient_name, h_initial, h_pat_id, h_pat_age_and_unit,h_pat_dob
Dim h_event_desc, h_onset_date, h_gender, h_institution, h_institutionID
Dim h_reference, h_keyword, h_journal, h_title, bFirstTime
Dim sSQL, sFormName
Dim l_temp, s_temp, LockStatus, DeleteStatus, CloseStatus, OpenedUser, TrueCaseId, lOpenCasePermit, lLockCasePermit
Dim lIsChecked, oRtype, asgnSField, CustomImport, CurrRecord,Filter
Dim bFilterCriteriaAdded
Dim sESMReportIdCaseNum
Dim sNoTran
Dim lIsJReport, selectedE2BAthorityId
Dim search_oGridControl
Dim search_sHtmlTemplate
Dim search_sTableName
Dim search_sFormName
Dim search_lStartRow
Dim search_lEndRow
Dim search_lCurrentPage
Dim search_lPageSize
Dim search_lPageRowCount
Dim search_lResultsRowCount
Dim search_bPaginate
Dim E2BDupSearch_oFS, sMsgCacheFileName
Dim selectedIsApplyNewFw

lIsChecked = 0
bFirstTime = 1
rb_report_status = GetLong(GetRequest("rb_report_status"), 1)
notes = GetRequest("notes")
esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
esm_status = GetRequest("esm_status")
search = GetLong(GetRequest("search"), 0)

sFormName="Frm_E2b"    
search_sFormName = sFormName
search_bPaginate    = CInt(NVL(Request(search_sFormName & "_Paginate").Item, 0))

bPrevSort = Request.Form ("PrevSort")
If (bPrevSort = "") Then
	bPrevSort = 1
End If

bCurrentSort = Request.Form ("CurrentSort")
bSortOrder = Request.Form ("SortOrder")
h_institution = Request.Form("h_institution")
h_institutionID = Request.Form("h_institutionID")
h_product_name = Request.Form("h_product_name")
h_report_type = Request.Form("h_report_type")
h_receipt_date = Request.Form("h_receipt_date")
h_generic_name = Request.Form("h_generic_name")
h_study_id = Request.Form("h_study_id")
h_center_id = Request.Form("h_center_id")
h_sal = Request.Form("h_sal")
h_first_name = Request.Form("h_first_name")
h_last_name = Request.Form("h_last_name")
h_suffix = Request.Form("h_suffix")
h_country_of_incidence = Request.Form("h_country_of_incidence")
h_state = Request.Form("h_state")
h_postal_code = Request.Form("h_postal_code")
h_patient_name = Request.Form("h_patient_name")
h_initial = Request.Form("h_initial")
h_pat_id = Request.Form("h_pat_id")
h_pat_age_and_unit = Request.Form("h_pat_age_and_unit")
h_event_desc = Request.Form("h_event_desc")
h_onset_date = Request.Form("h_onset_date")
h_gender = Request.Form("h_gender")
h_pat_dob = Request.Form("h_pat_dob")
h_reference = Request.Form("h_reference")
h_keyword = Request.Form("h_keyword")
h_journal = Request.Form("h_journal")
h_title = Request.Form("h_title")
CurrRecord = GetLong(GetRequest("CurrRecord"), 0)
LockStatus = GetLong(GetRequest("LockStatus"), 0)
DeleteStatus = GetLong(GetRequest("DeleteStatus"), 0)
CloseStatus = GetLong(GetRequest("CloseStatus"), 0)
OpenedUser = GetRequest("OpenedUser")
TrueCaseId = GetLong(GetRequest("CaseId"), 0)
lOpenCasePermit = GetXMLValueDirect(oSession, "CFG_USERS_ALLOW_CLOSING")
lLockCasePermit = GetXMLValueDirect(oSession, "CFG_USERS_ALLOW_LOCKING")
sESMReportIdCaseNum = GetString(GetRequest("selectedCaseNum"), "")
sNoTran = GetTranslationData("WL_CMN_NO_TRANSLATION")
lIsJReport = GetLong(GetRequest("IsJReport"), 0) ' 1 if Authority Id is 4 for the report else 0
selectedE2BAthorityId = GetLong(GetRequest("AuthorityId"), 0)
selectedIsApplyNewFw = GetLong(GetRequest("ApplyNewFw"), 0)

If (bCurrentSort = "") Then
	bCurrentSort = 1
End If
If (bPrevSort <> bCurrentSort) Then
    bSortOrder = 0
End If
If (bSortOrder = "") Then
	bSortOrder = 0
End If

sSql = "select value from cmn_profile where key = 'E2B_VIEW_PRINT'"
E2BViewType = ExecuteSQLReturnStr (sSQL, ErrorNum, Error)
CustomImport = "1"
if IsNullOrEmpty(CustomImport) then CustomImport = 0

Set oRecCaseList = Nothing
If (search = 1 or search_bPaginate = 1) Then
	if GetCaseSearchResults() then
		Set oRecCaseList = oOutCaseMsg.selectnodes("/MESSAGE/CASE_DATA")
		search_lPageRowCount = oRecCaseList.length
	else
		search_lPageRowCount = 0
		Response.Write "<script type=""text/javascript"" language=""javascript"">"
		Response.Write "window.parent.MessageBoxResEx('GENERAL_ERROR'," & glDisplayLang & ",'" & GetTranslationData("E2B_DUPL_SEARCH") & "'," & JavaScriptClean(Error) & ");" 
		Response.Write "</SCRIPT>"
	end if
	InitializeSearchResultsCache()
End If

Set oRecList = Nothing
if ExecIncomingE2b() then
	Set oRecList = oOutMsg.selectnodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")
	lNumRec = oRecList.length
	Error = ""
	ErrorNum = "0"
end if

GetDropDowns()
%>
<!-- Declaration of Page Scope variables Ends -->
<!-- Assign Values to Page Scope variables Starts -->
<body onload="initForm()" style="height: 100%">
	<div id="PopupMenu" class="popup-menu" onmouseover="onSelectHighlight(event)" onmouseout="onDeselect(event)"
		display:none>
	</div>
	<form id="<%= sFormName %>" name="<%=sFormName %>" method="post" class="form-100">
		<%Call BuildHiddenControlDirect("rb_report_status", rb_report_status)  %>
		<%Call BuildHiddenControlDirect("SelectedCaseNum", sESMReportIdCaseNum) %>
		<%Call BuildHiddenControlDirect("esm_report_id", esm_report_id) %>
		<%Call BuildHiddenControlDirect("esm_status", esm_status) %>
		<%Call BuildHiddenControlDirect("search", 0) %>
		<%Call BuildHiddenControlDirect("E2bViewType", E2BViewType) %>
		<%Call BuildHiddenControlDirect("start_date_value", "") %>
		<%Call BuildHiddenControlDirect("end_date_value", "") %>
		<%Call BuildHiddenControlDirect("PrevSort", bCurrentSort)  %>
		<%Call BuildHiddenControlDirect("CurrentSort", bCurrentSort) %>
		<%Call BuildHiddenControlDirect("submit_form", "") %>
		<%Call BuildHiddenControlDirect("SortOrder", bSortOrder) %>
		<%Call BuildHiddenControlDirect("h_institution", h_institution) %>
		<%Call BuildHiddenControlDirect("h_institutionID", h_institutionID) %>
		<%Call BuildHiddenControlDirect("h_product_name", h_product_name) %>
		<%Call BuildHiddenControlDirect("h_report_type", h_report_type) %>
		<%Call BuildHiddenControlDirect("h_receipt_date", h_receipt_date) %>
		<%Call BuildHiddenControlDirect("h_generic_name", h_generic_name) %>
		<%Call BuildHiddenControlDirect("h_study_id", h_study_id) %>
		<%Call BuildHiddenControlDirect("h_center_id", h_center_id) %>
		<%Call BuildHiddenControlDirect("h_sal", h_sal) %>
		<%Call BuildHiddenControlDirect("h_first_name", h_first_name) %>
		<%Call BuildHiddenControlDirect("h_last_name", h_last_name) %>
		<%Call BuildHiddenControlDirect("h_suffix", h_suffix) %>
		<%Call BuildHiddenControlDirect("h_country_of_incidence", h_country_of_incidence) %>
		<%Call BuildHiddenControlDirect("h_state", h_state) %>
		<%Call BuildHiddenControlDirect("h_postal_code", h_postal_code) %>
		<%Call BuildHiddenControlDirect("h_patient_name", h_patient_name) %>
		<%Call BuildHiddenControlDirect("h_initial", h_initial) %>
		<%Call BuildHiddenControlDirect("h_pat_id", h_pat_id) %>
		<%Call BuildHiddenControlDirect("h_pat_age_and_unit", h_pat_age_and_unit) %>
		<%Call BuildHiddenControlDirect("h_event_desc", h_event_desc) %>
		<%Call BuildHiddenControlDirect("h_onset_date", h_onset_date) %>
		<%Call BuildHiddenControlDirect("h_gender", h_gender) %>
		<%Call BuildHiddenControlDirect("h_pat_dob", h_pat_dob) %>
		<%Call BuildHiddenControlDirect("h_reference", h_reference) %>
		<%Call BuildHiddenControlDirect("h_keyword", h_keyword) %>
		<%Call BuildHiddenControlDirect("h_journal", h_journal) %>
		<%Call BuildHiddenControlDirect("h_title", h_title) %>
		<%Call BuildHiddenControlDirect("notes", "") %>
		<%Call BuildHiddenControlDirect("bFirstTime", bFirstTime) %>
		<%Call BuildHiddenControlDirect("isJReport", lIsJReport) %>
		<%Call BuildHiddenControlDirect("AuthorityId", selectedE2BAthorityId) %>
		<%Call BuildHiddenControlDirect("selectedIsApplyNewFw", selectedIsApplyNewFw) %>
		<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
		<!-- Outer Box Starts -->
		<table cellpadding="0" cellspacing="0" class="table" style="width: 100%; height:100%;"
			border="0">
			<tr style="height: 25px">
				<td class="section-header-middle">
				<%BuildLocalLabel("DUPLICATE_SEARCH").SetStyleSheet("label label-section").Render() %>
				</td>
			</tr>
			<tr valign="top" style="height:25px">
				<td class="padding-all">
					<table class="table inner-table" style="width: 100%;" cellpadding="0" cellspacing="0">
						<!-- Header Row Starts -->
						<tr style="height: 25px;">
							<td class="section-header-left"></td>
							<td class="section-header-middle">
								<% BuildLocalLabel("REP_INFO").SetStyleSheet("label label-section").Render()%>
							</td>
							<td class="section-header-middle" align="right" valign="middle">
								<%
					BuildImage("imgArgusMinimize", "/img/common/minimize_large.gif").Style("display:block;").OnClick("fn_ToggleSectionDisplay('imgArgusMaximize', 'imgArgusMinimize', 'TR_E2bIncoming')").TabIndex(1).Render() 
					BuildImage("imgArgusMaximize", "/img/common/maximize_large.gif").Style("display:none;").OnClick("fn_ToggleSectionDisplay('imgArgusMaximize', 'imgArgusMinimize', 'TR_E2bIncoming')").TabIndex(1).Render() 
								%>
							</td>
							<td class="section-header-right"></td>
						</tr>
						<!-- Header Row Ends -->
						<!--Data Row2 Starts-->
						<tr id="TR_E2bIncoming" valign="top" style="height: 225px">
							<td colspan="4" class="no-padding-left">
								<div style="width: 100%;">
									<table class="table inner-table border" cellpadding="3" cellspacing="0" style="width: 100%;"
										border="0">
										<col width="33%" />
										<col width="33%" />
										<col />
										<tr>
											<td>
												<% BuildLocalLabel("AGENCY").Render()%>
											</td>
											<td>
												<% BuildLocalLabel("ORIGINAL_CASE_NUM").Render()%>
											</td>
											<td>
												<% BuildLocalLabel("MSG_NUM").Render()%>
											</td>
										</tr>
										<tr>
											<td>
												<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "sender_agency", "", true, 1, "").Style("width:100%").Render()%>
											</td>
											<td>
												<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "originated_case", "", true, 2, "").Style("width:100%").Render()%>
											</td>
											<td>
												<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "message_number", "", true, 3, "").Style("width:100%").Render()%>
											</td>
										</tr>
										<!--Data Row2 Starts Here-->
										<tr>
											<td colspan="3" class="no-padding-left">
												<table class="table inner-table" cellpadding="3" cellspacing="0" style="width: 100%"
													border="0">
													<col width="9%" />
													<col width="4%" />
													<col width="9%" />
													<col width="12%" />
													<col width="9%" />
													<col width="5%" />
													<col width="10%" />
													<col width="11%" />
													<col width="10%" />
													<col width="10%" />
													<col />
													<tr>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_product_name = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_product_name", lIsChecked, false, 4, GetTranslationData("PROD_NAME")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "product_name", "", false, 5, "").SetMaxLength(250).Style("width:100%").Render()%>
														</td>
														<td style="float: left;" valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_report_type = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_report_type", lIsChecked, false, 6, GetTranslationData("REP_TYPE")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<select id="report_type_id" name="report_type_id" class="ddlist" style="width: 100%" tabindex="7">
																<option value="-1"></option>
																<%if oReportTypeList.length > 0 then
										for each oRtype in oReportTypeList%>
																<option value="<%=GetXMLValueDirect(oRtype, "LM_REPORT_TYPE_RPT_TYPE_ID") %>">
																	<%=Fn_Sanitize(GetDropDownOption(oRtype, "LM_REPORT_TYPE_REPORT_TYPE"))%>
																</option>
																<%   next
									end if%>
															</select>
														</td>
														<td valign="middle">
															<%lIsChecked = "0" %>
															<%If h_receipt_date = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_receipt_date", lIsChecked, false, 8, GetTranslationData("WL_CASE_RECEIPT_DATE")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%
															BuildControlLang(CTL_TEXTBOX_PRTL_DATE, "receipt_date", "", false, 9, "",glDisplayLang).Style("width:100%").Render()
															%>
														</td>
													</tr>
													<tr>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_generic_name = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_generic_name", lIsChecked, false, 10, GetTranslationData("WL_CASE_GENERIC_NAME_PT")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "generic_name", "", false, 11, "").SetMaxLength(1000).Style("width:100%").Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_study_id = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_study_id", lIsChecked, false, 12, GetTranslationData("STD_ID")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "study_id", "", false, 13, "").Style("width:100%").Render()%>
														</td>
														<td valign="middle">
															<%lIsChecked = "0" %>
															<%If h_center_id = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_center_id", lIsChecked, false, 14, GetTranslationData("CENTER_ID")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "center_id", "", false, 15, "").Style("width:100%").Render()%>
														</td>
													</tr>
													<tr style="padding-right: 5px">
														<td colspan="11">&nbsp;
														</td>
													</tr>
													<tr style="padding-right: 5px">
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_sal = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_sal", lIsChecked, false, 16, GetTranslationData("SAL")).Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "sal", "", false, 17, "").Style("width:100%").Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_first_name = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_first_name", lIsChecked, false, 18, GetTranslationData("FIRST_NAME")).Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "first_name", "", false, 19, "").Style("width:100%").Render()%>
														</td>
														<td style="float: left;" valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_country_of_incidence = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_country_of_incidence", lIsChecked, false, 20, GetTranslationData("WL_REP_COUNTRY_OF_INC")).Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_institution = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_institution", lIsChecked, false, 21, GetTranslationData("INSTITUTION")).Render()%>
														</td>
														<td style="float: right;" valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "institution", "", false, 22, "").Style("width:100%").Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_state = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_state", lIsChecked, false, 23, GetTranslationData("STATE_PLACE")).Render()%>
														</td>
														<td style="float: left;" valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "state", "", false, 24, "").Style("width:100%").Render()%>
														</td>
													</tr>
													<tr style="padding-right: 5px">
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_suffix = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_suffix", lIsChecked, false, 25, GetTranslationData("SUFFIX")).Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "suffix", "", false, 26, "").Style("width:100%").Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_last_name = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_last_name", lIsChecked, false, 27, GetTranslationData("LAST_NAME")).Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "last_name", "", false, 28, "").Style("width:100%").Render()%>
														</td>
														<td style="float: left;" valign="middle" colspan="2">
															<select name="country_id" class="ddlist" style="width: 100%;" tabindex="29">
																<option value="-1"></option>
																<% 
								if oCountryList.length >0 then
									For each oRType in oCountryList%>
																<option value="<%=GetXMLValueDirect(oRType, "LM_COUNTRIES_COUNTRY_ID") %>">
																	<%=Fn_Sanitize(GetDropDownOption(oRtype, "LM_COUNTRIES_COUNTRY"))%>
																</option>
																<% Next
								end if%>
															</select>
														</td>
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_institutionID = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_institutionID", lIsChecked, false, 30, GetTranslationData("INSTITUTION_ID")).Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "institutionID", "", false, 31, "").Style("width:100%").Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_postal_code = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_postal_code", lIsChecked, false, 32, GetTranslationData("POSTAL_CODE")).Render()%>
														</td>
														<td style="float: left;" valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "postal_code", "", false, 33, "").Style("width:100%").Render()%>
														</td>
													</tr>
													<tr>
														<td colspan="11">&nbsp;
														</td>
													</tr>
													<tr>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_patient_name = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_patient_name", lIsChecked, false, 34, GetTranslationData("PATIENT_NAME")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "patient_name", "", false, 35, "").Style("width:100%;").Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%lIsChecked = "0" %>
															<%If h_initial = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_initial", lIsChecked, false, 36, GetTranslationData("INITIALS")).Render()%>
														</td>
														<td valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "initial", "", false, 37, "").Style("width:100%;").Render()%>
														</td>
														<td valign="middle">
															<%lIsChecked = "0" %>
															<%If h_pat_id = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_pat_id", lIsChecked, false, 38, GetTranslationData("PAT_ID")).Render()%>
														</td>
														<td style="float: left;" valign="middle">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "pat_id", "", false, 39, "").Style("width:100%;").Render()%>
														</td>
														<td valign="middle">
															<%lIsChecked = "0" %>
															<%If h_pat_age_and_unit = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_pat_age_and_unit", lIsChecked, false, 40, GetTranslationData("AGE_UNITS")).Render()%>
														</td>
														<td colspan="2">
															<table width="100%" cellpadding="0" cellspacing="0">
																<tr>
																	<td valign="middle" style="width: 15%; padding-right: 2px" align="left">
																		<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "pat_age", "", false, 41, "").Style("width:100%").Render()%>
																	</td>
																	<td valign="middle" style="width: 35%" align="left">
																		<select name="pat_age_unit_id" class="ddlist" style="width: 100%;" tabindex="42">
																			<option value="-1">&nbsp;</option>
																			<%if oAgeUnitsList.length > 0 then 
											for each oRType in oAgeUnitsList%>
																			<option value="<%=GetXMLValueDirect(oRType, "LM_AGE_UNITS_AGE_UNIT_ID")%>">
																				<%=Fn_Sanitize(GetDropDownOption(oRType, "LM_AGE_UNITS_AGE_UNIT"))%>
																			</option>
																			<%next
										end if%>
																		</select>
																	</td>
																	<td valign="middle" align="left">
																		<%lIsChecked = "0" %>
																		<%If h_pat_dob = "1" Then lIsChecked = "1"  %>
																		<%BuildControlDirect(CTL_CHECKBOX, "chk_pat_dob", lIsChecked, false, 43, GetTranslationData("PAT_DOB")).Render()%>
																	</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_event_desc = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_event_desc", lIsChecked, false, 44, GetTranslationData("EVENT_DESC")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "event_desc", "", false, 45, "").Style("width:100%;").Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_onset_date = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_onset_date", lIsChecked, false, 46, GetTranslationData("ONSET_DATE")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlLang(CTL_TEXTBOX_PRTL_DATE, "onset_date", "", false, 47, "", glDisplayLang).Style("width:100%;").Render()%>
														</td>
														<td valign="middle">
															<%lIsChecked = "" %>
															<%If h_gender = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_gender", lIsChecked, false, 48, GetTranslationData("GENDER")).Render()%>
														</td>
														<td valign="middle">
															<select name="gender_id" class="ddlist" style="width: 100%" tabindex="49">
																<option value="-1">&nbsp;</option>
																<%if oGenderList.length > 0 then
									For each oRType in oGenderList %>
																<option value="<%=GetXMLValueDirect(oRType, "LM_GENDER_GENDER_ID")%>">
																	<%=Fn_Sanitize(GetDropDownOption(oRType, "LM_GENDER_GENDER"))%>                                                                    
																</option>
																<%next
									end if%>
															</select>

														</td>
														<td valign="middle">
															<%BuildControlLang(CTL_TEXTBOX_PRTL_DATE, "pat_dob", "", false, 50, "", glDisplayLang).Style("width:100%").Render()%>                                                            
														</td>
													</tr>
													<tr>
														<td colspan="11">&nbsp;
														</td>
													</tr>
													<tr>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_reference = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_reference", lIsChecked, false, 51, GetTranslationData("REF_NUM")).Render()%>
														</td>
														<td valign="middle" colspan="6">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "reference_number", "", false, 52, "").Style("width:100%;").Render()%>
														</td>
														<td valign="middle">
															<%lIsChecked = "0" %>
															<%If h_keyword = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_keyword", lIsChecked, false, 53, GetTranslationData("KEYWORD")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "keyword", "", false, 54, "").Style("width:100%;").Render()%>
														</td>
													</tr>
													<tr>
														<td valign="middle" colspan="2">
															<%lIsChecked = "0" %>
															<%If h_journal = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_journal", lIsChecked, false, 55, GetTranslationData("JOURNAL")).Render()%>
														</td>
														<td valign="middle" colspan="6">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "journal", "", false, 56, "").Style("width:100%;").Render()%>
														</td>
														<td valign="middle">
															<%lIsChecked = "0" %>
															<%If h_title = "1" Then lIsChecked = "1"  %>
															<%BuildControlDirect(CTL_CHECKBOX, "chk_title", lIsChecked, false, 57, GetTranslationData("TITLE_CHAR")).Render()%>
														</td>
														<td valign="middle" colspan="2">
															<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "title", "", false, 58, "").Style("width:100%;").Render()%>
														</td>
													</tr>
													<tr>
														<td valign="middle" colspan="11">
															<%BuildLocalLabel("NULLIFICATION_REASON").Render() %>
														</td>
													</tr>
													<tr>
														<td valign="middle" colspan="11">
															<%BuildControlDirect(CTL_TEXTAREA, "nullification_reason", "", true, 59, "").Style("width:100%").Render()%>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr class="tblheader-gray padding-all" height="25px">
                                            <td valign="middle" colspan="3">
                                                <table cellpadding="0" cellspacing="0" style="width:100%">
                                                    <tr>
												        <td class="align-left">
													        <%BuildButton("btnSelectAll", "SELECT_ALL", 60).Style("width:90px")_
						        .OnClick("fn_SelectE2bCheckBoxes(true);").Render()%>
												        <%BuildButton("btnDeSelectAll", "DESELECT_ALL", 61).Style("width:90px")_
						        .OnClick("fn_SelectE2bCheckBoxes(false);").Render()%>
												        </td>
												        <td class="align-right">
													        <%BuildButton("Btn_AcceptInitFup", "ACCEPT_AS_FOLLOW_UP", 62).Style("width:170px")_
						        .OnClick("fn_AcceptInitFup(1);").Render()%>
												        <%BuildButton("btnSearch", "SEARCH", 63).Style("width:60px")_
						        .OnClick("fn_Search(2);").Render()%>
												        <%BuildButton("btnView", "VIEW_E2B", 64).Style("width:80px")_
						        .OnClick("fn_ViewE2b();").Render() 
												        %>
												        <%BuildButton("Btn_AcceptCase", "BTN_ACCEPT_E2B", 65).Style("width:80px")_
						        .OnClick("fn_AcceptE2BCase(1);").Render() 
												        %>
												        <%BuildButton("Btn_RejectCase", "REJECT_E2B", 66).Style("width:80px")_
						        .OnClick("fn_RejectE2BCase();").Render() 
												        %>
												        <%BuildButton("Btn_Staging_Warning", "VIEW_WARNING", 67).Style("width:90px")_
						        .OnClick("fn_Staging_Warning();").Render() 
												        %>
												        <%BuildButton("Btn_Difference", "VIEW_DIFF", 68).Style("width:90px")_
						        .OnClick("fn_AcceptInitFup(0);").Render() 
                                                        %>
												        <%BuildButton("Btn_Close", "BTN_CLOSE", 69).Style("width:60px")_
						        .OnClick("window.close();").Render() 
												        %>
												        </td>
                                                    </tr>
                                                </table>
                                            </td>
										</tr>
									</table>
								</div>
							</td>
						</tr>
						<!--Data Row2 Ends-->
					</table>
				</td>
			</tr>
			<!-----------------CASE LIST -------------->
			<%If (search = 1 or search_bPaginate = 1) Then 
			Dim ANextSortOrder, AArrowStr
			%>
			<tr valign="top">
				<td class="padding-all">
					<div id="Div_SearchResult" style="width: 100%; height: 100%;">
					<!--#INCLUDE VIRTUAL="/E2B/E2BStatus/E2B_CaseList_Inc.asp" -->
					</div>
				</td>
			</tr>
            <%Else %>
			<tr valign="top">
				<td class="padding-all">
				</td>
			</tr>
			<%End If  %>
			<!--------------------->
		</table>
	</form>
</body>
</html>

<script language="VBScript" runat="Server">
function GetDropDowns()
	Dim oMessage
	Set oMessage = RetrieveDropdown("LM_REPORT_TYPE", ErrorNum, Error)
	Set oReportTypeList = oMessage.selectNodes ("/MESSAGE/TABLE_LM_REPORT_TYPE/LM_REPORT_TYPE")
	
	Set oMessage = RetrieveDropdown("LM_GENDER", ErrorNum, Error)
	Set oGenderList =  oMessage.selectNodes ("/MESSAGE/TABLE_LM_GENDER/LM_GENDER")
	'oMessage.save "c:\temp\GenderList.xml"
	
	Set oMessage = RetrieveDropdown("LM_AGE_UNITS", ErrorNum, Error)
	Set oAgeUnitsList = oMessage.selectNodes ("/MESSAGE/TABLE_LM_AGE_UNITS/LM_AGE_UNITS")
	'oMessage.save "c:\temp\TABLE_LM_AGE_UNITS.xml"
	
	Set oMessage = RetrieveDropdown("LM_COUNTRIES", ErrorNum, Error)
	Set oCountryList = oMessage.selectNodes ("/MESSAGE/TABLE_LM_COUNTRIES/LM_COUNTRIES")
	'//oMessage.save "c:\temp\oCountryList.xml"
end function

function GetCaseSearchResults()
	Dim temp, oInMessage, sSortOrder, sOrder, sField
	Dim stemp
	Dim PatNameArray
	Dim sTextSearchResults
	Dim lDisplayLang

	search_lCurrentPage = CInt(NVL(Request(search_sFormName & "_CurrentPage").Item, 1))
	search_lPageSize    = CInt(NVL(Request(search_sFormName & "_PageSize").Item, 100))
	search_lStartRow = ((search_lCurrentPage - 1) * search_lPageSize) + 1
	search_lEndRow = search_lCurrentPage * search_lPageSize

	Call CreateMessage (oInMessage, 300100027)

	Call SetXMLValueDirect (oInMessage, "GN_DB_LIST_START", search_lStartRow)
	Call SetXMLValueDirect (oInMessage, "GN_DB_LIST_END", search_lEndRow)
	Call SetXMLValueDirect (oInMessage, "GN_DB_LIST_LENGTH", 100)
	
	'Institution
	l_temp = Request.Form("chk_institution")
	stemp = Request.Form("institution") 
	If stemp = "-1" Then stemp = ""
	If l_temp > 0 and len(stemp) > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_INSTITUTION", stemp)
	End If
	'Institution ID
	l_temp = Request.Form("chk_institutionID")
	stemp = Request.Form("institutionID") 
	If stemp = "-1" Then stemp = ""
	If l_temp > 0 and len(stemp) > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_INSTITUTION_ID", stemp)
	End If
	'Report Type ID
	l_temp = Request.Form("chk_report_type")
	stemp = Request.Form("report_type_id") 
	If stemp = "-1" Then stemp = ""
	If l_temp > 0 and len(stemp) > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSM_RPT_TYPE_ID", stemp)
	End If
	'Receipt Date
	l_temp = Request.Form("chk_receipt_date")
	stemp = Request.Form("receipt_date")
	If (lIsJReport = 1) then
		lDisplayLang = cfCMN_LANG_JP
	Else
		lDisplayLang = cfCMN_LANG_EN
	End If
	stemp = fn_ConvertToArgusDateFormat(stemp, lDisplayLang)
	if not IsValidDate(stemp) then
		stemp = ""
	end if
	If l_temp > 0 and len(stemp) > 0 Then
		Call SetXMLValueDirect (oInMessage, "GN_STATUS_STRING", stemp)
	end if

	'PRODUCT
	l_temp = Request.Form("chk_product_name")
	s_temp = Request.Form("product_name")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPD_PRODUCT_NAME", s_temp)
	End If
	'Generic Name
	l_temp = Request.Form("chk_generic_name")
	s_temp = Request.Form("generic_name")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPD_GENERIC_NAME", s_temp)
	End If
	'Country
	l_temp = Request.Form("chk_country_of_incidence")
	s_temp = Request.Form("country_id")
	If s_temp = "-1" Then s_temp = ""
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "LM_COUNTRIES_COUNTRY_ID", s_temp)
	End If
	
	'REPORTER
	l_temp = Request.Form("chk_sal")
	s_temp = Request.Form("sal")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_PREFIX", s_temp)
	End If
	l_temp = Request.Form("chk_first_name")
	s_temp = Request.Form("first_name")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_FIRST_NAME", s_temp)
	End If
	l_temp = Request.Form("chk_last_name")
	s_temp = Request.Form("last_name")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_LAST_NAME", s_temp)
	End If
	l_temp = Request.Form("chk_state")
	s_temp = Request.Form("state")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_STATE", s_temp)
	End If
	l_temp = Request.Form("chk_postal_code")
	s_temp = Request.Form("postal_code")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_POSTCODE", s_temp)
	End If
	l_temp = Request.Form("chk_suffix")
	s_temp = Request.Form("suffix")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSRP_SUFFIX", s_temp)
	End If
	'Study ID
	l_temp = Request.Form("chk_study_id")
	s_temp = Request.Form("study_id")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSS_STUDY_NUM", s_temp)
		Call SetXMLValueDirect (oInMessage, "GN_GUI_RPT_TYPE", 1)
	End If
	
	'Center ID
	l_temp = Request.Form("chk_center_id")
	s_temp = Request.Form("center_id")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSS_CENTER_NAME", s_temp)
	End If
	
	'Patient
	l_temp = Request.Form("chk_pat_id")
	s_temp = Request.Form("pat_id")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPI_PAT_SUBJ_NUM", s_temp)
	End If
	
	l_temp = Request.Form("chk_patient_name")
	s_temp = Request.Form("patient_name")
	If len(s_temp) > 0  and l_temp > 0 Then 
		PatNameArray = split(s_temp," ")
		Call SetXMLValueDirect (oInMessage, "CSPI_PAT_FIRSTNAME", trim(PatNameArray(0)))
		If UBOUND(PatNameArray) > 0  Then
			Call SetXMLValueDirect (oInMessage, "CSPI_PAT_LASTNAME", trim(PatNameArray(1)))
		End If
	End If
	
	l_temp = Request.Form("chk_initial")
	s_temp = Request.Form("initial")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPI_PAT_INITIALS", s_temp)
	End If
	
	l_temp = Request.Form("chk_pat_age_and_unit")
	s_temp = Request.Form("pat_age")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPI_PAT_AGE", s_temp)
	end if
	
	l_temp = Request.Form("chk_pat_age_and_unit")
	s_temp = Request.Form("pat_age_unit_id") 'pat_age_unit_id
	If s_temp = "-1" Then s_temp = ""
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPI_AGE_UNIT_ID", s_temp)
	end if

	l_temp = Request.Form("chk_gender")
	s_temp = Request.Form("gender_id")		'gender_id
	If s_temp = "-1" Then s_temp = ""
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPI_GENDER_ID", s_temp)
	end if
	
	'Pat Dob
	l_temp = Request.Form("chk_pat_dob")
	s_temp = Request.Form("pat_dob")		
	if not IsValidDate(s_temp) then
		s_temp = ""
	Else
		s_temp =  GetArgusDate (s_temp)
	End If	
	
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSPI_PAT_DOB_PARTIAL", s_temp)
	End If
		
	'event
	l_temp = Request.Form("chk_event_desc")	
	s_temp = Request.Form("event_desc")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSE_DESC_CODED", s_temp)
	End If
	
	l_temp = Request.Form("chk_onset_date")
	s_temp = Request.Form("onset_date")
	if not IsValidDate(s_temp) then
		s_temp = ""
	Else
		s_temp =  GetArgusDate (s_temp)		
	end if
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "GN_CODE", s_temp) 'GN_CODE
	End If
	
	l_temp = Request.Form("chk_reference")
	s_temp = Request.Form("reference_number")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSR_REF_NO", s_temp)
	End If
	
	l_temp = Request.Form("chk_keyword")
	s_temp = Request.Form("keyword")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSNAT_KEYWORDS", s_temp)
	End If
	l_temp = Request.Form("chk_journal")
	s_temp = Request.Form("journal")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSCLI_JOURNAL", s_temp)
	End If
	l_temp = Request.Form("chk_title")
	s_temp = Request.Form("title")
	If len(s_temp) > 0 and l_temp > 0 Then
		Call SetXMLValueDirect (oInMessage, "CSCLI_TITLE", s_temp)
	End If

	If (bCurrentSort = "") Then
		bCurrentSort = 1
	End If

    if bPrevSort = bCurrentSort then
        if bSortOrder = 1 then
		    sSortOrder = " desc"
	    elseif bSortOrder = 0 then
		    sSortOrder = " asc"
	    end if
    else
		sSortOrder = " asc"
	end if
	Select Case (bCurrentSort)
		Case 1 : sField = "CASE_NUM"
		Case 2 : sField = "PAT_INITIALS"
		Case 3 : sField = "PAT_SUBJ_NUM"
		Case 4 : sField = "INIT_REPT_DATE"
		Case 5 : sField = "COUNTRY_ID"
		Case 6 : sField = "STUDY_NUM"
		Case 7 : sField = "PRODUCT_NAME"
		Case 8 : sField = "DESC_CODED"
		Case 9 : sField = "PROTOCOL_NUM"
		Case 10 : sField = "REPORT_TYPE"
		Case 11 : sField = "FULL_NAME"
		Case Else: sField = "INIT_REPT_DATE"
	End Select

	sOrder = " lower(" + sField & ") " & sSortOrder
		
	Call SetXMLValueDirect (oInMessage, "GN_GUI_LM_GENERAL_TEXT", sOrder)
	
	sTextSearchResults = ApplyTextSearch(oInMessage)
	Select Case sTextSearchResults
		Case ""
		Case "10599"
			' The OracleText index is missing
			ErrorNum = 10599
			Error = "The search engine index may be rebuilding.  Please try again in a few minutes."
			GetCaseSearchResults = False
		Case Else
			Set oOutCaseMsg = Server.CreateObject("MSXML2.DOMDocument.6.0")
			oOutCaseMsg.LoadXML sTextSearchResults
			'Merge Results output message and input message
			Call WriteMessageToCache("E2BDuplicateSearchResults", oOutCaseMsg, true)

			search_lResultsRowCount = GetXMLValueDirect(oOutCaseMsg, "/MESSAGE/CASE_DATA/CNT")
			If IsNullOrEmpty(search_lResultsRowCount) Then
				search_lResultsRowCount = 0
			Else
				search_lResultsRowCount = CLng(search_lResultsRowCount)
			End If

			GetCaseSearchResults = True
	End Select
end function

Function ApplyTextSearch(oInMessage)
	On Error Resume Next
	Dim oTextSearchControl, sSearchResult, sTemp, lTemp, sStartDate, sEndDate, sOrder
	Dim lStartRow, lEndRow
	Dim lUserId, sUserName, sPassword, sDatabase
	Dim sOnsetDate, bSearch
		
	lUserId = GetLong(oArgusUser.GetUserId(), -1)
	sUserName = oArgusUser.GetDbUserName()
	sPassword = oArgusUser.GetDbUserPwd()
	sDatabase = oArgusUser.GetDbName()
	Set oTextSearchControl = CreateTextSearch("DuplicateCaseSearch", lUserId, sUserName, sPassword, sDatabase)
	
	lTemp = GetLong(GetXMLValueDirect(oInMessage, "GN_DB_LIST_LENGTH"), -1)
	If lTemp <> -1 Then
		Call oTextSearchControl.SetLimit(1000)
	End If
	
	sStartDate = GetString(GetXMLValueDirect(oInMessage, "GN_DB_DATE_START"), "")
	sEndDate = GetString(GetXMLValueDirect(oInMessage, "GN_DB_DATE_END"), "")    
	If Len(sStartDate) > 0 and Len(sEndDate) > 0 Then
		Call oTextSearchControl.SetDateRange(sStartDate, sEndDate)
	End If
	
	lStartRow = GetXMLValueDirect(oInMessage, "GN_DB_LIST_START")
	lEndRow = GetXMLValueDirect(oInMessage, "GN_DB_LIST_END")
	Call oTextSearchControl.SetPaginationRange(lStartRow, lEndRow)
	
	lTemp = GetLong(GetXMLValueDirect(oInMessage, "GN_NUMBER3"), 0)
	If lTemp > 0 Then
	  Call oTextSearchControl.ApplySoundex(true)
	End If
	
	bFilterCriteriaAdded = false
	Call AddFilterCriteria(oTextSearchControl, "Report Type", GetXMLValueDirect(oInMessage, "CSM_RPT_TYPE_ID"))
	Call AddFilterCriteria(oTextSearchControl, "Initial Receipt Date", GetXMLValueDirect(oInMessage, "GN_STATUS_STRING"))
	Call oTextSearchControl.StartZone()
	Call AddFilterCriteria(oTextSearchControl, "Product Name", GetXMLValueDirect(oInMessage, "CSPD_PRODUCT_NAME"))
	Call AddFilterCriteria(oTextSearchControl, "Generic Name", GetXMLValueDirect(oInMessage, "CSPD_GENERIC_NAME"))
	Call oTextSearchControl.EndZone()
	Call AddFilterCriteria(oTextSearchControl, "Country of Incidence", GetXMLValueDirect(oInMessage, "LM_COUNTRIES_COUNTRY_ID"))   
	Call oTextSearchControl.StartZone()
	Call AddFilterCriteria(oTextSearchControl, "Reporter Prefix", GetXMLValueDirect(oInMessage, "CSRP_PREFIX"))
	Call AddFilterCriteria(oTextSearchControl, "Reporter Suffix", GetXMLValueDirect(oInMessage, "CSRP_SUFFIX"))
	Call AddFilterCriteria(oTextSearchControl, "Reporter First Name", GetXMLValueDirect(oInMessage, "CSRP_FIRST_NAME"))
	Call AddFilterCriteria(oTextSearchControl, "Reporter Last Name", GetXMLValueDirect(oInMessage, "CSRP_LAST_NAME"))
	Call AddFilterCriteria(oTextSearchControl, "Reporter State", GetXMLValueDirect(oInMessage, "CSRP_STATE"))
	Call AddFilterCriteria(oTextSearchControl, "Reporter Postal Code", GetXMLValueDirect(oInMessage, "CSRP_POSTCODE"))
	Call AddFilterCriteria(oTextSearchControl, "Reporter Country", GetXMLValueDirect(oInMessage, "CSRP_COUNTRY_ID"))  
	Call oTextSearchControl.EndZone()    
	Call AddFilterCriteria(oTextSearchControl, "Reporter Reference Number", GetXMLValueDirect(oInMessage, "CSR_REF_NO"))
	Call AddFilterCriteria(oTextSearchControl, "Reporter Institution", GetXMLValueDirect(oInMessage, "CSRP_INSTITUTION"))  
	Call AddFilterCriteria(oTextSearchControl, "Reporter Institution ID", GetXMLValueDirect(oInMessage, "CSRP_INSTITUTION_ID"))

	Call AddFilterCriteria(oTextSearchControl, "Study Number", GetXMLValueDirect(oInMessage, "CSS_STUDY_NUM"))
	Call AddFilterCriteria(oTextSearchControl, "Center Name", GetXMLValueDirect(oInMessage, "CSS_CENTER_NAME"))
	Call AddFilterCriteria(oTextSearchControl, "Keywords", GetXMLValueDirect(oInMessage, "CSNAT_KEYWORDS"))    
	Call AddFilterCriteria(oTextSearchControl, "Patient Subject Number", GetXMLValueDirect(oInMessage, "CSPI_PAT_SUBJ_NUM"))
	
	Call AddFilterCriteria(oTextSearchControl, "Patient Name", GetXMLValueDirect(oInMessage, "CSPI_PAT_INITIALS"))
	Call AddFilterCriteria(oTextSearchControl, "Patient First Name", GetXMLValueDirect(oInMessage, "CSPI_PAT_FIRSTNAME"))
	Call AddFilterCriteria(oTextSearchControl, "Patient Last Name", GetXMLValueDirect(oInMessage, "CSPI_PAT_LASTNAME"))
	
	Call AddFilterCriteria(oTextSearchControl, "Patient DOB", GetXMLValueDirect(oInMessage, "CSPI_PAT_DOB_PARTIAL"))
	Call AddFilterCriteria(oTextSearchControl, "Patient Age", GetXMLValueDirect(oInMessage, "CSPI_PAT_AGE"))
	Call AddFilterCriteria(oTextSearchControl, "Patient Age Unit", GetXMLValueDirect(oInMessage, "CSPI_AGE_UNIT_ID"))
	Call AddFilterCriteria(oTextSearchControl, "Patient Gender", GetXMLValueDirect(oInMessage, "CSPI_GENDER_ID"))
	Call oTextSearchControl.StartZone()
	Call AddFilterCriteria(oTextSearchControl, "Description Reported", GetXMLValueDirect(oInMessage, "CSE_DESC_CODED"))
	sOnsetDate = GetString(GetXMLValueDirect(oInMessage, "GN_CODE"), "")
	If Len(sOnsetDate) > 11 Then
		sOnsetDate = Left(sOnsetDate, 11)
	End If
	Call AddFilterCriteria(oTextSearchControl, "Event Onset", sOnsetDate)
	Call oTextSearchControl.EndZone()
	Call oTextSearchControl.StartZone()
	Call AddFilterCriteria(oTextSearchControl, "Journal", GetXMLValueDirect(oInMessage, "CSCLI_JOURNAL"))
	Call AddFilterCriteria(oTextSearchControl, "Title", GetXMLValueDirect(oInMessage, "CSCLI_TITLE"))
	Call oTextSearchControl.EndZone()
	sOrder = GetString(GetXMLValueDirect(oInMessage, "GN_GUI_LM_GENERAL_TEXT"), "")

	' The ORDER BY clause is slightly different for OracleText searches
	sOrder = Replace(sOrder, "CASE_NUM", "CSM_CASE_NUM")
	sOrder = Replace(sOrder, "PAT_INITIALS", "CSPI_PAT_INITIALS")
	sOrder = Replace(sOrder, "PAT_SUBJ_NUM", "CSPI_PAT_SUBJ_NUM")
	sOrder = Replace(sOrder, "INIT_REPT_DATE", "CSM_INIT_REPT_DATE")
	sOrder = Replace(sOrder, "COUNTRY_ID", "LM_COUNTRIES_COUNTRY")
	sOrder = Replace(sOrder, "STUDY_NUM", "CSS_STUDY_NUM")
	sOrder = Replace(sOrder, "PRODUCT_NAME", "CSPD_PRODUCT_NAME")
	sOrder = Replace(sOrder, "DESC_CODED", "CSE_DESC_CODED")
	sOrder = Replace(sOrder, "PROTOCOL_NUM", "CSS_PROTOCOL_NUM")
	sOrder = Replace(sOrder, "REPORT_TYPE", "LM_REPORT_TYPE_REPORT_TYPE")
	sOrder = Replace(sOrder, "FULL_NAME", "CSRP_FIRST_NAME")
	oTextSearchControl.OrderClause sOrder

	If bFilterCriteriaAdded Then
		sSearchResult = oTextSearchControl.Search()
		If Err.Number <> 0 Then
			ErrorNum = Err.Number
			Error = Err.description
			Err.Clear
		ElseIf sSearchResult = "" Then
			ErrorNum = -1
			Error = "Error in Text Search. Please contact System Administrator."
		End If
	Else
		sSearchResult = ""
		ErrorNum = -1
		Error = GetTranslationData("E2B_DUP_SRCH_CRITERIA_MISSING")
	End If
	ApplyTextSearch = sSearchResult
End Function

Sub AddFilterCriteria(oControl, sName, sValue)
	If Not IsNullOrEmpty(sValue) Then
		oControl.AddFilter sName, sValue
		bFilterCriteriaAdded = true
	End If
End Sub

function ChangeStatusE2bReports()
	Dim oMessage, oOutStatusMsg

	Call CreateMessage (oMessage, 300100262)
	Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", esm_report_id)
	Call SetXMLValueDirect (oMessage, "GN_UI_JUSTIFICATION", notes)
	Call SetXMLValueDirect (oMessage, "GN_STATUS_NUMBER", esm_status)
	
	Set oOutStatusMsg = ServiceRequest(oArgusSvr, oMessage, ErrorNum, Error)
	
	'Fix Issue 3510
	If ErrorNum <> 0 Then
		ChangeStatusE2bReports = FALSE
		Exit Function
	End If
	ChangeStatusE2bReports = TRUE	
end function

function ExecIncomingE2b()
	Dim oMessage
	Filter = "AND S.Report_Id =" & esm_report_id
	Call CreateMessage (oMessage, 300400005)
	Call SetXMLValueDirect (oMessage, "GN_GUI_LM_GENERAL_TEXT", Filter)
	Call SetXMLValueDirect(oMessage, "GN_LANG_PREF", glDisplayLang)	
	Call SetXMLValueDirect(oMessage, "GN_NUMBER1", lIsJReport)
	Call SetXMLValueDirect(oMessage, "GN_GUI_LM_GENERAL_ID", esm_report_id)

	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, ErrorNum, Error)	
	If ErrorNum <> 0 Then
		ExecIncomingE2b = FALSE
		Exit Function
	End If
	ExecIncomingE2b = TRUE
end function

function GetDropDownOption(oRec, col)
	Dim value

	value = GetXMLValue(oRtype, col)

	if (IsNullOrEmpty(value) and glDisplayLang = 1) then
		value = GetXMLValueDirect(oRtype, col)
		value = value + sNoTran
	end if

	GetDropDownOption = value
end function

Function InitializeSearchResultsCache()
	On Error Resume Next

	search_sFormName = "Frm_E2b"    
	' Initialize Search Results Grid
	search_sHtmlTemplate = ReadTextFromFile("/E2B/E2BStatus/CaseSearchMsgRowTemplate.tpl")
	
	Set search_oGridControl = Server.CreateObject("Relsys.Argus.Interop.Web.GridControl")
	search_oGridControl.SetTemplate(search_sHtmlTemplate)
	search_oGridControl.SetDataSource GetMessageCacheFileName("E2BDuplicateSearchResults", true), "CASE_DATA"
	
	search_oGridControl.AddDateColumn "CSM_INIT_REPT_DATE", "dd-MMM-yyyy"
	
	search_oGridControl.AddPlaceholder "ROW_CLASS", "row-normal", "row-alternate"
	search_oGridControl.AddPlaceholder "ROW_COLOR", "textbox-list", "textbox-list"
	
	search_lPageRowCount = search_oGridControl.GetRowCount()
		
	If (search_lEndRow > search_lPageRowCount) Then
		search_lEndRow = search_lPageRowCount
	End If	
	InitializeSearchResultsCache = True
End Function
</script>

<script type="text/javascript">
var parameters = window.dialogArguments;
var CurrRecord;
var aData = new Array();
var aRow;
var lTotalReports = 0;
var E2bReportID = -1;
var CompanyNumb = "";
var Encrypt_CompanyNumb = "";
var Encrypt_E2bReportID = "<%=query_Crypt.Encrypt(-1, GetEncryptKey()) %>";
var CountryID;
var sDate;
var blank, init_date, event_date, res_event_date, res_init_date, dt_date_start, dt_date_end;
var ebefore, l_default_before, eafter, rbefore, rafter, dt_first, dt_last;
var bAcceptButtonPress = true;
var selectedRow = -1;
var selectedCaseID = -1;
var selectedLockStatus = -1;
var selectedDeleteStatus = -1;
var selectedCaseNum = "";
var EsmInitialReport = "";
var openCasePermit = <%=JavaScriptSanitize(lOpenCasePermit) %>;
var lockCasePermit = <%=JavaScriptSanitize(lLockCasePermit) %>;
var sTitleIncomeE2b = '<%=GetTranslationData("INCOME_E2B")%>';
var sTitleIncomeRpt = '<%=GetTranslationData("INCOME_RPT")%>';
var lIsJReport      = <%=JavaScriptSanitize(lIsJReport) %>;
var lAthorityId      = <%=JavaScriptSanitize(selectedE2BAthorityId) %>;
var glDisplayLang   = '<%=glDisplayLang %>'
var lApplyNewFW = <%=JavaScriptSanitize(selectedIsApplyNewFw) %>;

async function fn_Show_Menu(event, iRow) 
{   
	//Call first to select the function
	await f_SelectRow(iRow);
	fn_SetMenu(event, 'DUPSEARCH');
}

function fn_load_menus() {
	//Dynamic Menus
	<%
		Dim sMenuAccess
		sMenuAccess = GetXMLValueDirect (oSession, "CFG_GROUPS_MENU_ACCESS") 
	%>
	var sectionMenuName = "DUPSEARCH";
	AddContextMenuItem('<%=GetTranslationData("CASE_SUMMARY")%>', 'fn_CaseSummary();', 'casesummary', sectionMenuName, false);
	<%If (Mid (sMenuAccess, 22, 1) = "1") Then %>
		AddContextMenuItem('<%=GetTranslationData("CASE_FORM_PRINT")%>', 'fn_CFPrint();', 'caseformprint', sectionMenuName, false);
	<%End If%>
	<%If (Mid (sMenuAccess, 139, 1) = "1") Then %>	
		AddContextMenuItem('<%=GetTranslationData("MED_SUMMARY_REP")%>', 'fn_MedSumRpt();', 'medsumreport', sectionMenuName, false);
	<%End If%>
}

function fn_CaseSummary()
{
	fn_OpenCaseSummary(selectedCaseID);
}

function fn_CFPrint()
{
	var strURL = "/CaseForm/Actions/CaseformPrint.asp?case_id=" + selectedCaseID + "&multiple_cases=0&case_num=" + fn_URLEncode(selectedCaseNum) + "&DSPLYLNG=" + glDisplayLang;
	var sDialogStyle = { dialogHeight: "480", dialogWidth: "720", resizable: false, scrollable: false };    
	fn_OpenModalDialog(strURL, window,sDialogStyle);
}

function fn_MedSumRpt()
{
	fn_OpenPrintMedicalSummary(selectedCaseID);
}

async function initForm()
{
	fn_load_menus();
	var lrecord, sTemp, state_id;
	CurrRecord = 0;
	CountryID = 0;
	lrecord = -1;
	CompanyNumb = "";
    Encrypt_CompanyNumb = "";
	E2bReportID = <%=JavaScriptSanitize(esm_report_id) %>;
    Encrypt_E2bReportID = "<%=query_Crypt.Encrypt(esm_report_id, GetEncryptKey()) %>";
	var lPatAge = 0;
	
	<%If search_lResultsRowCount >= 1000 And search = 1 Then %>
	await MessageBoxRes("E2B_FIND_DUPLICATE_LIMIT");
	<%End If %>
	
	if ("<%=ErrorNum %>" == "0")
	{
	//Change tag accordingly
	fn_RepStateChange(<%=JavaScriptSanitize(rb_report_status) %>);
	
	<% If not IsNullOrEmpty(oRecList) Then
		for each oRec in oRecList %>
			aRow = new Array();
			lrecord ++;
			
			lPatAge = <%=GetXMLValueDirect(oRec, "RPT_E2B_PATIENT_AGE")%>;
			if (lPatAge == 0)
				lPatAge = "";
			aRow[0] = <%= GetXMLValueDirect(oRec, "RPT_E2B_REPORT_ID") %>;
			aRow[1] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_MEDICINAL_PRODUCT")) %>;
			aRow[2] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_MESSAGE_SENDER_ID")) %>;
			aRow[3] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_RECEIPT_DATE")) %>;
			aRow[4] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORT_TYPE")) %>;
			aRow[5] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_ACTIVESUBSTANCE_NAME")) %>;
			aRow[6] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_SPONSOR_STUDY_NUMB")) %>;
			aRow[7] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORTER_TITLE")) %>;
			aRow[8] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORTER_GIVENAME")) %>;
			aRow[9] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORTER_FAMILY_NAME")) %>;
			aRow[10] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_COUNTRY")) %>;
			aRow[11] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORTER_STATE")) %>;
			aRow[12] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORTER_POSTCODE")) %>;
			aRow[13] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_PATIENT_INITIAL")) %>;
			if (lPatAge > 0)
				aRow[14] = lPatAge;
			else
				aRow[14] = "";
			aRow[15] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_PATIENT_AGE_UNIT")) %>;
			aRow[16] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_PRIMARY_SOURCE_REACTION")) %>;
			aRow[17] = <%= JavaScriptSanitize(Replace(GetXMLValueDirect(oRec, "RPT_E2B_REACTION_START_DATE"),"x","?")) %>;
			aRow[18] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_GENDER")) %>;
			aRow[19] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_AUTHORITYNUMB")) %>;
			aRow[20] = <%= JavaScriptClean(GetXMLValueDirect(oRec, "RPT_E2B_LITERATURE_REFERENCE")) %>;
			aRow[21] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_E2B_TYPE")) %>;
			aRow[22] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_SENDER_ORGANIZATION")) %>;
			aRow[23] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_COMPANYNUMB")) %>;
			aRow[24] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_LMESSAGENUMB")) %>;
			aRow[25] = <%= JavaScriptClean(GetXMLValueDirect(oRec, "RPT_E2B_NULLIFICATION_REASON")) %>;
			aRow[26] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORT_TYPE_ID")) %>;
			aRow[27] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_COUNTRY_ID")) %>;
			aRow[28] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_PATIENT_AGE_UNIT_ID")) %>;
			aRow[29] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_PATIENT_GENDER_ID")) %>;
			aRow[30] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_PATIENT_ID")) %>;
			aRow[31] = <%= JavaScriptSanitize(Replace(GetXMLValueDirect(oRec, "RPT_E2B_DATE_FROM"),"x","?")) %>;
			aRow[32] = <%= JavaScriptClean(GetXMLValueDirect(oRec, "CSCLI_JOURNAL")) %>;
			aRow[33] = <%= JavaScriptSanitize(GetXMLValueDirect(oRec, "RPT_E2B_REPORTER_INSTITUTION")) %>;
            aRow[34] = "<%= query_Crypt.Encrypt(GetXMLValueDirect(oRec, "RPT_E2B_COMPANYNUMB"), GetEncryptKey()) %>";
            aRow[35] = "<%= query_Crypt.Encrypt(GetXMLValueDirect(oRec, "RPT_E2B_REPORT_ID"), GetEncryptKey()) %>";
			aData[lrecord] = aRow;
		<% Next 
		End If%>

		lTotalReports = lrecord;	
		CurrRecord = <%=JavaScriptSanitize(CurrRecord)%>;
		if (lTotalReports > -1)
		{
			var sFirstTime = <%=JavaScriptSanitize(Request.Form ("bFirstTime")) %>;
			//fn_DisplayCurrentRow();
			if(sFirstTime == "")
			{  
				fn_DisplayCurrentRow();
				fn_initializeCheckBoxFirstTime();
			}
			else
			{
				fn_DisplayDataFromRequest();  
			}
		}
		else
		{
			sTemp = String(CurrRecord) + " of " + String(lTotalReports + 1);
			document.Frm_E2b.product_name.value = "";
			document.Frm_E2b.receipt_date.value = "";
			document.Frm_E2b.generic_name.value = "";
			document.Frm_E2b.study_id.value = "";
			document.Frm_E2b.center_id.value = "";
			document.Frm_E2b.sal.value = "";
			document.Frm_E2b.first_name.value = "";
			document.Frm_E2b.last_name.value = "";
			document.Frm_E2b.suffix.value = "";
			document.Frm_E2b.state.value = "";
			document.Frm_E2b.postal_code.value = "";
			document.Frm_E2b.patient_name.value = "";
			document.Frm_E2b.initial.value = "";
			document.Frm_E2b.pat_id.value = "";
			document.Frm_E2b.pat_age.value = "";
			document.Frm_E2b.event_desc.value = "";
			document.Frm_E2b.onset_date.value = "";
			document.Frm_E2b.reference_number.value = "";
			document.Frm_E2b.pat_dob.value = "";
			document.Frm_E2b.keyword.value = "";
			document.Frm_E2b.journal.value = "";
			document.Frm_E2b.title.value = "";    		
			document.Frm_E2b.country_id.value = "";
			document.Frm_E2b.report_type_id.value = "";
			document.Frm_E2b.pat_age_unit_id.value = "";
			document.Frm_E2b.gender_id.value = "";
			document.Frm_E2b.sender_agency.value = "";
			document.Frm_E2b.originated_case.value = "";
			document.Frm_E2b.message_number.value = "";
			document.Frm_E2b.nullification_reason.value = "";
			document.Frm_E2b.institution.value = "";
			document.Frm_E2b.institutionID.value = "";
		}
	}
	else
	{
		await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("E2B_INCOME_RPT")%>', <%=JavaScriptClean(Error) %>);
		window.close();
		return;
	}
}

function fn_DisplayDataFromRequest()
{
	document.Frm_E2b.product_name.value = <%=JavaScriptSanitize(Request.Form("product_name"))%>;
	document.Frm_E2b.receipt_date.value = <%=JavaScriptSanitize(Request.Form("receipt_date"))%>;
	document.Frm_E2b.generic_name.value = <%=JavaScriptSanitize(Request.Form("generic_name"))%>;
	document.Frm_E2b.study_id.value = <%=JavaScriptSanitize(Request.Form("study_id"))%>;
	document.Frm_E2b.center_id.value = <%=JavaScriptSanitize(Request.Form("center_id"))%>;
	document.Frm_E2b.sal.value = <%=JavaScriptSanitize(Request.Form("sal"))%>;
	document.Frm_E2b.first_name.value = <%=JavaScriptSanitize(Request.Form("first_name"))%>;
	document.Frm_E2b.last_name.value = <%=JavaScriptSanitize(Request.Form("last_name"))%>;
	document.Frm_E2b.suffix.value = <%=JavaScriptSanitize(Request.Form("suffix"))%>;
	document.Frm_E2b.state.value = <%=JavaScriptSanitize(Request.Form("state"))%>;
	document.Frm_E2b.postal_code.value = <%=JavaScriptSanitize(Request.Form("postal_code"))%>;
	document.Frm_E2b.patient_name.value = <%=JavaScriptSanitize(Request.Form("patient_name"))%>;
	document.Frm_E2b.initial.value = <%=JavaScriptSanitize(Request.Form("initial"))%>;
	document.Frm_E2b.pat_id.value = <%=JavaScriptSanitize(Request.Form("pat_id"))%>;
	document.Frm_E2b.pat_age.value = <%=JavaScriptSanitize(Request.Form("pat_age"))%>;
	
	document.Frm_E2b.event_desc.value = <%=JavaScriptSanitize(Request.Form("event_desc"))%>;
	document.Frm_E2b.onset_date.value = <%=JavaScriptSanitize(Request.Form("onset_date"))%>;
	document.Frm_E2b.reference_number.value = <%=JavaScriptSanitize(Request.Form("reference_number"))%>;
	document.Frm_E2b.keyword.value = <%=JavaScriptSanitize(Request.Form("keyword"))%>;
	document.Frm_E2b.journal.value = <%=JavaScriptSanitize(Request.Form("journal"))%>;
	document.Frm_E2b.title.value = <%=JavaScriptSanitize(Request.Form("title"))%>;
	document.Frm_E2b.country_id.value = <%=JavaScriptSanitize(Request.Form("country_id"))%>;
	document.Frm_E2b.report_type_id.value = <%=JavaScriptSanitize(Request.Form("report_type_id"))%>;
	document.Frm_E2b.pat_age_unit_id.value = <%=JavaScriptSanitize(Request.Form("pat_age_unit_id"))%>;
	document.Frm_E2b.gender_id.value = <%=JavaScriptSanitize(Request.Form("gender_id"))%>;
	document.Frm_E2b.pat_dob.value = <%=JavaScriptSanitize(Request.Form("pat_dob"))%>;
	E2bReportID = aRow[0];
	CompanyNumb = aRow[23];
    Encrypt_CompanyNumb = aRow[34];
    Encrypt_E2bReportID = aRow[35];
	document.Frm_E2b.sender_agency.value = <%=JavaScriptSanitize(Request.Form("sender_agency"))%>;
	document.Frm_E2b.originated_case.value = <%=JavaScriptSanitize(Request.Form("originated_case"))%>;
	document.Frm_E2b.message_number.value = <%=JavaScriptSanitize(Request.Form("message_number"))%>;
	document.Frm_E2b.nullification_reason.value = <%=JavaScriptClean(Request.Form("nullification_reason"))%>;
	document.Frm_E2b.institution.value = <%=JavaScriptClean(Request.Form("institution"))%>;
	document.Frm_E2b.institutionID.value = <%=JavaScriptClean(Request.Form("institutionID"))%>;
	//fn_Calc();
}

function fn_DisplayCurrentRow()
{
	var sTemp;
	aRow = aData[CurrRecord];
	
	document.Frm_E2b.product_name.value = aRow[1];
	document.Frm_E2b.receipt_date.value = aRow[3];
	
	document.Frm_E2b.generic_name.value = aRow[5];
	document.Frm_E2b.study_id.value = aRow[6];
	document.Frm_E2b.center_id.value = "";
	
	document.Frm_E2b.sal.value = aRow[7];
	document.Frm_E2b.first_name.value = aRow[8];
	document.Frm_E2b.last_name.value = aRow[9];
	document.Frm_E2b.suffix.value = "";
	document.Frm_E2b.state.value = aRow[11];
	document.Frm_E2b.postal_code.value = aRow[12];
	
	document.Frm_E2b.patient_name.value = "";
	document.Frm_E2b.initial.value = aRow[13];
	document.Frm_E2b.pat_id.value = aRow[30];
	document.Frm_E2b.pat_age.value = aRow[14];
	
	document.Frm_E2b.event_desc.value = aRow[16];
	document.Frm_E2b.onset_date.value = aRow[17];
	
	document.Frm_E2b.reference_number.value = "";
	document.Frm_E2b.keyword.value = "";
	document.Frm_E2b.journal.value = aRow[32];
	document.Frm_E2b.title.value = aRow[20];
	
	document.Frm_E2b.country_id.value = aRow[27];
	document.Frm_E2b.report_type_id.value = aRow[26];
	document.Frm_E2b.pat_age_unit_id.value = aRow[28];
	document.Frm_E2b.gender_id.value = aRow[29];
	document.Frm_E2b.pat_dob.value = aRow[31];
	E2bReportID = aRow[0];
	CompanyNumb = aRow[23];
    Encrypt_CompanyNumb = aRow[34];
    Encrypt_E2bReportID = aRow[35];
	document.Frm_E2b.sender_agency.value = aRow[2];
	document.Frm_E2b.originated_case.value = aRow[23];
	document.Frm_E2b.message_number.value = aRow[24];
	document.Frm_E2b.nullification_reason.value = aRow[25];
	document.Frm_E2b.institution.value = aRow[33];
	fn_Calc();
}

	async function fn_SortHeader(sort_field, sort_order)
{
	document.Frm_E2b.CurrentSort.value = sort_field;
	document.Frm_E2b.SortOrder.value = sort_order;
	await fn_Search();
}

async function fn_AcceptE2BCase(acceptAsReportIs) 
{
	bAcceptButtonPress = true;
	//if AcceptE2B button has been clicked then report will be accepted on the basis of its status
	if (acceptAsReportIs == 1)
		EsmInitialReport = ""; 
	await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "esm_report_id=" + E2bReportID);
}
async function fn_LockedReport()
{
		var xmlDoc = this.req.responseXML;
		var sLocked_User;
		var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
		if (sErrStr.length > 0) {
			await MessageBoxRes("GENERAL_ERROR", sTitleIncomeE2b, sErrStr);
			return;
		}

		var asLocked_User = xmlDoc.getElementsByTagName("LOCKED_USER");
		if (asLocked_User) {
			sLocked_User = GetTextContentFromXML(asLocked_User[0]);//.text;
			if (sLocked_User.length > 0)
				await MessageBoxRes("GENERAL_INFORMATION", sTitleIncomeE2b, sLocked_User);
			else {
				if (bAcceptButtonPress)
					await fn_AcceptE2BSingleCase();
				else {
					await fn_RejectSingleE2BCase();
				}
			}
		}
		else {
			if (bAcceptButtonPress)
				await fn_AcceptE2BSingleCase();
			else {
				await fn_RejectSingleE2BCase();
			}
		}
		return;
	}
async function fn_UnLockedReport()
{
	var xmlDoc = this.req.responseXML;
	var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
	if (sErrStr.length > 0)
	{
	   await MessageBoxRes("GENERAL_ERROR",sTitleIncomeE2b,sErrStr); 
		return;
	}
	return;
}

async function fn_AcceptE2BSingleCase() 
{
	var strURL, notes, status, lWarning, AcceptedCase, justnotes, PostSave = "";
	var CaseLocked = <%=JavaScriptSanitize(LockStatus) %>;
	var CaseDelete = <%=JavaScriptSanitize(DeleteStatus) %>;
	var CaseClosed = <%=JavaScriptSanitize(CloseStatus) %>;
	var OpenedUser = <%=JavaScriptSanitize(OpenedUser) %>;
	var lCaseId = <%=JavaScriptSanitize(TrueCaseId) %>;
	var esm_report_status = <%=JavaScriptSanitize(rb_report_status) %>;
	var titleIncomingE2B = '<%=GetTranslationData("INCOME_E2B")%>';
	var Msg = "";
	var sError = "";
	if (EsmInitialReport.length > 0)
	{
		CaseLocked = selectedLockStatus;
		CaseDelete = selectedDeleteStatus;
		CaseClosed = selectedCloseStatus;
		OpenedUser = selectedOpenedUser;
		lCaseId = selectedCaseID;
	}
	if (esm_report_status == 3 || esm_report_status == 5 || esm_report_status == 6 || EsmInitialReport.length > 0) //FollowUp
	{
		if (EsmInitialReport.length > 0)
		{
			strURL = "/E2B/Actions/E2B_AcceptFollowupE2BCase.asp?IsJReport=" + lIsJReport + "&pending=1&esm_report_id=" + E2bReportID + "&esm_initial_report_id=" + EsmInitialReport + "&case_id=" + lCaseId + "&CloseStatus=" + CaseClosed + "&LockStatus=" + CaseLocked + "&ApplyNewFw=" + lApplyNewFW;
		}
		else    
		{
			strURL = "/E2B/Actions/E2B_AcceptFollowupE2BCase.asp?IsJReport=" + lIsJReport + "&pending=1&esm_report_id=" + E2bReportID + "&case_id=" + lCaseId + "&CloseStatus=" + CaseClosed + "&LockStatus=" + CaseLocked + "&ApplyNewFw=" + lApplyNewFW;
		}	
		Msg = "ERROR";
		// if the case is opened by another user show the error message
		if (OpenedUser != "-99")
			await MessageBoxRes("INCOMERPT_ALREADY_IN_USE",'',OpenedUser);
			// if the case is deleted show the error message
		else if (CaseDelete == 1)
			await MessageBoxRes("E2B_DELCASE_UPDT_FAILED");
		else if (CaseClosed == 1 && openCasePermit == "0")
			await MessageBoxRes("INCOMERPT_CLOSECASE_RESTRICTION");
		else if (CaseLocked == 1 && lockCasePermit == "0")
			await MessageBoxRes("INCOMERPT_UNLOCKCASE_RESTRICTION");
		else
			Msg = "";
	}
	else if (esm_report_status == 4) //Nullification
	{
		strURL = "/E2B/Actions/E2B_AcceptNullificationE2BCase.asp?esm_report_id=" + E2bReportID + "&case_id=" + lCaseId + "&CloseStatus=" + CaseClosed;
		Msg = "ERROR";  
		// if the case is opened by another user show the error message
		if (OpenedUser != "-99")
			await MessageBoxRes("E2BNULL_CASE_IN_USE",'',OpenedUser);
			// if the case is deleted show the error message
		else if (CaseDelete == 1)
			await MessageBoxRes("E2B_DELCASE_UPDT_FAILED");
		else if (CaseClosed == 1 && openCasePermit == "0")
			await MessageBoxRes("INCOMERPT_CLOSECASE_RESTRICTION");
		else
			Msg = "";
	}
	else //Initial
		strURL = "/E2B/Actions/E2B_AcceptE2BCase.asp?IsJReport=" + lIsJReport + "&esm_report_id=" + E2bReportID + "&receipt_date=-1" + "&product_name=-1" + "&report_type_id=-1" + "&country_id=-1" + "&ApplyNewFw=" + lApplyNewFW;

	if (Msg.length == 0){
        var strStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };
        if (esm_report_status == 4) //Nullification
            strStyle = { dialogHeight: "640", dialogWidth: "480", resizable: false, scrollable: false };
		notes = await fn_OpenModalDialog(strURL, window, strStyle);
	}
	if (!notes)
	{
		await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + E2bReportID + "&esm_initial_report_id=" + EsmInitialReport);
	}
	else
	{
		AcceptedCase = "";
		//"1~101~CaseNum~accepted~postsave"
		Pos = notes.indexOf("~",1);
		if (Pos > 0)
		{
			lWarning = notes.substring(0,Pos);
			notes = notes.substring(Pos + 1);
			Pos = notes.indexOf("~",1);
			if (Pos > 0)
			{
				status = notes.substring(0,Pos);
				notes = notes.substring(Pos + 1);
				Pos = notes.indexOf("~",1);
				if (Pos > 0)
				{
					AcceptedCase = notes.substring(0,Pos);
					notes = notes.substring(Pos + 1)
					Pos = notes.indexOf("~",1);
					if (Pos > 0)
					{
						justnotes = notes.substring(0,Pos);
						if (justnotes.length < 1)
							return;
						PostSave = notes.substring(Pos + 1);
						if (status <= 0)
						{
							status = "";
						}
					}
				}
			}
			else
				status = "";
		}
		else
			status = "";

		if ((AcceptedCase.length > 1) && (AcceptedCase != "N/A") && (status != "-1") && (status != ""))
		{
			if (esm_report_status == 3 || esm_report_status == 5 || esm_report_status == 6 || esm_report_status == 4 || EsmInitialReport.length > 0)
			{
				sError = <%=JavaScriptSanitize(GetTranslationData("CASE_UPDATED_SUCCESSFULLY"))%>;
			}
			else
			{
				sError = <%=JavaScriptSanitize(GetTranslationData("CASE_ACCEPTED_AS"))%>;
			}
			if (PostSave.length > 0)
				sError = sError + "<br>" + PostSave;
			sError = sError.replace("case_num", AcceptedCase);
			await MessageBoxRes("GENERAL_INFORMATION",titleIncomingE2B, sError);
		}
		var retVal = notes + "#||#" + status
		setWindowReturnValue(retVal);
		window.close();	
	}
}

function fn_initializeCheckBoxFirstTime()
{
	//Institution
	if (document.Frm_E2b.institution.value != "")
	{
		document.Frm_E2b.chk_institution.checked = true;
	}
	//Institution ID
	if (document.Frm_E2b.institutionID.value != "")
	{
		document.Frm_E2b.chk_institutionID.checked = true;
	}
	//Report Type ID
	if (document.Frm_E2b.report_type_id.value != "" && document.Frm_E2b.report_type_id.value != "-1")
	{
		document.Frm_E2b.chk_report_type.checked = true;
	}
	//Receipt Date
	if (document.Frm_E2b.receipt_date.value != "")
	{
		document.Frm_E2b.chk_receipt_date.checked = true;
	}
	//PRODUCT
	if (document.Frm_E2b.product_name.value != "")
	{
		document.Frm_E2b.chk_product_name.checked = true;
	}
	//Generic Name
	if (document.Frm_E2b.generic_name.value != "")
	{
		document.Frm_E2b.chk_generic_name.checked = true;
	}
	//Country
	if (document.Frm_E2b.country_id.value != "")
	{
		document.Frm_E2b.chk_country_of_incidence.checked = true;
	}
	//REPORTER
	if (document.Frm_E2b.first_name.value != "")
	{
		document.Frm_E2b.chk_first_name.checked = true;
	}
	//Last name
	if (document.Frm_E2b.last_name.value != "")
	{
		document.Frm_E2b.chk_last_name.checked = true;
	}
	//State
	if (document.Frm_E2b.state.value != "")
	{
		document.Frm_E2b.chk_state.checked = true;
	}
	//Postalcode
	if (document.Frm_E2b.postal_code.value != "")
	{
		document.Frm_E2b.chk_postal_code.checked = true;
	}
	//Study ID
	if (document.Frm_E2b.study_id.value != "")
	{
		document.Frm_E2b.chk_study_id.checked = true;
	}
	//Center ID
	if (document.Frm_E2b.center_id.value != "")
	{
		document.Frm_E2b.chk_center_id.checked = true;
	}
	//Pat Id
	if (document.Frm_E2b.pat_id.value != "")
	{
		document.Frm_E2b.chk_pat_id.checked = true;
	}
	//Pat Ini
	if (document.Frm_E2b.initial.value != "")
	{
		document.Frm_E2b.chk_initial.checked = true;
	}
	//pat age
	if (document.Frm_E2b.pat_age.value != "")
	{
		document.Frm_E2b.chk_pat_age_and_unit.checked = true;
	}
	//Gender
	if (document.Frm_E2b.gender_id.value != "" && document.Frm_E2b.gender_id.value != "-1")
	{
		document.Frm_E2b.chk_gender.checked = true;
	}
	//Pat Dob
	if (document.Frm_E2b.pat_dob.value != "")
	{
		document.Frm_E2b.chk_pat_dob.checked = true;
	}
	//event
	if (document.Frm_E2b.event_desc.value != "")
	{
		document.Frm_E2b.chk_event_desc.checked = true;
	}
	//onset date
	if (document.Frm_E2b.onset_date.value != "")
	{
		document.Frm_E2b.chk_onset_date.checked = true;
	}
	//Ref
	if (document.Frm_E2b.reference_number.value != "")
	{
		document.Frm_E2b.chk_reference.checked = true;
	} 
	//Key
	if (document.Frm_E2b.keyword.value > "")
	{
		document.Frm_E2b.chk_keyword.checked = true;
	} 
	//journal
	if (document.Frm_E2b.journal.value != "")
	{
		document.Frm_E2b.chk_journal.checked = true;
	}
	//title
	if (document.Frm_E2b.title.value != "")
	{
		document.Frm_E2b.chk_title.checked = true;
	}
	// sal 
	if (document.Frm_E2b.sal.value != "")
	{
		document.Frm_E2b.chk_sal.checked = true;
	}
}

async function fn_RejectE2BCase() 
{
	bAcceptButtonPress = false;
	await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "esm_report_id=" + E2bReportID);
}

async function fn_RejectSingleE2BCase()
{
		var strURL, notes, esm_report_status;
		esm_report_status = <%=JavaScriptSanitize(rb_report_status) %>;
		if (esm_report_status == 3 || esm_report_status == 5 || esm_report_status == 6) //FollowUp
			strURL = "/E2B/Actions/RejectFollowupE2BCase.asp?fType=1";
		else if (esm_report_status == 4) //Nullification
		{
			strURL = "/E2B/Actions/RejectFollowupE2BCase.asp?fType=2";
		}
		else  //Initial
			strURL = "/E2B/Actions/RejectE2BCase.asp";
		var sDialogStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };
		results = await fn_OpenModalDialog(strURL, window, sDialogStyle);
		if (results) {
			notes = results;
		}
		if (!notes) {
			await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + E2bReportID);
		}
		else {
			setWindowReturnValue(notes);
			document.Frm_E2b.notes.value = notes;
			window.close();
		}
	}

	function fn_ViewE2b() {
		if (E2bReportID > 0) {
			var strParam, strURL;
            var strEncodedParam;
			strParam = "ProductID:-1" + ";LicenseID:-1" + ";CountryID:-1" + ";AgencyID:-1" + ";Draft:-1" + ";TimeFrame:-1" + ";E2BViewType:-1" + ";IncomingE2b:1" + ";ClickViewReport:0" + ";AuthorityId:" + lAthorityId + ";ApplyNewFw:" + lApplyNewFW;
            strEncodedParam = fn_URLEncode(';CaseID:') + fn_URLEncode(Encrypt_CompanyNumb) + fn_URLEncode(';ReportID:') + "<%=Server.URLEncode(query_Crypt.Encrypt(-1, GetEncryptKey())) %>" + fn_URLEncode(';E2bReportID:') + fn_URLEncode(Encrypt_E2bReportID);
			strURL = "/E2B/E2bViewer/E2bViewer.asp?<%=GetRequestKeyValue()%>";
			window.open(fn_common_open(strURL, strParam, strEncodedParam));
		}
	}

function fn_RepStateChange(state_id)
{
		if (state_id == "3" || state_id == "5" || state_id == "6")	//Followup
		{
		fn_getElementByName("Btn_AcceptInitFup").disabled = true;
		fn_getElementByName("Btn_Difference").disabled = false;
		if (fn_getElementByName("Btn_Staging_Warning"))
			fn_getElementByName("Btn_Staging_Warning").disabled = false;
	} else if (state_id == "4")		//Nullification
	{
		fn_getElementByName("Btn_AcceptInitFup").disabled = true;
		fn_getElementByName("Btn_Difference").disabled = true;
		if (fn_getElementByName("Btn_Staging_Warning"))
			fn_getElementByName("Btn_Staging_Warning").disabled = true;
	} else 		//Initial
	{
		fn_getElementByName("Btn_AcceptInitFup").disabled = true;
		fn_getElementByName("Btn_Difference").disabled = true;
		if (fn_getElementByName("Btn_Staging_Warning"))
			fn_getElementByName("Btn_Staging_Warning").disabled = false;
		}
	}

function fn_Calc()
{
<%if (glDisplayLang = cfCMN_LANG_JP) then %>
	if (document.Frm_E2b.receipt_date.value && document.Frm_E2b.receipt_date.value.length > 0 )
		document.Frm_E2b.receipt_date.value = fn_ConvertDate(document.Frm_E2b.receipt_date.value, "minus", 0);
		
	if (document.Frm_E2b.onset_date.value && document.Frm_E2b.onset_date.value.length > 0)
		document.Frm_E2b.onset_date.value = fn_GetLocalDate(document.Frm_E2b.onset_date.value);
		
	if (document.Frm_E2b.pat_dob.value && document.Frm_E2b.pat_dob.value.length > 0)
		document.Frm_E2b.pat_dob.value = fn_ConvertDate(document.Frm_E2b.pat_dob.value, "minus", 0);
<%end if%>
}

function fn_calc_res(sDate)
{
	var day, month, year;
	day = sDate.substring(0,2);
	month = sDate.substring(3,6);
	year = sDate.substring(7, sDate.length);
	
	if (day == "??" && month == "???" && year == "0000")
		return -1;
		
	if (day == "??" && month == "???")
		return 5;
	
	if (day == "??")
		return 7;
}

function fn_ConvertDate(sDate, change, amount)
{
	var sDay, sMonth, sYear, newmonth;
	var date2, parsdate, newdate, day, month, year;
	sDay = sDate.substring(0,2);
	sMonth = sDate.substring(3,6);
	sYear = sDate.substring(7,11);

	if (sMonth == "???")  
		newmonth = "06";
	else	
	{
		switch (sMonth) {
			case "JAN": newmonth = "01"; break;
			case "FEB": newmonth = "02"; break;
			case "MAR": newmonth = "03"; break;
			case "APR": newmonth = "04"; break;
			case "MAY": newmonth = "05"; break;
			case "JUN": newmonth = "06"; break;
			case "JUL": newmonth = "07"; break;
			case "AUG": newmonth = "08"; break;
			case "SEP": newmonth = "09"; break;
			case "OCT": newmonth = "10"; break;
			case "NOV": newmonth = "11"; break;
			case "DEC": newmonth = "12"; break;
			default: 
			  break;
		}
	}
		
	if (sDay == "??")
		sDay = "15"

    if (parseInt(sYear) && parseInt(sDay)) {
	    if (change == "minus")
	    {
		    var date1 = new Date(sYear, newmonth-1, sDay - amount);
	    }
	    else
	    {
		    var date1 = new Date(sYear, newmonth-1, sDay - (-amount));
        }
    }

    if (date1) {
	    date2 = date1.toString();
	    parsdate = date2.split(' ');
	    if (parsdate[2].length == 1)
		    day = "0" + parsdate[2];
	    else
		    day = parsdate[2];
	    month = parsdate[1].toUpperCase();
	    year = parsdate[3]; 
	    newdate = day + "-" + month + "-" + year;
        newdate = fn_GetLocalDate(newdate); //Convert Date to Original Format
    } else
        newdate = "";
	return newdate;
}

function AddBaseTargetTag()
{
	var oHeadColl = document.all.tags('HEAD');
	var oBase=document.createElement("BASE");
	oBase.target = "_self";
	oHeadColl[0].appendChild(oBase);
}

	async function fn_Search(parm)
{
	var esm_report_status;
	var bIsCriteriaChecked;	
	bIsCriteriaChecked = fn_IsCheckBoxSelected();
	if (!bIsCriteriaChecked)
	{        
		await MessageBoxRes("E2B_DUP_SRCH_CRITERIA_MISSING");
		return;
	}
	
	showLoading();
	if (document.Frm_E2b.chk_institution.checked == true)
		document.Frm_E2b.h_institution.value = 1;
	else
		document.Frm_E2b.h_institution.value = 0;

	if (document.Frm_E2b.chk_institutionID.checked == true)
		document.Frm_E2b.h_institutionID.value = 1;
	else
		document.Frm_E2b.h_institutionID.value = 0;	

	if (document.Frm_E2b.chk_product_name.checked == true)
		document.Frm_E2b.h_product_name.value = 1;
	else
		document.Frm_E2b.h_product_name.value = 0;
		
	if (document.Frm_E2b.chk_report_type.checked == true)
		document.Frm_E2b.h_report_type.value = 1;
	else
		document.Frm_E2b.h_report_type.value = 0;
		
	if (document.Frm_E2b.chk_receipt_date.checked == true)
		document.Frm_E2b.h_receipt_date.value = 1;
	else
		document.Frm_E2b.h_receipt_date.value = 0;

	if (document.Frm_E2b.chk_generic_name.checked == true)
		document.Frm_E2b.h_generic_name.value = 1;
	else
		document.Frm_E2b.h_generic_name.value = 0;
	
	if (document.Frm_E2b.chk_study_id.checked == true)
		document.Frm_E2b.h_study_id.value = 1;
	else
		document.Frm_E2b.h_study_id.value = 0;
	
	if (document.Frm_E2b.chk_center_id.checked == true)
		document.Frm_E2b.h_center_id.value = 1;
	else
		document.Frm_E2b.h_center_id.value = 0;
		
	if (document.Frm_E2b.chk_sal.checked == true)
		document.Frm_E2b.h_sal.value = 1;
	else
		document.Frm_E2b.h_sal.value = 0;

	if (document.Frm_E2b.chk_first_name.checked == true)
		document.Frm_E2b.h_first_name.value = 1;
	else
		document.Frm_E2b.h_first_name.value = 0;

	if (document.Frm_E2b.chk_last_name.checked == true)
		document.Frm_E2b.h_last_name.value = 1;
	else
		document.Frm_E2b.h_last_name.value = 0;

	if (document.Frm_E2b.chk_suffix.checked == true)
		document.Frm_E2b.h_suffix.value = 1;
	else
		document.Frm_E2b.h_suffix.value = 0;

	if (document.Frm_E2b.chk_country_of_incidence.checked == true)
		document.Frm_E2b.h_country_of_incidence.value = 1;
	else
		document.Frm_E2b.h_country_of_incidence.value = 0;

	if (document.Frm_E2b.chk_state.checked == true)
		document.Frm_E2b.h_state.value = 1;
	else
		document.Frm_E2b.h_state.value = 0;	
			
	if (document.Frm_E2b.chk_postal_code.checked == true)
		document.Frm_E2b.h_postal_code.value = 1;
	else
		document.Frm_E2b.h_postal_code.value = 0;

	if (document.Frm_E2b.chk_patient_name.checked == true)
		document.Frm_E2b.h_patient_name.value = 1;
	else
		document.Frm_E2b.h_patient_name.value = 0;			

	if (document.Frm_E2b.chk_initial.checked == true)
		document.Frm_E2b.h_initial.value = 1;
	else
		document.Frm_E2b.h_initial.value = 0;
		
	if (document.Frm_E2b.chk_pat_id.checked == true)
		document.Frm_E2b.h_pat_id.value = 1;
	else
		document.Frm_E2b.h_pat_id.value = 0;			

	if (document.Frm_E2b.chk_pat_age_and_unit.checked == true)
		document.Frm_E2b.h_pat_age_and_unit.value = 1;
	else
		document.Frm_E2b.h_pat_age_and_unit.value = 0;

	if (document.Frm_E2b.chk_event_desc.checked == true)
		document.Frm_E2b.h_event_desc.value = 1;
	else
		document.Frm_E2b.h_event_desc.value = 0;			

	if (document.Frm_E2b.chk_onset_date.checked == true)
		document.Frm_E2b.h_onset_date.value = 1;
	else
		document.Frm_E2b.h_onset_date.value = 0;

	if (document.Frm_E2b.chk_gender.checked == true)
		document.Frm_E2b.h_gender.value = 1;
	else
		document.Frm_E2b.h_gender.value = 0;
		
	//Pat Dob
	if (document.Frm_E2b.chk_pat_dob.checked == true)
		document.Frm_E2b.h_pat_dob.value = 1;
	else
		document.Frm_E2b.h_pat_dob.value = 0;
	
	if (document.Frm_E2b.chk_reference.checked == true)
		document.Frm_E2b.h_reference.value = 1;
	else
		document.Frm_E2b.h_reference.value = 0;			

	if (document.Frm_E2b.chk_keyword.checked == true)
		document.Frm_E2b.h_keyword.value = 1;
	else
		document.Frm_E2b.h_keyword.value = 0;
		
	if (document.Frm_E2b.chk_journal.checked == true)
		document.Frm_E2b.h_journal.value = 1;
	else
		document.Frm_E2b.h_journal.value = 0;
		
	if (document.Frm_E2b.chk_title.checked == true)
		document.Frm_E2b.h_title.value = 1;
	else
		document.Frm_E2b.h_title.value = 0;	
				
	document.Frm_E2b.search.value=1;
    if (document.Frm_E2b.Frm_E2b_CurrentPage != null && parm == 2) // Only reset currentpage value on Search button click. 
        document.Frm_E2b.Frm_E2b_CurrentPage.value = 1;
	document.Frm_E2b.rb_report_status.value= <%=JavaScriptSanitize(rb_report_status) %>;
	fn_ValidateSubmitForm(document.Frm_E2b);
}	

	async function fn_Warning() {
		var esm_report_status;
		var Pos, notes, warning, status;

		esm_report_status = <%=JavaScriptSanitize(rb_report_status) %>;
		if (esm_report_status == 1) //Initial
			strURL = "/E2B/Misc/GetLoadE2bWarning.asp?check_warning=1&esm_report_id=" + E2bReportID
		else if (esm_report_status == 3 || esm_report_status == 6 || esm_report_status == 5) //FollowUp
			strURL = "/E2B/Misc/GetLoadE2bWarning.asp?check_warning=1&esm_report_id=" + E2bReportID
		else if (esm_report_status == 4) //Nullification
			strURL = "/E2B/Misc/GetLoadE2bWarning.asp?check_warning=1&esm_report_id=" + E2bReportID
		showLoading();
		var sDialogStyle = { dialogHeight: "560", dialogWidth: "840", resizable: false, scrollable: false };
		notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);
		Pos = notes.indexOf("~", 1);
		if (Pos > 0) {
			warning = notes.substring(0, Pos); //either 0 or 1
			status = notes.substring(Pos + 1); //status 102 or 101
			if (warning == 0)
				await MessageBoxRes("E2BCHECK_NO_WARNING");
		}
	}
	async function fn_Staging_Warning()
{
	var esm_report_status;
	var Pos, notes, warning, status;
	var custom_import = <%=JavaScriptSanitize(CustomImport) %>;
	esm_report_status = <%=JavaScriptSanitize(rb_report_status) %>;
	E2bReportID = <%=JavaScriptSanitize(esm_report_id) %>;
    Encrypt_E2bReportID = "<%=query_Crypt.Encrypt(esm_report_id, GetEncryptKey()) %>";
	if (custom_import == "1")
	{
		if (esm_report_status == 1) //Initial
			strURL = "/E2B/Misc/GetLoadE2bStagingWarning.asp?check_warning=1&custom_import=1&esm_report_id=" + E2bReportID
		else if (esm_report_status == 3 || esm_report_status == 5 || esm_report_status == 6) //FollowUp
			strURL = "/E2B/Misc/GetLoadE2bStagingWarning.asp?check_warning=1&custom_import=1&esm_report_id=" + E2bReportID
		else if (esm_report_status == 4) //Nullification
			strURL = "/E2B/Misc/GetLoadE2bStagingWarning.asp?check_warning=1&custom_import=1&esm_report_id=" + E2bReportID
	}
	else
	{
		if (esm_report_status == 1) //Initial
			strURL = "/E2B/Misc/GetLoadE2bWarning.asp?check_warning=1&esm_report_id=" + E2bReportID
		else if (esm_report_status == 3 || esm_report_status == 5 || esm_report_status == 6) //FollowUp
			strURL = "/E2B/Misc/GetLoadE2bWarning.asp?check_warning=1&esm_report_id=" + E2bReportID
		else if (esm_report_status == 4) //Nullification
			strURL = "/E2B/Misc/GetLoadE2bWarning.asp?check_warning=1&esm_report_id=" + E2bReportID
	}
	showLoading();
	var sDialogStyle = { dialogHeight: "560", dialogWidth: "840", resizable: false, scrollable: false };
	notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);
	if ((notes != undefined) && (notes !== sSessionTimeOutDialogReturn))
	{
		Pos = notes.indexOf("~",1);
		if (Pos > 0)
		{
			warning = notes.substring(0,Pos); //either 0 or 1
			status = notes.substring(Pos + 1); //status 102 or 101
			if (warning == 0)
			    await MessageBoxRes("E2BCHECK_NO_WARNING");
		}
	}
}
async function fn_difference()
{
	var esm_report_status;
    var lCaseId = <%=JavaScriptSanitize(TrueCaseId) %>;
	esm_report_status = <%=JavaScriptSanitize(rb_report_status) %>;
	E2bReportID = <%=JavaScriptSanitize(esm_report_id) %>;
    Encrypt_E2bReportID = "<%=query_Crypt.Encrypt(esm_report_id, GetEncryptKey()) %>";
	if (esm_report_status == 3 || esm_report_status == 5 || esm_report_status == 6 || EsmInitialReport.length > 0) //FollowUp
	{
        if (selectedCaseID > 0) {
            lCaseId = selectedCaseID;
        }
		
		strURL = "/e2b/e2bimport/Ajax_E2BDifferenceReport.asp?IsJReport=" + lIsJReport + "&report_id=" + E2bReportID + "&case_id=" + lCaseId + "&ApplyNewFw=" + lApplyNewFW ;
		if (EsmInitialReport != "" )	
			strURL += "&initialreport_id=" + EsmInitialReport;		
	}	
	else if (esm_report_status == 4) //Nullification
		strURL = "/E2B/Misc/AjaxGetFollowUpDiffData.asp";
	else
	{
		return;
	}
	showLoading();
	await loadArgusMessage(strURL, fn_ViewPDF, "esm_report_id=" + E2bReportID, fn_HandleError);
}

async function fn_ViewPDF()
{
	var xmlDoc = this.req.responseXML;
	var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
	var return_value = "0";
	var CaseLocked;
	var CaseDelete;
	var CaseClose;
	var OpenedUser;
	var sDocId = "";
	var asDocId = xmlDoc.getElementsByTagName("GN_REPORT_IDENTIFIER");
	if (asDocId && (asDocId.length > 0))
		sDocId = GetTextContentFromXML(asDocId[0]);
			
	if (sErrStr.length > 0)
	{
		if(sDocId.length > 0) 
		{
			var lResponse = await MessageBoxRes("E2B_COMPARE_NOTVALID");
			if(lResponse == MB_YES)
			{
				fn_ViewDocument(sDocId, "");
			}
		} 
		else 
		{
		    await MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("VIEW_FU_DIFF")%>',sErrStr) ;
		}
		hideLoading();
	}
	else
	{
		try
		{
		oNode = xmlDoc.getElementsByTagName("USER_ID");
		var lUserID = GetTextContentFromXML(oNode[0]);
		oNode = xmlDoc.getElementsByTagName("CASE_ID");
		var lCaseID = GetTextContentFromXML(oNode[0]);
		oNode = xmlDoc.getElementsByTagName("PREVIOUS_REPORT_ID");
		var lPrevID = GetTextContentFromXML(oNode[0]);
		oNode = xmlDoc.getElementsByTagName("CURRENT_REPORT_ID");
		var lCurrentID = GetTextContentFromXML(oNode[0]);
			if ((EsmInitialReport.length > 0) || (selectedCaseID > 0))
			{
				CaseLocked = selectedLockStatus;
				CaseDelete = selectedDeleteStatus;
				CaseClose = selectedCloseStatus;
				OpenedUser = selectedOpenedUser;
			}
			else
			{
				CaseLocked = <%=JavaScriptSanitize(LockStatus) %>;
				CaseDelete = <%=JavaScriptSanitize(DeleteStatus) %>;
				CaseClose = <%=JavaScriptSanitize(CloseStatus) %>;
				OpenedUser = <%=JavaScriptSanitize(OpenedUser) %>;
			}
			var strStyle = { dialogHeight: "960", dialogWidth: "1440", resizable: false, scrollable: false };
			var strURL = "/E2B/Incoming/E2b_ViewDifferences.asp?IsJReport=" + lIsJReport + "&case_id=" + lCaseID + "&user_id=" + lUserID + "&DocId=" + sDocId + "&current_id=" + lCurrentID + "&ApplyNewFw=" + lApplyNewFW;
			strURL += "&prev_id=" + lPrevID + "&LockStatus=" + CaseLocked + "&DeleteStatus=" + CaseDelete;
			strURL +=  "&CloseStatus=" + CaseClose + "&OpenedUser=" + fn_URLEncode(OpenedUser) + "&esm_initial_report_id=" + EsmInitialReport;
			strURL +=  "&report_case_num=" + fn_URLEncode("<%=sESMReportIdCaseNum%>") + "&selected_case_num=" + fn_URLEncode(selectedCaseNum);
			return_value = await fn_OpenModalDialog(strURL, window, strStyle);
			hideLoading();
		}
		catch(err)
		{
			fn_HandleError();
		}

		if (return_value == "1")
		{
			if (EsmInitialReport.length > 0)
				await loadArgusMessage("/E2B/Incoming/Ajax_E2BUpdateStatus.asp", fn_CheckUpdateStatusError, "status=8&esm_report_id=" + EsmInitialReport, fn_HandleError);
		}
		else if (return_value == "2") //accepted or rejected
		{
			window.close();
		}
	}
}

async function fn_CheckUpdateStatusError()
{
	var xmlDoc = this.req.responseXML;
	var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
	if (sErrStr.length > 0)
	{
	    await MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("E2B_INCOME_1")%>',sErrStr); 
	}
}

async function f_SelectRow(iRow)
{
	if (iRow != selectedRow)
	{
		selectedCaseID	 = eval("document.all.case_id_"		+ iRow).value;
		selectedCaseNum  = eval("document.all.case_num_"	+ iRow).value;
		await loadArgusMessage("/E2B/Incoming/Ajax_PreAcceptE2B.asp", fn_PreAcceptResult, "CaseId=" + selectedCaseID);
		
		f_SetRowColor(iRow, ccf_row_color_selected);
		if (selectedRow != -1)
		{
			if (selectedRow%2 == 0)
			{
				f_SetRowColor(selectedRow, "#E9E9F3");
			}
			else
			{
				f_SetRowColor(selectedRow, "#FFFFFF");
			}
		}
		selectedRow = iRow;
		var sReportStatus = <%=JavaScriptSanitize(rb_report_status) %>;
		if (sReportStatus == "1" || sReportStatus == "3" || sReportStatus == "5" || sReportStatus == "6")
		{
			fn_getElementByName("Btn_AcceptInitFup").disabled = false;
			fn_getElementByName("Btn_Difference").disabled = false;
		}
	}
}
function f_SetRowColor(iRow, colorValue)
{
	eval ("document.all.data_row_"			    + iRow).style.background = colorValue;
}

async function fn_PreAcceptResult()
{
	var xmlDoc = this.req.responseXML; 
	var sError = fn_GetAjaxErrorMsg(xmlDoc);
	
	if (sError.length > 0) {
	   await MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("BTN_ACCEPT_E2B")%>',"Error Ocurred:<br><br>" + sError); 
		return;
	}
	
	selectedLockStatus = xmlDoc.getElementsByTagName("CSM_DATE_LOCKED");
	if (GetTextContentFromXML(selectedLockStatus[0]).length > 0)
		selectedLockStatus = 1;
	else
		selectedLockStatus = 0;
	
	selectedDeleteStatus = xmlDoc.getElementsByTagName("CSM_STATE_ID");
	selectedDeleteStatus = GetTextContentFromXML(selectedDeleteStatus[0]);
	
	selectedCloseStatus = xmlDoc.getElementsByTagName("CSM_CLOSE_DATE");
	   if (GetTextContentFromXML(selectedCloseStatus[0]).length > 0)
		selectedCloseStatus = 1;
	else
		selectedCloseStatus = 0;
			
	selectedOpenedUser = xmlDoc.getElementsByTagName("CFG_USERS_USER_FULLNAME");
	  selectedOpenedUser = GetTextContentFromXML(selectedOpenedUser[0]);
}

async function fn_AcceptInitFup(pIniAsFup)
{
	var vAns = 0;
	var esm_report_status;
	var Msg = "";
	var CaseLocked = 0;
	var CaseDelete = 0;
	var CaseClosed = 0;
	var OpenedUser = "-99";
	esm_report_status = <%=JavaScriptSanitize(rb_report_status) %>;
	
	if (pIniAsFup == 1)
	{
		CaseLocked = selectedLockStatus;
		CaseDelete = selectedDeleteStatus;
		CaseClosed = selectedCloseStatus;
		OpenedUser = selectedOpenedUser;
		if (esm_report_status == 1)
		{
			vAns = await MessageBoxRes("E2BINCOME_ADD_E2BFU","",selectedCaseNum);
		}
		else if (esm_report_status == 3 || esm_report_status == 5 || esm_report_status == 6)
		{
			var sESMReportIdCaseNum = "";    		
			sESMReportIdCaseNum = <%=JavaScriptSanitize(sESMReportIdCaseNum) %>;
			if (selectedCaseNum.toLowerCase() != sESMReportIdCaseNum.toLowerCase())
				vAns = await MessageBoxRes("E2BINCOME_Q_ACCEPT_FU","",selectedCaseNum,sESMReportIdCaseNum);
			else
				vAns = MB_YES;    	    
		}
	}
	else
	{
		if (esm_report_status == 1)
			vAns = MB_YES;
		else
		{
			vAns = MB_NO;
		}
	}
	//vAns = 1 (Yes), 2 (No)
	if (vAns == MB_YES)
	{
		if (OpenedUser != "-99")
		    await MessageBoxRes("INCOMERPT_ALREADY_IN_USE","",OpenedUser);
		// if the case is deleted show the error message
		else if (CaseDelete == 1)
		    await MessageBoxRes("E2B_DELCASE_UPDT_FAILED");
		else if (CaseClosed == 1 && openCasePermit == "0")
		    await MessageBoxRes("INCOMERPT_CLOSECASE_RESTRICTION");
		else if (CaseLocked == 1 && lockCasePermit == "0")
		    await MessageBoxRes("INCOMERPT_UNLOCKCASE_RESTRICTION");
		else
		{
			var strURL = "/E2B/Incoming/Ajax_E2BUpdateFollowup.asp";
			var Param;
			if (pIniAsFup == 1)
			{
				Param = "e2b_type=3&esm_report_id=" + E2bReportID + "&accept_init_fup=" + pIniAsFup;
			}	        
			else
			{
				Param = "e2b_type=NULL&esm_report_id=" + E2bReportID + "&accept_init_fup=" + pIniAsFup;
			}
			Param = Param + "&case_id=" + selectedCaseID;
			Param = Param + "&IsJReport=" + lIsJReport;
			Param = Param + "&case_num=" + fn_URLEncode(selectedCaseNum);
			Param = Param + "&ApplyNewFw=" + lApplyNewFW;
			showLoading();
			await loadArgusMessage( strURL, fn_CheckError, Param);
		}
	}
	else if (vAns == MB_NO && pIniAsFup == 0)
	{
		await fn_difference();
	}
}

async function fn_CheckError()
{
	var xmlDoc = this.req.responseXML;
	hideLoading();

	EsmInitialReport = "";
	var sErrStr = xmlDoc.getElementsByTagName("ERROR_STRING");
	var sReportId = xmlDoc.getElementsByTagName("ESM_INITIAL_REPORT");
	var sIniAsFup = xmlDoc.getElementsByTagName("CLICK_INIT_FUP");
	var sDocId = "";

	if (sErrStr)
	{
	   sErrStr = GetTextContentFromXML(sErrStr[0]);
	}

	if (sErrStr.length > 0)
	{
		var asDocId = xmlDoc.getElementsByTagName("GN_REPORT_IDENTIFIER");
		if (asDocId && (asDocId.length > 0))
			sDocId = GetTextContentFromXML(asDocId[0]);

		if(sDocId.length > 0) 
		{
			var lResponse = await MessageBoxRes("E2B_COMPARE_NOTVALID");
			if(lResponse == MB_YES)
			{
				fn_ViewDocument(sDocId, "");
			}
		} 
		else 
		{
		    await MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("E2B_INCOME_1")%>',sErrStr);
		}
	}
	else
	{
		if (sReportId)
			EsmInitialReport = GetTextContentFromXML(sReportId[0]);
		if (sIniAsFup)
			sIniAsFup = GetTextContentFromXML(sIniAsFup[0]);
		if (sIniAsFup == "1")
		{    
			await fn_AcceptE2BCase();
		}
		else
		{                        
			await fn_difference();
			EsmInitialReport = "";            
		}
	}
}
async function fn_HandleError() {
	await MessageBoxRes("E2BVWR_ERROR");
	hideLoading();
}
function CloseSearchCaseList(caseNum)
{
	setWindowReturnValue("#||#caseNum#||#"+ caseNum); 
	window.close();    
}

function fn_SelectE2bCheckBoxes(bSelect) 
{ 
	var div = document.getElementById('TR_E2bIncoming');     
	var inputElements = div.getElementsByTagName('input');   
	for (var i = 0; i < inputElements.length; i++)
	{
		var myElement = inputElements[i];
		if (myElement.type == "checkbox") 
		{
		   var elmName = myElement.name;
		   if ((elmName.indexOf('chk_')) > -1 ) 
		   {
				myElement.checked = bSelect;
		   }
		}
	}   
} 

function fn_IsCheckBoxSelected()
{    
	var div = document.getElementById('TR_E2bIncoming');     
	var inputElements = div.getElementsByTagName('input');
	var isChecked = false; 
	for (var i = 0; i < inputElements.length; i++)
	{
		var myElement = inputElements[i];
		if (myElement.type == "checkbox") 
		{
		   var elmName = myElement.name;
		   if ((elmName.indexOf('chk_')) > -1 ) 
		   {
				if(myElement.checked)
				{
					isChecked = true;
					break;
				}
		   }
		}
	}
	return isChecked;  
}

//--></script>

<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
