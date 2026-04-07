<!--#INCLUDE VIRTUAL="/Nav/AGHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<!-- #INCLUDE VIRTUAL="/E2B/Incoming/E2B_SubTab_inc.asp" -->
<!-- Page Comments Starts -->
<%
' Description  : E2B Pending
'******************************************************************************
' Author       : Pankaj Gautam
' Page         : E2BPending.asp
'******************************************************************************
' Revision History
' Date		Author	   Description
' 28AUG2006 PS         Original
'******************************************************************************
%>
<html>
<!-- Page Comments Ends -->
<!DOCTYPE html>
<!-- HTML Header Starts -->
<head>
	<!-- Page Title -->
	<title>Incoming ICSR Reports - Pending</title>
	<!-- Include Stylesheet here -->
	<link href="/CSS/relsys.css" type="text/css" rel="stylesheet" />
	<!-- Client Library Includes Starts -->
	<script type="text/javascript" src="/js/Menu/jmenu.js"></script>
	<script type="text/javascript" src="/js/E2B/E2B_Common.js"></script>
	<style type="text/css">
		TABLE.inner-table {
			table-layout: fixed;
		}

			TABLE.inner-table {
				outline: none;
			}

	</style>
	<!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
' Grid Control
Dim e2bPendingStsRpt_oGridControl
Dim e2bPendingStsRpt_sHtmlTemplate
Dim e2bPendingStsRpt_sTableName
Dim e2bPendingStsRpt_sFormName
Dim e2bPendingStsRpt_lStartRow
Dim e2bPendingStsRpt_lEndRow
Dim e2bPendingStsRpt_lCurrentPage
Dim e2bPendingStsRpt_lPageSize
Dim e2bPendingStsRpt_lPageRowCount
Dim e2bPendingStsRpt_lResultsRowCount
' General
Dim e2bPendingStsRpt_bPaginate
Dim e2bPendingStsRpt_sMessageType
DIm e2bPendingStsRpt_sAgency
Dim e2bPendingStsRpt_lRadioBtn
Dim e2bPendingStsRpt_sMessageTo
Dim e2bPendingStsRpt_sMessageFrom
Dim e2bPendingStsRpt_lErrorNum
Dim e2bPendingStsRpt_sErrorDesc
Dim e2bPendingStsRpt_oContactMsg
Dim e2bPendingStsRpt_oContactList
Dim e2bPendingStsRpt_oContact
Dim e2bPendingStsRpt_oOutMsg
Dim e2bPendingStsRpt_oRecList
Dim e2bPendingStsRpt_oRec
Dim e2bPendingStsRpt_lNumRec
Dim e2bPendingStsRpt_lCurrRec
Dim e2bPendingStsRpt_lUserId		
Dim e2bPendingStsRpt_oOutMessage
Dim e2bPendingStsRpt_oReportList
Dim e2bPendingStsRpt_oReport
Dim e2bPendingStsRpt_bViewAll
Dim e2bPendingStsRpt_bViewGroup
Dim e2bPendingStsRpt_bPrevSort
Dim e2bPendingStsRpt_bCurrentSort
Dim e2bPendingStsRpt_bSortOrder
Dim e2bPendingStsRpt_sSortOrder
Dim e2bPendingStsRpt_sOrder
Dim e2bPendingStsRpt_sArrow
Dim e2bPendingStsRpt_lNextSortOrder
Dim e2bPendingStsRpt_sField
Dim e2bPendingStsRpt_lOffSet
Dim e2bPendingStsRpt_lGmtOffSet
Dim e2bPendingStsRpt_oDateRangeMsg
Dim e2bPendingStsRpt_oDateRangeList
Dim e2bPendingStsRpt_lRangeList
Dim e2bPendingStsRpt_oDate
Dim e2bPendingStsRpt_sDateFrom
Dim e2bPendingStsRpt_sDateTo
Dim e2bPendingStsRpt_bSearch
Dim e2bPendingStsRpt_lMessageId
Dim e2bPendingStsRpt_sOpenCase
Dim e2bPendingStsRpt_TradingPrtnr                'Trading Partner
Dim e2bPendingStsRpt_ProdName                    'Product Name
Dim e2bPendingStsRpt_GenericName                 'Generic Name      
Dim e2bPendingStsRpt_Status                      'Report Status        
Dim e2bPendingStsRpt_MsgType                     'Message Type 
Dim e2bPendingStsRpt_HdnMsgType_Desc       
Dim e2bPendingStsRpt_RptType                     'Report Type     
Dim e2bPendingStsRpt_DateType                    'Date Type     
Dim e2bPendingStsRpt_Year                        'Year
Dim e2bPendingStsRpt_Month                       'Month       
Dim e2bPendingStsRpt_Day                         'Year
Dim e2bPendingStsRpt_FormatedDate                'Date Format 
Dim e2bTransStsRpt_lGmtOffSet     
Dim ReportingAgencyKey
Dim ReportingAgencyValue
Dim oMessageIn, Filter
Dim sFormName
Dim sRequestXML
Dim sResponseXML	
Dim E2BViewType, sSql
Dim notes,rb_report_status,esm_report_id,esm_status,esm_report_id_splitted, iArray
Dim hdn_Product,hdn_Generic,e2bPendingStsRpt_oProdGenRadio
Dim e2bPendingStsRpt_oStatusList, e2bPendingStsRpt_oStatusMsg, oMessageCriteria
Dim lError, sError, custom_import, lOpenCasePermit, lLockCasePermit
Dim lmDateRangeFieldName
Dim oMessageDdl, oMessageType, sMTSQL
Dim sUserFullName
Dim sCopyDateFrom
Dim sCopyDateTo
Dim sProdName
Dim sDataType
Dim sAgencyList
Dim sReportType
Dim sStatus    
Dim lCacheId
%>

<!-- Declaration of Page Scope variables Ends -->
<!-- Assign Values to Page Scope variables Starts -->
<!-- Assign Values to Page Scope variables Ends -->
<!-- Page Processing Starts -->
<%
	' Initialize Page Processing
	rb_report_status = GetLong(GetRequest("rb_report_status"), 1)
	notes = GetRequest("notes")
	esm_report_id = cfCmn_FindRegEx(GetRequest("esm_report_id"), "0-9,")
	esm_status = GetLong(GetRequest("esm_status"), 0)
	custom_import = "1"
	lOpenCasePermit = GetXMLValueDirect(oSession, "CFG_USERS_ALLOW_CLOSING")
	lLockCasePermit = GetXMLValueDirect(oSession, "CFG_USERS_ALLOW_LOCKING")
	
	'If multiple rows are selected then we have to split the string    
	If Len(rb_report_status) = 0 Then
		rb_report_status = 1
	Else
		If notes <> "" Then
			If esm_status > 0 then
				esm_report_id_splitted = Split(esm_report_id,",")
				For iArray=0 to UBound(esm_report_id_splitted)
					esm_report_id = esm_report_id_splitted(iArray)
					If ChangeStatusE2bReports() = False Then
						Response.Write "<script type='text/javascript'>"
						Response.Write "MessageBoxResEx('GENERAL_WARNING'," & glDisplayLang & ",'" & GetTranslationData("E2B_REJECT")& "','" & Fn_Sanitize(e2bPendingStsRpt_sErrorDesc) & "');"
						Response.Write "</SCRIPT>"
					End If
				Next
			End If
		End If
	End If
 
	e2bPendingStsRpt_Initialize()
	E2BViewType = GetValueFromCMN_PROFILEOnKey ("E2B_VIEW_PRINT")
	If E2BViewType = "" Then E2BViewType = "0"

	'CFG_MESSAGE_TYPE
	sMTSQL = "select id, root_element description, root_element description_j from cfg_message_type where display = 1 order by id"
	Set oMessageDdl = ExecuteSQL (sMTSQL, "81810001, 81850006, 81850306", lError, lError)
	Set oMessageType = oMessageDdl.selectNodes ("/MESSAGE/TABLE_CFG_MESSAGE_TYPE/CFG_MESSAGE_TYPE")
%>
<!-- Page Display Starts -->
<body onload="Init_Frm();">
	<div id="PopupMenu" class="popup-menu" onmouseover="onSelectHighlight(event)" onmouseout="onDeselect(event)"
		display:none>
	</div>
	<!-- Case Title Starts -->
	<table class="table" cellspacing="0" cellpadding="0" width="100%" height="100%" border="0">
		<tr style="height: 25px;">
			<td>
				<!--#INCLUDE VIRTUAL="/Nav/AGToolbar_inc.asp" -->
			</td>
		</tr>
		<tr>
			<td class="valign-top">
				<form id="<%= e2bPendingStsRpt_sFormName %>" name="<%= e2bPendingStsRpt_sFormName %>"
					method="post" action="/E2B/Incoming/E2BPending.asp" class="form-100">
					<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
					<table class="table" cellspacing="0" cellpadding="0" style="width: 100%; height: 100%;">
						<tr style="height: 25px;">
							<td class="padding-left-right">
								<%BuildLocalLabel ("INCOMING_E2B_REPORTS").SetStyleSheet("label label-page-header").Render()%>
							</td>
						</tr>
						<tr>
							<td valign="top">
								<table class="table" cellspacing="0" cellpadding="0" id="Table_Outer_Box"
									style="width: 100%; height: 100%">
									<tr valign="top" style="height: 75px;">
										<td>
											<%Call BuildHiddenControlDirect("E2bViewType", E2BViewType) %>
											<%Call BuildHiddenControlDirect("OffSet", "") %>
											<%Call BuildHiddenControlDirect("GmtOffSet", e2bPendingStsRpt_lGmtOffSet)  %>
											<%Call BuildHiddenControlDirect("PrevSort", e2bPendingStsRpt_bCurrentSort) %>
											<%Call BuildHiddenControlDirect("CurrentSort", e2bPendingStsRpt_bCurrentSort) %>
											<%Call BuildHiddenControlDirect("SortOrder", e2bPendingStsRpt_bSortOrder) %>
											<%Call BuildHiddenControlDirect("sOrder", e2bPendingStsRpt_sOrder) %>
											<%Call BuildHiddenControlDirect("bSearch", "") %>
											<%Call BuildHiddenControlDirect("notes", "") %>
											<%Call BuildHiddenControlDirect("hdn_Message_type_Desc", "ichicsr") %>
											<%Call BuildHiddenControlDirect("UserCacheId", lCacheId) %>
											<table class="table" cellspacing="0" cellpadding="0" width="100%" id="Table_Tab_Box">
												<!-- Section Header - General Starts -->
												<tr>
													<td valign="top">
														<% gSecHead_sLabel = GetTranslationData("SEARCH_CRT")%>
														<!-- #INCLUDE VIRTUAL="/Common/SectionHeader_inc.asp" -->
													</td>
												</tr>
												<!-- Section Header - General Ends -->
												<!-- Section Contents - General Starts -->
												<tr class="row">
													<td valign="top">
														<table class="table border-blue" width="100%" id="Table_Section1_Cont" style="background-color: White">
															<tr style="height: 50px">
																<td width="30%" valign="top">
																	<table cellpadding="0" cellspacing="0" width="100%">
																		<tr>
																			<td>
																				<%BuildLocalLabel("TRADING_PARTNER").Render() %>
																			</td>
																			<td style="float: right;">
																				<%BuildControlDirect(CTL_BUTTON,"btnFilter",GetTranslationData("REP_FILTER"),false,0,"").OnClick("fn_addDestination();").Style("Width:60px").Render()
																			  BuildControlDirect(CTL_BUTTON,"btnClear",GetTranslationData("REMOVE_ALL"),false,0,"").OnClick("fn_ClearAll();").Style("Width:70px").Render() %>
																			</td>
																		</tr>
																		<tr>
																			<td colspan="2">
																				<select name="ReportingAgency" class='ddlist-multi' size='5' tabindex='1' style='width: 100%; overflow: auto;'
																					multiple='multiple'>
																					<%If instr(ReportingAgencyKey,",") > 0 Then
																				Dim ary, x, ArrVaue
																				ary = split(ReportingAgencyKey, ",")
																				ArrVaue = split(ReportingAgencyValue, "--*--")
																				For x=0 to ubound(ary)
																					Response.Write("<option value='" & Fn_Sanitize(ary(x)) & "'>" &  Fn_Sanitize(ArrVaue(x)) & "</option>")
																				Next
																			 Else
																				response.Write("<option value='" & Fn_Sanitize(ReportingAgencyKey) & "'>" & Fn_Sanitize(ReportingAgencyValue) & "</option>")   
																			 End If %>
																				</select>
																				<%Call BuildHiddenControlDirect("ReportingAgencyKey", ReportingAgencyKey)
																			  Call BuildHiddenControlDirect("ReportingAgencyValue", ReportingAgencyValue) %>
																			</td>
																		</tr>
																	</table>
																</td>
																<td>
																	<table class="table" width="100%" cellpadding="2" cellspacing="0" style="table-layout: fixed" border="0">
																		<col width="40%" />
																		<col width="20%" />
																		<col width="40%" />
																		<tr>
																			<td>
																				<% BuildControlDirect (CTL_RADIOBUTTON, "chk_product_name", e2bPendingStsRpt_oProdGenRadio, false, 0, _
																					 GetTranslationData("PROD_NAME") & ":0;" & GetTranslationData("GENERIC_NAME") & ":1") _
																					 .SetStyleSheet("label label-no-padding") _
																					 .onClick("document.all.selProduct.value = ''") _
																					 .Render() %>
																			</td>
																			<td>
																				<%BuildLocalLabel("MSG_TYPE").Render() %>
																			</td>
																			<td>
																				<span class="label label-no-padding"><%BuildLocalLabel("REP_TYPE").Render()%></span>
																			</td>
																		</tr>
																		<tr>
																			<td>
																				<input type="text" style="width: 80%" class="textbox" id="selProduct" name="selProduct"
																					value="" tabindex="3" maxlength="2000">
																				<% BuildControlDirect(CTL_BUTTON, "btn_select", GetTranslationData("SELECT"), false, 4, "") _
																					   .OnClick("e2b_select();") _
																					   .Render() %>
																			</td>
																			<td>
																				<%BuildListDirectFromXML("selectMessageType", "1", false, 4, oMessageType, "CFG_MESSAGE_TYPE_ID", "CFG_MESSAGE_TYPE_DESCRIPTION").AddTopOption("All:-1")_
																					.Style("width:100%").OnChange("fn_MessageTypeChange()").Render() %>
																			</td>
																			<td>
																				<select name="selRType" id="selRType" class="ddlist" style="width: 100%">
																					<option value="-1"><%=GetTranslationData("ALL")%></option>
																					<option value="1"><%=GetTranslationData("INITIAL")%> </option>
																					<option value="3"><%=GetTranslationData("E2B_RPT_TYPE_FOLLOW_UP")%></option>
																					<option value="4"><%=GetTranslationData("E2B_RPT_TYPE_NULLIFICATION")%></option>
																					<option value="5"><%=GetTranslationData("WL_REP_DOWNGRADE")%></option>
																					<option value="7"><%=GetTranslationData("URGENT_REPORT")%></option>
																				</select>
																			</td>
																		</tr>
																		<tr>
																			<td colspan="3">
																				<table class="table" width="100%" cellpadding="2" cellspacing="0" style="table-layout: fixed" border="0">
																					<col width="20%" />
																					<col width="20%" />
																					<col width="20%" />
																					<col width="20%" />
																					<col width="20%" />
																					<tr>
																						<td>
																							<%BuildLocalLabel("DATE_RANGE").Render() %>
																						</td>
																						<td>
																							<%BuildLocalLabel("DATEFROM").Render() %>
																						</td>
																						<td>
																							<%BuildLocalLabel("DATETO").Render() %>
																						</td>
																						<td>
																							<%BuildLocalLabel("RANGE").Render() %>
																						</td>
																					</tr>
																					<tr>
																						<td>
																							<select style="width: 100%" class="ddlist" name="DateType" tabindex="5">
																								<option value="to_date(m.messageheader.messagedate,'yyyy-mm-dd hh24:mi:ss')"><%=GetTranslationData("DATE_RANGE_TRANS_DATE")%></option>
																								<option value="to_date(e.edi_complete_date,'yyyy-mm-dd')"><%=GetTranslationData("DATE_RANGE_MDN_SENT_DATE")%></option>
																								<option value="to_date(s.RECEIPTDATE,'yyyy-mm-dd')"><%=GetTranslationData("DATE_RANGE_CASE_RCPT_DATE")%></option>
																								<option value="to_date(m.DATE_RECEIVED + (<%=e2bTransStsRpt_lGmtOffSet%>/24))"><%=GetTranslationData("DATE_RANGE_INTERCHNG_PROC_DATE")%></option>
																							</select>
																						</td>
																						<td>
																							<input type="text" style="width: 100%" class="textbox" id="DateFrom" name="DateFrom"
																								value="<%=Fn_Sanitize(e2bPendingStsRpt_sDateFrom) %>" onchange="DateExit(this, false, false)"
																								onblur="DateBlur(this, false, false)" onkeydown="document.Frm_E2B.RangeList.value = -1;"
																								tabindex="6" />
																						</td>
																						<td>
																							<input type="text" style="width: 100%" class="textbox" id="DateTo" name="DateTo"
																								value="<%=Fn_Sanitize(e2bPendingStsRpt_sDateTo) %>" onchange="DateExit(this, false, false)"
																								onblur="DateBlur(this, false, false)" onkeydown="document.Frm_E2B.RangeList.value = -1;"
																								tabindex="7" />
																						</td>
																						<td>
																							<select style="width: 100%" class="ddlist" name="RangeList" onchange="f_e2bPendingStsRpt_ChangeDateRange(this.value);"
																								tabindex="5">
																								<%For Each e2bPendingStsRpt_oDate in e2bPendingStsRpt_oDateRangeList%>
																								<option value="<%=GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID")%>"
																									<%If e2bPendingStsRpt_lRangeList = GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID") then%>Selected<%End If%>>
																									<%=Fn_Sanitize(GetXMLValueDirect(e2bPendingStsRpt_oDate,lmDateRangeFieldName))%></option>
																								<%Next%>
																								<option value="-1" <%If "-1" = e2bPendingStsRpt_lRangeList Then%>Selected<%End If%>>
																									<%=GetTranslationData("CUST_DT_RANGE")%></option>
																							</select>
																						</td>
																						<td>
																							<%BuildButton("btnRetrieve", "BULK_RPT_RETRIEVE", 9).Style("width:60px")_
																												.OnClick("fn_Search(2);").Render()%>
																						</td>
																					</tr>
																				</table>
																			</td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>
													</td>
												</tr>
											</table>
										</td>
									</tr>
                                    <tr class="spacer""><td></td></tr>
									<tr style="height: 20px;">
										<!-- Pending/Prpcessed Sub Tab Section Starts -->
										<td>
											<%
											Dim cfPendingProcessedTabs
											Set cfPendingProcessedTabs = BuildTabs ("CF_PENDING", 2, CONST_TABS_TOP, 1)
											cfPendingProcessedTabs.AddTab GetTranslationData("BULK_RPT_PENDING"), "f_CFN_SwitchTab(1)"
											cfPendingProcessedTabs.AddTab GetTranslationData("PROCESSED"), "f_CFN_SwitchTab(2)"
											cfPendingProcessedTabs.SetTabWidth "150px"
											cfPendingProcessedTabs.Render
											%>
										</td>
										<!-- Pending/Processed Sub Tab Section Ends -->
									</tr>
									<tr>
										<td valign="top">
											<table class="table" cellspacing="0" cellpadding="0" width="100%" height="100%" id="Table3"
												border="0" valign="top">
												<!-- Section Header - Open Cases Search Results Starts -->
												<tr style="height: 25px;">
													<td style="margin-top: 0px; margin-bottom: 0px" valign="top">
														<% 
														gSecHead_sGridName = e2bPendingStsRpt_sFormName
														gSecHead_lGridRowCount = e2bPendingStsRpt_lResultsRowCount
														%>
														<!-- #INCLUDE VIRTUAL="/Common/SectionHeaderGrid_inc.asp" -->
													</td>
												</tr>
												<!-- Section Header - Open Cases Search Results Ends -->
												<!-- Section Column Headers - Open Cases Search Results Starts -->
												<tr>
													<td class="no-padding-left">
														<table class="table" cellspacing="0" cellpadding="0" style="width: 100%; height: 100%">
															<thead>
																<tr style="height: 60px">
																	<td>
																		<div id="Div1" class="table-scroll-hide" style="width: 100%; overflow-y: scroll;">
																			<table class="grd-table-header" cellspacing="0" cellpadding="2">
																				<col width="5%" />
																				<col width="15%" />
																				<col width="12%" />
																				<col width="18%" />
																				<col width="20%" />
																				<col width="20%" />
																				<col />
																				<tr class="tblheader-lightblue" style="height: 60px">
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 50%" align="right">
																									<input type="checkbox" onclick="f_e2bPendingStsRpt_SelectUnselectAll(this)" />
																								</td>
																								<td style="width: 50%" align="left">
																									&nbsp;
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(1, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("TRADING_PARTNER").Render() %> </span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 1) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(2, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("INITIAL_FU_NULLIFICATION").Render() %> </span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 2) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(3, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("WORLD_WIDE_UNIQUE_NUM").Render() %> </span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 3) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(4, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("SENDER_CASE_NUMBER").Render() %> </span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 4) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(5, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("TRANSMISSION_SENT").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 5) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(6, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("DATE_RANGE_MDN_SENT_DATE").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 6) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(7, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("DATE_RANGE_INTERCHNG_PROC_DATE").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 7) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">&nbsp;
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(8, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("CASE_RECPT_DATE").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 8) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(9, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("EXP_RPT_COUN_INCI").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 9) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(10, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("REP_TYPE").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 10) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(11, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("IMPORTED_CASE_NUMBER").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 11) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(12, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("PROD_NAME").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 12) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(13, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("GENERIC_NAME").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 13) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>&nbsp;
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(14, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("CURRENT_WF_STATE").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 14) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(15, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("EVENT_PT").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 15) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(16, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("EVENT_LLT").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 16) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>&nbsp;
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(17, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("SITE_ASSIGNED").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 17) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(18, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("PAT_INITIALS").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 18) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(19, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLabelDirect(GetTranslationData("STD_ID") & " / " & GetTranslationData("PAT_ID")).Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 19) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(20, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("REPTR_TYPE").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 20) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"
																										style="cursor: pointer" onclick="fn_SortHeader(21, <%= e2bPendingStsRpt_lNextSortOrder %>);"><%BuildLocalLabel("REPORTER").Render()%></span>
																									<% If (e2bPendingStsRpt_bCurrentSort = 21) Then %>
																									<img src="<%=e2bPendingStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																				</tr>
																			</table>
																		</div>
																	</td>
																</tr>
															</thead>
															<tbody valign="top">
																<tr>
																	<td class="no-padding-left">
																		<div id="Div2" class="table-scroll" style="height: 100%; overflow-y: scroll">
																			<table id="" class="grd-table-body" cellspacing="0" cellpadding="2">
																				<col width="5%" />
																				<col width="15%" />
																				<col width="12%" />
																				<col width="18%" />
																				<col width="20%" />
																				<col width="20%" />
																				<col />
																				<%        
																					If gSecHead_lGridRowCount>0 Then
                                                                                        Response.Write(e2bPendingStsRpt_oGridControl.Render(1, e2bPendingStsRpt_lPageRowCount))
                                                                                    End If
																				%>
																			</table>
																		</div>
																	</td>
																</tr>
															</tbody>
														</table>
													</td>
												</tr>
												<!-- Section Contents - Open Cases Search Results Ends -->
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr class="tblheader-gray" style="height: 25px;">
							<td align="right">
								<%If (e2bPendingStsRpt_lResultsRowCount > 0) Then %>
								<input class="button button-large" type='button' value='<%=GetTranslationData("REJECT_E2B")%>' onclick="fn_RejectMultipleE2B()" />
								<input class="button button-large" type='button' value='<%=GetTranslationData("BTN_ACCEPT_E2B")%>' onclick="fn_AcceptMultipleE2B()" />
								<input class="button button-large" type='button' value='<%=GetTranslationData("PRINT_LIST")%>' onclick="fn_PrintList();" />
								<%End If %>
							</td>
						</tr>
					</table>
				</form>
			</td>
		</tr>
	</table>
</body>
</html>
<!-- Page Display Ends -->
<!-- Javascript Functions Starts -->

<script type="text/javascript">
	var selectedCaseId;
	var rptTitle = "";
	var caseNum = "";
	var checkboxString = "";
	var selectedMsgNum = -1;
	var selectedRegRptID = -1;
	var selectedCaseNum = "0";
	var selectedAgency = "";
	var selectedStatus = "";
	var selectedTransId = -1;
	var selectedBlobsize = -1;
	var AjaxRemoveTx = 0;
	var AjaxReTx = 0;
	var lTotalReports = 0;
	var CompanyNumb = "";
	var selectedE2BAthorityId = "-1"
	var selectedisHL7Profile = "1"
	var selectedIsApplyNewFw = "0"
    var selectedE2bReportID = "-1";
	var selectedEncryptE2bReportID = "<%=query_Crypt.Encrypt(-1, GetEncryptKey()) %>";
	var isJReport = 0;
	var selectedWorldWideUNum = "-1";
    var selectedEncryptWorldWideUNum = "<%=query_Crypt.Encrypt(-1, GetEncryptKey()) %>";
	var selectedReportTypeIFN = "-1";
	var selectedCaseDeleteStatus = "0";
	var selectedCaseLockStatus = "0";
	var selectedCaseCloseStatus = "0";
	var selectedCaseOpenedUser = "-99";
	var pending_action = 8;
	var stitleE2bIncoming = '<%=GetTranslationData("INCOME_E2B")%>';
	var stitleE2bPendRpt = '<%=GetTranslationData("PEND_RPT")%>';

	function fn_MessageTypeChange(){
		var mesTypedll = fn_getElementByName("selectMessageType");
			   var selectedText = GetTextContentFromXML(mesTypedll.options[mesTypedll.options.selectedIndex]);//.text;
		fn_getElementByName("hdn_Message_type_Desc").value = selectedText;
	}

	//Constant Definition
	//pending_action = 1 (ICSR Viewer)
	//pending_action = 2 (View Error)
	//pending_action = 3 (Duplicate Search)
	//pending_action = 4 (Reject ICSR)
	//pending_action = 5 (Accept ICSR)
	//pending_action = 6 (Reject ICSRs)
	//pending_action = 7 (Accept ICSRs)
	//pending_action = 8 (Print List)
	function fn_Show_Menu(e, iRow) 
	{  
		//Call first to select the function
		f_e2bPendingStsRpt_SelectRow(e.target, iRow);
		fn_SetMenu(e, 'PENDING');
		if (selectedReportTypeIFN == 1)
			ShowContextMenuItem('e2bselectiveintake');
		else
			HideContextMenuItem('e2bselectiveintake');

		if(selectedIsApplyNewFw == 1)
			ShowContextMenuItem('e2bviewvalidation');
		else
			HideContextMenuItem('e2bviewvalidation');
	}

	function fn_load_menus() {
		//Dynamic Menus    
		var sectionMenuName = "PENDING";
		var icsrViewer = '<%=GetTranslationData("ICSR_VIEWER")%>';
		var viewErrOrWarningMsg = '<%=GetTranslationData("ERR_WARN_MSG")%>';
		var duplicateSearch = '<%=GetTranslationData("DUPLICATE_SEARCH")%>';
		var rejectE2B = '<%=GetTranslationData("REJECT_E2B")%>';
		var acceptE2B = '<%=GetTranslationData("BTN_ACCEPT_E2B")%>';
		var selectiveAcceptance = '<%=GetTranslationData("SELECTIVE_ACCEPTANCE")%>';
		var viewValidation = '<%=GetTranslationData("VIEW_VALIDATION")%>';
	
		AddContextMenuItem(icsrViewer, 'fn_ViewE2b();', 'e2bviewer', sectionMenuName, false);
		AddContextMenuItem(viewErrOrWarningMsg, 'fn_Warning();', 'e2berror', sectionMenuName, false);
		AddContextMenuItem(duplicateSearch, 'fn_DuplicateSearch();', 'e2bdupsearch', sectionMenuName, false);
		AddContextMenuItem(rejectE2B, 'fn_RejectE2BCase();', 'e2breject', sectionMenuName, false);
		AddContextMenuItem(acceptE2B, 'fn_AcceptE2BCase();', 'e2baccept', sectionMenuName, false);
		AddContextMenuItem(selectiveAcceptance,'fn_SelectiveIntake();','e2bselectiveintake',sectionMenuName,true);
		AddContextMenuItem(viewValidation,'fn_ViewValidation();','e2bviewvalidation',sectionMenuName,false);
	}

	async function fn_ViewValidation() 
	{
		var strURL;
		if (selectedE2bReportID > 0)
		{
			strURL = "/e2b/e2bimport/Ajax_E2BICSRValidation.asp?case_num=" + selectedWorldWideUNum + "&IsJReport=" + isJReport + "&report_id=" + selectedE2bReportID + "&<%=GetRequestKeyValue()%>";
		   await loadArgusMessage(strURL, fn_ViewValidationPDF);
		}   
	}

	async function fn_ViewValidationPDF()
	{
		var xmlDoc = this.req.responseXML;
		var asDocId = xmlDoc.getElementsByTagName("GN_REPORT_IDENTIFIER");
		var sError = xmlDoc.getElementsByTagName("ERROR_TXT");

		if (asDocId && (asDocId.length > 0))
			fn_ViewDocument(GetTextContentFromXML(asDocId[0]), "");
		else
			await MessageBoxRes("E2BCHECK_ERROR", "", GetTextContentFromXML(sError[0]));
	}

	async function fn_SelectiveIntake()
    {
		showLoading();
		var strURL = "/E2B/Misc/AjaxGetFollowUpDiffData.asp?IsJReport=" + isJReport + "&current_report_id=" + selectedE2bReportID;
		await loadArgusMessage(strURL, fn_CallBack_SelectiveIntake);
	}

	async function fn_CallBack_SelectiveIntake()
    {
		hideLoading();
		var strStyle = { dialogHeight: "950", dialogWidth: "1440", resizable: false, scrollable: false };
		var strURL = "/E2B/Incoming/E2b_ViewDifferences.asp?IsJReport=" + isJReport+ "&current_id=" + selectedE2bReportID + "&ApplyNewFw=" + selectedIsApplyNewFw;
		var oResults = await fn_OpenModalDialog(strURL, window, strStyle);
	
		if (oResults == null || oResults == undefined || oResults == 1) 
		{
			await loadArgusMessage("/E2B/Incoming/Ajax_DeleteE2BDiffData.asp", fn_AfterDelete, "esm_report_id=" + selectedE2bReportID);			
			fn_Search(1); 
		}
		else
		{
			if(oResults !== sSessionTimeOutDialogReturn)
			{
				var oResultsArray = oResults.split("#||#");
				if (oResultsArray[0]=="Accepted") //Case of Accept
				{
					var AcceptNotes = oResultsArray[1];
					var AcceptStatus = oResultsArray[2];
					document.Frm_E2B.notes.value = AcceptNotes;
					document.Frm_E2B.action = "/E2B/Incoming/E2bPending.asp?rb_report_status=" + selectedReportTypeIFN + "&esm_report_id=" + selectedE2bReportID + "&esm_status=" + AcceptStatus + "&<%=GetRequestKeyValue()%>";
					fn_Search(1);
				}
				else if (oResultsArray[0]=="Rejected")  //Case of Reject
				{
					var AcceptNotes = oResultsArray[1];
					document.Frm_E2B.notes.value = AcceptNotes;
					document.Frm_E2B.action = "/E2B/Incoming/E2bPending.asp?rb_report_status=" + selectedReportTypeIFN + "&esm_report_id=" + selectedE2bReportID + "&esm_status=103&<%=GetRequestKeyValue()%>";
					fn_Search(1);
				}
			}
		}
	}

	async function fn_AfterDelete()
	{
		var oXML = this.req.responseXML; 
		var sError = fn_GetAjaxErrorMsg(oXML);
		if (sError.length > 0)             // Check if the returned message has error
			await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("IMPORT")%>', sError);
		await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + selectedE2bReportID);
	}
			
	function fn_Search(bRefreshCache)
	{
        showLoading();
        if (bRefreshCache == 2) { // Only reset CurrentPage value on Search button clicked.
            document.all.bSearch.value = 1;
            document.Frm_E2B.Frm_E2B_CurrentPage.value = 1;
        }
        else { 
            document.all.bSearch.value = bRefreshCache;
        }
		fn_ValidateSubmitForm(document.forms[0]);
	}

	function fn_GetSelecetedIds()
	{
		var rowCount = <%=e2bPendingStsRpt_lPageRowCount%>;
		var selectedReportIds = "";
		var selectedWorldWideIds = "";
		var selectedReportTypes = "";
		var selectedRowCount = 0;
	
		for (var i=1;i<=rowCount;i++) 
		{
			if (eval("document.all.checkbox_" + i).checked == true)
			{
				var reportId = eval("document.all.E2BTpl_REPORT_ID_" + i).value;			
				var worldwideId = eval("document.all.E2BTpl_WORLDWIDEUNIQUENO_" + i).value;			
				var reporttype = eval("document.all.E2BTpl_REPORTTYPEIFN_" + i).value;			
				selectedReportIds = selectedReportIds + reportId + "|";			
				selectedWorldWideIds = selectedWorldWideIds + worldwideId + "|";			
				selectedReportTypes = selectedReportTypes + reporttype + "|";			
				selectedRowCount += 1;
			}
		}   
	}

	async function fn_RejectMultipleE2B()
    {
        if (!fn_ValidateForm("document.<%= e2bPendingStsRpt_sFormName %>"))
            return;
		pending_action = 6;

		var rowCount = <%=e2bPendingStsRpt_lPageRowCount%>;
		if (rowCount == "0") return;
		var selectedReportIds = "";
		var selectedRowCount = 0;
	
		for (var i=1;i<=rowCount;i++) 
		{
			if (eval("document.all.checkbox_" + i).checked == true)
			{
				var reportId = eval("document.all.E2BTpl_REPORT_ID_" + i).value;			
				selectedReportIds = selectedReportIds + reportId + ",";			
				selectedRowCount += 1;
			}
		}
		if (selectedRowCount == 0)
		{
			await MessageBoxRes("E2BREJECT_MIN_RPT_ONE");
			return false;
		}
	
		if(selectedReportIds.lastIndexOf(",") > 0)
		{
			selectedReportIds = selectedReportIds.substring(0,selectedReportIds.length -1);
		}
		await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp",fn_LockedReport, "bulk=1&esm_report_id=" + selectedReportIds);
	}

	async function fn_RejectMultipleE2BReports(selectedReportIds)
	{    
		var strURL, notes;
	
		if (selectedReportIds.length > 0)
		{
			strURL = "/E2B/Actions/RejectE2BCaseMultiple.asp";
			var sDialogStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };
			notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);
			
			if (!notes)
			{
				await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + selectedReportIds);
			}
			else
			{
				if(notes !== sSessionTimeOutDialogReturn)
				{
					document.Frm_E2B.notes.value = notes;
					document.Frm_E2B.action = "/E2B/Incoming/E2bPending.asp?rb_report_status=0&esm_report_id=" + selectedReportIds + "&esm_status=103&<%=GetRequestKeyValue()%>";
					fn_Search(1);
				}
			}
		}
	}

	async function fn_AcceptMultipleE2B()
	{
		var strURL, notes, esm_report_status;
		var rowCount = <%=e2bPendingStsRpt_lPageRowCount%>;
		if (rowCount == "0") return;
		var selectedReportIdsIni = "";
		var selectedReportIdsFollow = "";
		var showFolloupPopup = 0;
		var selectedReportIdsNull = "";
		var selectedReportTypes = "";
		var selectedRowCount = 0;
		var CaseLocked = 0;
		var sRptForm = "";
		var chkCLResults = -1;
        var CaseDelete = 0;

        if (!fn_ValidateForm("document.<%= e2bPendingStsRpt_sFormName %>"))
            return;

		<%If GetXMLValueDirect(oSession, "NUMBERING_AUTO") <> "1" Then%>
		await MessageBoxRes("AUTONUMBERING_PREREQUSITE_UNMET");
		return;
		<%End If %>

		for (var i=1;i<=rowCount;i++) 
		{
			if (eval("document.all.checkbox_" + i).checked == true)
			{
				if( eval("document.all.E2BTpl_REPORTTYPEIFN_ID_" + i).value == "1") //Initial
				{
					var reportIdIni = eval("document.all.E2BTpl_REPORT_ID_" + i).value;
					selectedReportIdsIni =  selectedReportIdsIni + reportIdIni + ",";
				}
				if( eval("document.all.E2BTpl_REPORTTYPEIFN_ID_" + i).value == "3" || eval("document.all.E2BTpl_REPORTTYPEIFN_ID_" + i).value == "5" || eval("document.all.E2BTpl_REPORTTYPEIFN_ID_" + i).value == "6") //Follow Up
				{
					if (!(eval("document.all.E2BTpl_IMPORTED_CASE_NUM_" + i).value))
					{
						var reportIdIni = eval("document.all.E2BTpl_REPORT_ID_" + i).value;
						selectedReportIdsIni =  selectedReportIdsIni + reportIdIni + ",";
					}
					else
					{
						if (eval("document.all.CASE_STATE_" + i).value == "0" && eval("document.all.CASE_OPENED_USER_" + i).value == "-99")
						{
							if (eval("document.all.CASE_CURRENT_STATE_" + i).value == "0")
							{
								var reportIdFollow = eval("document.all.E2BTpl_REPORT_ID_" + i).value;
								if(eval("document.all.E2BTpl_HL7P_" + i).value == "0" && isJReport != 1){
									showFolloupPopup = 1;
								}
								selectedReportIdsFollow =  selectedReportIdsFollow + reportIdFollow + ",";
							}
							else
							{
								CaseDelete = 1;
								eval("document.all.checkbox_" + i).checked = false;
							}
						}
						else
						{
							CaseLocked = 1;
							eval("document.all.checkbox_" + i).checked = false;
						}
					}
				}	
				if( eval("document.all.E2BTpl_REPORTTYPEIFN_ID_" + i).value == "4") //Nullification
				{
					if ((eval("document.all.CASE_CLOSED_STATE_" + i).value == "0") && (eval("document.all.CASE_OPENED_USER_" + i).value == "-99"))
					{
						if (eval("document.all.CASE_CURRENT_STATE_" + i).value == "0")
						{
							var reportIdIniNull = eval("document.all.E2BTpl_REPORT_ID_" + i).value;
							selectedReportIdsNull =  selectedReportIdsNull + reportIdIniNull + ",";	
						}
						else
						{
							CaseDelete = 1;
							eval("document.all.checkbox_" + i).checked = false;
						}
					}
					else
					{
						CaseLocked = 1;
						eval("document.all.checkbox_" + i).checked = false;
					}
				}	
				selectedRowCount += 1;
			}
		
		}
	
		if (selectedRowCount == 0)
		{
			await MessageBoxRes("E2BPEND_MIN_ACCEPT_RPT_ONE");
			return;
		}
	
		var E2BViewType = <%=JavaScriptSanitize(E2BViewType) %>;
		var pInput = new Array();
		if (selectedReportIdsIni.length > 0)
		{
			if(selectedReportIdsIni.lastIndexOf(",") > 0)
			{
				selectedReportIdsIni = selectedReportIdsIni.substring(0,selectedReportIdsIni.length -1);
			}
		}
		if (selectedReportIdsFollow.length > 0)
		{
			if(selectedReportIdsFollow.lastIndexOf(",") > 0)
			{
				selectedReportIdsFollow = selectedReportIdsFollow.substring(0,selectedReportIdsFollow.length -1);
			}
		}
		if (selectedReportIdsNull.length > 0)
		{
			if(selectedReportIdsNull.lastIndexOf(",") > 0)
			{
				selectedReportIdsNull = selectedReportIdsNull.substring(0,selectedReportIdsNull.length -1);
			}
		
		}
		if (selectedReportIdsIni.length > 0)
			pInput[0] = selectedReportIdsIni;
		else
			pInput[0] = "";
		if (selectedReportIdsFollow.length > 0)
			pInput[1] = selectedReportIdsFollow;
		else
			pInput[1] = "";
		if (selectedReportIdsNull.length > 0)
			pInput[2] = selectedReportIdsNull;
		else
			pInput[2] = "";
			
		var text = "";
		if (CaseLocked == 1 && CaseDelete == 0)		
			text = '<%=GetTranslationData("CASE_LOCKED_ARCHIVED_READONLY_MSG")%>';
		else if (CaseLocked == 0 && CaseDelete == 1)
			text = '<%=GetTranslationData("CASE_DELETED_MSG")%>';
		else if (CaseLocked == 1 && CaseDelete == 1)
			text = '<%=GetTranslationData("CASE_LOCKED_ARCHIVED_DELETED_READONLY_MSG")%>';
		else
			text = "";    
		if (text.length > 0)
			await MessageBoxRes("GENERAL_WARNING", stitleE2bPendRpt, text);
	
		if (selectedReportIdsFollow.length > 0 && showFolloupPopup == 1)
		{
			strURL = "/E2B/Misc/E2B_FollowupFormat.asp"
			var sDialogStyle = { dialogHeight: "100", dialogWidth: "300", resizable: false, scrollable: false };
			E2BViewType = await fn_OpenModalDialog(strURL, window, sDialogStyle);
		}
		if (selectedReportIdsIni.length > 0 || selectedReportIdsFollow.length > 0 || selectedReportIdsNull.length > 0)
		{
			var ajaxParms;
			ajaxParms = "pinput0=" + pInput[0];
			ajaxParms += "&pinput1=" + pInput[1];
			ajaxParms += "&pinput2=" + pInput[2];
			await loadArgusMessage( "/E2B/Incoming/Ajax_E2BPendingReportIDs.asp", null, ajaxParms);
		
			strURL = "/E2B/Incoming/E2B_BulkIncomingResult.asp?E2BViewType=" + E2BViewType;
			var sDialogStyle = { dialogHeight: "700", dialogWidth: "900", resizable: false, scrollable: false };
			await fn_OpenModalDialog(strURL, window, sDialogStyle);
		}
		window.close();
		if(remainingTime > 0)
			fn_Search(1);
	}

	<%
	'**************************************************************************************************
	' Author	    : Pankaj Gautam
	' Called From   : On clicking on any row
	' Parameters	: Row Id of that particular row
	' Returns	    : None
	' Description	: Shows Menu on right click after selecting the Clicked Row
	'**************************************************************************************************
	%>
	function fn_ItemMenu_OnClick(lRowId)
	{
		selectedRow = lRowId;
		showMenu(window.MN_Menu);
	}

	<%
	'**************************************************************************************************
	' Author	    : Pankaj Gautam
	' Called From   : On changing the date range value
	' Parameters	: Date range value
	' Returns	    : None
	' Description	: Sets the date from and date to values according to the date range selected
	'**************************************************************************************************
	%>
		async function f_e2bPendingStsRpt_ChangeDateRange(lValue)
	{
		var sStart, sStop;
		sStart = document.Frm_E2B.DateFrom.value;
		sStop  = document.Frm_E2B.DateTo.value;
	
		if (lValue == -1)
		{
			sStart = <%=JavaScriptSanitize(e2bPendingStsRpt_sDateFrom) %>;
			sStop = <%=JavaScriptSanitize(e2bPendingStsRpt_sDateTo) %>;
			document.Frm_E2B.RangeList.value = -1;		
			document.Frm_E2B.DateFrom.value = sStart;
			document.Frm_E2B.DateTo.value = sStop;
			<%
				If (IsDate(e2bPendingStsRpt_sDateTo) and IsDate(e2bPendingStsRpt_sDateFrom)) Then 
					If (CDate(e2bPendingStsRpt_sDateTo) < CDate(e2bPendingStsRpt_sDateFrom)) Then 
			%>
				    await MessageBoxRes("E2B_DATES_NOT_VALID");
			<% 
					End If
			End If
		%>	
			}
		else
		{
			switch (lValue)
			{
			<%	
				Set e2bPendingStsRpt_oDateRangeList = e2bPendingStsRpt_oDateRangeMsg.selectnodes ("/MESSAGE/TABLE_LM_DATE_RANGES/LM_DATE_RANGES")
			For Each e2bPendingStsRpt_oDate in e2bPendingStsRpt_oDateRangeList
			e2bPendingStsRpt_sDateFrom = fn_date_from_iso(GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_FROM_DATE"), 8, false)
			e2bPendingStsRpt_sDateTo  = fn_date_from_iso(GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_TO_DATE"), 8, false)
	%>		
			case "<%= GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID") %>": 
			sStart = "<%=e2bPendingStsRpt_sDateFrom%>"; 
			sStop = "<%=e2bPendingStsRpt_sDateTo%>"; 
			document.Frm_E2B.RangeList.value = "<%= GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID") %>"; 
			break;
			<% 
				Next 
			%>
			}
	}
	
	document.Frm_E2B.DateFrom.value = sStart;
	document.Frm_E2B.DateTo.value = sStop;
	}

<%
	'**************************************************************************************************
	' Author	    : Pankaj Gautam
	' Called From   : On selecting a row in the result set
	' Parameters	: row index
	' Returns	    : None
	' Description	: set the parameters value according to the row selected
	'**************************************************************************************************
	%>
	function f_e2bPendingStsRpt_SelectRow(oRow, iRow)
	{
		if (iRow != selectedRow)
		{
			//Now We have to save values of the field which i need on row click
			selectedCaseId = eval("document.all.CASE_ID_" + iRow).value;
			selectedE2bReportID = eval("document.all.E2BTpl_REPORT_ID_"	            + iRow).value;
            selectedEncryptE2bReportID = eval("document.all.E2BTpl_ENCRYPT_REPORT_ID_"	            + iRow).value;
			selectedWorldWideUNum = eval("document.all.E2BTpl_WORLDWIDEUNIQUENO_"	+ iRow).value;
            selectedEncryptWorldWideUNum = eval("document.all.E2BTpl_ENCRYPT_WORLDWIDEUNIQUENO_"	+ iRow).value;
			selectedReportTypeIFN = eval("document.all.E2BTpl_REPORTTYPEIFN_ID_"	    + iRow).value;
			selectedCaseDeleteStatus = eval("document.all.CASE_CURRENT_STATE_" + iRow).value;
			selectedCaseLockStatus = eval("document.all.CASE_STATE_" + iRow).value;
			selectedCaseCloseStatus = eval("document.all.CASE_CLOSED_STATE_" + iRow).value;
			selectedCaseOpenedUser = eval("document.all.CASE_OPENED_USER_" + iRow).value;
			selectedE2BAthorityId = eval("document.all.RPT_E2B_AUTHORITY_ID_" + iRow).value;
			selectedisHL7Profile = eval("document.all.E2BTpl_HL7P_" + iRow).value;
			selectedIsApplyNewFw = eval("document.all.E2BTpl_APPLY_NEW_FW_" + iRow).value;
			if (selectedE2BAthorityId == 4)
			{
				isJReport = 1
			}
		
			if (eval("document.all.E2BTpl_IMPORTED_CASE_NUM_" + iRow).value)
				selectedCaseNum = eval("document.all.E2BTpl_IMPORTED_CASE_NUM_" + iRow).value;
			else
			{
				selectedCaseNum = "0";
				selectedReportTypeIFN = 1;
			}
			fn_highlightCurrentRow(oRow);
		
			selectedRow = iRow;
		}
	}

	<%
	'**************************************************************************************************
	' Author	    : Pankaj Gautam
	' Called From   : f_e2bPendingStsRpt_SelectRow
	' Parameters	: Row Index, New Background Color
	' Returns	    : None
	' Description	: Set the background color for the specified row
	'**************************************************************************************************
	%>	
		function f_e2bPendingStsRpt_SetRowColor(iRow, colorValue)
		{
			//Setting row color
			eval ("document.all.data_row_"			                  + iRow).style.background = colorValue;
		
			//Setting vaue of read only text boxes so that they will feel like label
			eval ("document.all.checkbox_"			                  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_ROWNUM_"			              + iRow).style.background = colorValue;
			//Row 1
			eval ("document.all.E2BTpl_TRADING_PARTNER_"			  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_REPORTTYPEIFN_"			      + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_WORLDWIDEUNIQUENO_"			  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_SENDERCASENO_"			      + iRow).style.background = colorValue;
			//Row 2
			eval ("document.all.E2BTpl_TRANSMISSIONSENTSENT_"		  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_EDI_TRANSMIT_DATE_"		  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_INTERCHANGE_PROCESSED_DATE_"	  + iRow).style.background = colorValue;
			//Row 3
			eval ("document.all.E2BTpl_CASE_RECEIPT_DATE_"		      + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_COUNTRY_"		              + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_REPORT_TYPE_"		          + iRow).style.background = colorValue;
			//Row 4
			eval ("document.all.E2BTpl_PRODUCT_NAME_"			      + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_GENERIC_NAME_"			      + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_not_used1_"	                  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_WORKFLOW_STATE_NAME_"          + iRow).style.background = colorValue;
			//Row 5
			eval ("document.all.E2BTpl_EVENT_PT_"	                  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_EVENT_LLT_"	                  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_not_used2_"	                  + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_SITES_DESC_"	                  + iRow).style.background = colorValue;
			//Row 6
			eval ("document.all.E2BTpl_PAT_INI_"			          + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_STYDYID_PATID_"			      + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_REPORTER_TYPE_"			      + iRow).style.background = colorValue;
			eval ("document.all.E2BTpl_REPORTER_"			          + iRow).style.background = colorValue;
		}

		<%
		'**************************************************************************************************
	' Author	   : Pankaj Gautam
	' Called From  : When Print button is clicked
	' Parameters   : None
	' Returns	   : None
	' Description  : gets the IDs of the selected checkboxes on the page
	'**************************************************************************************************
	%>
	async function f_e2bPendingStsRpt_GetCheckBoxString() 
	{
		var rowCount = <%=e2bPendingStsRpt_lPageRowCount%>;
		var selectedIds = "";
		var selectedRowCount = 0;
		
		for (var i=1;i<=rowCount;i++) {
			if (eval("document.all.checkbox_" + i).checked == true){
				var id = eval("document.all.checkbox_" + i).value;			
				selectedIds = selectedIds + id + ",";			
				selectedRowCount += 1;
			}
		}
		
		if (selectedRowCount == 0)
		{
			await MessageBoxRes("E2BTRANS_MIN_RECORD_ONE");
			return false;
		}
		
		if (selectedIds.length > 0) 
			selectedIds = selectedIds.substring(0, selectedIds.length-1);
			
		f_e2bPendingStsRpt_Print(selectedIds);
	}

	<%
	'**************************************************************************************************
	' Author	   : Pankaj Gautam
	' Called From  : When Trading Partner filter button is clicked to add New Destination
	' Parameters   : None
	' Returns	   : None
	' Description  : 
	'**************************************************************************************************
	%>
	async function fn_addDestination()
	{
        var strURL, strStyle, strResult;
		
		strStyle = {
			dialogHeight: "480",
			dialogWidth: "720",
			resizable: false,
			scrollable: false
        };
        if (!fn_ValidateForm("document.<%= e2bPendingStsRpt_sFormName %>"))
            return;
		strURL   = "/Reports/BRBF/BBF_Agency_Filter.asp?E2BImport=ICSR&DSPLYLNG=" + glDisplayLang;
		strResult = await fn_OpenModalDialog(strURL, window, strStyle);
	}

	async function fn_DuplicateSearch()
	{
		//if report is process by another user
		pending_action = 3;
		await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "check_lock_status=1&esm_report_id=" + selectedE2bReportID);
	}

	async function fn_CallDuplicateSearch()    
	{ 
		var strURL = "/E2B/Incoming/E2bIncomingReports.asp?DSPLYLNG=" + glDisplayLang + "&IsJReport=" + isJReport + "&rb_report_status=" + selectedReportTypeIFN + "&esm_report_id=" + selectedE2bReportID + "&LockStatus=" + selectedCaseLockStatus + "&DeleteStatus=" + selectedCaseDeleteStatus + "&CloseStatus=" + selectedCaseCloseStatus + "&OpenedUser=" + selectedCaseOpenedUser + "&CaseId=" + selectedCaseId + "&SelectedCaseNum=" + selectedCaseNum + "&AuthorityId=" + selectedE2BAthorityId + "&ApplyNewFw=" + selectedIsApplyNewFw;        
		var strStyle = { dialogHeight: "720", dialogWidth: "1080", resizable: false, scrollable: false };
		var oResults = await fn_OpenModalDialog(strURL, window, strStyle);
			   if (oResults == null || oResults == undefined) 
		{
			fn_Search(1); 
		}
		else
		{
			if (oResults !== sSessionTimeOutDialogReturn)
			{
				var oResultsArray = oResults.split("#||#");
			
				if (oResultsArray.length=="3") //Case of Open Case
				{
					var CaseNum = oResultsArray[2];
					await fn_OpenCaseFormNum(CaseNum);
				}
				else if (oResultsArray.length=="2") //Case of Accept
				{
					var AcceptNotes = oResultsArray[0];
					var AcceptStatus = oResultsArray[1];
					document.Frm_E2B.notes.value = AcceptNotes;
					document.Frm_E2B.action = "/E2B/Incoming/E2bPending.asp?rb_report_status=" + selectedReportTypeIFN + "&esm_report_id=" + selectedE2bReportID + "&esm_status=" + AcceptStatus + "&<%=GetRequestKeyValue()%>";
					fn_Search(1);
				}
				else if (oResultsArray.length=="1")  //Case of Reject
				{
					var AcceptNotes = oResultsArray[0];
					document.Frm_E2B.notes.value = AcceptNotes;
					document.Frm_E2B.action = "/E2B/Incoming/E2bPending.asp?rb_report_status=" + selectedReportTypeIFN + "&esm_report_id=" + selectedE2bReportID + "&esm_status=103&<%=GetRequestKeyValue()%>";
					fn_Search(1);
				}
			}
		}
	}

	function fn_ClearAll()
    {
        if (!fn_ValidateForm("document.<%= e2bPendingStsRpt_sFormName %>"))
            return;
		document.all.ReportingAgency.length= 0;
		document.all.ReportingAgencyKey.value='';
		document.all.ReportingAgencyValue.value='';
	}

	<%
	'**************************************************************************************************
	' Author	   : Pankaj Gautam
	' Called From  : Onload function of the Page
	' Parameters   : None
	' Returns	   : None
	' Description  : It initialises the search criteria
	'**************************************************************************************************
		%>
		async function Init_Frm()
	{
		fn_load_menus();
		
		if(<%=JavaScriptSanitize(trim(e2bPendingStsRpt_oProdGenRadio))%> == "0")
		{
			if(<%=JavaScriptSanitize(trim(e2bPendingStsRpt_ProdName)) %> != "")
				document.all.selProduct.value = <%=JavaScriptSanitize(e2bPendingStsRpt_ProdName)%>;
		} 
		else
		{
			if(<%=JavaScriptSanitize(trim(e2bPendingStsRpt_GenericName)) %> != "")
				document.all.selProduct.value = <%=JavaScriptSanitize(e2bPendingStsRpt_GenericName)%>;
		}
				
		if(<%=JavaScriptSanitize(trim(e2bPendingStsRpt_RptType)) %> != "")
			document.all.selRType.value = <%=JavaScriptSanitize(e2bPendingStsRpt_RptType) %>;

					
		if(<%=JavaScriptSanitize(trim(e2bPendingStsRpt_DateType)) %> != "")
			document.all.DateType.value = <%=JavaScriptSanitize(e2bPendingStsRpt_DateType) %>; 
		
		if(<%=JavaScriptSanitize(trim(e2bPendingStsRpt_MsgType)) %> != "")
			document.all.selectMessageType.value = <%=JavaScriptSanitize(e2bPendingStsRpt_MsgType) %>; 
			

		fn_MessageTypeChange();

		await f_e2bPendingStsRpt_ChangeDateRange(document.Frm_E2B.RangeList.value);
	}
	
	/*'************************************************************************************
	' Author       : Pankaj Gautam
	' Page         : E2BPending.asp
	' Description  : This function invokes the ajax function which fetches the url of 
	'                the report generated when the user tries to print the current view
	'************************************************************************************
	' Revision History
	' Date		    Author		Description
	' 6-JULY-2006     PS        Initial Version
	'************************************************************************************/
	async function fn_PrintList()
	{
		
        var strURL;
        if (!fn_ValidateForm("document.<%= e2bPendingStsRpt_sFormName %>"))
            return;
		strURL = "/Lookup/ReportGenerateAjaxIntrim.asp?ReportType=3&resultsxml=Frm_E2BPending_Response&cacheId=" + fn_URLEncode(document.Frm_E2B.UserCacheId.value);
		strURL += "&gmtOffSet=<%= Server.URLEncode(e2bPendingStsRpt_lGmtOffSet) %>";
		strURL += "&displayLang=<%= Server.URLEncode(glDisplayLang) %>";
		strURL += "&userFullName=<%= Server.URLEncode(sUserFullName) %>";
		strURL += "&dateFrom=<%= Server.URLEncode(sCopyDateFrom) %>";
		strURL += "&dateTo=<%= Server.URLEncode(sCopyDateTo) %>";
		strURL += "&prodName=<%= Server.URLEncode(sProdName) %>";
		strURL += "&dateType=<%= Server.URLEncode(sDataType) %>";
		strURL += "&agencyList=<%= Server.URLEncode(sAgencyList) %>";
		strURL += "&strReportType=<%= Server.URLEncode(sReportType) %>";
		strURL += "&status=<%= Server.URLEncode(sStatus) %>";
		strURL += "&cfgMessageTypeDesc=<%= Server.URLEncode(e2bPendingStsRpt_HdnMsgType_Desc) %>";
		var sDialogStyle = { dialogHeight: "200", dialogWidth: "350", resizable: false, scrollable: false };
		await fn_OpenModalDialog(strURL, window, sDialogStyle);
	}
	
	async function fn_QuickyOpen(sCaseNum)
	{
		var sCaseNum;
		showLoading();
		await loadArgusMessage( "/Worklist/Status/Ajax_CheckForCase.asp?CaseNum=" + fn_URLEncode(sCaseNum), fn_callbackCaseNum);
		if(chkCLResults == 0)
			await MessageBoxRes("CASE_SEARCH_FALIED");
		else 
			await fn_OpenCaseFormNum(sCaseNum);
		return false;
	}

	async function fn_callbackCaseNum()
	{
		var xmlDoc = this.req.responseXML;
		var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
		var results;
		hideLoading();
		if (sErrStr.length > 0)
		{
			await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("OPEN_CASE")%>', sErrStr); 
			chkCLResults = 0;
			return;
		}
		results = GetTextContentFromXML(xmlDoc.lastChild.attributes[0]);
		if (results.length > 0)
			chkCLResults = parseInt(results);
	}

	function openE2BViewer(repID,CompanyNumb)
	{
		fn_ViewE2b(CompanyNumb, <%=JavaScriptSanitize(E2bViewType)%>, repID);
	}

	function fn_ViewE2b(CompanyNumb,E2bViewType,E2bReportID)
	{
		if (E2bReportID > 0)
		{
			strURL = "/E2B/E2bViewer/E2bViewer.asp?CaseID=" + fn_URLEncode(CompanyNumb) + "&ReportID=<%=Server.URLEncode(query_Crypt.Encrypt(-1, GetEncryptKey())) %>" + "&ProductID=-1" + "&LicenseID=-1" + "&CountryID=-1" + "&AgencyID=-1" + "&Draft=-1" + "&TimeFrame=-1" + "&E2BViewType=" + E2bViewType +"&IncomingE2b=1&E2bReportID=" + fn_URLEncode(E2bReportID) + "&ClickViewReport=0&<%=GetRequestKeyValue()%>";
			window.open(strURL);
		}
	}
	
	<%
	'**************************************************************************************************
	' Author	    : Pankaj Gautam
	' Called From   : 
	' Parameters	: Calling control
	' Returns	    : None
	' Description	: 
	'**************************************************************************************************
	%>	
	function f_e2bPendingStsRpt_SelectUnselectAll(oCheckbox)
	{
	   
		if(<%=e2bPendingStsRpt_lPageRowCount %> < 1) return;
		<%	Dim lIndex
		For lIndex = 1 To e2bPendingStsRpt_lPageRowCount
		%>
			document.all.checkbox_<%=lIndex %>.checked = oCheckbox.checked;
		<%
		Next
		%>
	}	

	async function fn_AcceptE2BCase() 
	{
		var notes, rowCount;
		pending_action = 5;
		
		var rowCount = <%=e2bPendingStsRpt_lPageRowCount%>;
		if (rowCount < 0) return;
		notes = "";
				
		await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "esm_report_id=" + selectedE2bReportID);
	}
	
	async function fn_LockedReport()
	{
		var xmlDoc = this.req.responseXML;
		var sLocked_User;
		var sesm_report_ids = "";
		
		var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
		if (sErrStr.length > 0)
		{
			await MessageBoxRes("GENERAL_ERROR",stitleE2bIncoming,sErrStr) ;
			return;
		}
				
		var asLocked_User = xmlDoc.getElementsByTagName("LOCKED_USER");
		var asEsmReportIDs= xmlDoc.getElementsByTagName("MODIFY_REPORT_IDS");
		if (asLocked_User)
		{
			sLocked_User = GetTextContentFromXML(asLocked_User[0]);
			if (asEsmReportIDs)
				sesm_report_ids = GetTextContentFromXML(asEsmReportIDs[0]);
			if (sLocked_User.length > 0)
				await MessageBoxRes("GENERAL_INFORMATION",stitleE2bIncoming,sLocked_User); 
			if ((pending_action == 6) || (sLocked_User.length <= 0 && pending_action != 6))
			{
				if (pending_action == 2)
				{
					fn_CallWarning();
				}
				else if (pending_action == 3)
				{
					fn_CallDuplicateSearch();
				}
				else if (pending_action == 4)
				{
					await fn_RejectSingleE2BCase();
				}
				else if (pending_action == 5)
				{
					await fn_AcceptE2BSingleCase();
				}
				else if (pending_action == 6)
				{
					await fn_RejectMultipleE2BReports(sesm_report_ids);
				}
				else
				{
					await MessageBoxRes("E2BINCOME_NO_COND_FOR_FURTHER");
				}
			}
		}
		else
		{
			if (pending_action == 2)
			{
				fn_CallWarning();
			}
			else if (pending_action == 3)
			{
				fn_CallDuplicateSearch();
			}
			else if (pending_action == 4)
			{
				await fn_RejectSingleE2BCase();
			}
			else if (pending_action == 5)
			{
				await fn_AcceptE2BSingleCase();
			}
			else if (pending_action == 6)
			{
				await fn_RejectMultipleE2BReports(sesm_report_ids);
			}
			else
			{
				await MessageBoxRes("E2BINCOME_NO_COND_FOR_FURTHER");
			}
		}
		return;
	}
	async function fn_UnLockedReport()
	{
		var xmlDoc = this.req.responseXML;
		var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
		if (sErrStr.length > 0)
			await MessageBoxRes("GENERAL_ERROR",stitleE2bIncoming,sErrStr); 
		return;
	}

	async function fn_AcceptE2BSingleCase() 
	{
		var strURL, esm_report_status, notes, status, lWarning, AcceptedCase,rowCount, PostSave = "";
		var openCasePermit = <%=JavaScriptSanitize(lOpenCasePermit) %>;
		var lockCasePermit = <%=JavaScriptSanitize(lLockCasePermit) %>;
		var rowCount = <%=e2bPendingStsRpt_lPageRowCount%>;
		var titleIncomingE2B = '<%=GetTranslationData("INCOME_E2B")%>';
		var sError = "";
		var justnotes = "";
		if (rowCount < 0) return;
		notes = "";
		var sCaseOrCases,sIsOrAre;
		if (selectedReportTypeIFN == 3 || selectedReportTypeIFN == 5 || selectedReportTypeIFN == 6 ) //FollowUp
		{
			strURL = "/E2B/Actions/E2B_AcceptFollowupE2BCase.asp?IsJReport=" + isJReport + "&pending=1&esm_report_id=" + selectedE2bReportID + "&case_id=" + selectedCaseId + "&CloseStatus=" + selectedCaseCloseStatus + "&LockStatus=" + selectedCaseLockStatus + "&ApplyNewFw=" + selectedIsApplyNewFw;
			// if the case is opened by another user show the error message
			if (selectedCaseOpenedUser != "-99")
			{
				strURL = "";
				await MessageBoxRes("INCOMERPT_ALREADY_IN_USE",'',selectedCaseOpenedUser); 
			}
			else if (selectedCaseDeleteStatus == 1) // if the case is deleted show the error message
			{
				strURL = "";
				await MessageBoxRes("E2B_DELCASE_UPDT_FAILED");
			}
			else if (selectedCaseCloseStatus == 1) // if the case is closed show the error message
			{
				if (openCasePermit == "0")
				{
					strURL = "";
					await MessageBoxRes("INCOMERPT_CLOSECASE_RESTRICTION");
				}
			}
			else if (selectedCaseLockStatus == 1) // if the case is locked show the error message
			{
				if (lockCasePermit == "0")
				{
					strURL = "";
					await MessageBoxRes("INCOMERPT_UNLOCKCASE_RESTRICTION");
				}
			}
		}
		else if (selectedReportTypeIFN == 4) //Nullification
		{
			strURL = "/E2B/Actions/E2B_AcceptNullificationE2BCase.asp?esm_report_id=" + selectedE2bReportID + "&case_id=" + selectedCaseId + "&CloseStatus=" + selectedCaseCloseStatus;
			// if the case is opened by another user show the error message
			if (selectedCaseOpenedUser != "-99")
			{
				strURL = "";
				await MessageBoxRes("E2BNULL_CASE_IN_USE","",selectedCaseOpenedUser); 
			}
				// if the case is deleted show the error message
			else if (selectedCaseDeleteStatus == 1)
			{
				strURL = "";
				await MessageBoxRes("E2B_DELCASE_UPDT_FAILED");
			}
			else if (selectedCaseCloseStatus == 1)
			{
				if (openCasePermit == "0")
				{
					strURL = "";
					await MessageBoxRes("INCOMERPT_CLOSECASE_RESTRICTION");
				}
			}
		}
		else //Initial
			strURL = "/E2B/Actions/E2B_AcceptE2BCase.asp?IsJReport=" + isJReport + "&esm_report_id=" + selectedE2bReportID + "&receipt_date=-1" + "&product_name=-1" + "&report_type_id=-1" + "&country_id=-1" + "&ApplyNewFw=" + selectedIsApplyNewFw;

		if (strURL != "") {
			var sDialogStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };
			if (selectedReportTypeIFN == 4) //Nullification
				sDialogStyle = { dialogHeight: "640", dialogWidth: "480", resizable: false, scrollable: false };
			notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);
		}
		if (!notes)
		{
			await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + selectedE2bReportID);
		}
		else
		{
			if(notes !== sSessionTimeOutDialogReturn)
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
							notes = notes.substring(Pos + 1);
							Pos = notes.indexOf("~",1);
							if (Pos > 0)
							{
								justnotes = notes.substring(0,Pos);
								PostSave = notes.substring(Pos + 1);
								if (justnotes.length < 1)
									return;
								if (status <= 0)
								{
									status = "";
								}
							}  
							else 
								justnotes = notes;
						}
					}
					else
						status = "";
				}
				else
					status = "";

				if (status != "102")
				{
					if (selectedReportTypeIFN == 3 || selectedReportTypeIFN == 4 || selectedReportTypeIFN == 5 || selectedReportTypeIFN == 6) {
						var sCaseNum = AcceptedCase;
						if(sCaseNum.toLowerCase() == "undefined") sCaseNum = "";
						await MessageBoxRes("E2BINCOME_CASE_UPLOAD_FAILED","",sCaseNum);
					} else
						await MessageBoxRes("E2BINCOME_CASE_NOT_ACCEPTD");
					await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + selectedE2bReportID);
				}
				else
				{
					if ((AcceptedCase.length > 1) && (AcceptedCase != "N/A") && (status == "102"))
					{
						if (selectedReportTypeIFN == 3 || selectedReportTypeIFN == 4 || selectedReportTypeIFN == 5 || selectedReportTypeIFN == 6)
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
				}
				document.Frm_E2B.notes.value = justnotes;
				document.Frm_E2B.action = "/E2B/Incoming/E2bPending.asp?rb_report_status=" + selectedReportTypeIFN + "&esm_report_id=" + selectedE2bReportID + "&esm_status=" + status + "&bSearch=1&<%=GetRequestKeyValue()%>";
				fn_ValidateSubmitForm(document.Frm_E2B);
			}
		}
	}

	async function fn_RejectE2BCase() 
	{
		pending_action = 4;
		
		await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "esm_report_id=" + selectedE2bReportID);
	}

	async function fn_RejectSingleE2BCase() {
		var strURL, notes, esm_report_status;
	
		if (selectedReportTypeIFN == 3 || selectedReportTypeIFN == 5 || selectedReportTypeIFN == 6) //FollowUp
			strURL = "/E2B/Actions/RejectFollowupE2BCase.asp?fType=1";
		else if (selectedReportTypeIFN == 4) //Nullification
		{
			strURL = "/E2B/Actions/RejectFollowupE2BCase.asp?fType=2";
		}
		else  //Initial
			strURL = "/E2B/Actions/RejectE2BCase.asp";
		var sDialogStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };
		notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);
		if (!notes)
		{
			await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + selectedE2bReportID);
		}
		else
		{
			if(notes !== sSessionTimeOutDialogReturn)
			{
				document.Frm_E2B.notes.value = notes;
				document.Frm_E2B.action = "/E2B/Incoming/E2bPending.asp?rb_report_status=" + selectedReportTypeIFN + "&esm_report_id=" + selectedE2bReportID + "&esm_status=103&bSearch=1&<%=GetRequestKeyValue()%>";
				fn_ValidateSubmitForm(document.Frm_E2B);
			}
		}
	}
	
	//To view E2B Report
	function fn_ViewE2b()
	{
		if (selectedE2bReportID > 0)
		{
			strURL = "/E2B/E2bViewer/E2bViewer.asp?CaseID=" + fn_URLEncode(selectedEncryptWorldWideUNum) + "&ReportID=<%=Server.URLEncode(query_Crypt.Encrypt(-1, GetEncryptKey())) %>" + "&ProductID=-1" + "&LicenseID=-1" + "&CountryID=-1" + "&AgencyID=-1" + "&Draft=-1" + "&TimeFrame=-1" + "&E2BViewType=-1" +"&IncomingE2b=1&E2bReportID=" + fn_URLEncode(selectedEncryptE2bReportID) + "&ClickViewReport=0" + "&AuthorityId=" + selectedE2BAthorityId + "&ApplyNewFw=" + selectedIsApplyNewFw + "&<%=GetRequestKeyValue()%>";
			window.open(strURL);
		}
	}
	   
	//To view Warning	
	async function fn_Warning()
	{
		pending_action = 2;
		await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "check_lock_status=1&esm_report_id=" + selectedE2bReportID);

	}
	async function fn_CallWarning()
	{
		var esm_report_status,strURL;
		var Pos, notes, warning, status;
		var custom_import = <%=JavaScriptSanitize(custom_import) %>;
		if(selectedE2bReportID < 0 ) return;
		if (custom_import == "1")
			strURL = "/E2B/Misc/GetLoadE2bStagingWarning.asp?check_warning=1&custom_import=1&esm_report_id=" + selectedE2bReportID
		else
			strURL = "/E2B/Misc/GetLoadE2bWarning.asp?check_warning=1&esm_report_id=" + selectedE2bReportID
		showLoading();
		var sDialogStyle = { dialogHeight: "560", dialogWidth: "840", resizable: false, scrollable: false };
		notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);
		if((notes) && (notes !== sSessionTimeOutDialogReturn))
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

	async function e2b_select() {
		var strURL, strStyle, oOutput, lCriteria;
		var aParam = new Array();

		lCriteria = 1;
		if (document.all.chk_product_name[1].checked) {
			lCriteria = 2;
		}
		aParam[0] = document.all.selProduct.value;

		strURL = "/E2b/E2bImport/FindProductName.asp?DSPLYLNG=" + glDisplayLang + "&criteria=" + lCriteria;
		strStyle = { dialogHeight: "320", dialogWidth: "480", resizable: false, scrollable: false };
		oOutput = await fn_OpenModalDialog(strURL, aParam, strStyle);
		if (oOutput) {
			document.all.selProduct.value = oOutput;
		}
	}
</script>
<!-- Javascript Functions Ends -->

<!-- Server-Side Functions Starts -->
<script language="vbscript" runat="server">

'**************************************************************************************************
' Author	    : Pankaj Gautam
' Called From   : On Page load
' Parameters	: None
' Returns	    : None
' Description	: Initializes all page level variables
'**************************************************************************************************
Function e2bPendingStsRpt_Initialize()
	e2bPendingStsRpt_SetLMTblDateRangeFldName()
	
	e2bPendingStsRpt_sFormName = "Frm_E2B"
	e2bPendingStsRpt_lCurrRec = 0
	e2bPendingStsRpt_lPageRowCount = 0
	
	sFormName = "Frm_E2BPending"
	sRequestXML = sFormName & "_Request"
	sResponseXML = sFormName & "_Response"		
	
	e2bPendingStsRpt_bPrevSort	   = CInt(NVL(Request("PrevSort").Item, "1"))
	e2bPendingStsRpt_bCurrentSort  = CInt(NVL(Request("CurrentSort").Item, "8"))               'Default Sort
	e2bPendingStsRpt_bSortOrder	   = CInt(NVL(Request("SortOrder").Item, "0"))
	e2bPendingStsRpt_bSearch	   = CInt(NVL(Request("bSearch").Item, "0"))                   'By default It will fetch up all the data
	'e2bPendingStsRpt_bViewAll	   = CInt(NVL(Request("bViewAll").Item, "0"))
	e2bPendingStsRpt_bPaginate	   = CInt(NVL(Request(e2bPendingStsRpt_sFormName & "_Paginate").Item, "0"))
	
	e2bPendingStsRpt_sDateFrom	   = NVL(Request("DateFrom").Item, "")
	e2bPendingStsRpt_sDateTo	   = NVL(Request("DateTo").Item, "")
	e2bPendingStsRpt_lRangeList	   = NVL(Request("RangeList").Item, "1")
	
	e2bTransStsRpt_lGmtOffSet	   = CDbl(NVL(Request("GmtOffSet").Item, "0"))
	If e2bTransStsRpt_lGmtOffSet   = 0 then e2bTransStsRpt_lGmtOffSet = GetXMLValueDirect (oSession, "GMT")
	
	e2bPendingStsRpt_TradingPrtnr  = NVL(Request("ReportingAgency").Item, "")
	
	'Getting value of selected radio button
	e2bPendingStsRpt_oProdGenRadio = NVL(Request("chk_product_name"), "0")
	
	If e2bPendingStsRpt_oProdGenRadio = "0" Then
		e2bPendingStsRpt_ProdName = NVL(Request("selProduct").Item, "")
	Else
		e2bPendingStsRpt_GenericName = NVL(Request("selProduct").Item, "")    
	End If
	
	e2bPendingStsRpt_RptType     = NVL(Request("selRType").Item, "")
	e2bPendingStsRpt_MsgType     = NVL(Request("selectMessageType").Item, "")
	e2bPendingStsRpt_HdnMsgType_Desc = NVL(Request("hdn_Message_type_Desc").Item, "")
	e2bPendingStsRpt_DateType    = NVL(Request("DateType").Item, "")
	ReportingAgencyKey = Request.Form("ReportingAgencyKey")
	ReportingAgencyValue = Request.Form("ReportingAgencyValue")
   
	'Called From Menu
	'Getting user preferences from XML
	If trim(request.QueryString("CallFrom")) = "Menu" then
		e2bPendingStsRpt_Status = GetUserPreferences("RPT_E2B_STATUS_TEXT_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		e2bPendingStsRpt_RptType = GetUserPreferences("RPT_E2B_REPORT_TYPE_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		e2bPendingStsRpt_oProdGenRadio = GetUserPreferences("RPT_PROD_GEN_RADIO_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		'If not getting any node in XML like user 1st time then select 0 "Product as Default"
		if e2bPendingStsRpt_oProdGenRadio = "" Then
			e2bPendingStsRpt_oProdGenRadio = "0"
		End If
			
		If e2bPendingStsRpt_oProdGenRadio = "0" Then
			e2bPendingStsRpt_ProdName = GetUserPreferences("RPT_E2B_MEDICINAL_PRODUCT_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		Else
			e2bPendingStsRpt_GenericName = GetUserPreferences("RPT_E2B_ACTIVESUBSTANCE_NAME_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		End If
		
		e2bPendingStsRpt_sDateFrom = fn_date_from_iso(GetUserPreferences("RPT_E2B_DATE_FROM_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc), 8, false)
		e2bPendingStsRpt_sDateTo = fn_date_from_iso(GetUserPreferences("RPT_E2B_DATE_TO_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc), 8, false)
		e2bPendingStsRpt_DateType = GetUserPreferences("RPT_E2B_DATE_TYPE_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		
		e2bPendingStsRpt_lRangeList = GetUserPreferences("RPT_E2B_DATE_RANGE_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		ReportingAgencyKey = GetUserPreferences("ReportingAgencyKey_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		ReportingAgencyValue = GetUserPreferences("ReportingAgencyValue_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		e2bPendingStsRpt_sOrder = GetUserPreferences("E2BPEN_SORT_ORDER_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		e2bPendingStsRpt_bCurrentSort = GetUserPreferences("E2BPEN_SORT_CURRENT_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		e2bPendingStsRpt_bSortOrder = GetUserPreferences("E2BPEN_SORT_ASC_DESC_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		e2bPendingStsRpt_MsgType = GetUserPreferences("RPT_E2B_MESSAGE_TYPE_ID_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		'e2bPendingStsRpt_HdnMsgType_Desc = GetUserPreferences("RPT_E2B_MESSAGE_TYPE_DESC_PEN", e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
		'If It is getting no crossponding Node
		If e2bPendingStsRpt_bSortOrder ="" Then
			e2bPendingStsRpt_bSortOrder = 0
		End If 
		If IsNullOREmpty(e2bPendingStsRpt_bCurrentSort) Then
			e2bPendingStsRpt_bCurrentSort = 8
		End If
	End If
   
	e2bPendingStsRpt_lOffSet = CInt(NVL(Request("OffSet").Item, "0"))		
	'Getting DropDown values
	e2bPendingStsRpt_GetDropDowns()
	
	If (e2bPendingStsRpt_bSortOrder = 0) Then	'ASC
		e2bPendingStsRpt_sArrow = "/img/common/order_up.gif"
		e2bPendingStsRpt_lNextSortOrder = 1
	Else
		e2bPendingStsRpt_sArrow = "/img/common/order_down.gif"
		e2bPendingStsRpt_lNextSortOrder = 0
	End If

	e2bPendingStsRpt_GetE2bPendingResult()
	e2bPendingStsRpt_InitializeSearchResultsCache()
End Function

'**************************************************************************************************
' Author	    : Pankaj Gautam
' Called From   : On Page load
' Parameters	: None
' Returns	    : None
' Description	: populates the date range drop down
'**************************************************************************************************
Function e2bPendingStsRpt_GetDropDowns()	
	Set e2bPendingStsRpt_oDateRangeMsg = RetrieveDropdownNoCache("LM_DATE_RANGES", e2bPendingStsRpt_lErrorNum, e2bPendingStsRpt_sErrorDesc)
	Set e2bPendingStsRpt_oDateRangeList = e2bPendingStsRpt_oDateRangeMsg.selectnodes("/MESSAGE/TABLE_LM_DATE_RANGES/LM_DATE_RANGES")
	
	'Take Default DateRange as Last 7 Days
	If e2bPendingStsRpt_lRangeList = -2 Then
		For Each e2bPendingStsRpt_oDate in e2bPendingStsRpt_oDateRangeList
			e2bPendingStsRpt_lRangeList = GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID")
			e2bPendingStsRpt_sDateFrom = fn_date_from_iso(GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_FROM_DATE"), 8, false)
			e2bPendingStsRpt_sDateTo  = fn_date_from_iso(GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_TO_DATE"), 8, false)
			If UCase(Trim(GetXMLValueDirect(e2bPendingStsRpt_oDate, "LM_DATE_RANGES_RANGE_NAME"))) = UCase("Last 7 Days") Then Exit For
		Next
	End if
End Function

'**************************************************************************************************
' Author	    : Pankaj Gautam
' Called From   : On search button click
' Parameters	: None
' Returns	    : None
' Description	: Fetches the entire data in the grid according to the search parameters
'**************************************************************************************************
Function e2bPendingStsRpt_GetE2bPendingResult()
	Dim oMessage, oMessage1, oMessageTypeDesc, oMessageTypeDdl, lMError, sMError
	Dim sqlSelect
	Dim sqlWhere
	Dim sSql, sSearchName, messageTypeDesc, messageTypeId
	Dim dateFormat
	
	sqlSelect = ""
	sqlWhere = ""	
	
	If isNullOrEmpty(e2bPendingStsRpt_bCurrentSort) Then
		e2bPendingStsRpt_bCurrentSort = 8                 'Default sort will be on Initial receipt date 
	End If

	If (e2bPendingStsRpt_bSortOrder = 1) Then
		e2bPendingStsRpt_sSortOrder = "DESC"
	Else
		e2bPendingStsRpt_bSortOrder = 0
		e2bPendingStsRpt_sSortOrder = "ASC"
	End If
	
	'Based on these value query string will get created in where clause
	Select Case (e2bPendingStsRpt_bCurrentSort)              
		Case 1  : e2bPendingStsRpt_sField = "upper(TradingPartner)" 
		Case 2  : e2bPendingStsRpt_sField = "upper(ReportTypeIFN)" 
		Case 3  : e2bPendingStsRpt_sField = "upper(companynumb)"
		Case 4  : e2bPendingStsRpt_sField = "upper(SenderCaseNum)"
		Case 5  : e2bPendingStsRpt_sField = "m.messageheader.messagedate"
		Case 6  : e2bPendingStsRpt_sField = "e.edi_complete_date"
		Case 7  : e2bPendingStsRpt_sField = "m.DATE_RECEIVED"
		Case 8  : e2bPendingStsRpt_sField = "s.RECEIPTDATE"
		Case 9  : e2bPendingStsRpt_sField = "upper(country)"
		Case 10 : e2bPendingStsRpt_sField = "upper(reporttype)"
		Case 11 : e2bPendingStsRpt_sField = "upper(ImportedCaseNumber)"
		Case 12 : e2bPendingStsRpt_sField = "upper(ProductName)"
		Case 13 : e2bPendingStsRpt_sField = "upper(GenericName)"
		Case 14 : e2bPendingStsRpt_sField = "upper(state_name)"
		Case 15 : e2bPendingStsRpt_sField = "upper(eventpt)"
		Case 16 : e2bPendingStsRpt_sField = "upper(eventllt)"
		Case 17 : e2bPendingStsRpt_sField = "upper(site_desc)"
		Case 18 : e2bPendingStsRpt_sField = "upper(patientinitial)"
		Case 19 : e2bPendingStsRpt_sField = "upper(StudyIdPatId)"
		Case 20 : e2bPendingStsRpt_sField = "upper(reportertype)"
		Case 21 : e2bPendingStsRpt_sField = "upper(Reporter)"
	End Select
	
	'If Call is not from Menu then save the preferences in XML
	if trim(request.QueryString("CallFrom")) <> "Menu" then
		e2bPendingStsRpt_sOrder =  e2bPendingStsRpt_sField & " " & e2bPendingStsRpt_sSortOrder
	End If
	sqlWhere = ""
	sSearchName = "" 
	
	If e2bPendingStsRpt_oProdGenRadio = "0" Then
		IF trim(e2bPendingStsRpt_ProdName) <> "" Then
			sSearchName = e2bPendingStsRpt_ProdName
			sqlWhere = sqlWhere & " AND (CASE WHEN :P_DISPLAY_LANGUAGE = 1 THEN UPPER(TRIM(NVL2(fgd.PRODUCT_NAME_J, fgd.PRODUCT_NAME_J, fgd.PRODUCT_NAME))) ELSE UPPER(TRIM(fgd.PRODUCT_NAME)) END) =  UPPER(:SEARCH_NAME) " 
		End IF
	Else
		IF trim(e2bPendingStsRpt_GenericName) <> "" Then
			sSearchName = e2bPendingStsRpt_GenericName
			sqlWhere = sqlWhere & " AND (CASE WHEN :P_DISPLAY_LANGUAGE = 1 THEN UPPER(TRIM(NVL2(fgd.GENERIC_NAME_J, fgd.GENERIC_NAME_J, fgd.GENERIC_NAME))) ELSE UPPER(TRIM(fgd.GENERIC_NAME)) END) =  UPPER(:SEARCH_NAME) " 
		End IF    
	End If
	
	IF trim(e2bPendingStsRpt_Status) <> "" Then
	   sqlWhere = sqlWhere & " AND s.Status =" & GetLong(e2bPendingStsRpt_Status, 0)
	End IF
	
	IF trim(e2bPendingStsRpt_RptType) <> "-1" Then
		IF trim(e2bPendingStsRpt_RptType) = "4" Then
			sqlWhere = sqlWhere & " AND E2B_Type in (4, 6)"
		Else
			IF trim(e2bPendingStsRpt_RptType) <> "7" Then
				sqlWhere = sqlWhere & " AND E2B_Type =" & GetLong(e2bPendingStsRpt_RptType, 0)
			Else
				sqlWhere = sqlWhere & " AND FGD.URGENT_REPORT = 1 "
			End If
	End IF
	End IF
	
	IF trim(e2bPendingStsRpt_MsgType) <> "-1"   Then 
		messageTypeDesc = trim(e2bPendingStsRpt_HdnMsgType_Desc)
		messageTypeId = GetLong(e2bPendingStsRpt_MsgType, 1) ' Defaulted to ICHICSR
		sqlWhere = sqlWhere & " AND (m.MESSAGEHEADER.BATCHMESSAGETYPE = :MESSAGE_TYPE_ID OR m.MESSAGEHEADER.MESSAGETYPE = :MESSAGE_TYPE_DESC) "
	End IF

	IF trim(ReportingAgencyValue) <> "" Then
		sqlWhere = sqlWhere & " AND s.agency_id IN (" & cfCmn_FindRegEx(ReportingAgencyKey, "0-9,") & ") " 
	End IF
	
	IF trim(e2bPendingStsRpt_DateType) <> "" Then
		IF (IsDate(e2bPendingStsRpt_sDateFrom) AND IsDate(e2bPendingStsRpt_sDateTo)) then
			dateFormat = GetDateFormatForFromToDateFields ()
		
			sqlWhere = sqlWhere & " AND trunc(" & cfCmn_FindRegEx(e2bPendingStsRpt_DateType, "-A-Z0-9._/\\()+':, ") & ") >= TO_DATE('" & e2bPendingStsRpt_sDateFrom & "','"& dateFormat & "')"
		
			sqlWhere = sqlWhere & " AND trunc(" & cfCmn_FindRegEx(e2bPendingStsRpt_DateType, "-A-Z0-9._/\\()+':, ") & ") <= TO_DATE('" & e2bPendingStsRpt_sDateTo & "','"& dateFormat & "')"
		End If
	End IF
	
	If glUserType = gUSERTYPE_EN Then
		sqlWhere = sqlWhere & " AND cp.authority_id <> 4"
	End If    

	Filter = "  "  & sqlWhere
	If trim(e2bPendingStsRpt_sOrder) <> "" Then
		Filter = Filter & " order by " & e2bPendingStsRpt_sOrder
	End If

	e2bPendingStsRpt_lPageSize     = CInt(NVL(Request(e2bPendingStsRpt_sFormName & "_PageSize").Item, "100"))
	e2bPendingStsRpt_lCurrentPage  = CInt(NVL(Request(e2bPendingStsRpt_sFormName & "_CurrentPage").Item, "1"))
	e2bPendingStsRpt_lStartRow = ((e2bPendingStsRpt_lCurrentPage - 1) * e2bPendingStsRpt_lPageSize) + 1
	e2bPendingStsRpt_lEndRow = e2bPendingStsRpt_lCurrentPage * e2bPendingStsRpt_lPageSize

	sUserFullName = GetXMLValueDirect (oSession, "CFG_USERS_USER_FULLNAME")
	Call CreateMessage (oMessageIn, 300400081)
	Call SetXMLValueDirect (oMessageIn, "GN_GUI_LM_GENERAL_TEXT", Filter) 'Adding Filter Text
	Call SetXMLValueDirect (oMessageIn, "GN_THIS_IS_ASP", 1)
	Call SetXMLValueDirect (oMessageIn, "GN_RPT_GMT_OFFSET", e2bTransStsRpt_lGmtOffSet)
	Call SetXMLValueDirect (oMessageIn, "CFG_USERS_USER_FULLNAME", sUserFullName)
	Call SetXMLValueDirect (oMessageIn, "CFG_USERS_PROD_SECURITY", GetXMLValueDirect (oSession, "CFG_USERS_PROD_SECURITY"))
	Call SetXMLValueDirect (oMessageIn, "CFG_USERS_STUDY_SECURITY", GetXMLValueDirect (oSession, "CFG_USERS_STUDY_SECURITY"))
	Call SetXMLValueDirect (oMessageIn, "LM_PRODUCT_PROD_NAME", sSearchName)
	Call SetXMLValueDirect (oMessageIn, "GN_STRING2", messageTypeDesc)
	Call SetXMLValueDirect (oMessageIn, "GN_STRING3", messageTypeId)
	Call SetXMLValueDirect (oMessageIn, "GN_DB_LIST_START", e2bPendingStsRpt_lStartRow)
	Call SetXMLValueDirect (oMessageIn, "GN_DB_LIST_END", e2bPendingStsRpt_lEndRow)
	
	'Saving search Criteria In Xml
	If ReportingAgencyValue = "" Then
		sAgencyList = GetTranslationData("ALL")
		Call SetXMLValueDirect (oMessageIn, "RPT_E2B_LOCAL_COMPANY_NAME", sAgencyList)
	Else
		sAgencyList = ReportingAgencyValue
		Call SetXMLValueDirect (oMessageIn, "RPT_E2B_LOCAL_COMPANY_NAME", sAgencyList)
	End If    
	sReportType = Replace(Replace(Replace(Replace(Replace(replace(e2bPendingStsRpt_RptType,"-1",GetTranslationData("ALL")),"1",GetTranslationData("INITIAL")),"3",GetTranslationData("E2B_RPT_TYPE_FOLLOW_UP")),"4",GetTranslationData("E2B_RPT_TYPE_NULLIFICATION")),"5",GetTranslationData("WL_REP_DOWNGRADE")),"7",GetTranslationData("URGENT_REPORT"))
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_REPORT_TYPE", sReportType)
	Call SetXMLValueDirect (oMessageIn, "GN_STRING1",Replace(Replace(Replace(Replace(Replace(replace(e2bPendingStsRpt_RptType,"-1",GetTranslationData("ALL")),"1",GetTranslationData("INITIAL")),"3",GetTranslationData("E2B_RPT_TYPE_FOLLOW_UP")),"4",GetTranslationData("E2B_RPT_TYPE_NULLIFICATION")),"5",GetTranslationData("WL_REP_DOWNGRADE")),"7",GetTranslationData("URGENT_REPORT")))
  
	Call SetXMLValueDirect (oMessageIn, "CFG_MESSAGE_TYPE_ID", e2bPendingStsRpt_MsgType)

	Call SetXMLValueDirect (oMessageIn, "CFG_MESSAGE_TYPE_DESCRIPTION", e2bPendingStsRpt_HdnMsgType_Desc)

	If e2bPendingStsRpt_oProdGenRadio = 0 Then
		sProdName = Replace(e2bPendingStsRpt_ProdName,"-1",GetTranslationData("ALL"))
		sStatus = ""
		Call SetXMLValueDirect (oMessageIn, "GN_GUI_PROD_NAME_1", sProdName)
		Call SetXMLValueDirect (oMessageIn, "RPT_E2B_STATUS_TEXT", sStatus)
	Else
		sProdName = ""
		sStatus = Replace(e2bPendingStsRpt_GenericName,"-1","All")
		Call SetXMLValueDirect (oMessageIn, "RPT_E2B_STATUS_TEXT", sStatus)
		Call SetXMLValueDirect (oMessageIn, "GN_GUI_PROD_NAME_1", sProdName)
	End If
	
	IF e2bPendingStsRpt_DateType= "to_date(m.messageheader.messagedate,'yyyy-mm-dd hh24:mi:ss')" THEN
		sDataType =  GetTranslationData("DATE_RANGE_TRANS_DATE")
		Call SetXMLValueDirect (oMessageIn, "GN_RPT_DATE_STRING", sDataType )
	ElseIf e2bPendingStsRpt_DateType= "to_date(e.edi_complete_date,'yyyy-mm-dd')" THEN
		sDataType = GetTranslationData("DATE_RANGE_MDN_SENT_DATE")
		Call SetXMLValueDirect (oMessageIn, "GN_RPT_DATE_STRING", sDataType  )
	ElseIf e2bPendingStsRpt_DateType= "to_date(s.RECEIPTDATE,'yyyy-mm-dd')" THEN  
		sDataType = GetTranslationData("DATE_RANGE_CASE_RCPT_DATE")
		Call SetXMLValueDirect (oMessageIn, "GN_RPT_DATE_STRING", sDataType  )  
	ELSE
		sDataType = GetTranslationData("DATE_RANGE_INTERCHNG_PROC_DATE")
		Call SetXMLValueDirect (oMessageIn, "GN_RPT_DATE_STRING",  sDataType )
	END IF
	
	sCopyDateFrom = e2bPendingStsRpt_sDateFrom
	sCopyDateTo = e2bPendingStsRpt_sDateTo
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_DATE_FROM", e2bPendingStsRpt_sDateFrom)
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_DATE_TO", e2bPendingStsRpt_sDateTo)
	Call SetXMLValueDirect (oMessageIn, "GN_DISPLAY_LANGUAGE", glDisplayLang)
	
	Set e2bPendingStsRpt_oOutMessage = ServiceRequest(oArgusSvr, oMessageIn, e2bPendingStsRpt_lErrorNum, e2bPendingStsRpt_sErrorDesc)
			
	If e2bPendingStsRpt_lErrorNum <> 0 Then
		Call ExecuteErrorPage (e2bPendingStsRpt_lErrorNum, e2bPendingStsRpt_sErrorDesc)
	End If

	lCacheId = GetXMLValueDirect(e2bPendingStsRpt_oOutMessage, "GN_DB_CACHE_ID")
	e2bPendingStsRpt_lResultsRowCount = GetXMLValueDirect(e2bPendingStsRpt_oOutMessage, "TABLE_RPT_E2B/GN_DB_LIST_LENGTH")
	If IsNullOrEmpty(e2bPendingStsRpt_lResultsRowCount) Then
		e2bPendingStsRpt_lResultsRowCount = 0
	Else
		e2bPendingStsRpt_lResultsRowCount = CLng(e2bPendingStsRpt_lResultsRowCount)
	End If
			
	WriteMessageToCache sResponseXML, e2bPendingStsRpt_oOutMessage, true
	
	'Save Selection  Criteria
	Call SetUserPreferences("RPT_E2B_STATUS_TEXT_PEN", e2bPendingStsRpt_Status,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_REPORT_TYPE_PEN", e2bPendingStsRpt_RptType,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_PROD_GEN_RADIO_PEN", e2bPendingStsRpt_oProdGenRadio,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	
	If  e2bPendingStsRpt_oProdGenRadio = "0" Then
		Call SetUserPreferences("RPT_E2B_MEDICINAL_PRODUCT_PEN", e2bPendingStsRpt_ProdName,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Else
		Call SetUserPreferences("RPT_E2B_ACTIVESUBSTANCE_NAME_PEN", e2bPendingStsRpt_GenericName,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	End If
	
	Call SetUserPreferences("RPT_E2B_DATE_FROM_PEN", e2bPendingStsRpt_sDateFrom,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_DATE_TO_PEN", e2bPendingStsRpt_sDateTo,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_DATE_TYPE_PEN", e2bPendingStsRpt_DateType,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_DATE_RANGE_PEN", e2bPendingStsRpt_lRangeList,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("ReportingAgencyKey_PEN", ReportingAgencyKey,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("ReportingAgencyValue_PEN", ReportingAgencyValue,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("E2BPEN_SORT_ORDER_PEN", e2bPendingStsRpt_sOrder,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("E2BPEN_SORT_CURRENT_PEN", e2bPendingStsRpt_bCurrentSort,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("E2BPEN_SORT_ASC_DESC_PEN", e2bPendingStsRpt_bSortOrder,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_MESSAGE_TYPE_ID_PEN", e2bPendingStsRpt_MsgType,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	'Call SetUserPreferences("RPT_E2B_MESSAGE_TYPE_DESC_PEN", e2bPendingStsRpt_HdnMsgType_Desc,e2bPendingStsRpt_lErrorNum,e2bPendingStsRpt_sErrorDesc)
	
	E2BViewType = GetValueFromCMN_PROFILEOnKey ("E2B_VIEW_PRINT")	
End Function


'**************************************************************************************************
' Author	    : Pankaj Gautam
' Called From   : On Reload after clicking Search or Sort
' Parameters	: None
' Returns	    : None
' Description	: Initializes the Grid component with the XML message
'**************************************************************************************************	
Function e2bPendingStsRpt_InitializeSearchResultsCache()

	Dim sIconNonExprep, sEmptyImage
	Dim e2bPendingStsRpt_sHtmlTemplate

	' Set paths for status icons
	sIconNonExprep = "/img/CF/zoom.gif"
	sEmptyImage = "/img/CF/Empty.gif"

	' Initialize Search Results Grid
	e2bPendingStsRpt_sHtmlTemplate = ReadTextFromFile("E2BPendingTemplate.tpl")	
	
	Set e2bPendingStsRpt_oGridControl = Server.CreateObject("Relsys.Argus.Interop.Web.GridControl")
    e2bPendingStsRpt_oGridControl.SetEncryptKey(GetEncryptKey())
	e2bPendingStsRpt_oGridControl.SetTemplate(e2bPendingStsRpt_sHtmlTemplate)
	e2bPendingStsRpt_oGridControl.SetDataSource GetMessageCacheFileName(sResponseXML, true), "RPT_E2B"
	
	e2bPendingStsRpt_oGridControl.AddPlaceholder "ROW_CLASS", "row-normal", "row-alternate"
	e2bPendingStsRpt_oGridControl.AddPlaceholder "ROW_COLOR", "#FFFFFF", "#E9E9F3"
			
	e2bPendingStsRpt_lPageRowCount = 0
	e2bPendingStsRpt_lPageRowCount = e2bPendingStsRpt_oGridControl.GetRowCount()
	
	If (e2bPendingStsRpt_lEndRow > e2bPendingStsRpt_lPageRowCount) Then
		e2bPendingStsRpt_lEndRow = e2bPendingStsRpt_lPageRowCount
	End If
		
	e2bPendingStsRpt_lCurrRec = e2bPendingStsRpt_lPageRowCount Mod e2bPendingStsRpt_lPageSize
	If (e2bPendingStsRpt_lEndRow - e2bPendingStsRpt_lStartRow + 1) = e2bPendingStsRpt_lPageSize Then
		e2bPendingStsRpt_lCurrRec = e2bPendingStsRpt_lPageSize
	End If
End Function

Function ChangeStatusE2bReports()
	Dim oMessage, oOutStatusMsg

	Call CreateMessage (oMessage, 300100262)
	Call SetXMLValueDirect (oMessage, "GN_GUI_NEW_SAVE_ID", esm_report_id)
	Call SetXMLValueDirect (oMessage, "GN_UI_JUSTIFICATION", notes)
	Call SetXMLValueDirect (oMessage, "GN_STATUS_NUMBER", esm_status)
	Call SetXMLValueDirect (oMessage, "CSAC_DATE_DONE", fn_date_from_iso_gmt(TodayNowGMT(), 8, true))
		
	Set oOutStatusMsg = ServiceRequest(oArgusSvr, oMessage, e2bPendingStsRpt_lErrorNum, e2bPendingStsRpt_sErrorDesc)
	
	'Fix Issue 3510
	If e2bPendingStsRpt_lErrorNum <> 0 Then
		ChangeStatusE2bReports = FALSE
		Exit Function
	End If
	ChangeStatusE2bReports = TRUE	
End Function

Function e2bPendingStsRpt_SetLMTblDateRangeFldName()
	If glUserType = gUSERTYPE_JP Then
		lmDateRangeFieldName = "LM_DATE_RANGES_RANGE_NAME_J"
	Else
		lmDateRangeFieldName = "LM_DATE_RANGES_RANGE_NAME"
	End If
End Function 

Function GetDateFormatForFromToDateFields ()
	Dim dateFormat
	If glUserType = gUSERTYPE_JP Then
		dateFormat = "yyyy/mm/dd"
	Else
		dateFormat = "dd-mm-yyyy"
	End If
	GetDateFormatForFromToDateFields = dateFormat
End Function

</script>
<!-- Server-Side Functions Ends -->

<!-- #INCLUDE VIRTUAL="/Nav/AGFooter_inc.asp" -->
