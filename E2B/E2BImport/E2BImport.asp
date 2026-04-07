<!--#INCLUDE VIRTUAL="/Nav/AGHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<!-- #INCLUDE VIRTUAL="/E2B/Incoming/E2B_SubTab_inc.asp" -->
<!-- Page Comments Starts -->
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : E2BImport.asp
' Description  : E2B Import
'******************************************************************************
' Revision History
' Date		Author		Description
' 28AUG2006 PS         Original
'******************************************************************************
%>
<html>
<!-- Page Comments Ends -->
<!DOCTYPE html>
<!-- HTML Header Starts -->
<head>
	<!-- Page Title -->
	<title><%= GetTranslationData("E2B_PROCES_PROCESSED_E2B_REPORT") %></title>
	<!-- Include Stylesheet here -->
	<link href="/CSS/relsys.css" type="text/css" rel="stylesheet" />
	<!-- Client Library Includes Starts -->
	<script type="text/javascript" src="/js/Menu/jmenu.js"></script>
	<script type="text/javascript" src="/js/E2B/E2B_Common.js"></script>
	<style type="text/css">
		TABLE.inner-table {
			outline: none;
		}
	</style>
	<!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
' Grid Control
Dim e2bImportStsRpt_oGridControl
Dim e2bImportStsRpt_sHtmlTemplate
Dim e2bImportStsRpt_sTableName
Dim e2bImportStsRpt_sFormName
Dim e2bImportStsRpt_lStartRow
Dim e2bImportStsRpt_lEndRow
Dim e2bImportStsRpt_lCurrentPage
Dim e2bImportStsRpt_lPageSize
Dim e2bImportStsRpt_lPageRowCount
Dim e2bImportStsRpt_lResultsRowCount

' General
Dim e2bImportStsRpt_bPaginate
Dim e2bImportStsRpt_sMessageType
DIm e2bImportStsRpt_sAgency
Dim e2bImportStsRpt_lRadioBtn
Dim e2bImportStsRpt_sMessageTo
Dim e2bImportStsRpt_sMessageFrom
Dim e2bImportStsRpt_lErrorNum
Dim e2bImportStsRpt_sErrorDesc
Dim e2bImportStsRpt_oContactMsg
Dim e2bImportStsRpt_oContactList
Dim e2bImportStsRpt_oContact
Dim e2bImportStsRpt_oOutMsg
Dim e2bImportStsRpt_oRecList
Dim e2bImportStsRpt_oRec
Dim e2bImportStsRpt_lNumRec
Dim e2bImportStsRpt_lCurrRec
Dim e2bImportStsRpt_lUserId		
Dim e2bImportStsRpt_oOutMessage
Dim e2bImportStsRpt_oReportList
Dim e2bImportStsRpt_oReport
Dim e2bImportStsRpt_bViewAll
Dim e2bImportStsRpt_bViewGroup
Dim e2bImportStsRpt_bPrevSort
Dim e2bImportStsRpt_bCurrentSort
Dim e2bImportStsRpt_bSortOrder
Dim e2bImportStsRpt_sSortOrder
Dim e2bImportStsRpt_sOrder
Dim e2bImportStsRpt_sArrow
Dim e2bImportStsRpt_lNextSortOrder
Dim e2bImportStsRpt_sField
Dim e2bImportStsRpt_lOffSet
Dim e2bImportStsRpt_lGmtOffSet
Dim e2bImportStsRpt_oDateRangeMsg
Dim e2bImportStsRpt_oDateRangeList
Dim e2bImportStsRpt_lRangeList
Dim e2bImportStsRpt_oDate
Dim e2bImportStsRpt_sDateFrom
Dim e2bImportStsRpt_sDateTo
Dim e2bImportStsRpt_bSearch
Dim e2bImportStsRpt_lMessageId
Dim e2bImportStsRpt_sOpenCase
Dim e2bImportStsRpt_TradingPrtnr 
Dim e2bImportStsRpt_ProdName     
Dim e2bImportStsRpt_Status       
Dim e2bImportStsRpt_RptType     
Dim e2bImportStsRpt_DateType     
Dim e2bImportStsRpt_Year
Dim e2bImportStsRpt_Month
Dim e2bImportStsRpt_Day   
Dim e2bImportStsRpt_FormatedDate  
Dim e2bTransStsRpt_lGmtOffSet     
Dim ReportingAgencyKey
Dim ReportingAgencyValue
Dim oMessageIn, Filter, OrderBy
Dim sFormName
Dim sRequestXML
Dim sResponseXML	
Dim E2BViewType, sSql
Dim e2bImportStsRpt_oStatusList, e2bImportStsRpt_oStatusMsg, oMessageCriteria
Dim sUserFullName
Dim sCopyDateFrom
Dim sCopyDateTo
Dim sProdName
Dim sDataType
Dim sAgencyList
Dim sReportType
Dim sStatus
Dim lCacheId

Dim e2bImportStsRpt_bPostBack
Dim lmDateRangeFieldName
Dim SuppressACKtransmission
%>
<!-- Declaration of Page Scope variables Ends -->
<!-- Assign Values to Page Scope variables Starts -->
<!-- Assign Values to Page Scope variables Ends -->
<!-- Page Processing Starts -->
<%
	e2bImportStsRpt_lPageRowCount = 0
	e2bImportStsRpt_Initialize()
	E2BViewType = GetValueFromCMN_PROFILEOnKey("E2B_VIEW_PRINT")
	if E2BViewType = "" then E2BViewType = "0"
%>
<!-- Page Display Starts -->
<body onload="Init_Frm();">
	<table class="table" cellspacing="0" cellpadding="0" border="0" style="width: 100%; height: 100%">
		<tr style="height: 25px;">
			<td>
				<!--#INCLUDE VIRTUAL="/Nav/AGToolbar_inc.asp" -->
			</td>
		</tr>
		<tr>
			<td class="valign-top">
				<form id="<%= e2bImportStsRpt_sFormName %>" name="<%= e2bImportStsRpt_sFormName %>"
					method="post" action="/E2B/E2BImport/E2BImport.asp" class="form-100">
					<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
					<table class="table" cellspacing="0" cellpadding="0" style="width: 100%; height: 100%;">
						<tr style="height: 25px;">
							<td class="padding-left-right">
								<%BuildLocalLabel ("E2B_PROCES_PROCESSED_E2B_REPORT").SetStyleSheet("label label-page-header").Render()%>
							</td>
						</tr>
						<tr>
							<td valign="top">
								<table class="table" cellspacing="0" cellpadding="0" id="Table_Outer_Box"
									style="width: 100%; height: 100%">
									<tr valign="top" style="height: 75px;">
										<td>
											<%Call BuildHiddenControlDirect("OffSet", "") %>
											<%Call BuildHiddenControlDirect("GmtOffSet", e2bImportStsRpt_lGmtOffSet) %>
											<%Call BuildHiddenControlDirect("PrevSort", e2bImportStsRpt_bCurrentSort) %>
											<%Call BuildHiddenControlDirect("CurrentSort", e2bImportStsRpt_bCurrentSort) %>
											<%Call BuildHiddenControlDirect("SortOrder", e2bImportStsRpt_bSortOrder) %>
											<%Call BuildHiddenControlDirect("sOrder", e2bImportStsRpt_sOrder) %>
											<%Call BuildHiddenControlDirect("bSearch", "") %>
											<%Call BuildHiddenControlDirect("PostBack", "1") %>
											<%Call BuildHiddenControlDirect("UserCacheId", lCacheId) %>

											<table class="table" cellspacing="0" cellpadding="0" width="100%" id="Table_Tab_Box">
												<!-- Section Header - General Starts -->
												<tr>
													<td valign="top">
														<% gSecHead_sLabel = GetTranslationData("SEARCH_CRT") %>
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
																				<span class="label label-no-padding"><%BuildLocalLabel("TRADING_PARTNER").Render()%></span>
																			</td>
																			<td style="float: right;">
																				<%BuildControlDirect(CTL_BUTTON, "btnFilter", GetTranslationData("REP_FILTER"), false, 0, "").OnClick("fn_addDestination();").Style("Width:60px").Render()
																			  BuildControlDirect(CTL_BUTTON, "btnClear", GetTranslationData("REMOVE_ALL"), false, 0, "").OnClick("fn_ClearAll();").Style("Width:70px").Render() %>
																			</td>
																		</tr>
																		<tr>
																			<td colspan="2">
																				<select name="ReportingAgency" class='ddlist-multi' size='5' tabindex='1' style='width: 100%; overflow: auto;'
																					multiple='multiple'>
																					<%If instr(ReportingAgencyKey,",") > 0 then
																				Dim ary, x, ArrVaue
																				ary = split(ReportingAgencyKey, ",")
																				ArrVaue = split(ReportingAgencyValue, "--*--")
																				For x=0 to ubound(ary)
																					Response.Write("<option value=""" & Fn_Sanitize(ary(x)) & """>" & Fn_Sanitize(ArrVaue(x)) & "</option>")
																				Next
																			 Else
																					Response.Write("<option value=""" & Fn_Sanitize(ReportingAgencyKey) & """>" & Fn_Sanitize(ReportingAgencyValue) & "</option>")   
																			 End If%>
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
																				<%BuildLocalLabel("PROD_NAME").Render() %>
																			</td>
																			<td>
																				<%BuildLocalLabel("IMPORT_STATUS").Render() %>
																			</td>
																			<td>
																				<%BuildLocalLabel("REP_TYPE").Render() %>
																			</td>
																		</tr>
																		<tr>
																			<td>
																				<input type="text" style="width: 80%" class="textbox" id="selProduct" name="selProduct"
																					value="<%=Fn_Sanitize(e2bImportStsRpt_ProdName) %>" tabindex="3" maxlength="2000">
																				<% BuildControlDirect(CTL_BUTTON, "btn_select", GetTranslationData("SELECT"), false, 4, "") _
																					   .OnClick("e2b_select();") _
																					   .Render() %>
																			</td>
																			<td>
																				<select name="selStatus" id="selStatus" class="ddlist" style="width: 100%">
																					<option value="102"><%=GetTranslationData("IMP_STATUS_SUCCESSFUL")%> </option>
																					<option value="103"><%=GetTranslationData("IMP_STATUS_FAILURE")%></option>
																				</select>
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
																					<col width="15%" />
																					<col width="15%" />
																					<col width="15%" />
																					<col width="15%" />
                                                                                    <col width="30%" />
																					<col width="10%" />
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
																						<td></td>
                                                                                        <td></td>
																					</tr>
																					<tr>
																						<td>
																							<select style="width: 100%" class="ddlist" name="DateType" tabindex="5">
																								<option value="to_date(NVL(s.transmissiondate,to_number(SUBSTR(s.TRANSMISSIONDATER3,1,8))),'yyyy-mm-dd')"><%=GetTranslationData("DATE_RANGE_TRANS_DATE") %></option>
																								<option value="to_date(s.DATE_IMPORTED + (<%=e2bTransStsRpt_lGmtOffSet%>/24))"><%=GetTranslationData("DATE_RANGE_IMP_DATE")%></option>
																								<option value="to_date(m.DATE_RECEIVED + (<%=e2bTransStsRpt_lGmtOffSet%>/24))"><%=GetTranslationData("DATE_RANGE_INTERCHNG_PROC_DATE")%></option>
																							</select>
																						</td>
																						<td>
																							<input type="text" style="width: 100%" class="textbox" id="DateFrom" name="DateFrom"
																								value="<%=Fn_Sanitize(e2bImportStsRpt_sDateFrom) %>" onchange="DateExit(this, false, false)"
																								onblur="DateBlur(this, false, false)" onkeydown="document.Frm_E2B.RangeList.value = -1;"
																								tabindex="6">
																						</td>
																						<td>
																							<input type="text" style="width: 100%" class="textbox" id="DateTo" name="DateTo"
																								value="<%=Fn_Sanitize(e2bImportStsRpt_sDateTo) %>" onchange="DateExit(this, false, false)"
																								onblur="DateBlur(this, false, false)" onkeydown="document.Frm_E2B.RangeList.value = -1;"
																								tabindex="7">
																						</td>
																						<td>
																							<select style="width: 100%" class="ddlist" name="RangeList" onchange="f_e2bImportStsRpt_ChangeDateRange(this.value);"
																								tabindex="5">
																								<%For Each e2bImportStsRpt_oDate in e2bImportStsRpt_oDateRangeList%>
																								<option value="<%=GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID")%>"
																									<%If e2bImportStsRpt_lRangeList = GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID") then%>Selected<%End If%>>
																									<%=Fn_Sanitize(GetXMLValueDirect(e2bImportStsRpt_oDate,lmDateRangeFieldName))%></option>
																								<%Next%>
																								<option value="-1" <%If "-1" = e2bImportStsRpt_lRangeList Then%>Selected<%End If%>><%=GetTranslationData("CUST_DT_RANGE")%></option>
																							</select>
																						</td>
                                                                                        <td>
                                                                                            <%BuildControlDirect(CTL_CHECKBOX, "chkSuppressACKtransmission",SuppressACKtransmission, false, 0, "").Render()%>
                                                                                            <%BuildLabelDirect(GetTranslationData("REPORTS_WITH_ACK_SUPPRESSION")).Render()%>
                                                                                          
                                                                                        </td>
																						<td>
																							<%BuildButton("btnRetrieve", "BULK_RPT_RETRIEVE", 9).Style("width:60px")_
																							.OnClick("fn_Search(1);").Render()%>
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
												<!-- Section Contents - General Ends -->
											</table>
										</td>
									</tr>
                                    <tr class="spacer"><td></td></tr>
									<tr style="height: 20px;">
										<!-- Pending/Prpcessed Sub Tab Section Starts -->
										<td>
										<%
										Dim cfPendingProcessedTabs 
										Set cfPendingProcessedTabs = BuildTabs ("CF_PENDING", 2, CONST_TABS_TOP, 2)
										cfPendingProcessedTabs.AddTab GetTranslationData("BULK_RPT_PENDING"), "f_CFN_SwitchTab(1)"
										cfPendingProcessedTabs.AddTab GetTranslationData("PROCESSED"), "f_CFN_SwitchTab(2)"
										cfPendingProcessedTabs.SetTabWidth "150px"
										cfPendingProcessedTabs.Render
										%>
										</td>
										<!-- Pending/Prpcessed Sub Tab Section Ends -->
									</tr>
									<tr>
										<td valign="top">
											<table class="table" cellspacing="0" cellpadding="0" width="100%" height="100%" id="Table3"
												border="0" valign="top">
												<!-- Section Header - Open Cases Search Results Starts -->
												<tr width="100%" style="height: 25px;">
													<td style="margin-top: 0px; margin-bottom: 0px" valign="top">
														<% 
														gSecHead_sGridName = e2bImportStsRpt_sFormName
														gSecHead_lGridRowCount = e2bImportStsRpt_lResultsRowCount
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
																<tr style="height: 40px">
																	<td>
																		<div id="Div1" class="table-scroll-hide" style="width: 100%; overflow-y: scroll;">
																			<table class="grd-table-header" cellspacing="0" cellpadding="2">
																				<col width="16%" />
																				<col width="15%" />
																				<col width="23%" />
																				<col width="22%" />
																				<col width="16%" />
																				<col />
																				<tr class="tblheader-lightblue" style="height: 42px">
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(1, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("ORIGINATED_CASE_NUM").Render()%> </span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 1) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(2, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("INITIAL_FU_NULLIFICATION").Render()%> </span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 2) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(3, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("TRADING_PARTNER").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 3) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(4, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("WORLD_WIDE_UNIQUE_NUM").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 4) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(5, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("IMP_STATUS_WARNING_ERRORS").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 5) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then%>
																										style="cursor: pointer" onclick="fn_SortHeader(6, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("CASE_NUM_IMPORTED_AS").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 6) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(7, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("ACCEPTED_REJECTED_BY").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 7) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(8, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("NOTES").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 8) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<td style="width: 100%">
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(9, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("INTERCHANGE_DATE").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 9) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding" <% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) Then %>
																										style="cursor: pointer" onclick="fn_SortHeader(10, <%= e2bImportStsRpt_lNextSortOrder %>);"
																										<% End If %>><%BuildLocalLabel("DATE_IMPORTED_REJECTED").Render()%></span>
																									<% If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPaginate = 1)) And (e2bImportStsRpt_bCurrentSort = 10) Then %>
																									<img src="<%=e2bImportStsRpt_sArrow %>" />
																									<% End If %>
																								</td>
																							</tr>
																						</table>
																					</td>
																					<td class="grd-header">
																						<table class="table" cellspacing="0" cellpadding="0" width="100%">
																							<tr>
																								<col width="65%" />
																								<td>
																									<span class="label label-no-padding"><%BuildLocalLabel("ACK").Render()%></span>
																								</td>
																								<td>
																									<span class="label label-no-padding"><%BuildLocalLabel("EDI").Render()%></span>
																								</td>
																							</tr>
																							<tr>
																								<td>
																									<span class="label label-no-padding"><%BuildLocalLabel("GEN").Render()%></span>
																								</td>
																								<td>
																									<span class="label label-no-padding"><%BuildLocalLabel("OUT").Render()%></span>
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
																	<td class="no-padding-left" bgcolor="white">
																		<div id="Div2" class="table-scroll" style="height: 100%; overflow-y: scroll">
																			<table id="" class="grd-table-body" cellspacing="0" cellpadding="2">
																				<col width="16%" />
																				<col width="15%" />
																				<col width="23%" />
																				<col width="22%" />
																				<col width="16%" />
																				<col />
																				<%        
																			If ((e2bImportStsRpt_bSearch = 1) Or (e2bImportStsRpt_bPostBack = "1")) Then
																				if gSecHead_lGridRowCount>0 then
																				 Response.Write(e2bImportStsRpt_oGridControl.RenderFullXML())
																				end if
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
							<td style="padding-left:5px">
								<table cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td style="width:100px">
											<span class="label label-no-padding"><%=GetTranslationData("STATE_LEGEND")%></span>
										</td>
										<td style="width:20px;">
											<img src="/img/E2B/Pending.gif" />
										</td>
										<td style="width:60px;">
											<span class="label label-no-padding"><%=GetTranslationData("IMP_STATUS_PENDING")%></span>
										</td>
										<td style="width:20px;">
											<img src="/img/E2B/Warning.gif" />
										</td>
										<td style="width:60px;">
											<span class="label label-no-padding"><%=GetTranslationData("IMP_STATUS_WARNINGS")%></span>
										</td>
										<td style="width:20px;">
											<img src="/img/E2B/OK.gif" />
										</td>
										<td style="width:60px;">
											<span class="label label-no-padding"><%=GetTranslationData("IMP_STATUS_SUCCESS")%></span>
										</td>
										<td style="width:20px;">
											<img src="/img/E2B/Error.gif" />
										</td>
										<td style="width:60px;">
											<span class="label label-no-padding"><%=GetTranslationData("IMP_STATUS_FAILED")%></span>
										</td>
										<td align="right" class="padding-left-right">
										    <% If (e2bImportStsRpt_lResultsRowCount > 0) Then %>
	    									<input class="button button-large" type='button' value='<%=GetTranslationData("PRINT_LIST")%>' onclick="fn_PrintList();" />
										    <% End If %>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</form>
			</td>
		</tr>
	</table>
	<!-- Case Title Ends -->
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
var selectedCaseNum = "";
var selectedAgency = "";
var selectedStatus = "";
var selectedTransId = -1;
var selectedBlobsize = -1;
var AjaxRemoveTx = 0;
var AjaxReTx = 0;
function fn_Search(bRefreshCache)
{
	showLoading();
	document.all.bSearch.value = bRefreshCache;
	fn_ValidateSubmitForm(document.forms[0]);
}
function fn_SortHeader(sort_field, sort_order) {
	document.all.CurrentSort.value = sort_field;
	document.all.SortOrder.value = sort_order;
	fn_Search(1);
}

async function fn_QuickyOpen(sCaseNum) {
	var sCaseNum;chkCLResults
	showLoading();
	await loadArgusMessage("/Worklist/Status/Ajax_CheckForCase.asp?CaseNum=" + fn_URLEncode(sCaseNum), fn_callbackCaseNum);
		
	if (chkCLResults == 0)
		await MessageBoxRes("CASE_SEARCH_FALIED");
	else 
		await fn_OpenCaseFormNum(sCaseNum);
	hideLoading();
	return false;
}

async function fn_callbackCaseNum()
{
	var xmlDoc = this.req.responseXML;
	var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
	var results;
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

function e2bImportStsRpt_CalBolbSize(blobsize)
{
   var min_blobsize;
   var lblobsize = parseInt(blobsize,10);
   
   min_blobsize = 9999;
   if ((lblobsize * 1) < (min_blobsize * 1))
	  min_blobsize = lblobsize;
   return min_blobsize;
}

<%
'**************************************************************************************************
' Author	    : Gaurav Naresh Mittal
' Called From   : On changing the date range value
' Parameters	: Date range value
' Returns	    : None
' Description	: Sets the date from and date to values according to the date range selected
'**************************************************************************************************
%>
async function f_e2bImportStsRpt_ChangeDateRange(lValue)
{
	var sStart, sStop;
	sStart = document.Frm_E2B.DateFrom.value;
	sStop  = document.Frm_E2B.DateTo.value;
	
	if (lValue == -1)
	{
		sStart = <%=JavaScriptSanitize(e2bImportStsRpt_sDateFrom) %>;
		sStop = <%=JavaScriptSanitize(e2bImportStsRpt_sDateTo) %>;
		document.Frm_E2B.RangeList.value = -1;		
		document.Frm_E2B.DateFrom.value = sStart;
		document.Frm_E2B.DateTo.value = sStop;
		<%
			If (IsDate(e2bImportStsRpt_sDateTo) and IsDate(e2bImportStsRpt_sDateFrom)) Then 
				If (CDate(e2bImportStsRpt_sDateTo) < CDate(e2bImportStsRpt_sDateFrom)) Then 
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
				Set e2bImportStsRpt_oDateRangeList = e2bImportStsRpt_oDateRangeMsg.selectnodes ("/MESSAGE/TABLE_LM_DATE_RANGES/LM_DATE_RANGES")
				For Each e2bImportStsRpt_oDate in e2bImportStsRpt_oDateRangeList
					e2bImportStsRpt_sDateFrom = fn_date_from_iso(GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_FROM_DATE"), 8, false)
					e2bImportStsRpt_sDateTo  = fn_date_from_iso(GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_TO_DATE"), 8, false)
			%>		
					case "<%= GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID") %>": 
						sStart = "<%=e2bImportStsRpt_sDateFrom%>"; 
						sStop = "<%=e2bImportStsRpt_sDateTo%>"; 
						document.Frm_E2B.RangeList.value = "<%= GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID") %>"; 
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
' Author	    : Gaurav Naresh Mittal
' Called From   : f_e2bImportStsRpt_SelectRow
' Parameters	: Row Index, New Background Color
' Returns	    : None
' Description	: Set the background color for the specified row
'**************************************************************************************************
%>	
function f_e2bImportStsRpt_SetRowColor(iRow, colorValue)
{
	eval ("document.all.data_row_"			    + iRow).style.background = colorValue;
	eval ("document.all.case_num_"			    + iRow).style.background = colorValue;
	eval ("document.all.agency_name_"			+ iRow).style.background = colorValue;
	eval ("document.all.edi_transmit_date_"		+ iRow).style.background = colorValue;
	eval ("document.all.transmit_date_"			+ iRow).style.background = colorValue;
	eval ("document.all.local_company_name_"	+ iRow).style.background = colorValue;
	eval ("document.all.status_text_"			+ iRow).style.background = colorValue;
}

<%
'**************************************************************************************************
' Author	   : Gaurav Naresh Mittal
' Called From  : When Print button is clicked
' Parameters   : None
' Returns	   : None
' Description  : gets the IDs of the selected checkboxes on the page
'**************************************************************************************************
%>
async function f_e2bImportStsRpt_GetCheckBoxString() 
{
	var rowCount = "<%=e2bImportStsRpt_lPageRowCount%>";
	var selectedIds = "";
	var selectedRowCount = 0;
	
	for (var i = 1; i <= <%=e2bImportStsRpt_lCurrRec %>; i++) {
		if (eval("document.all.checkbox_" + i).checked == true){
			var id = eval("document.all.checkbox_" + i).value;			
			selectedIds = selectedIds + id + ",";			
			selectedRowCount += 1;
		}
	}
	
	if (selectedRowCount == 0)
	{
		await MessageBoxRes("E2BTRANS_MIN_RECORD_ONE");
		return;
	}
	
	if (selectedIds.length > 0) 
		selectedIds = selectedIds.substring(0, selectedIds.length-1);
		
	f_e2bImportStsRpt_Print(selectedIds);
}

async function fn_addDestination()
{
    var strURL, strStyle, strResult;
    if (!fn_ValidateForm("document.<%= e2bImportStsRpt_sFormName %>"))
        return;
	strStyle = { dialogHeight: "480", dialogWidth: "720", resizable: false, scrollable: false };
	strURL   = "/Reports/BRBF/BBF_Agency_Filter.asp?E2BImport=ICSR&DSPLYLNG=" + glDisplayLang;
	strResult = await fn_OpenModalDialog(strURL, window, strStyle);
}

function fn_ClearAll()
{
    if (!fn_ValidateForm("document.<%= e2bImportStsRpt_sFormName %>"))
        return;
	document.all.ReportingAgency.length = 0;
	document.all.ReportingAgencyKey.value = "";
	document.all.ReportingAgencyValue.value = "";
}

async function Init_Frm()
{
	if(<%=JavaScriptSanitize(trim(e2bImportStsRpt_ProdName))%> != "")
		document.all.selProduct.value = <%=JavaScriptSanitize(e2bImportStsRpt_ProdName)%>;
	if(<%=JavaScriptSanitize(trim(e2bImportStsRpt_Status)) %> != "")
		document.all.selStatus.value = <%=JavaScriptSanitize(e2bImportStsRpt_Status) %>;
	if(<%=JavaScriptSanitize(trim(e2bImportStsRpt_RptType)) %> != "")
		document.all.selRType.value = <%=JavaScriptSanitize(e2bImportStsRpt_RptType) %>;
	if(<%=JavaScriptSanitize(trim(e2bImportStsRpt_DateType)) %> != "")
		document.all.DateType.value = <%=JavaScriptSanitize(e2bImportStsRpt_DateType)%>;
	
  await f_e2bImportStsRpt_ChangeDateRange(document.Frm_E2B.RangeList.value);
}

var ShowErrorText;
function GetErrorText()
{
	return ShowErrorText;
}

async function ShowError(idx)
{
	var  strURL, retObj, e;
	var sParams = new Object();
	eval("e = document.all.HDNERROR_"+idx+".value");
	ShowErrorText = e;
	strURL = "/E2B/E2BImport/showMessage.asp";
	sParams.window = this;
	var sDialogStyle = { dialogHeight: "480", dialogWidth: "720", resizable: false, scrollable: false };
	retObj = await fn_OpenModalDialog(strURL, sParams, sDialogStyle);
}

/*'************************************************************************************
' Author       : Prakash Singh
' Page         : ExpeditedReport.asp
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
    if (!fn_ValidateForm("document.<%= e2bImportStsRpt_sFormName %>"))
            return;
	strURL = "/Lookup/ReportGenerateAjaxIntrim.asp?ReportType=3&resultsxml=Frm_E2BImport_Response&cacheId=" + fn_URLEncode(document.all.UserCacheId.value);
	strURL += "&gmtOffSet=<%= Server.URLEncode(e2bTransStsRpt_lGmtOffSet) %>";
	strURL += "&displayLang=<%= Server.URLEncode(glDisplayLang) %>";
	strURL += "&userFullName=<%= Server.URLEncode(sUserFullName) %>";
	strURL += "&dateFrom=<%= Server.URLEncode(sCopyDateFrom) %>";
	strURL += "&dateTo=<%= Server.URLEncode(sCopyDateTo) %>";
	strURL += "&prodName=<%= Server.URLEncode(sProdName) %>";
	strURL += "&dateType=<%= Server.URLEncode(sDataType) %>";
	strURL += "&agencyList=<%= Server.URLEncode(sAgencyList) %>";
	strURL += "&strReportType=<%= Server.URLEncode(sReportType) %>";
    strURL += "&status=<%= Server.URLEncode(sStatus) %>";
    if ("<%= Server.URLEncode(SuppressACKtransmission) %>"== "1")
        strURL += "&strSuppressACKtransmission=Yes";
    else
        strURL += "&strSuppressACKtransmission=No";
    var sDialogStyle = { dialogHeight: "200", dialogWidth: "350", resizable: false, scrollable: false };
	await fn_OpenModalDialog(strURL, window, sDialogStyle);
}

function openE2BViewer(repID, encryptRepID, CompanyNumb)
{
	 fn_ViewE2b(CompanyNumb, <%=JavaScriptSanitize(E2bViewType)%>, repID , encryptRepID);
}

function fn_ViewE2b(CompanyNumb, E2bViewType, E2bReportID, EncryptE2bReportID)
{
	if (E2bReportID > 0)
	{
		strURL = "/E2B/E2bViewer/E2bViewer.asp?CaseID=" + fn_URLEncode(CompanyNumb) + "&ReportID=<%=Server.URLEncode(query_Crypt.Encrypt(-1, GetEncryptKey())) %>" + "&ProductID=-1" + "&LicenseID=-1" + "&CountryID=-1" + "&AgencyID=-1" + "&Draft=-1" + "&TimeFrame=-1" + "&E2BViewType=-1" + "&IncomingE2b=1&E2bReportID=" + fn_URLEncode(EncryptE2bReportID) + "&ClickViewReport=0&<%=GetRequestKeyValue()%>";
		window.open(strURL);
	}
}

async function e2b_select() {
	var strURL, strStyle, oOutput;
	var aParam = new Array();

	aParam[0] = document.all.selProduct.value;
	strURL = "/E2b/E2bImport/FindProductName.asp?DSPLYLNG=" + glDisplayLang + "&criteria=0";
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
' Author	    : Gaurav Naresh Mittal
' Called From   : On Page load
' Parameters	: None
' Returns	    : None
' Description	: Initializes all page level variables
'**************************************************************************************************
Function e2bImportStsRpt_Initialize()
	e2bImportStsRpt_SetLMTableDateRangeFieldNameForDropDownValues()

	e2bImportStsRpt_sFormName = "Frm_E2B"
	e2bImportStsRpt_lCurrRec = 0
	
	sFormName = "Frm_E2BImport"
	sRequestXML = sFormName & "_Request"
	sResponseXML = sFormName & "_Response"		

	e2bImportStsRpt_bPrevSort	= CInt(NVL(Request("PrevSort").Item, "1"))
	e2bImportStsRpt_bCurrentSort = CInt(NVL(Request("CurrentSort").Item, "3"))
	e2bImportStsRpt_bSortOrder	= CInt(NVL(Request("SortOrder").Item, "0"))
	e2bImportStsRpt_bSearch		= CInt(NVL(Request("bSearch").Item, "0"))
	
	e2bImportStsRpt_bPostBack = Request.Form ("PostBack")
	
	e2bImportStsRpt_bPaginate	= CInt(NVL(Request(e2bImportStsRpt_sFormName & "_Paginate").Item, "0"))
	e2bImportStsRpt_lPageSize    = CInt(NVL(Request(e2bImportStsRpt_sFormName & "_PageSize").Item, "100"))
	
	e2bImportStsRpt_lCurrentPage = CInt(NVL(Request(e2bImportStsRpt_sFormName & "_CurrentPage").Item, "1"))
		
	e2bImportStsRpt_sDateFrom	= NVL(Request("DateFrom").Item, "")
	e2bImportStsRpt_sDateTo		= NVL(Request("DateTo").Item, "")
	e2bImportStsRpt_lRangeList	= NVL(Request("RangeList").Item, "1")
	SuppressACKtransmission= CInt(NVL(Request("chkSuppressACKtransmission").Item, "0"))
	e2bTransStsRpt_lGmtOffSet	= CDbl(NVL(Request("GmtOffSet").Item, "0"))
	If e2bTransStsRpt_lGmtOffSet = 0 Then e2bTransStsRpt_lGmtOffSet = GetXMLValueDirect (oSession, "GMT")
	
	e2bImportStsRpt_TradingPrtnr = NVL(Request("ReportingAgency").Item, "")
	e2bImportStsRpt_ProdName    = NVL(Request("selProduct").Item, "")
	e2bImportStsRpt_Status      = NVL(Request("selStatus").Item, "")
	e2bImportStsRpt_RptType     = NVL(Request("selRType").Item, "")
	e2bImportStsRpt_DateType    = NVL(Request("DateType").Item, "")
	ReportingAgencyKey = Request.Form("ReportingAgencyKey")
	ReportingAgencyValue = Request.Form("ReportingAgencyValue")
   
	If trim(request.QueryString("CallFrom")) = "Menu" Then
		e2bImportStsRpt_Status = GetUserPreferences("RPT_E2B_STATUS_TEXT", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		e2bImportStsRpt_RptType = GetUserPreferences("RPT_E2B_REPORT_TYPE", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		e2bImportStsRpt_ProdName = GetUserPreferences("RPT_E2B_MEDICINAL_PRODUCT", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		e2bImportStsRpt_sDateFrom = GetUserPreferences("RPT_E2B_DATE_FROM", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		e2bImportStsRpt_sDateTo= GetUserPreferences("RPT_E2B_DATE_TO", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		e2bImportStsRpt_DateType = GetUserPreferences("RPT_E2B_DATE_TYPE", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		e2bImportStsRpt_lRangeList = GetUserPreferences("RPT_E2B_DATE_RANGE", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		ReportingAgencyKey = GetUserPreferences("ReportingAgencyKey", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
		ReportingAgencyValue = GetUserPreferences("ReportingAgencyValue", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
        SuppressACKtransmission=GetUserPreferences("ACK_TRANS_SUPPRESSED", e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	End If
	
	e2bImportStsRpt_lOffSet		= CInt(NVL(Request("OffSet").Item, "0"))
	e2bImportStsRpt_lGmtOffSet	= CDbl(NVL(Request("GmtOffSet").Item, "0"))
	If e2bImportStsRpt_lGmtOffSet = 0 Then e2bImportStsRpt_lGmtOffSet = GetXMLValueDirect (oSession, "GMT")
	
	If e2bImportStsRpt_bSearch = 1 OR e2bImportStsRpt_bPostBack = "1" Then
		e2bImportStsRpt_GetWorklistResult()
		Set e2bImportStsRpt_oReportList = e2bImportStsRpt_oOutMessage.selectnodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")	
	End If
	
	e2bImportStsRpt_GetDropDowns()
	
	If (e2bImportStsRpt_bSortOrder = 0) Then	'ASC
		e2bImportStsRpt_sArrow = "/img/common/order_up.gif"
		e2bImportStsRpt_lNextSortOrder = 1
	Else
		e2bImportStsRpt_sArrow = "/img/common/order_down.gif"
		e2bImportStsRpt_lNextSortOrder = 0
	End If
	if ((e2bImportStsRpt_bSearch = 1) OR (e2bImportStsRpt_bPostBack = "1")) Then
		e2bImportStsRpt_InitializeSearchResultsCache()
	End If
End Function

'**************************************************************************************************
' Author	    : Gaurav Naresh Mittal
' Called From   : On Page load
' Parameters	: None
' Returns	    : None
' Description	: populates the date range drop down
'**************************************************************************************************
Function e2bImportStsRpt_GetDropDowns()
	Set e2bImportStsRpt_oDateRangeMsg = RetrieveDropdownNoCache("LM_DATE_RANGES", e2bImportStsRpt_lErrorNum, e2bImportStsRpt_sErrorDesc)
	Set e2bImportStsRpt_oDateRangeList = e2bImportStsRpt_oDateRangeMsg.selectnodes("/MESSAGE/TABLE_LM_DATE_RANGES/LM_DATE_RANGES")
	If e2bImportStsRpt_lRangeList = -2 Then
		For Each e2bImportStsRpt_oDate in e2bImportStsRpt_oDateRangeList
			e2bImportStsRpt_lRangeList = GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_RANGE_ID")
			e2bImportStsRpt_sDateFrom = fn_date_from_iso(GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_FROM_DATE"), 8, false)
			e2bImportStsRpt_sDateTo  = fn_date_from_iso(GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_TO_DATE"), 8, false)
			If UCase(Trim(GetXMLValueDirect(e2bImportStsRpt_oDate, "LM_DATE_RANGES_RANGE_NAME"))) = UCase("Last 7 Days") Then Exit For
		Next
	End if
End Function

'**************************************************************************************************
' Author	    : Gaurav Naresh Mittal
' Called From   : On search button click
' Parameters	: None
' Returns	    : None
' Description	: Fetches the data to be displayed on the screen according to the search
'                 parameters
'**************************************************************************************************
Function e2bImportStsRpt_GetWorklistResult()
	Dim oMessage, oMessage1
	Dim sqlSelect
	Dim sqlWhere
	Dim sSql, sSearchName
	Dim dateFormat
	
	sqlSelect = ""
	sqlWhere = ""	
	
	If (e2bImportStsRpt_bCurrentSort = "") Then
		e2bImportStsRpt_bCurrentSort = 3
	End If

	If (e2bImportStsRpt_bSortOrder = 1) Then
		e2bImportStsRpt_sSortOrder = "DESC"
	Else
		e2bImportStsRpt_bSortOrder = 0
		e2bImportStsRpt_sSortOrder = "ASC"
	End If
	
	Select Case (e2bImportStsRpt_bCurrentSort)
		Case 1 : e2bImportStsRpt_sField = "upper(SafetyReportId1)"
		Case 2 : e2bImportStsRpt_sField = "upper(e2b_type)"
		Case 3 : e2bImportStsRpt_sField = "upper(Sender_Agency)"
		Case 4 : e2bImportStsRpt_sField = "upper(Companynumb)"
		Case 5 : e2bImportStsRpt_sField = "status_text"
		Case 6 : e2bImportStsRpt_sField = "upper(cm.case_num)"
		Case 7 : e2bImportStsRpt_sField = "upper(USR.USER_FULLNAME)"
		Case 8 : e2bImportStsRpt_sField = "upper(Notes)"
		Case 9 : e2bImportStsRpt_sField = "m.date_received"
		Case 10 : e2bImportStsRpt_sField = "date_imported"
	End Select
	
	e2bImportStsRpt_sOrder =  e2bImportStsRpt_sField & " " & e2bImportStsRpt_sSortOrder
	
	sqlWhere = ""

	sSearchName = ""
    If SuppressACKtransmission = "1" Then
		sqlWhere = sqlWhere & " AND cr.SUPPRESS_ACK_TRANSMIT=1"
	End If
	If trim(e2bImportStsRpt_ProdName) <> "" Then
		sSearchName = e2bImportStsRpt_ProdName
		sqlWhere = sqlWhere & " AND (CASE WHEN :P_DISPLAY_LANGUAGE = 1 THEN UPPER(TRIM(NVL2(erd.PRODUCT_NAME_J, erd.PRODUCT_NAME_J, erd.PRODUCT_NAME))) ELSE UPPER(TRIM(erd.PRODUCT_NAME)) END ) =  UPPER(:SEARCH_NAME) "     
	End If
	If trim(e2bImportStsRpt_Status) <> "" Then
	   sqlWhere = sqlWhere & " AND S.Status =" & GetLong(e2bImportStsRpt_Status, 0)
	End If
	If trim(e2bImportStsRpt_RptType) <> "-1" Then
		IF trim(e2bImportStsRpt_RptType) = "4" Then
			sqlWhere = sqlWhere & " AND E2B_Type in (4, 6)"
		Else
			IF trim(e2bImportStsRpt_RptType) <> "7" Then   
				sqlWhere = sqlWhere & " AND E2B_Type =" & GetLong(e2bImportStsRpt_RptType, 0)
			Else
				sqlWhere = sqlWhere & " AND erd.URGENT_REPORT = 1 "
			End If
		End IF
	End If
	If trim(ReportingAgencyValue) <> "" Then
		sqlWhere = sqlWhere & " AND S.agency_id IN (" & cfCmn_FindRegEx(ReportingAgencyKey, "0-9,")  & ") " 
	End If
	If trim(e2bImportStsRpt_DateType) <> "" Then        
		If (IsDate(e2bImportStsRpt_sDateFrom) AND IsDate(e2bImportStsRpt_sDateTo)) Then            
			dateFormat = GetDateFormatForFromToDateFields()
			sqlWhere = sqlWhere & " AND trunc(" & cfCmn_FindRegEx(e2bImportStsRpt_DateType, "-A-Z0-9._/\\()+':, ") & ") >= TO_DATE('" & e2bImportStsRpt_sDateFrom & "','" & dateFormat & "')"
			sqlWhere = sqlWhere & " AND trunc(" & cfCmn_FindRegEx(e2bImportStsRpt_DateType, "-A-Z0-9._/\\()+':, ") & ") <= TO_DATE('" & e2bImportStsRpt_sDateTo & "','" & dateFormat & "')"            
		End If        
	End If
		
	If glUserType = gUSERTYPE_EN Then
		sqlWhere = sqlWhere & " AND cp.authority_id <> 4"
	End If

	Filter = "  "  & sqlWhere 
	OrderBy = " order by " & e2bImportStsRpt_sOrder
	sUserFullName = GetXMLValueDirect (oSession, "CFG_USERS_USER_FULLNAME")
	sCopyDateFrom = e2bImportStsRpt_sDateFrom
	sCopyDateTo = e2bImportStsRpt_sDateTo
	sProdName = Replace(e2bImportStsRpt_ProdName,"-1",GetTranslationData("ALL"))
	Call CreateMessage (oMessageIn, 300400079)
	Call SetXMLValueDirect (oMessageIn, "GN_GUI_LM_GENERAL_TEXT", Filter)
	Call SetXMLValueDirect (oMessageIn, "GN_STRING1", OrderBy)
	Call SetXMLValueDirect (oMessageIn, "GN_THIS_IS_ASP", 1)
	Call SetXMLValueDirect (oMessageIn, "GN_RPT_GMT_OFFSET", e2bTransStsRpt_lGmtOffSet)
	Call SetXMLValueDirect (oMessageIn, "CFG_USERS_USER_FULLNAME", sUserFullName)
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_DATE_FROM", e2bImportStsRpt_sDateFrom)
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_DATE_TO", e2bImportStsRpt_sDateTo)
	Call SetXMLValueDirect (oMessageIn, "GN_GUI_PROD_NAME_1", sProdName)
	Call SetXMLValueDirect (oMessageIn, "CFG_USERS_PROD_SECURITY", GetXMLValueDirect (oSession, "CFG_USERS_PROD_SECURITY"))
	Call SetXMLValueDirect (oMessageIn, "CFG_USERS_STUDY_SECURITY", GetXMLValueDirect (oSession, "CFG_USERS_STUDY_SECURITY"))
	Call SetXMLValueDirect (oMessageIn, "LM_PRODUCT_PROD_NAME", sSearchName)
	
	If left(ucase(mid(e2bImportStsRpt_DateType,11)), len("DATE_IMPORTED"))= "DATE_IMPORTED" Then
		sDataType =  GetTranslationData("DATE_RANGE_IMP_DATE")
		Call SetXMLValueDirect (oMessageIn, "GN_RPT_DATE_STRING", sDataType  )
	ElseIf left(ucase(mid(e2bImportStsRpt_DateType,11)), len("DATE_RECEIVED"))= "DATE_RECEIVED" Then
		sDataType = GetTranslationData("DATE_RANGE_INTERCHNG_PROC_DATE")
		Call SetXMLValueDirect (oMessageIn, "GN_RPT_DATE_STRING",  sDataType)
	Else
		sDataType =  GetTranslationData("DATE_RANGE_TRANS_DATE")
		Call SetXMLValueDirect (oMessageIn, "GN_RPT_DATE_STRING", sDataType)
	End If

	If ReportingAgencyValue = "" Then
		sAgencyList = GetTranslationData("ALL")
		Call SetXMLValueDirect (oMessageIn, "RPT_E2B_LOCAL_COMPANY_NAME", sAgencyList)
	Else
		sAgencyList = replace(ReportingAgencyValue,"'" ," ")
		Call SetXMLValueDirect (oMessageIn, "RPT_E2B_LOCAL_COMPANY_NAME", sAgencyList)
	End If

	sReportType = Replace(Replace(Replace(Replace(Replace(replace(e2bImportStsRpt_RptType,"-1",GetTranslationData("ALL")),"1",GetTranslationData("INITIAL")),"3",GetTranslationData("E2B_RPT_TYPE_FOLLOW_UP")),"4",GetTranslationData("E2B_RPT_TYPE_NULLIFICATION")),"5",GetTranslationData("WL_REP_DOWNGRADE")),"7",GetTranslationData("URGENT_REPORT"))
	sStatus = Replace(Replace(e2bImportStsRpt_Status,"102",GetTranslationData("IMP_STATUS_SUCCESSFUL")),"103",GetTranslationData("IMP_STATUS_FAILURE"))
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_REPORT_TYPE", sReportType)
	Call SetXMLValueDirect (oMessageIn, "RPT_E2B_STATUS_TEXT", sStatus)
	Call SetXMLValueDirect (oMessageIn, "GN_DISPLAY_LANGUAGE", glDisplayLang)

	If (e2bImportStsRpt_bSearch = "1" ) Then		
		e2bImportStsRpt_lCurrentPage = 1
	else
		e2bImportStsRpt_lCurrentPage = CInt(NVL(Request(e2bImportStsRpt_sFormName & "_CurrentPage").Item, 1))
	end if
	e2bImportStsRpt_lPageSize    = CInt(NVL(Request(e2bImportStsRpt_sFormName & "_PageSize").Item, 100))
	e2bImportStsRpt_lStartRow = ((e2bImportStsRpt_lCurrentPage - 1) * e2bImportStsRpt_lPageSize) + 1
	e2bImportStsRpt_lEndRow = e2bImportStsRpt_lCurrentPage * e2bImportStsRpt_lPageSize
	
	Call SetXMLValueDirect (oMessageIn, "GN_DB_LIST_START",e2bImportStsRpt_lStartRow)
	Call SetXMLValueDirect (oMessageIn, "GN_DB_LIST_END", e2bImportStsRpt_lEndRow)
	Call SetXMLValueDirect (oMessageIn, "GN_DISPLAY_LANGUAGE", glDisplayLang)		

	Set e2bImportStsRpt_oOutMessage = ServiceRequest(oArgusSvr, oMessageIn, e2bImportStsRpt_lErrorNum, e2bImportStsRpt_sErrorDesc)

	If e2bImportStsRpt_lErrorNum <> 0 Then
		Call ExecuteErrorPage (e2bImportStsRpt_lErrorNum, e2bImportStsRpt_sErrorDesc)
	End If
		
	WriteMessageToCache sResponseXML, e2bImportStsRpt_oOutMessage, true
	
	lCacheId = GetXMLValueDirect(e2bImportStsRpt_oOutMessage, "GN_DB_CACHE_ID")
	e2bImportStsRpt_lResultsRowCount = GetXMLValueDirect(e2bImportStsRpt_oOutMessage, "GN_DB_LIST_LENGTH")
	If IsNullOrEmpty(e2bImportStsRpt_lResultsRowCount) Then 
		e2bImportStsRpt_lResultsRowCount = 0
	Else
		e2bImportStsRpt_lResultsRowCount = CLng(e2bImportStsRpt_lResultsRowCount)
	End If
	
	Call SetUserPreferences("RPT_E2B_STATUS_TEXT", e2bImportStsRpt_Status,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_REPORT_TYPE", e2bImportStsRpt_RptType,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_MEDICINAL_PRODUCT", e2bImportStsRpt_ProdName,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_DATE_FROM", e2bImportStsRpt_sDateFrom,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_DATE_TO", e2bImportStsRpt_sDateTo,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_DATE_TYPE", e2bImportStsRpt_DateType,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("RPT_E2B_DATE_RANGE", e2bImportStsRpt_lRangeList,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("ReportingAgencyKey", ReportingAgencyKey,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	Call SetUserPreferences("ReportingAgencyValue", ReportingAgencyValue,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
    Call SetUserPreferences("ACK_TRANS_SUPPRESSED", SuppressACKtransmission,e2bImportStsRpt_lErrorNum,e2bImportStsRpt_sErrorDesc)
	E2BViewType = GetValueFromCMN_PROFILEOnKey ("E2B_VIEW_PRINT")
End Function

'**************************************************************************************************
' Author	    : Gaurav Naresh Mittal
' Called From   : On Reload after clicking Search or Sort
' Parameters	: None
' Returns	    : None
' Description	: Initializes the Grid component with the XML message
'**************************************************************************************************	
Function e2bImportStsRpt_InitializeSearchResultsCache()
	Dim sIconNonExprep, sEmptyImage
	Dim e2bImportStsRpt_sHtmlTemplate
	Dim sIconNotStarted, sIconPending, sIconSuccess, sIconError, sIconWarnings
	' Set paths for status icons
	sIconNonExprep = "/img/CF/zoom.gif"
	sEmptyImage = "/img/CF/Empty.gif"
	
	' Set paths for status icons
	sIconNotStarted = "/img/E2B/NotStarted.gif"
	sIconPending = "/img/E2B/Pending.gif"
	sIconSuccess = "/img/E2B/OK.gif"
	sIconError = "/img/E2B/Error.gif"
	sIconWarnings = "/img/E2B/Warning.gif"

	' Initialize Search Results Grid
	e2bImportStsRpt_sHtmlTemplate = ReadTextFromFile("E2BImportTemplate.tpl")	
	
	Set e2bImportStsRpt_oGridControl = Server.CreateObject("Relsys.Argus.Interop.Web.GridControl")
	e2bImportStsRpt_oGridControl.SetTemplate(e2bImportStsRpt_sHtmlTemplate)
    e2bImportStsRpt_oGridControl.SetEncryptKey(GetEncryptKey())
	e2bImportStsRpt_oGridControl.SetDataSource GetMessageCacheFileName(sResponseXML, true), "RPT_E2B"
	e2bImportStsRpt_oGridControl.AddPlaceholder "ROW_CLASS", "row-normal", "row-alternate"
	e2bImportStsRpt_oGridControl.AddPlaceholder "ROW_COLOR", "#FFFFFF", "#E9E9F3"
	e2bImportStsRpt_oGridControl.AddRule "E2B_LENGTH_CHECK_WARNING", "Warning", "EICON", sIconNonExprep
	e2bImportStsRpt_oGridControl.AddRule "E2B_LENGTH_CHECK_WARNING", "No Warning", "EICON", sEmptyImage
	e2bImportStsRpt_oGridControl.AddRule "E2B_LENGTH_CHECK_WARNING", "Error", "EICON", sIconNonExprep
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_GIVENAME", "Pending", "ICON_1", sIconPending 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_GIVENAME", "Success", "ICON_1", sIconSuccess 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_GIVENAME", "Warning", "ICON_1", sIconWarnings 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_GIVENAME", "Rejected", "ICON_1", sIconError 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_GIVENAME", "Error", "ICON_1", sIconError 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_FAMILY_NAME", "Pending", "ICON_2", sIconPending 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_FAMILY_NAME", "Generated", "ICON_2", sIconPending 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_FAMILY_NAME", "Failure", "ICON_2", sIconError 
	e2bImportStsRpt_oGridControl.AddRule "RPT_E2B_REPORTER_FAMILY_NAME", "Success", "ICON_2", sIconSuccess 
	
	If (e2bImportStsRpt_bSearch = "1" ) Then		
		e2bImportStsRpt_lCurrentPage = 1
		gSecHead_lGrid_CurrentPage = 1
	else
		e2bImportStsRpt_lCurrentPage = CInt(NVL(Request(sFormName & "_CurrentPage").Item, 1))
	end if 
	
	e2bImportStsRpt_lPageRowCount = e2bImportStsRpt_oGridControl.GetRowCount()
	If (e2bImportStsRpt_lEndRow > e2bImportStsRpt_lPageRowCount) Then
		e2bImportStsRpt_lEndRow = e2bImportStsRpt_lPageRowCount
	End If	
	
	e2bImportStsRpt_lCurrRec = e2bImportStsRpt_lPageRowCount Mod e2bImportStsRpt_lPageSize
	If (e2bImportStsRpt_lCurrRec = 0) Then
		e2bImportStsRpt_lCurrRec = e2bImportStsRpt_lPageSize
	End If
End Function

Function csCsl_Initialize()
	Dim oFS,sMsgCacheFileName
	Set oFS = Server.CreateObject("Scripting.FileSystemObject")      
	sMsgCacheFileName = GetMessageCacheFileName("CaseSearchResults", True)
	 
	If ((csCsl_bSearch = "1") And (csCsl_bPaginate = 0) ) Then
		csCsl_lCurrentPage = 1
		GetSearchResults()	  
		csCsl_InitializeSearchResultsCache()
		checkifcacheexists = "1" 
	ElseIf (oFS.FileExists(sMsgCacheFileName) = True) Then
		Set csCsl_oSearchResults = ReadMessageFromCache ("CaseSearchResults", True)
		csCsl_InitializeSearchResultsCache()
		checkifcacheexists = "1" 
		Set oResultList = csCsl_oSearchResults.SelectNodes ("/MESSAGE/TABLE_CASE_MASTER/CASE_MASTER")
		iResultCount = oResultList.Length
	Else
	   Set csCsl_oSearchResults = Nothing
	   checkifcacheexists = "0" 
	   iResultCount = "-1"
	End If
	Set oFS = Nothing    
End Function

Function e2bImportStsRpt_SetLMTableDateRangeFieldNameForDropDownValues()
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
