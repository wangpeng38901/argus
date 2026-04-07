<!-- #INCLUDE VIRTUAL="/Include/CheckLoginLite_inc.asp" -->
<!-- Section Column Headers - Open Cases Search Results Starts -->
<%
    Dim bWhite
    If (bPrevSort = bCurrentSort) Then
        If (bSortOrder = 0) Then	'ASC
            sArrow = "order_up.gif"
            lNextSortOrder = 1
        Else
            sArrow = "order_down.gif"
            lNextSortOrder = 0
        End If
    Else
        sArrow = "order_up.gif"
        lNextSortOrder = 1
    End If
    lIsJReport = GetLong(GetRequest("IsJReport"), 0) ' 1 if Authority Id is 4 for the report else 0

    AArrowStr = "<img src='/img/common/" & sArrow & "' />"
    bWhite ="White"
%>
<table class="table" cellpadding="0" cellspacing="0" style="width: 100%; height:100%">
	<tr style="height: 25px">
		<td>
			<table class="table" cellspacing="0" cellpadding="0" style="width: 100%; height:100%" border="0" id="CA_Header_Table">
				<!-- Section Header - Open Cases Search Results Starts -->
				<tr width="100%">
					<td style="margin-top: 0px; margin-bottom: 0px" valign="top">
						<%  
							gSecHead_sGridName = search_sFormName
							gSecHead_lGridRowCount = search_lResultsRowCount
						%>
						<!-- #INCLUDE VIRTUAL="/Common/SectionHeaderGrid_inc.asp" -->
					</td>
				</tr>
				<!-- Section Header - Open Cases Search Results Ends -->
			</table>
		</td>
	</tr>
	<!-- Grid Control Ends -->
	<!-- Section Column Headers - Open Cases Search Results Starts -->
	<tr>
		<td class="no-padding-left">
			<div id="Div_CaseAssign" style="width: 100%; height: 100%;">
				<table class="table" cellspacing="0" cellpadding="0" style="width: 100%; height: 100%;">
					<thead>
						<tr style="height: 36px">
							<td>
								<div id="DivTable1" class="table-scroll-hide" style="width: 100%; overflow-y: scroll;">
									<table id="result_table" class="grd-table-header" cellspacing="0" cellpadding="2">
										<col width="20%" />
										<col width="15%" />
										<col width="15%" />
										<col width="30%" />
										<col />

										<tr class="tblheader-lightblue">
											<th class="grd-header">
												<table cellpadding='0' cellspacing='0' width="100%">
													<tr style="cursor: pointer">
														<td width="100%" onclick="fn_SortHeader(1, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("CASE_NUMBER").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 1) Then response.Write(trim(AArrowStr))  %>
														</td>
													</tr>
													<tr style="cursor: pointer">
														<td width="100%" onclick="fn_SortHeader(2, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("PAT_INITIALS").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 2) Then response.Write(trim(AArrowStr))  %>
														</td>
													</tr>
												</table>
											</th>
											<th class="grd-header">
												<table cellpadding='0' cellspacing='0' width="100%">
													<tr style="cursor: pointer">
														<td width="100%" onclick="fn_SortHeader(9, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("PROJECT_ID").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 9) Then response.Write(trim(AArrowStr))  %>
														</td>
													</tr>
													<tr style="cursor: pointer">
														<td width="100%" onclick="fn_SortHeader(6, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("STD_ID").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 6) Then response.Write(trim(AArrowStr))  %>
														</td>
													</tr>

												</table>
											</th>
											<th style="cursor: pointer" class="grd-header">
												<table cellpadding='0' cellspacing='0' width="100%">
													<tr width="100%">
														<td width="100%" onclick="fn_SortHeader(4, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("DATE").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 4) Then response.Write(AArrowStr) %>
														</td>
													</tr>
													<tr width="100%">
														<td width="100%" onclick="fn_SortHeader(5, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("COUNTRY").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 5) Then response.Write(AArrowStr) %>
														</td>
													</tr>

												</table>
											</th>
											<th style="cursor: pointer" class="grd-header">
												<table cellpadding='0' cellspacing='0' width="100%">
													<tr>
														<td width="100%" onclick="fn_SortHeader(7, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("PROD").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 7) Then response.Write(AArrowStr) %>
														</td>
													</tr>
													<tr>
														<td width="100%" onclick="fn_SortHeader(8, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("EVENT").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 8) Then response.Write(AArrowStr) %>
														</td>
													</tr>
												</table>
											</th>
											<th style="cursor: pointer" class="grd-header">
												<table cellpadding='0' cellspacing='0' width="100%">
													<tr>
														<td width="100%" onclick="fn_SortHeader(10, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("REPORT_TYPE").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 10) Then response.Write(AArrowStr) %>
														</td>
													</tr>
													<tr>
														<td width="100%" onclick="fn_SortHeader(11, '<%=lNextSortOrder%>');">
															<%BuildLocalLabel("REPORTER").SetStyleSheet("label label-no-padding").Render()%>
															<%If (bCurrentSort = 11) Then response.Write(AArrowStr) %>
														</td>
													</tr>
												</table>
											</th>
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
										<col width="20%" />
										<col width="15%" />
										<col width="15%" />
										<col width="30%" />
										<col />

										<%
								IF search_lPageRowCount <= 0 THEN
									Response.Write("<tr><td colspan=5 class='label align-center'>" & GetTranslationDataForNoDataFound("NO_CASE_FOUND") & "<td></tr>")
								ELSE
									Response.Write(search_oGridControl.Render(1, search_lPageRowcount))
								END IF
										%>
									</table>
								</div>
							</td>
						</tr>
					</tbody>
				</table>
			</div>
		</td>
	</tr>
</table>
<!-- Section Column Headers - Open Cases Search Results Starts -->
<script language="VBScript" runat="Server">
function GetGridValue(oRec, col)
Dim value
if (glDisplayLang = cfCMN_LANG_JP) then
	value = GetXMLValueDirect(oRec, col + "_J")
	if (IsNullOrEmpty(value)) then
		value = GetXMLValueDirect(oRec, col)
		if (Not IsNullOrEmpty(value)) then 'only when english value exists that "(no translation)" will be appended 
			value = value + sNoTran
		end if
	end if
else
	value = GetXMLValueDirect(oRec, col)
end if
GetGridValue = value
end function 

Function GetFormattedDate(oRec, col)
Dim value
	if (glDisplayLang = cfCMN_LANG_JP) then
		value = GetXMLValueDirect(oRec, col + "_J")
		if (IsNullOrEmpty(value)) then
			value = GetXMLValueDirect(oRec, col) 'get english value
		end if
	else
		value = GetXMLValueDirect(oRec, col)
	end if
	GetFormattedDate = fn_date_from_iso_default_precise(fn_ConvertToArgusDateFormat(value, glDisplayLang), "", False, False, False)
End Function

</script>
