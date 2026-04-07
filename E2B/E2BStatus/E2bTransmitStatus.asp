<!--#INCLUDE VIRTUAL="/Nav/AGHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : PrakashSingh
' Page         : RegulartoryRptListing.asp
' Description  :  
'******************************************************************************
' Revision History
' Date		Author		Description
' 11jul2006 Prakash    Original
'******************************************************************************
%>
<!DOCTYPE html>
<html style="overflow-x: hidden;">
<head>
	<!-- Page Title -->
	<title>Argus Safety</title>
	<!-- Include Stylesheet here -->
	<link rel="stylesheet" href="/css/Relsys.css" />
	<!-- Client Library Includes Starts -->
	<script type="text/javascript" src="/js/Menu/JMenu.js"></script>
	<!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%   
   
	Dim sFormName
	Dim Rpt_Search
	Dim Rpt_lPageRowCount
	Dim Rpt_lResultsRowCount
	Dim Rpt_lStartRow
	Dim Rpt_lEndRow
	Dim Rpt_bPostBack
	Dim Rpt_sHtmlTemplate
	Dim Rpt_oGridControl
	Dim Rpt_lCurrentPage
	Dim Rpt_lPageSize
	Dim Rpt_blinded
   
	Dim lErrNo, sError  
	Dim oDateRangeMsg, oDateRangeList, oDateRange
	Dim oContactMsg, oContactList, oContact
	Dim Agency, DateFrom, DateTo, MessageFrom, MessageTo, MsgType
	Dim oOutMsg, oRecList, oRec, lNumRec, lContactRec
	Dim bPrevSort, bCurrentSort, bSortOrder, sArrow, lNextSortOrder
	Dim RadioBtn, RangeList, TypeList, sSortOrder, sField
	Dim lOffSet, lGmtOffSet, search
	Dim sSortString
	Dim	lID, sStart, sStop
	Dim lPrevSort, lSortOrder, sArrowStr
	Dim lCurrentSort, oSearchResults 
	Dim lCacheId

	'This Page will be English for Japanese Users
	glDisplayLang = cfCMN_LANG_EN

	Rpt_bPostBack = Request.Form ("PostBack")
	lSortOrder = Request.Form ("SortOrder")
	lPrevSort = Request.Form ("PrevSortField")
	lCurrentSort = Request.Form ("CurrentSortField")
	Rpt_Search	= GetLong(GetRequest("SearchClicked"), 1)
	Agency = GetLong(GetRequest("Agency"), -1)
	DateFrom = GetRequest("DateFrom")
	DateTo = GetRequest("DateTo")
	MessageFrom = GetString(GetRequest("MessageFrom"), "")
	MessageTo = GetString(GetRequest("MessageTo"), "")
	MsgType = GetRequest("MessageType")
	search = GetRequest("bSearch")
	bPrevSort = Request.Form ("PrevSort")
	bCurrentSort = Request.Form ("CurrentSort")
	bSortOrder = Request.Form ("SortOrder")
	RadioBtn = GetLong(GetRequest("RadioBtn"), -1)
	RangeList = GetRequest("rangelist")
	TypeList = GetRequest("typelist")
	lOffSet = GetLong(GetRequest("OffSet"), 0)
	lGmtOffSet = GetLong(GetRequest("GmtOffSet"), 0)
	lNumRec = 0 
	sFormName="Frm_E2B"

	If Rpt_bPostBack = "1" Then 
		Rpt_bPostBack = True
	Else
		Rpt_bPostBack = False
	End If

	If IsNullOrEmpty(DateFrom) Then
		DateFrom = "00-MMM-0000"
	End If

	If IsNullOrEmpty(DateTo) Then
		DateFrom = "00-MMM-0000"
	End If

	If TypeList = "" Then
		TypeList = "(Any)"
	End If
	If RangeList = "" Then
		RangeList = 1
	End If    

	If (lPrevSort = "") Then
		lPrevSort = 1
	End If
	If (lCurrentSort = "") Then
		lCurrentSort = 1
	End If

	If (lSortOrder = "") Then
		lSortOrder = 1
	End If
	
	If IsNullOrEmpty(lSortOrder) Then lSortOrder = 0
	
	If not GetDropDowns() Then
		Call ExecuteErrorPage(lErrNo, sError)
	End If

	If Not IsNullOrEmpty(MsgType) Then 
		GetResults()
		'Initialize the Grid Control for display
		Rpt_InitializeSearchResultsCache()
	End If

	sArrow = "order_up.gif"
	lNextSortOrder = 1
	If (lPrevSort = lCurrentSort) Then
		If (lSortOrder = 0) Then	'ASC
			sArrow = "order_down.gif"
			lNextSortOrder = 1
		Else
			sArrow = "order_up.gif"
			lNextSortOrder = 0
		End If
	Else
		sArrow = "order_up.gif"
		lNextSortOrder = 0
	End If
	sArrowStr = "<img src='/img/common/" & sArrow & "' />"
%>
<!-- Declaration of Page Scope variables Ends -->

<!-- Assign Values to Page Scope variables Starts -->
<body class="no-margin" onload="init_Form()">
    <table class="table fixed-table" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%;">
		<tr style="height: 25px">
			<td>
				<!-- #INCLUDE VIRTUAL="/Nav/AGToolbar_inc.asp" -->
			</td>
		</tr>
		<tr>
			<td class="valign-top">
	        <form id="<%= sFormName %>" name="<%=sFormName %>" action="" method="post" class="form-100">
		        <%Call BuildHiddenControlDirect("selected_case_num", "") %>
		        <%Call BuildHiddenControlDirect("PrevSortField", lCurrentSort) %>
		        <%Call BuildHiddenControlDirect("CurrentSortField", lCurrentSort) %>
		        <%Call BuildHiddenControlDirect("CurrentSortFieldClause", sSortString) %>
		        <%Call BuildHiddenControlDirect("SortOrder", lSortOrder) %>
		        <%Call BuildHiddenControlDirect("SortString", "") %>
		        <%Call BuildHiddenControlDirect("Rpt_Search", "0") %>
		        <%Call BuildHiddenControlDirect("SearchClicked", "0") %>
		        <%Call BuildHiddenControlDirect("selected_row", "")  %>
		        <%Call BuildHiddenControlDirect("PostBack", "1") %>
		        <%Call BuildHiddenControlDirect("Agency", Agency) %>
		        <%Call BuildHiddenControlDirect("TypeList", TypeList) %>
		        <%Call BuildHiddenControlDirect("RadioBtn", RadioBtn) %>
		        <%Call BuildHiddenControlDirect("GmtOffSet", "") %>
		        <%Call BuildHiddenControlDirect("OffSet", "") %>
		        <%Call BuildHiddenControlDirect("UserCacheId", lCacheId) %>
		        <!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
	            <table class="table" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%">
		            <tr style="height: 25px">
			            <td class="padding-left-right">
				            <%BuildLabelDirect("ICSR Transmit Status").SetStyleSheet("label label-page-header").Render() %>
			            </td>
		            </tr>
		            <tr style="height: 25px">
			            <td>
				            <!--#INCLUDE VIRTUAL="/E2B/E2BStatus/Transmit_Filter_inc.asp" -->
			            </td>
		            </tr>
		            <tr style="height: 5px">
			            <td></td>
		            </tr>
		            <!--#INCLUDE VIRTUAL="/E2B/E2BStatus/Transmit_Result_inc.asp" -->
		            <tr class="tblheader-gray" style="height: 25px">
			            <td align="right">
				            <%BuildButtonDirect("btnPrint", "Print List", 100).Style("width:80px")_
					            .OnClick("fn_PrintList();").Disable(Rpt_lPageRowCount=0).Render()%>
			            </td>
		            </tr>
	            </table>
            </form>
            </td>
		</tr>
	</table>
</body>
</html>

<script type="text/vbscript" language="vbscript" runat="server">
Function GetDropDowns()
	Dim oMessage
	Call CreateMessage (oMessage, 300400003)
	Set oContactMsg = ServiceRequest(oArgusSvr, oMessage, lErrNo, sError)
	
	If lErrNo <> 0 Then
		GetDropDowns = false
		Exit Function
	End If
	Set oContactList = oContactMsg.SelectNodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")
	
	Set oDateRangeMsg = RetrieveDropdownNoCache("LM_DATE_RANGES", lErrNo, sError)
	If lErrNo <> 0 Then
		GetDropDowns = false
		Exit Function
	End If
	Set oDateRangeList = oDateRangeMsg.SelectNodes("/MESSAGE/TABLE_LM_DATE_RANGES/LM_DATE_RANGES")
	
	GetDropDowns = true
End Function

Function GetResults()
	Dim oMessage
	Dim dtDateFrom, dtDateTo, sAgencyName
	Dim oList, oRecord
	
	dtDateFrom = DateFrom
	dtDateTo = DateTo

	sAgencyName = "ALL"
	If Agency > 0 Then
		Set oList = oContactMsg.SelectNodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B[LM_REGULATORY_CONTACT_AGENCY_ID = " & Agency & "]")
		For Each oRecord in oList
			sAgencyName = GetXMLValueDirect(oRecord, "LM_REGULATORY_CONTACT_AGENCY_NAME")
			Exit For
		Next 
	End If

	Rpt_lCurrentPage = CInt(NVL(Request(sFormName & "_CurrentPage").Item, 1))
	Rpt_lPageSize    = CInt(NVL(Request(sFormName & "_PageSize").Item, 100))
	Rpt_lStartRow = ((Rpt_lCurrentPage - 1) * Rpt_lPageSize) + 1
	Rpt_lEndRow = Rpt_lCurrentPage * Rpt_lPageSize
	
	Call CreateMessage (oMessage, 300400001)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_AGENCY_NAME", sAgencyName)
	If RadioBtn = 0 Then
		Call SetXMLDate (oMessage, "GN_RPT_START_DATE", DateFrom)
		Call SetXMLDate (oMessage, "GN_RPT_END_DATE", DateTo)
	Elseif RadioBtn = 1 Then
		Call SetXMLValueDirect (oMessage, "RPT_E2B_MESSAGE_FROM", cfCmn_FindRegEx(MessageFrom, "0-9"))
		Call SetXMLValueDirect (oMessage, "RPT_E2B_MESSAGE_TO", cfCmn_FindRegEx(MessageTo, "0-9"))
	End If
	Call SetXMLValueDirect (oMessage, "RPT_E2B_MESSAGE_TYPE", MsgType)
	Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet) 'local browser gmt offset	
	Call SetXMLValueDirect (oMessage, "GN_PRINT", 0)

	Call SetXMLValueDirect (oMessage, "GN_DB_LIST_START", Rpt_lStartRow)
	Call SetXMLValueDirect (oMessage, "GN_DB_LIST_END", Rpt_lEndRow)
		
	'add the sort string
	sSortString = fn_GetSortString()
	Call SetXMLValueDirect (oMessage, "GN_GUI_LM_GENERAL_TEXT", sSortString)
	Set oSearchResults = ServiceRequest(oArgusSvr, oMessage, lErrNo, sError)

	If lErrNo <> 0 Then
		Call ExecuteErrorPage (lErrNo, sError)
	End If

	Rpt_lResultsRowCount = GetXMLValueDirect(oSearchResults, "TABLE_RPT_E2B/GN_DB_LIST_LENGTH")
	lCacheId = GetXMLValueDirect(oSearchResults, "TABLE_RPT_E2B/GN_DB_CACHE_ID")
	If IsNullOrEmpty(Rpt_lResultsRowCount) Then
		Rpt_lResultsRowCount = 0
	Else
		Rpt_lResultsRowCount = CLng(Rpt_lResultsRowCount)
	End If
	
	If Not oSearchResults is Nothing Then
		Set oRecList = oSearchResults.SelectNodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")
		lNumRec = oRecList.length
	End If
	' Save to Message in the Cache folder
	WriteMessageToCache "E2BTransmitResultSet", oSearchResults, true
	DateFrom = dtDateFrom
	DateTo = dtDateTo
	GetResults = True
End Function

'This is same function in E2bTransmitStatusPrint.asp, any changes please handle there as well. Purpose: to handle sort in print as well.
Function fn_GetSortString()
	Dim sSortOrder, sSortField

	'decide the sort order
	If (lCurrentSort = "") Then
		lCurrentSort = 1
	End If

	If lPrevSort = lCurrentSort Then
		If lSortOrder = 0 Then
			sSortOrder = " desc"
		Elseif lSortOrder = 1 Then
			sSortOrder = " asc"
		End If
	Else
		sSortOrder = " asc"
		lSortOrder = 1
	End If
	
	'decide the sort field
	Select Case (lCurrentSort)
		Case 1 : sSortField = "TYP"
		Case 2 : sSortField = "NUM_REPORTS"
		Case 3 : sSortField = "AGENCY"
		Case 4 : sSortField = "LMESSAGENUMB"
		Case 5 : sSortField = "RMESSAGENUMB"
		Case 6 : sSortField = "FILENAME"
		Case 7 : sSortField = "TRANSMIT_DATE"
		Case 8 : sSortField = "EDI_TRACKING_ID"
		Case 9 : sSortField = "EDI_TRANSMIT_DATE"
		case 10 : sSortField = "ESM_AE_STATUS_NAME"
		Case 11 : sSortField = "EDI_COMPLETE_DATE"
		case 12 : sSortField = "EDI_CONTROL_NO"
		Case Else: sSortField = "TYP"			
	End Select
	
	If lCurrentSort <> 1 Then
	   sSortString = sSortField & " "  & sSortOrder & ", UPPER(TYP) "
	Else
		sSortString = sSortField & " "  & sSortOrder 
	End If
	fn_GetSortString = sSortString
End function

Function Rpt_InitializeSearchResultsCache()	    
	Dim sIconLock, sIconUnLock
	Dim FollowupStr
	' Initialize Search Results Grid
	Rpt_sHtmlTemplate = ReadTextFromFile("E2BTransmitStatus.tpl")	
	Set Rpt_oGridControl = Server.CreateObject("Relsys.Argus.Interop.Web.GridControl")

	Rpt_oGridControl.SetTemplate(Rpt_sHtmlTemplate)

	Rpt_oGridControl.SetDataSource GetMessageCacheFileName("E2BTransmitResultSet", true), "RPT_E2B"
	Rpt_oGridControl.AddPlaceholder "ROW_CLASS", "row-normal", "row-alternate"
	Rpt_oGridControl.AddPlaceholder "ROW_COLOR", "textbox-list", "textbox-list"
	Rpt_oGridControl.AddRule "RPT_E2B_NUMB_REPORTS", "-1", "RPT_E2B_NUMB_REPORTS", ""
	
	Rpt_lPageRowCount = Rpt_oGridControl.GetRowCount()

	If (Rpt_lEndRow > Rpt_lPageRowCount) Then
		Rpt_lEndRow = Rpt_lPageRowCount
	End If

End Function
</script>

<script type="text/javascript">
function fn_Sort(sort_field, sort_order){
	var strError;
	var bPostBack;
	bPostBack = <%=JavaScriptSanitize(Rpt_bPostBack) %>; 
	<%If (gSecHead_lGridRowCount <= 0) Then %>
	return false;
	<%end if %>
	
	document.Frm_E2B.CurrentSortField.value = sort_field;
	document.Frm_E2B.SortOrder.value = sort_order;
	fn_ValidateSubmitForm(document.Frm_E2B);
}
</script>
<!--#INCLUDE VIRTUAL="/Nav/AGFooter_inc.asp" -->
