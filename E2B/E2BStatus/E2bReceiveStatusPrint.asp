<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : E2bTransmitStatusPrint.asp
' Description  : Print the Data for E2B Report that is Transmitted from Argus in PDF
'******************************************************************************
' Revision History
' Date		Author		Description
' 15jun2006 Prakash     Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title>Receive Status Print List</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
Dim oMessage, bSuccess, lViewAll, sSortString, oOutMsg, sDocId, sUserID, lSortOrder, lPrevSort, lCurrentSort
Dim RangeID, RangeName, oDateRangeMsg, oDateRangeList, oDate, IID, lOffSet, lGmtOffSet
Dim lErrNo, sError
Dim Agency, RadioBtn, DateFrom 
Dim DateTo, TypeList
Dim MessageFrom, MessageTo
Dim lCacheId
bSuccess = true

Agency = GetLong(GetRequest("Agency"), -1)
RadioBtn = GetLong(GetRequest("RadioBtn"), -1)
DateFrom = GetRequest("DateFrom")
DateTo = GetRequest("DateTo")
TypeList = GetRequest("typelist")
RangeID = GetRequest("rangelist")
MessageFrom = GetLong(GetRequest("MessageFrom"), 0)
MessageTo = GetLong(GetRequest("MessageTo"), 0)
lCacheId = GetRequest("cacheId")

lOffSet = GetLong(GetRequest("OffSet"), 0)
lGmtOffSet = GetLong(GetRequest("GmtOffSet"), 0)
lSortOrder = Request("SortOrder")
lPrevSort = Request("PrevSortField")
lCurrentSort = Request("CurrentSortField")
If RadioBtn = 0 Then
	MessageTo = ""
	MessageFrom = ""
ElseIf RadioBtn = 1 Then
	DateFrom = ""
	DateTo = ""
End If
If TypeList = "" Then
	TypeList = "(Any)"
End If

RangeName = "Custom Range Name"
Set oDateRangeMsg = RetrieveDropdownNoCache("LM_DATE_RANGES", lErrNo, sError)
Set oDateRangeList = oDateRangeMsg.selectnodes("/MESSAGE/TABLE_LM_DATE_RANGES/LM_DATE_RANGES")
For each oDate in oDateRangeList
	IID = GetXMLValueDirect(oDate, "LM_DATE_RANGES_RANGE_ID")
	If IID = RangeID Then
		RangeName = GetXMLValueDirect(oDate, "LM_DATE_RANGES_RANGE_NAME")
		Exit For
	end if
Next

'Fetch the Agency Name from Agency Id
Dim oContact, oContactList, sAgencyName
sAgencyName = "ALL"
If Agency > 0 Then
	Call CreateMessage (oMessage, 300400003)
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lErrNo, sError)	
	Set oContactList = oOutMsg.SelectNodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B[LM_REGULATORY_CONTACT_AGENCY_ID = " & Agency & "]")
	For Each oContact in oContactList
		sAgencyName = GetXMLValueDirect(oContact, "LM_REGULATORY_CONTACT_AGENCY_NAME")
		Exit For
	Next 
End If

Call CreateMessage (oMessage, 300400002)
Call SetXMLValueDirect (oMessage, "RPT_E2B_AGENCY_NAME", sAgencyName)
Call SetXMLValueDirect (oMessage, "RPT_E2B_DATE_FROM", DateFrom)
Call SetXMLValueDirect (oMessage, "RPT_E2B_DATE_TO", DateTo)
Call SetXMLValueDirect (oMessage, "RPT_E2B_MESSAGE_FROM", MessageFrom)
Call SetXMLValueDirect (oMessage, "RPT_E2B_MESSAGE_TO", MessageTo)
Call SetXMLValueDirect (oMessage, "RPT_E2B_MESSAGE_TYPE", TypeList)

If RadioBtn = 0 Then
	Call SetXMLDate (oMessage, "GN_RPT_START_DATE", fn_ZuluStandard(DateFrom, lOffSet))
	Call SetXMLDate (oMessage, "GN_RPT_END_DATE", fn_ZuluStandard(DateTo, lOffSet))
End If
Call SetXMLValueDirect (oMessage, "LM_DATE_RANGES_RANGE_NAME", RangeName)		
Call SetXMLValueDirect (oMessage, "GN_PRINT", 1)
Call SetXMLValueDirect (oMessage, "GN_DB_CACHE_ID", lCacheId)
Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
'add the sort string
sSortString = fn_GetSortString()
Call SetXMLValueDirect (oMessage, "GN_GUI_LM_GENERAL_TEXT", sSortString)
Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)

Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lErrNo, sError)
If (lErrNo <> 0) Then	
	bSuccess = False
Else
	sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
End If

Function fn_GetSortString()
	Dim sSortOrder, sSortField

	'decide the sort order
	If (lCurrentSort = "") Then
		lCurrentSort = 1
	End If

	If lPrevSort = lCurrentSort then
		If lSortOrder = 0 then
			sSortOrder = " desc"
		ElseIf lSortOrder = 1 then
			sSortOrder = " asc"
		End If
	Else
		sSortOrder = " asc"
	End If
	
	'decide the sort field
	Select Case (lCurrentSort)
		Case 1 : sSortField = "TYP"
		Case 2 : sSortField = "SENDER_ID"
		Case 3 : sSortField = "MSG_ID"
		Case 4 : sSortField = "MESSAGENUMB"
		Case 5 : sSortField = "REPORTS"
		Case 6 : sSortField = "REJECT_CNT"
		Case 7 : sSortField = "RECEIVE_FILENAME"
		Case 8 : sSortField = "RECEIVE_DATE"
		Case 9 : sSortField = "ESM_AE_STATUS_NAME"
		case 10 : sSortField = "EDI_CONTROL_NO"
		Case Else: sSortField = "TYP"
	End Select

	If lCurrentSort <> 1 Then
	   sSortString = sSortField & " "  & sSortOrder 
	Else
		sSortString = sSortField & " "  & sSortOrder 
	End If
	fn_GetSortString = sSortString
End function

%>
<body onload="fn_init();">
</body>
</html>

<script type="text/javascript">
    async function fn_init()
	{
		<%If bSuccess Then %>
		fn_ViewDocument("<%=sDocId%>", "window");
		<%Else %>
            await MessageBoxResEx("EL_PRINT_ERROR",<%=glDisplayLang %>);
		<%End If %>
		window.close();
	}
</script>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
