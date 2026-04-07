<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : BulkAcceptE2B.asp
' Description  : Bulk Accept E2B Case
'******************************************************************************
' Revision History
' Date		   Author		Description
' 15MAY2006 	Umar       	Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title>Bulk Accept ICSR Case</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<% Response.write "<BR><TABLE width='100%'><TR><td><center>Loading Data. Please Wait...<BR><img name='status' src='/img/Common/status.gif'></center></TD></TR></TABLE></span>"
%>
<html>
<head>
    <script type="text/javascript">
        parent.dLoading.innerHTML = "";
    </script>
    <%
Dim oOutMessage, oInMessage, ErrorNo, ErrorStr, bSuccess, E2b_type, E2BViewType, followup_reports, followup_notes, nullification_reports
Dim initial_reports, initial_notes, nullification_notes, sFilePathName, sDocId, ReportsArray
Dim QueueSize, CustomImport, lError, sError

QueueSize = 100 ' Batch Size
initial_reports = cfCmn_FindRegEx(GetRequest("initial"), "0-9,")
initial_notes = GetRequest("initial_notes")
followup_reports = cfCmn_FindRegEx(GetRequest("followup"), "0-9,")
followup_notes = GetRequest("followup_notes")
nullification_reports = cfCmn_FindRegEx(GetRequest("nullification"), "0-9,")
nullification_notes = GetRequest("nullification_notes")
E2b_type = GetLong(GetRequest("E2b_Type"), 0)
E2BViewType = GetRequest("E2BViewType")
CustomImport = "1"
if IsNullOrEmpty(CustomImport) then CustomImport = "0"

Sub ResponseStatus(count,total,from)
'********************************************************************
' Author       : Saurabh K
' Parameters   : Reports Processed, Total Reports
' Called from  : ProcessInitial, ProcessFollowup, ProcessNullification
' Description  : Updates the status message
'********************************************************************
' Revision History
' Date		Author		Description
' 20FEB2006 Saurabh K   Original
'********************************************************************
Dim innerhtml, str
if from="INITIAL" Then
	str = GetTranslationData("PROCESSED_X_OUT_OF_Y_INITIAL_REPORTS") 
elseif from="FOLLOWUP" Then
	str = GetTranslationData("PROCESSED_X_OUT_OF_Y_FOLLOWUP_REPORTS") 
elseif from="NULL" Then
	str = GetTranslationData("PROCESSED_X_OUT_OF_Y_NULLIFICATION_REPORTS") 
else
	str = "Processing Reports..."
end if

str = Replace(str,"x",cstr(count))
str = Replace(str,"y",cstr(total))
	
innerhtml = "<br><br><br><BR><TABLE width='100%'><TR><td><center>" & _
            str & _
            "<BR><img name='status' src='/img/Common/status.gif'>" & _
			"</center></TD></TR></TABLE>"
    %>

    <script type="text/javascript">
        parent.dLoading.innerHTML = "<%=innerhtml%>"
    </script>

    <%
Response.Flush
End Sub

bSuccess = False
If (E2b_type = 1) Then
	bSuccess = ProcessInitial()
ElseIf (E2b_type = 3 Or E2b_type = 5 Or E2b_type = 6)   Then
	bSuccess = ProcessFollowup()
ElseIf (E2b_type = 4) Then
	bSuccess = ProcessNullification()
Else
	bSuccess = False
	ErrorStr = "Error Occurred. Trying to import invalid cases"
End If

If (bSuccess) Then
	sDocId = GetXMLValueDirect (oOutMessage, "GN_REPORT_IDENTIFIER")
End If
    %>
    <%If (bSuccess) Then %>

    <script type="text/javascript">
        var sURL = "/ArgusNet/Common/DocViewer.aspx?DocId=<%=Server.URLEncode(sDocId) %>&<%=GetRequestKeyValue()%>";
        document.location = sURL;
        parent.dLoading.innerHTML = "";
    </script>

    <% End If %>
    <body leftmargin="0" topmargin="0" onload="fn_CheckErrors();">
        <%=Fn_Sanitize(ErrorStr) %>

        <script lang="javascript" type="text/javascript">
    var E2bReports;
    async function fn_CheckErrors()
    {
      	<%If (E2b_type = 1) Then %>
	        E2bReports = <%=JavaScriptSanitize(initial_reports) %>;
        <%ElseIf (E2b_type = 3 Or E2b_type = 5 Or E2b_type = 6) Then %>
	        E2bReports = <%=JavaScriptSanitize(followup_reports) %>;
        <%ElseIf (E2b_type = 4) Then%>
	        E2bReports = <%=JavaScriptSanitize(nullification_reports) %>;
	    <%Else %>
	        E2bReports = "-1";
	    <% End If %>;
        await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + E2bReports);
    }
    
        async function fn_UnLockedReport() {
            var xmlDoc = this.req.responseXML;
            var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
            if (sErrStr.length > 0) {
                await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("INCOME_E2B")%>', sErrStr);
            }
            return;
        }
        </script>

    </body>
</html>

<script language="vbscript" runat="server">
Function ReportsQueue(RArray, Num, QueueSize, byref Completed)
'********************************************************************
' Author       : Saurabh K
' Parameters   : Report Ids Array, Batch number, Batch size, By Ref - End of data
' Called from  : ProcessInitial
' Description  : Returns next batch of QueueSize reports; comma separated
'********************************************************************
' Revision History
' Date		Author		Description
' 20FEB2006 Saurabh K   Original
'********************************************************************
	Dim current_reports, i, min, max, length
	ReportsQueue = ""
	length = UBound(RArray)+1
		
	min = Num*QueueSize
	If ( length > (Num+1)*QueueSize ) Then
		max = ((Num+1)*QueueSize)-1
		Completed = 0
	Else
		max = length - 1
		Completed = 1
	End If	
		
	For i = min to max
		ReportsQueue = ReportsQueue & "," & RArray(i)
	Next
	ReportsQueue = Right(ReportsQueue,Len(ReportsQueue)-1)
	
End Function

Function ProcessInitial()
'********************************************************************
' Author       : Suchita
' Parameters   : 
' Called from  : This Form
' Description  : Sends the initial reports to database for acceptance
'********************************************************************
' Revision History
' Date		Author		Description
' 11JUL2002 Suchita     Original
' 20FEB2006 SaurabhK    Batch Processing implemented
'********************************************************************

	Dim current_reports, total_reports
	Dim Num, Completed, count, FirstTime
	
	Num = 0
	Completed = 0
	FirstTime = 1
	
	ReportsArray = Split(initial_reports, ",")
	total_reports = UBound(ReportsArray)+1
	
	While Completed = 0
		current_reports = ReportsQueue(ReportsArray,Num,QueueSize,Completed)
		
		Call CreateMessage (oInMessage, 300400011)
		Call SetXMLValueDirect (oInMessage, "GN_RPT_STRING", current_reports)
		Call SetXMLValueDirect (oInMessage, "GN_UI_JUSTIFICATION", initial_notes)
		Call SetXMLValueDirect (oInMessage, "CSM_SID", GetXMLValueDirect(oSession, "GN_UI_DBNAME"))
		Call SetXMLValueDirect (oInMessage, "CSM_SITE_ID", GetXMLValueDirect (oSession, "CFG_USERS_SITE_ID"))
		Call SetXMLValueDirect (oInMessage, "RPT_E2B_E2B_TYPE", E2b_type)
		Call SetXMLValueDirect (oInMessage, "GN_STATUS_NUMBER", FirstTime*2+Completed)
		Call SetXMLValueDirect (oInMessage, "GN_RPT_GMT_OFFSET", GetXMLValueDirect(oSession, "GMT"))
		Call SetXMLValueDirect (oInMessage, "RPT_E2B_VIEW_STAGING_WARNING", CustomImport)
		Call SetXMLValueDirect (oInMessage, "GN_GEN_SAVE_REPORT", 1)
		
		FirstTime = FirstTime * 0
	
		Set oOutMessage = ServiceRequest(oArgusSvr, oInMessage, ErrorNo, ErrorStr)
		If (ErrorNo <> 0) Then
			If IsNullOrEmpty(ErrorStr) Then 
				ErrorStr = "Error Occurred. Please Contact System Administrator"
			End If
			ProcessInitial = False
			Exit Function
		End If
		
		Num = Num + 1
		
		If Completed = 0 Then
			count = Num*QueueSize
		Else
			count = total_reports	
		End If
		Call ResponseStatus(count,total_reports,"INITIAL")
	Wend
	ProcessInitial = True
End Function

Function ProcessFollowup()
'********************************************************************
' Author       : Suchita
' Parameters   : 
' Called from  : This Form
' Description  : Sends the followup reports to database for acceptance
'********************************************************************
' Revision History
' Date		Author		Description
' 11JUL2002 Suchita     Original
' 20FEB2006 SaurabhK    Use Argusvr2a Object, Status message shows follow-up reports
' 11MAR2008 UR          Passing Justification Id to save in case followup
'********************************************************************

	Dim current_reports, total_reports, QueueSizeFup, temp_notes_array, followup_notes_id
	Dim Num, Completed, count, FirstTime, pdf_file_name

	Num = 0
	Completed = 0
	FirstTime = 1
	QueueSizeFup = 25
	pdf_file_name = ""
	ReportsArray = Split(followup_reports, ",")
	total_reports = UBound(ReportsArray)+1
	temp_notes_array = Split(followup_notes, sSeparatorSSO)
	followup_notes_id  = temp_notes_array(0)
	followup_notes = temp_notes_array(1)
	
	While Completed = 0
		current_reports = ReportsQueue(ReportsArray,Num,QueueSizeFup,Completed)
    		
		Call CreateMessage (oInMessage, 300400012)
	    Call SetXMLValueDirect (oInMessage, "GN_RPT_STRING", current_reports)
	    Call SetXMLValueDirect (oInMessage, "LM_JUSTIFICATIONS_JUSTIFICATION_ID", followup_notes_id)
	    Call SetXMLValueDirect (oInMessage, "GN_UI_JUSTIFICATION", followup_notes)
	    Call SetXMLValueDirect (oInMessage, "RPT_E2B_E2B_TYPE", E2b_type)
	    Call SetXMLValueDirect (oInMessage, "GN_STATUS_NUMBER", FirstTime*2+Completed)
	    Call SetXMLValueDirect (oInMessage, "RPT_E2B_VIEW_TYPE", E2BViewType)
	    Call SetXMLValueDirect (oInMessage, "GN_RPT_GMT_OFFSET", GetXMLValueDirect(oSession, "GMT"))
		Call SetXMLValueDirect (oInMessage, "RPT_E2B_FILENAME", pdf_file_name)

        If Completed = 1 Then 'Last Request
    		Call SetXMLValueDirect (oInMessage, "GN_GEN_SAVE_REPORT", 1)
        End If

		FirstTime = FirstTime * 0	
		Set oOutMessage = ServiceRequest(oArgusSvr, oInMessage, ErrorNo, ErrorStr)

		If (ErrorNo <> 0) Then
			If IsNullOrEmpty(ErrorStr) Then 
				ErrorStr = "Error Occurred. Please Contact System Administrator"
			End If
			ProcessFollowup = False
			Exit Function
		End If
		pdf_file_name = GetXMLValueDirect (oOutMessage, "GN_RPT_STRING")
		Num = Num + 1		
		If Completed = 0 Then
			count = Num*QueueSizeFup
		Else
			count = total_reports	
		End If
		Call ResponseStatus(count,total_reports,"FOLLOWUP")

	Wend
	ProcessFollowup = True
End Function

Function ProcessNullification()
'********************************************************************
' Author       : Suchita
' Parameters   : 
' Called from  : This Form
' Description  : Sends the nullification reports to database for acceptance
'********************************************************************
' Revision History
' Date		Author		Description
' 11JUL2002 Suchita     Original
' 20FEB2006 SaurabhK    Use Argusvr2a Object, Status message shows follow-up reports
'********************************************************************
	Call ResponseStatus(0,0,"NULL")
	Dim Pos
	Pos = InStrRev(nullification_notes,"~")
	If (Pos > 0) then
	    nullification_notes = Mid(nullification_notes,Pos + 1)
	End If
	Call CreateMessage (oInMessage, 300400013)
	Call SetXMLValueDirect (oInMessage, "GN_RPT_STRING", nullification_reports)
	Call SetXMLValueDirect (oInMessage, "GN_UI_JUSTIFICATION", nullification_notes)
	Call SetXMLValueDirect (oInMessage, "RPT_E2B_E2B_TYPE", E2b_type)
	Call SetXMLValueDirect (oInMessage, "GN_RPT_GMT_OFFSET", GetXMLValueDirect(oSession, "GMT"))
    Call SetXMLValueDirect (oInMessage, "GN_GEN_SAVE_REPORT", 1)
    	
	Set oOutMessage = ServiceRequest(oArgusSvr, oInMessage, ErrorNo, ErrorStr)
	If (ErrorNo <> 0) Then
		If IsNullOrEmpty(ErrorStr) Then 
			ErrorStr = "Error Occurred. Please Contact System Administrator"
		End If
		ProcessNullification = False
		Exit Function
	End If
	ProcessNullification = True
End Function
</script>

<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
