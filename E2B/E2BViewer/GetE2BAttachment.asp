<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>

<html>
<head>
    <title>Argus Safety</title>
    <link rel="stylesheet" href="/css/Relsys.css" />
</head>

<%
'**************************************************************************************************
' Author		: Surya Kant Pandey
' Page			: GetE2BAttachment.asp
'**************************************************************************************************
' Revision History
' Date			Author		Description
' 22-JAN-14		Surya Kant 	Original
'**************************************************************************************************
%>

<%
Dim lAttach_id, IncomingE2b
Dim oInMessage, oOutMsg, lErrNo, sPgError
Dim sDocId
Dim bSuccess, lError, sError

bSuccess = True

lAttach_id = GetLongAsStr(GetRequest("attach_id"))
IncomingE2b = GetLong(GetRequest("IncomingE2b"), 0)
If (lAttach_id = "" ) Then
    sPgError = GetTranslationData("ERR_IN_FETCH_ATTACH")
	bSuccess = false
End If

If (bSuccess) Then
	bSuccess = GetAttachment()
End If

%>
<script type="text/javascript">
    async function initForm()
    {
        <% If(bSuccess = false) Then %>
            await MessageBoxRes("GENERAL_WARNING", '<%=GetTranslationData("VIEW_ATTACH")%>', <%=JavaScriptClean(sPgError) %>); 
        setWindowReturnValue("Error");
        window.close();
        <%else%>
        fn_ViewDocument("<%=sDocId%>", "");
        window.close();
        <%End If %>
        }
</script>


<html>
<body onload="initForm();">
</body>
</html>

<script language="VBScript" runat="Server">

Function GetAttachment()
    Dim strSQL, sErrMsg, sTemp, sError, lError
    Dim lSeqNum,documentumid
	
    If len(lAttach_id) > 0 Then
		documentumid = ""
        If IncomingE2b = 0 Then
            'check to see if documentum ID exists
	        Call SetParameter("ID", lAttach_id, PARAM_NUMBER)
	        documentumid = ExecuteSQLReturnStr("select documentum_id from CASE_NOTES_E2B_ATTACH where ID = :ID", lError, sError)
        End If
        		
        If Len(documentumid) > 0 Then
		    Call CreateMessage (oInMessage, 300200285)
		    Call SetXMLValueDirect (oInMessage, "CSNAT_USER_ID", oArgusUser.GetUserId())
		    Call SetXMLValueDirect (oInMessage, "CSNAT_DOCUMENTUM_ID", documentumid)
        Else
		    Call CreateMessage (oInMessage, 300500011)
		    Call SetXMLValueDirect (oInMessage, "GN_NUMBER1",lAttach_id)
            Call SetXMLValueDirect (oInMessage, "GN_NUMBER2",IncomingE2b)
	    End If
        Call SetXMLValueDirect (oInMessage, "GN_GEN_SAVE_REPORT",1)
		'oInMessage.save "c:\openAttach.xml"
		Set oOutMsg = ServiceRequest(oArgusSvr, oInMessage, lError, sError)
		'oOutMsg.save "c:\oOutopenAttach.xml"

		if lError <> 0 then
			GetAttachment = False
		else
    		sDocId = GetXMLValueDirect (oOutMsg, "GN_REPORT_IDENTIFIER")
			GetAttachment = True
		end if
    End If
End Function

</script>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
