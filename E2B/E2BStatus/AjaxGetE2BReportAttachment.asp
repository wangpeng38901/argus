<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
Dim Attachment_Id
Dim oMessage, oOutMsg
Dim sDocId
Dim bSuccess
Dim ErrorNum, Error, cf_sErrStr

bSuccess = True
Attachment_Id = GetLong(GetRequest("Attachment_Id"), 0)
ErrorNum = 0 

If (Attachment_Id = "") Then
	ErrorNum = 1
	Error = "Insufficient information. Attachment Id not received"
	bSuccess = false
End If

If (bSuccess) Then
	bSuccess = GetE2BAttachmentFile()
End If

If ErrorNum <> 0 Then
	cf_sErrStr = ConstructAjaxErrorMessage(ErrorNum, Error)    
Else
	If IsNullOrEmpty(sDocId) Then
		cf_sErrStr = ConstructAjaxErrorMessage("-1", "Failed to receive ICSR attachment file name")
	else
		cf_sErrStr = "<MESSAGE><GN_REPORT_IDENTIFIER>" & sDocId & "</GN_REPORT_IDENTIFIER></MESSAGE>"
	end if        
End If
Response.Write cf_sErrStr
%>

<script language="VBScript" runat="Server">
Function GetE2BAttachmentFile()
	Dim oMessage
	ErrorNum = 0
	Error = ""
	Call CreateMessage (oMessage, 300400089)
	Call SetXMLValueDirect (oMessage, "ESM_ATTACHMENT_ESM_ATTACHMENT_ID", Attachment_Id)						
	Call SetXMLValueDirect (oMessage, "GN_GEN_SAVE_REPORT", 1)						
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, ErrorNum, Error)
	
	ErrorNum = GetXMLValueDirect(oOutMsg, "GN_ERROR_NUMBER")
	If (ErrorNum <> 0) Then
		Error = GetXMLValueDirect(oOutMsg, "GN_ERROR_STRING")
		GetE2BAttachmentFile = FALSE
		Exit Function
	Else
		sDocId = GetXMLValueDirect(oOutMsg, "GN_REPORT_IDENTIFIER")
	End If		
	GetE2BAttachmentFile = TRUE		
End Function
</script>



