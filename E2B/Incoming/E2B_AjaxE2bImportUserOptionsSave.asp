<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
    Dim sNotChecked, sChecked, oMessage, lError, sError, oOutMsg, sRecUndelete, sRecDelete
    
    sChecked = cfCmn_FindRegEx(GetRequest("checked"), "0-9,")
    sNotChecked = cfCmn_FindRegEx(GetRequest("not_checked"), "0-9,")
    sRecUndelete = cfCmn_FindRegEx(GetRequest("rec_undelete"), "0-9,")
    sRecDelete = cfCmn_FindRegEx(GetRequest("rec_delete"), "0-9,")
    
    Call CreateMessage (oMessage, 501200028) ' MID_bus_app_e2b_import_save
    Call SetXMLValueDirect (oMessage, "ESM_DIFFERENCE_REPORT_SEQ_NUM_CHECKED", sChecked)
    Call SetXMLValueDirect (oMessage, "ESM_DIFFERENCE_REPORT_SEQ_NUM_NOT_CHECKED", sNotChecked)
    Call SetXMLValueDirect (oMessage, "ESM_DIFFERENCE_REPORT_REC_DELETE", sRecDelete)
    Call SetXMLValueDirect (oMessage, "ESM_DIFFERENCE_REPORT_REC_UNDELETE", sRecUndelete)
    Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lError, sError)
    
    If lError <> 0 Then
        Response.Write ConstructAjaxErrorMessage(lError, sError)
    Else
        Response.Write oOutMsg.xml
    End If
%>

