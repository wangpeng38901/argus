<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<!-- Declaration of Page Scope variables Starts -->
<%
Dim oMessage, lError, sError
Dim lCaseId
Dim sSQL, strResult
%>
<!-- Declaration of Page Scope variables Ends -->

<!-- Assign Values to Page Scope variables Starts -->
<%
    lError = 0
    sError = ""
    lCaseId = GetLong(GetRequest("CaseId"), -1)
%>
<!-- Assign Values to Page Scope variables Ends -->

<!-- Page Processing Starts -->
<%
Call SetParameter("P_CASE_ID", lCaseId, PARAM_NUMBER)
sSQL = "Select decode(cm.state_id,1,1,0), cm.date_locked, cm.close_date, nvl(cu.user_fullname,'-99') " &_
    " from case_master cm, case_lock_info lock_info, cfg_users cu " &_
    " where cm.case_id = :P_CASE_ID" &_
    " and cm.case_id = lock_info.case_id(+) " &_
    " and lock_info.user_id = cu.user_id(+) " &_
    " and lock_info.lock_status(+) = 1"
strResult = "2110018, 2140075, 2140074, 6550009"
Set oMessage = ExecuteSQL(sSQL, strResult, lError, sError)

If lError <> 0 Then
    Response.Write ConstructAjaxErrorMessage(lError, sError)
Else
    Response.Write oMessage.xml    
End If
%>
<!-- Page Processing Ends -->
