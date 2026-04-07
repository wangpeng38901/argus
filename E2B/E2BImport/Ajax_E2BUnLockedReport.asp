<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : Ajax_E2BUnLockedReport.asp
' Description  : Called for E2BPending.ASP 
'                
'                
'******************************************************************************
' Revision History
' Date		Author		Description
' 22Apr2007  Umar        Ajax version
'******************************************************************************
%>
<script runat="SERVER" language="VBSCRIPT">
Dim esm_report_id, strSQL, lErrNo, sError, sReturnXML, oMessage, sCurrentUserId, locked_user_id
Dim MyArray, esm_report, esm_initial_report_id, iDoNotClearInterTables

esm_report_id = cfCmn_FindRegEx(GetRequest("esm_report_id"), "0-9,")
sCurrentUserId = oArgusUser.GetUserId()
esm_initial_report_id = GetLong(GetRequest("esm_initial_report_id"), 0)
iDoNotClearInterTables = GetLong(GetRequest("do_not_clear_inter_tables"), 0)

if esm_initial_report_id > 0 then
    Call SetParameter("P_REPORT_ID", esm_initial_report_id, PARAM_NUMBER)
    CALL UpdateSQL("update SAFETYREPORT set status = 8, e2b_type_accept_as = NULL WHERE report_id = :P_REPORT_ID",lErrNo,sError)
end if
MyArray = Split(esm_report_id,",")

For Each esm_report In MyArray
    Call SetParameter("P_REPORT_ID", esm_report, PARAM_NUMBER)
    strSQL = "select locked_user_id from safetyreport where report_id = :P_REPORT_ID"
    locked_user_id = ExecuteSQLReturnStr(strSQL, lErrNo, sError)
    if lErrNo = 0 then
        locked_user_id = GetString(locked_user_id, "-1")
        if locked_user_id = "0" then locked_user_id = "-1"
        if locked_user_id = sCurrentUserId then
            Call SetParameter("P_REPORT_ID", esm_report, PARAM_NUMBER)
            strSQL = "update safetyreport set locked_user_id = Null where report_id = :P_REPORT_ID"
            Set oMessage = UpdateSQL(strSQL, lErrNo, sError)
            
            If(iDoNotClearInterTables = 0) Then
                Call SetParameter("USER_NUM", sCurrentUserId, PARAM_NUMBER)
                strSQL = "begin ESM_IMP_DIFF_REPORT.p_clear_inter_tables_all (:USER_NUM, :P_REPORT_ID); end;"
        	    Set oMessage = UpdateSQL(strSQL, lErrNo, sError)
            End If
        end if
    else
        Exit For
    end if
Next    

sReturnXML = ConstructAjaxErrorMessage(lErrNo, sError)
Response.Write sReturnXML
</script>
