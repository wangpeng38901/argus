<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : Ajax_E2BLockedReport.asp
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
Dim esm_report_id, strSQL, lErrNo, sError, sReturnXML, locked_user_name, sUserID, oMessage, locked_user_id, bulk, MyArray
Dim new_esm_report_ids, esm_report, esm_report_changed, check_lock_status

esm_report_id = cfCmn_FindRegEx(GetRequest("esm_report_id"), "0-9,")
bulk = GetLong(GetRequest("bulk"), 0)
check_lock_status = GetLong(GetRequest("check_lock_status"), 0)

locked_user_name = ""
esm_report_changed = "0"
new_esm_report_ids = ""
sUserID = oArgusUser.GetUserId()

MyArray = Split(esm_report_id,",")

For Each esm_report In MyArray
    Call SetParameter("P_REPORT_ID", esm_report, PARAM_NUMBER)
    strSQL = "SELECT Locked_User_Id FROM SAFETYREPORT WHERE STATUS = 101 AND Report_Id = :P_REPORT_ID UNION SELECT User_Id FROM SAFETYREPORT WHERE Report_Id = :P_REPORT_ID  AND STATUS IN (102,103,104) "
    locked_user_id = ExecuteSQLReturnStr(strSQL, lErrNo, sError)        
    if lErrNo = 0 then
        locked_user_id = GetString(locked_user_id, "-1")
        if locked_user_id = "0" then locked_user_id = "-1"
        
        if locked_user_id <> sUserID and locked_user_id <> "-1" then
            esm_report_changed = "1" 'Skip this esm_report_id, because this is locked by other user
            Call SetParameter("USER_NUM", locked_user_id, PARAM_NUMBER)
            strSQL = "select decode(user_fullname,null,user_name,user_fullname) username "
            strSQL = strSQL & " from cfg_users "            
            strSQL = strSQL & " where user_id = :USER_NUM"
            locked_user_name = ExecuteSQLReturnStr (strSQL, lErrNo, sError)
            if lErrNo = 0 then                
                locked_user_name = Replace(GetTranslationData("SIMUL_PROCESS"), "[1]", locked_user_name)
            end if
        else
    		if new_esm_report_ids = "" then
                new_esm_report_ids = esm_report
            else
                new_esm_report_ids = new_esm_report_ids & "," & esm_report
            end if
            if check_lock_status <> "1" then
                Call SetParameter("USER_NUM", sUserID, PARAM_NUMBER)
                Call SetParameter("P_REPORT_ID", esm_report, PARAM_NUMBER)
                strSQL = "update safetyreport set locked_user_id = :USER_NUM where report_id = :P_REPORT_ID"
                Set oMessage = UpdateSQL(strSQL, lErrNo, sError)
            end if
        end if
    end if
Next
if bulk = "1" and esm_report_changed = "1" then
   locked_user_name = GetTranslationData("BULK_EXCLUDE_LOCK")
end if
sReturnXML = "<MESSAGE><ERROR_NUM>" &  ConvertXMLSpecialChars(lErrNo) & "</ERROR_NUM><ERROR_STRING>" &  ConvertXMLSpecialChars(sError) & "</ERROR_STRING><LOCKED_USER>" & locked_user_name & "</LOCKED_USER><MODIFY_REPORT_IDS>" & new_esm_report_ids & "</MODIFY_REPORT_IDS></MESSAGE>"
Response.Write sReturnXML
</script>
