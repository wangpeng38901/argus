<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : Ajax_ReturnE2BReportId.asp
' Description  : Called for E2BImport.ASP 
'                
'                
'******************************************************************************
' Revision History
' Date		Author		Description
' 30un2006 Prakash     Ajax version
'******************************************************************************
%>
<script runat="SERVER" language="VBSCRIPT">
Dim Saf_Report_Id, Reg_Report_Id, strSQL, lLocked, lState, lErrNo, sError
Saf_Report_Id = GetRequest("Safety_report_id")
Call SetParameter("P_REPORT_ID", Saf_Report_Id, PARAM_NUMBER)
strSQL = "SELECT reg_report_id FROM CMN_REG_REPORTS WHERE ESM_REPORT_ID =:P_REPORT_ID"
Reg_Report_Id = ExecuteSQLReturnStr (strSQL, lErrNo, sError)

If IsNullOrEmpty(Reg_Report_Id) Then
    lErrNo =  -1
    sError = "No data found for this report." 
End if


If (lErrNo <> 0) Then 
    Response.Write ConstructAjaxErrorMessage(lErrNo, sError)
Else
    Response.Write "<MESSAGE cnt=""" & Reg_Report_Id & """></MESSAGE>"          
End if            
</script>
