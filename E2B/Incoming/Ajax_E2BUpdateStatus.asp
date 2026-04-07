<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
  Dim esm_report_id, e2b_status, lError, sError, sReturn, sSQL
  
  esm_report_id = GetRequest("esm_report_id")
  e2b_status = GetRequest("status")
  
  Call SetParameter("E2B_STATUS", e2b_status, PARAM_NUMBER)
  Call SetParameter("P_REPORT_ID", esm_report_id, PARAM_NUMBER)
  sSQL = "update SAFETYREPORT set e2b_type_accept_as = NULL, status = :E2B_STATUS where report_id = :P_REPORT_ID"
  Call UpdateSQL(sSQL,lError,sError)

  sReturn = ConstructAjaxErrorMessage(lError, sError)
  Response.Write sReturn
%>


