<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Ankur Gupta
' Page         : Ajax_E2BActionItem.asp
' Description  : Called from E2B_AcceptNullificationE2BCase.asp to create Action Item
'                
'                
'******************************************************************************
' Revision History
' Date          Author      Description
' 03Sep2008     AnkurG      Ajax version
'******************************************************************************
Dim esm_report_id, bulk, act_type_id, user_grp_id, due_date, sDesc
Dim lError, sError, oMessage, oOutMsg

esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
bulk = GetLong(GetRequest("bulk"), 0)
act_type_id = GetLongAsStr(GetRequest("action_type"))
user_grp_id = GetLongAsStr(GetRequest("user_group"))
due_date = GetRequest("due_date")
sDesc = GetRequest("desc")

Call CreateMessage (oMessage, 300400134) 'MID_db_app_create_action_item
Call SetXMLValueDirect (oMessage, "GN_STRING1", esm_report_id)
Call SetXMLValueDirect (oMessage, "CFG_GROUPS_GROUP_ID", user_grp_id)
Call SetXMLValueDirect (oMessage, "CSAC_CODE", act_type_id)
Call SetXMLValueDirect (oMessage, "CSAC_DESCRIPTION", sDesc)
Call SetXMLValueDirect (oMessage, "GN_STRING2", due_date)
Call SetXMLValueDirect (oMessage, "GN_DISPLAY_LANGUAGE", glDisplayLang)

Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lError, sError)
Response.Write ConstructAjaxErrorMessage(lError, sError)

%>
