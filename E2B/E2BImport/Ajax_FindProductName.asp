<!-- #INCLUDE VIRTUAL="/Nav/AJAXHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Eric Popejoy
' Page         : Ajax_FindProductName.asp
' Description  : Called from FindProductName.asp to scan the ICSR Imports
'                
'                
'******************************************************************************
' Revision History
' Date          Author      Description
' 17FEB2014     EricP       initial version
'******************************************************************************

Dim sSearch, lType, lLanguage
Dim lError, sError, oMessage, oOutMsg


sSearch = Left(GetRequest("SEARCH"), 2000)
lType = GetLong(GetRequest("TYPE"), 0)
lLanguage = GetLong(GetRequest("DSPLYLNG"), 0)

Call CreateMessage (oMessage, 300200306) 'MID_db_app_E2b_find_product_name
Call SetXMLValueDirect (oMessage, "GN_DESC", sSearch)
Call SetXMLValueDirect (oMessage, "GN_DB_STATE", lType)
Call SetXMLValueDirect (oMessage, "GN_DISPLAY_LANGUAGE", lLanguage)

Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lError, sError)
If lError <> 0 Then
    Response.Write ConstructAjaxErrorMessage(lError, sError)
Else
    Response.Write oOutMsg.xml
End If
%>
