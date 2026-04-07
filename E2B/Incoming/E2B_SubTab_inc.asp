<!-- #INCLUDE VIRTUAL="/Include/CheckLoginLite_inc.asp" -->
<%
'******************************************************************************
' Author       : Pankaj G
' Page         : E2bImport.asp
' Description  : Pending/Processed report
'                This page is used in order to creare and switch between Pending/Processed reportc 
'               
'******************************************************************************
' Revision History
' Date		    Author		Description
' 30-NOV-2006   Pankaj G  Initial Version
'******************************************************************************
%>
<script type="text/javascript">
function f_CFN_SwitchTab(iTab)
{
    var sURL;
    switch (iTab) {
        case 1:
            sURL = "/E2B/Incoming/E2BPending.asp?CallFrom=Menu&<%=GetRequestKeyValue()%>"; break;
        case 2:
            sURL = "/E2B/E2BImport/E2BImport.asp?CallFrom=Menu&<%=GetRequestKeyValue()%>"; break;
        default:
            sURL = "/E2B/Incoming/E2BPending.asp?CallFrom=Menu&<%=GetRequestKeyValue()%>"; break;
    }
    showLoading();
    document.location = sURL;
}
</script>
