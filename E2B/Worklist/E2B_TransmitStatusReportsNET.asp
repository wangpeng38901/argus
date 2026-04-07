<!-- #INCLUDE VIRTUAL="/Nav/AGHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<!DOCTYPE html>
<% 
    Dim ParamNet
    Dim strUrlNet 
    
    strUrlNet = "/ArgusNet/Worklist/BulkE2bTransmitReportsPage.aspx"    
    ParamNet = "GMT=" & GetXMLValueDirect(oSession, "GMT")
   
    If Not IsNullOrEmpty(Request.QueryString("message_id")) Then
        ParamNet = "&message_id=" & GetLong(Request.QueryString("message_id"), -1)
    End If
    If Not IsNullOrEmpty(Request.QueryString("reg_report_id")) Then
        ParamNet = "&reg_report_id=" & GetLong(Request.QueryString("reg_report_id"), -1)
    End If
    If Not IsNullOrEmpty(Request.QueryString("FailureState")) Then
        If Len(ParamNet) > 0 Then
            ParamNet = ParamNet + "&FailureState=" & Server.URLEncode(Request.QueryString("FailureState"))
        End If
    End If  
    If Not IsNullOrEmpty(Request.QueryString("report_type")) Then
       ParamNet = ParamNet + "&report_type=" & Server.URLEncode(Request.QueryString("report_type"))
    End If   
    strUrlNet = strUrlNet & "?" & ParamNet & "&" & GetRequestKeyValue()
%>
<html>
<head>
    <!-- Page Title -->
    <title>Bulk ICSR Transmit Reports</title>
    <link rel="stylesheet" href="/css/Relsys.css" />
</head>
<body>
    <table cellpadding="0" cellspacing="0" style="padding: 0; height: 100%; width: 100%; table-layout: fixed; margin: 0;">
        <tr height="25px">
            <td>
                <!--#INCLUDE VIRTUAL="/Nav/AGToolbar_inc.asp" -->
            </td>
        </tr>
        <tr>
            <td class="valign-top">
                <iframe scrolling="yes" frameborder="0" src="<%=strUrlNet%>" height="100%" width="100%"
                    style="border: 0;" marginheight="0" marginwidth="0"></iframe>
            </td>
        </tr>
    </table>
</body>
<!-- Page Display Ends -->
</html>
<!-- #INCLUDE VIRTUAL="/Nav/AGFooter_inc.asp" -->

<script type="text/javascript">
    showLoading();
</script>

