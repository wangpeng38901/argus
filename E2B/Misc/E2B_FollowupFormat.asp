<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : E2B_FollowupFormat.asp
' Description  : Take the Input from User for Report Format
'******************************************************************************
' Revision History
' Date				Author		Description
' 08JAN2007 		Umar       	Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title><%=GetTranslationData("E2B_FOLLOW_UP_REPORT_FORM")%></title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<!-- Declaration of Page Scope variables Ends -->
<!-- Assign Values to Page Scope variables Starts -->
<!-- Assign Values to Page Scope variables Ends -->
<!-- Page Processing Starts -->
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body>
    <form id="E2B_FollowupFormat" name="E2B_FollowupFormat" class="no-margin">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <table class="table" cellspacing="0" cellpadding="0" width="100%" height="100%" border="0">
            <tr style="height: 20px;" valign="top">
                <td class="section-header-middle" id="HeaderText" name="HeaderText">
                    <% BuildLocalLabel("E2B_FOLLOW_UP_REPORT_FORM").SetStyleSheet("label label-section").Render()%>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <%BuildControlDirect(CTL_RADIOBUTTON, "rdbutton", "1", false, 1, "  CIOMS:1;  MedWatch:0").OnClick("fn_clickradiobtn();").Render() %>
                </td>
            </tr>
            <tr class="tblheader-gray" height="25px">
                <td align="center">
                    <%BuildButtonDirect("imgOk", "OK", 2).Style("width:50px").OnClick("fn_OK();").Render() %>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
<!-- Page Display Ends -->
<!-- Javascript Functions Starts -->

<script type="text/javascript">
var sRptForm = "1";
function fn_clickradiobtn()
{
	if (document.getElementById("rdbutton_0").checked)
		sRptForm = "1";
	else
		sRptForm = "2";
}
function fn_OK()
{
	setWindowReturnValue(sRptForm);
	window.close();
}
</script>

<!-- Javascript Functions Ends -->
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
