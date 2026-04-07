<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title>Argus Safety</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body onload="fn_fillin()">
    <form name="Frm" id="Frm" action="" method="post">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <!-- Outer Box Starts -->
        <table class="table" cellpadding="0" cellspacing="0" style="width:100%; height:100%">
            <tr height="25px">
                <td class="section-header-middle">
                    <%BuildLocalLabel("DESC").SetStyleSheet("label label-section").Render() %>
                </td>
            </tr>
            <tr>
                <td class="padding-all">
                    <%BuildControlDirect(CTL_TEXTAREA, "Desc", request.QueryString("ErrorDetails"), true, 3, "").SetRowCols(8, 65).Style("width:100%;height:100%").SetMaxLength(32767).OnKeypress("return false;").Render()%>
                </td>
            </tr>
            <tr class="tblheader-gray" height="25px">
                <td align="center" valign="middle">
                    <%BuildButtonDirect("btnOk", "OK", 1).Style("width:50px")_
                        .OnClick("window.close();").Render()%>
                </td>
            </tr>
        </table>
        <!-- Outer Box Ends -->
    </form>
</body>
</html>
<!-- Page Display Ends -->
<!-- Javascript Functions Starts -->

<script type="text/javascript">
var objParams = window.dialogArguments;
var oParent = objParams.window;

function fn_fillin()
{
    document.all.Desc.value = oParent.GetErrorText();
}
</script>

<!-- Javascript Functions Ends -->
<!-- Server-Side Functions Starts -->
<!-- Server-Side Functions Ends -->
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
