<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'********************************************************************
' Author      : Eric Popejoy
' Page        : FindProductName.asp
' Description : Select a product from the ESM Schema
'********************************************************************
' Revision History
' Date		Author		Description
' 06FEB2014	Popejoy		Original 7.0.3.003
'********************************************************************
%>
<%
Dim sTitle, lCriteria, sTitleCodeName
lCriteria = GetLong(GetRequest("criteria"), 0)
If lCriteria <> 2 Then
    sTitleCodeName = "PROD_NAME"
Else
    sTitleCodeName = "GENERIC_NAME"
End If
%>
<html>
<head>
    <title><%=GetTranslationData(sTitleCodeName)%></title>
    <link rel="stylesheet" href="/css/Relsys.css" />
</head>

<body onload="fn_init()">
    <table class="table border" cellspacing="0" cellpadding="0" width="100%" height="100%">
        <tr style="height: 25px">
            <td class="section-header-middle">
                <%BuildLocalLabel(sTitleCodeName).SetStyleSheet("label label-section").Render() %>
            </td>
        </tr>
        <tr style="height: 25px">
            <td>&nbsp;
            <input type="text" style="width: 80%" class="textbox" id="selProduct" name="selProduct"
                value="" maxlength="2000" tabindex="1">
                &nbsp;
            <% BuildControlDirect(CTL_BUTTON, "btn_search", GetTranslationData("SEARCH"), false, 2, "") _
                       .OnClick("fn_search();") _
                       .Render() %>
            </td>
        </tr>
        <tr>
            <td class="padding-all">
                <div class="table-scroll" id="DIV_NAME_LIST" style="width: 100%; height: 100%">
                    <table id="NAME_LIST" width="100%" class="table">
                    </table>
                </div>
            </td>
        </tr>
        <tr class="tblheader-gray" style="height: 25px">
            <td align="center" valign="middle">
            <%BuildButton("bt_select", "OK", 3).Disable(true).Style("width:50px").OnClick("fn_select()").Render()  %>
            <%BuildButton("bt_cancel", "CANCEL", 4).Style("width:60px").OnClick("window.close()").Render() %>
            </td>
        </tr>
    </table>

</body>
</html>

<script type="text/javascript">
    var oCurrentRow = null;
    var sCurrentClass = "";

    function fn_init() {
        var parameters = window.dialogArguments; 
        document.all.selProduct.value = parameters[0];
    }

    function fn_SelectRow(oRow)
    {
        if(oCurrentRow)
            oCurrentRow.className = sCurrentClass;
        oCurrentRow = oRow;
        sCurrentClass = oRow.className;
        oRow.className = "row-select";
        fn_getElementByName("bt_select").disabled = false;
    }

    function fn_select() {
        var sSelectProduct = oCurrentRow.children[0].innerText;
        setWindowReturnValue(sSelectProduct);
        window.close();
    }

    async function fn_search() {
        var sTableFill;
        var sValue = document.all.selProduct.value;
        sTableFill = "<tr height='120px'><td align='middle'><%=GetTranslationData("BULK_RPT_PROCESS") %></td></tr>";
        DIV_NAME_LIST.innerHTML = "<table id='NAME_LIST' width='100%' class='table'>" + sTableFill + "</table>";
        oCurrentRow = null;
        sCurrentClass = "";
        obj = fn_getElementByName("bt_select");
        obj.disabled = true;
        var sParams = "SEARCH=" + fn_URLEncode(sValue) + "&DSPLYLNG=" + glDisplayLang + "&TYPE=<%=lCriteria %>";
        await loadArgusMessage("/E2B/E2BImport/Ajax_FindProductname.asp", fn_search_results, sParams);
    }

    async function fn_search_results() {
        var xmlDoc = this.req.responseXML;
        var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
        var sTableFill;
        sTableFill = "<tr height='120px'><td align='middle'><%=GetTranslationData("NO_DATA_WAS_FOUND") %></td></tr>";
        if (sErrStr.length > 0) {
            await MessageBoxRes("GENERAL_ERROR", "", sErrStr);
        }
        else {
            var sFile = xmlDoc.getElementsByTagName("GN_RPT_STRING")
            if (sFile.length > 0)
                if (GetTextContentFromXML(sFile[0]).length > 0) {
                    sTableFill = GetTextContentFromXML(sFile[0]);
                }
        }
        DIV_NAME_LIST.innerHTML = "<table id='NAME_LIST' width='100%' class='table'>" + sTableFill + "</table>";
    }

</script>

<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
