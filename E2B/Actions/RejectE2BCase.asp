<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : RejectE2BCase.asp
' Description  : This page is called form worklist/worklistbulktransmit.asp 
'******************************************************************************
' Revision History
' Date		Author		Description
' 15jun2006 Prakash     Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title><%=GetTranslationData("REJECT_INITIAL")%></title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <script type="text/javascript" src="/js/Common/CommonReAuthFun.js"></script>
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<%	Dim justi_sTitle, justi_oJustificationMsg, justi_oJustification, justi_oJustificationList, justi_row
	Dim field_id, justlen
    Dim bReAuthEnabled, bOIDCReAuthMode

    bReAuthEnabled = IsReAuthEnabled()
    bOIDCReAuthMode = IsOIDCReAuthMode()
	
	ProcessInput()%>

<script type="text/javascript">
var bPasswordOK = false;
var l_attempt = 0;
var strError;
var selected = 0;
var objParent = window.dialogArguments;
var l_delaytime = <%=GetRequestDelayTime() %>;

async function fn_password_verify()
{
	if (l_delaytime > 0) {
		hideLoading();
	}
    var sErrMsg = fn_GetAjaxErrorMsg(this.req.responseXML); 
    if (sErrMsg.length > 0)
    {
        bPasswordOK = false;
        await MessageBoxRes("GENERAL_WARNING", '<%=GetTranslationData("E2B_REJECT_CASE")%>', sErrMsg);
         <% If bReAuthEnabled Then %>
               document.all.info_pending.style.display = 'block';
               document.all.info_done.style.display = 'none';
               document.all.ReAuth.value = "";
               document.all.code.value = "";
               document.all.state.value = "";
         <% End If %>
    }
    else
    {
        bPasswordOK = true;
    }  
}

async function fn_KeyPress()
{
	if (window.event.keyCode == 13 && l_attempt < 3)
	{
		document.all.Btn_Ok.focus();
		await fn_ClickOK();
		return;
	}
}

function fn_ClickCancel()
{
	setWindowReturnValue("");
	window.close();
}

async function fn_ClickOK()
{
	var s_password, s_notes;
	var sURL, retVal;
	
	s_password = document.all.password.value;
	s_notes = document.all.notes.value;

	<%If bReAuthEnabled Then %>
    if (document.all.ReAuth.value.length == 0) {
	    await MessageBoxRes("RE_AUTH_JUSTIFICTN_PROVIDE_PASS");
	    return -1;
	}
	<%Else %>
    if (s_password.length == 0) {
        await MessageBoxRes("JUSTIFICTN_PROVIDE_PASS");
        document.all.password.focus();
	    return -1;
	}  
	<%End If %>

	if (s_notes.length == 0) {
		await MessageBoxRes("ACCEPT_CASE_NOTES_REQ");
		document.all.notes.focus();
		return -1;
	}
    var sParams = "password=" + fn_URLEncode(document.all.password.value) + "&FormName=Reject ICSR Report";
     <% If bReAuthEnabled And bOIDCReAuthMode Then %>
            sParams = sParams + "&code=" + fn_URLEncode(document.all.code.value) + "&state=" + fn_URLEncode(document.all.state.value);
     <% End If %>

	
	if (l_delaytime > 0) {
			showLoading();
			await sleep(l_delaytime);
	}

    await loadArgusMessage("/CaseForm/Dialogs/CF_AjaxPasswordVerify.asp", fn_password_verify, sParams);	
    
	if 	(!bPasswordOK)
	{
		document.all.password.focus();
		document.all.password.select();
		l_attempt ++;
		if (l_attempt < <%=GetPwdTriesCount()%>) return;
        await MessageBoxRes("E2B_REJECTCASE_ATTMPT_EXCEED"); 
        await fn_SessionTimeout();
        window.close();
        return;
    }
	setWindowReturnValue(s_notes);
	window.close();
}

function fn_select(item) {
	if (selected > 0) {
		eval("document.all.td" + selected).style.background = ccf_row_color_normal;
	}
	selected = item;
	eval("document.all.td" + selected).style.background = ccf_row_color_selected;
	str_value = eval("document.all.td" + selected).innerText;
    <% If glDisplayLang = cfCMN_LANG_JP Then %>
	    str_value = fn_ParseJapaneseJustification(str_value);
    <% End If %>
	document.all.notes.value = str_value;
}

function fn_KeyDownInList(event)
{
    var lKey = event.keyCode;
    var sCharacter = String.fromCharCode(lKey);
    var lPickRow = 1;
    var sJustification;
    
    sCharacter = sCharacter.toUpperCase();
    for (lRow = 1; lRow <= <%=justi_oJustificationList.length %>; lRow++) {
        sJustification = eval("document.all.td" + lRow).innerText;
        sJustification = sJustification.toUpperCase().substring(0,1);
        if (sCharacter >= sJustification) {
            lPickRow = lRow;
        }
    }
    fn_select(lPickRow);
}

</script>

<!-- Declaration of Page Scope variables Starts -->
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body>
    <form action="/E2B/Actions/RejectE2BCase.asp" method="post" name="RejectForm" class="margin_0_override">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <!-- Outer Box Starts -->
        <table class="table border" width="100%" height="100%" cellspacing="0" cellpadding="0">
            <tr height="25px">
                <td class="section-header-middle">
                    <%BuildLocalLabel("REJECT_INITIAL").SetStyleSheet("label label-section").Render()%>
                </td>
            </tr>
            <tr>
                <td class="padding-all">
                    <table class="table inner-table" width="100%" height="100%" cellspacing="0" cellpadding="0">
                        <col width="25%" />
                        <col />
                        <tr height="50px">
                            <td colspan="2">
                                <%BuildLocalLabel("REJECT_E2B_CASE_INSTRUCTIONS").Render()%>
                            </td>
                        </tr>
                        <tr height="20px">
                            <td>
                                <%BuildLocalLabel("USER_NAME").Render()%>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "txtUserName", GetXMLValueDirect(oSession, "CFG_USERS_USER_FULLNAME"), true, -1, "").style("width:100%").Render()%>
                            </td>
                        </tr>
                        <% If Not bReAuthEnabled Then %>
                        <tr style="height: 20px">
                            <td>
                                <%BuildLocalLabel("PASSWORD").Render()%>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_PASSWORDBOX, "password", "", false, 1, "").style("width:100%").Render()%>
                            </td>
                        </tr>
                        <%Else %>
                        <tr style="height: 25px">
                            <td>
                                <a href="#" id="reauthenticate" onclick="javascript:fn_ReAuthenticate('<%=GetReAuthUrl() %>');return false;" tabindex="1"><%=GetTranslationData("CLICK_TO_AUTHORIZE") %></a>
                            </td>
                            <td>
                                <%BuildImage("info_pending", "/img/common/unauthorized.png").Style("display:block").Render() 
                                        BuildImage("info_done", "/img/common/authorized.png").Style("display:none").Render()%>
                            </td>
                        </tr>
                        <%Call BuildHiddenControlDirect("password", "")%>
                        <%Call BuildHiddenControlDirect("ReAuth", "")%>
                        <%Call BuildHiddenControlDirect("code", "")%>
                        <%Call BuildHiddenControlDirect("state", "")%>
                        <%End If %>
                        <tr height="20px">
                            <td>
                                <%BuildLocalLabel("DATE").Render()%>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_DATETIME, "txtDate", fn_date_from_iso_gmt(TodayNowGMT(), 8, true), true, -1, "").Render()%>
                            </td>
                        </tr>
                        <tr height="20px">
                            <td colspan="2">
                                <%BuildLocalLabel("NOTES").Render()%>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <%BuildControlDirect(CTL_TEXTAREA, "notes", "", false, 3, "").SetRowCols(8, 65).Style("width:100%;height=100%").OnKeyUp("fn_maxlength(this, 200)").Render()%>
                            </td>
                        </tr>
                        <tr height="20px">
                            <td colspan="2">
                                <%BuildLocalLabel("E2B_CASE_STD_MSG").Render()%>
                            </td>
                        </tr>
                        <tr height="40%">
                            <td colspan="2">
                                <div tabname="PatientHistory" class="Table-Scroll" style="height: 100%"
                                    onkeydown="fn_KeyDownInList(event);">
                                    <table width="100%" border="0" cellspacing="0">
                                        <%justi_row = 0
                                Dim strRaw, lID, iPos, strJustification, strJustification_j, strTemp
                                for each justi_oJustification in justi_oJustificationList
                                    strRaw = GetString(GetXMLValueDirect(justi_oJustification, "GN_GUI_LM_GENERAL_TEXT"), "")
                                    iPos = instr(strRaw, "&")
                                    If iPos > 0 Then
                                        lID = left(strRaw, iPos - 1)
                                        strJustification = mid(strRaw, iPos + 1)
                                    Else
                                        lID = 0
                                        strJustification = strRaw
                                    End If
                                    iPos = instr(strJustification, "$#*")
                                    strJustification_j = ""
                                    If iPos > 0 Then
                                        strJustification_j = mid(strJustification, iPos + 3)
                                        strJustification = left(strJustification, iPos - 1)
                                    End If
                                                
                                    If glDisplayLang = cfCMN_LANG_JP Then
                                        strTemp = strJustification
                                        strJustification = strJustification_j
                                        strJustification_j = strTemp
                                    End If
                                    justi_row = justi_row + 1%>
                                        <tr>
                                            <td class="ddlist alc-header" id="td<%=justi_row%>" onclick="fn_select(<%=justi_row%>);">
                                                <%=Fn_Sanitize(strJustification)%>
                                            </td>
                                        </tr>
                                        <%next%>
                                    </table>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr align="center" class="tblheader-gray" height="25">
                <td>
                    <%BuildButton("b_Ok", "OK", 4).Style("width:50px").onClick("fn_ClickOK()").Render()%>
                    <%BuildButton("cancelbtn", "CANCEL", 5).Style("Width:60px").onClick("window.close()").Render()%>
                </td>
            </tr>
        </table>
        <!-- Outer Box Ends -->
    </form>
</body>
</html>

<script runat="SERVER" language="VBSCRIPT">
Sub ProcessInput()
	Dim oMessage
    Dim lErrNo, sError
	field_id = 25150108
	justlen = 2000
	
	If (justi_sTitle = "") Then justi_sTitle = "Action Justification"

	If Len(field_id) > 0 Then
		Call CreateMessage (oMessage, 300200267)	' MID_db_app_ASP_special_SQL
		Call SetXMLValueDirect (oMessage, "GN_GUI_LM_GENERAL_ID", 4)
		Call SetXMLValueDirect (oMessage, "GN_GUI_PARAMETER", field_id)
		Set justi_oJustificationMsg = ServiceRequest(oArgusSvr, oMessage, lErrNo, sError)
		If lErrNo <> 0 Then
		    Call ExecuteErrorPage(lErrNo, sError)
		else
		    Set justi_oJustificationList = justi_oJustificationMsg.selectNodes ("/MESSAGE/TABLE_WEB_ARGUS_SESSION/WEB_ARGUS_SESSION")
		end if
	End If
End Sub

</script>

<!-- Page Display Ends -->
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
