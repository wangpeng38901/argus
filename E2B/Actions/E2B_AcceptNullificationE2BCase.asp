<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : AcceptNullificationE2BCase.asp
' Description  : Accept Nullification E2B Case
'******************************************************************************
' Revision History
' Date		   Author		Description
' 15MAY2006 	Umar       	Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title><%=GetTranslationData("ACPT_NULLIFICATION_E2B_CASE")%></title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <script type="text/javascript" src="/js/Common/CommonReAuthFun.js"></script>
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
	Dim strSQL, CaseNum, esm_report_id, case_id, CaseCloseStatus
	Dim Bulk, lError, sError
    Dim bReAuthEnabled,bOIDCReAuthMode
%>
<%	Dim justi_sTitle, justi_oJustificationMsg, justi_oJustification, justi_oJustificationList, justi_row
	Dim field_id, justlen

    bReAuthEnabled = IsReAuthEnabled()
    bOIDCReAuthMode = IsOIDCReAuthMode()
	
	ProcessInput()%>
<!-- Declaration of Page Scope variables Ends -->
<!-- Assign Values to Page Scope variables Starts -->
<!-- Assign Values to Page Scope variables Ends -->
<!-- Page Processing Starts -->
<%
    Bulk = GetLong(GetRequest("Bulk"), 0)
	esm_report_id = GetLong(GetRequest("esm_report_id"), -1)
	case_id = GetLong(GetRequest("case_id"), 0)
	CaseCloseStatus = GetLong(GetRequest("CloseStatus"), -1)

	If (Bulk = 0) Then 
	    Call SetParameter("P_REPORT_ID", esm_report_id, PARAM_NUMBER)
	    strSQL = "select c.case_num from safetyreport s, case_master c "
	    strSQL = strSQL & "where s.report_id = :P_REPORT_ID and s.case_xref = c.case_id"
	    CaseNum = ExecuteSQLReturnStr (strSQL, lError, sError)
	    
	    If CaseNum = "" Then
		    strSQL = "SELECT c.case_num FROM case_master c, safetyreport a "
            strSQL = strSQL & "WHERE a.report_id = (SELECT MAX (b.report_id) "
            strSQL = strSQL & "FROM safetyreport b "
            strSQL = strSQL & "WHERE b.status = 102 AND b.reportacknowledgment.reportacknowledgmentcode = '01' "
            strSQL = strSQL & "AND (upper(b.companynumb) = (SELECT upper(ss.companynumb) FROM safetyreport ss "
            strSQL = strSQL & "WHERE ss.report_id = :P_REPORT_ID and b.sender_agency=ss.sender_agency) OR upper(b.authoritynumb) = "
            strSQL = strSQL & "(SELECT upper(ss.authoritynumb) FROM safetyreport ss "
            strSQL = strSQL & "WHERE ss.report_id = :P_REPORT_ID and b.sender_agency=ss.sender_agency))) "
            strSQL = strSQL & "AND a.case_xref = c.case_id "
		    CaseNum = ExecuteSQLReturnStr (strSQL, lError, sError)
        End If
	End If
%>
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body>
    <%Call BuildHiddenControlDirect("esm_report_id", esm_report_id) %>
    <table class="table border-blue inner-table" style="width: 100%; height: 100%" cellpadding="0"
        cellspacing="0">
        <tr height="25px">
            <td class="section-header-middle">
                <% BuildLocalLabel("ACPT_NULLIFICATION_E2B_CASE").SetStyleSheet("label label-section").Render()%>
            </td>
        </tr>
        <tr valign="top">
            <td class="padding-all">
                <table class="table inner-table" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%">
                    <col width="25%" />
                    <col />
                    <tr height="75px">
                        <td colspan="2">
                            <% BuildLocalLabel("ACPT_NULLIFICATIN_E2B_CASE_INSTRUCTIONS").Render()
                            %>
                        </td>
                    </tr>
                    <tr height="20px">
                        <td>
                            <%BuildLocalLabel("USER_NAME").Render()%>
                        </td>
                        <td>
                            <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "txtUserName", GetXMLValueDirect(oSession, "CFG_USERS_USER_FULLNAME"), true, -1, "").Style("width:100%").Render()%>
                        </td>
                    </tr>
                    <% If Not bReAuthEnabled Then %>
                    <tr style="height: 20px">
                        <td>
                            <%BuildLocalLabel("PASSWORD").Render()%>
                        </td>
                        <td>
                            <%
                            BuildControlDirect(CTL_PASSWORDBOX, "password", "", false, 1, "")._
                            Style("width:100%").SetMaxLength("30").onKeyPress("fn_KeyPress()").Render() %>
                        </td>
                    </tr>
                    <%Else %>
                    <tr style="height: 20px">
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
                    <tr height="25px">
                        <td>
                            <% BuildLocalLabel("DATE").Render()%>
                        </td>
                        <td>
                            <%BuildControlDirect(CTL_TEXTBOX_DATETIME, "txtDate", fn_date_from_iso_gmt(TodayNowGMT(), 8, true), true, -1, "").Render()%>
                        </td>
                    </tr>
                    <tr height="20px">
                        <td>
                            <% BuildLocalLabel("ACT_ITEMS").Render()%>
                        </td>
                        <td>
                            <table class="table" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td width="150px">
                                        <%BuildControl("CSAC_CODE", "action_type", -1, false, 2, "True").style("width:100%").Onchange("on_ActionItem_change(this)").Render()%>
                                    </td>
                                    <td>&nbsp;&nbsp;<%BuildLocalLabel("DUE_IN").Render()%>&nbsp;&nbsp;
                                        <%BuildControlDirect(CTL_TEXTBOX_NUMERIC, "due_in", "", true, 3, "").SetMaxLength(3).Style("width:30%").Render()%>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr height="20px">
                        <td>
                            <%BuildLocalLabel("USER_GROUP").Render()%>
                        </td>
                        <td>
                            <%BuildTypeAheadDirectQuick("CSCLG_GROUP_ID", "user_group", -1, false, 4, 0,"True").Style("width:100%").Render()%>
                        </td>
                    </tr>
                    <tr height="20px">
                        <td colspan="2">
                            <% BuildLocalLabel("NOTES").Render()%>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <%BuildControlDirect(CTL_TEXTAREA, "notes", "", false, 5, "").SetRowCols(8, 65).Style("width:100%").onKeyUp("fn_maxlength(this,190)").Render() %>
                        </td>
                    </tr>
                    <tr height="20px">
                        <td colspan="2">
                            <%BuildLocalLabel("E2B_CASE_STD_MSG").Render()%>
                        </td>
                    </tr>
                    <tr height="40%">
                        <td colspan="2">
                            <div tabname="PatientHistory" class="Table-Scroll" style="width: 100%; height: 100%"
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
        <tr class="tblheader-gray" height="25px">
            <td align="center">
                <%BuildButton("btn_OK", "OK", 6).Style("width:50px").OnClick("fn_ClickOK();").Render() %>
                <%BuildButton("btn_Cancel", "CANCEL", 7).Style("Width:60px").OnClick("fn_ClickCancel();").Render() %>
            </td>
        </tr>
    </table>
</body>
</html>
<!-- Page Display Ends -->
<!-- Javascript Functions Starts -->

<script type="text/javascript">
var objParent = window.dialogArguments;
var l_attempt = 0;
var bPasswordOK = false;
var bSaveFollowup = false;
var selected = 0;
var l_OldAIValue = 0;
var l_delaytime = <%=GetRequestDelayTime() %>;

async function fn_KeyPress()
{
	if (window.event.keyCode == 13 && l_attempt < 3)
	{
		document.all.btn_OK.focus();
		await fn_ClickOK();
		return;
	}
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

function fn_ClickCancel()
{
	window.close();
}

async function fn_ClickOK()
{
	var sURL, retVal;
	var s_password, s_notes;
	var strURL, ret, iStatus;
	var CloseStatus = <%=JavaScriptSanitize(CaseCloseStatus) %>;
	var CaseID = <%=JavaScriptSanitize(case_id) %>;
	var dialog_result = "1";
	
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
    var sParams = "password=" + fn_URLEncode(s_password) + "&FormName=Accept ICSR Nullification Report";
    <% If bReAuthEnabled And bOIDCReAuthMode Then %>
            sParams = sParams + "&code=" + fn_URLEncode(document.all.code.value) + "&state=" + fn_URLEncode(document.all.state.value);
    <% End If %>

	
	if (l_delaytime > 0) {
			showLoading();
			await sleep(l_delaytime);
	}

    await loadArgusMessage("/CaseForm/Dialogs/CF_AjaxPasswordVerify.asp", fn_password_verify, sParams);
    
	if (!bPasswordOK)
	{
		document.all.password.focus();
		document.all.password.select();
		l_attempt ++;
		if (l_attempt < <%=GetPwdTriesCount()%>) return;
        await MessageBoxRes("NULLIFICATION_PASS_TRY_EXCEED"); 
        await fn_SessionTimeout();
        window.close();
        return;
    }
    
	<% If (Bulk = 0) Then %>
        if (CloseStatus == 1)
	        dialog_result = await fn_ReOpenCase(CaseID);
        
        if (dialog_result == "1")
        {
		    strURL = "/E2B/Misc/AjaxSaveFollowUpDiffData.asp";
            await loadArgusMessage(strURL, fn_CheckErrors, "E2bType=4&case_id=" + CaseID + "&esm_report_id=" + document.all.esm_report_id.value + "&notes=" + fn_URLEncode(document.all.notes.value));
            
		    if (bSaveFollowup)
		    {
			    await fn_create_actionitem();
			    iStatus = "102";
                await MessageBoxRes("NULLIFICATION_DEL_ERROR", "", <%=JavaScriptSanitize(CaseNum) %>);
		    }	
		    else
		    {
			    iStatus = "-1";
                await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + document.all.esm_report_id.value);
		    }
		}
    <% Else %>
	    await fn_create_actionitem();
	    iStatus = "102";
	<% End If %>
	if (dialog_result == "1")
	{
	    setWindowReturnValue("0" + "~" + iStatus + "~" + " " + "~" + s_notes);
	    window.close();
	}
}

async function fn_create_actionitem()
{
    var sActionItemURL;
    var act_type_id = document.getElementById("action_type").value;
    var user_grp_id = document.getElementById("user_group").value;
	var due_in = fn_getElementByName("due_in").value;
    if (due_in == "")
        due_in = 0;
    var sesm_report_ids = <%=JavaScriptSanitize(esm_report_id) %>;
    var lBulk = <%=JavaScriptSanitize(Bulk)%>;
    var sNotes;
    var sNullText = '<%=GetTranslationData("NULLIFICATION")%>';
    
    if (act_type_id > 0)
    {
		var str_date = fn_GetArgusDate(fn_getElementByName("txtDate").value);
        var str_months = "JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC";
        var yr_num, mo_num, day_num, hour_num, min_num;
        yr_num = str_date.substr(7,4);
	    mo_num = str_months.search(str_date.substr(3,3)) / 4;
	    day_num = str_date.substr(0,2);
	    hour_num = str_date.substr(12,2);
	    min_num = str_date.substr(15,2);
	    var oDate = new Date(yr_num, mo_num, day_num, hour_num, min_num, "00");
        oDate.setDate(oDate.getDate() + parseInt(due_in));
        str_date = oDate.getDate() + "-" + (oDate.getMonth()+1) + "-" + oDate.getFullYear() + " " + oDate.getHours() + ":" + oDate.getMinutes();
        
        sNotes = fn_URLEncode(sNullText + ": " + document.all.notes.value);
        
        sActionItemURL = "/E2B/E2BImport/Ajax_E2BActionItem.asp?action_type=" + act_type_id + "&user_group=" + user_grp_id + "&due_date=" + str_date + "&desc=" + sNotes + "&esm_report_id=" + sesm_report_ids + "&bulk=" + lBulk;
        await loadArgusMessage(sActionItemURL, fn_ActionItem);
        
    }
}

async function fn_ActionItem() {
        var xmlDoc = this.req.responseXML; 
        var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
        if (sErrStr.length > 0)
            await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("PEND_RPT")%>', sErrStr);
    }

async function fn_password_verify() {
		if (l_delaytime > 0) {
			hideLoading();
		}
        var sErrMsg = fn_GetAjaxErrorMsg(this.req.responseXML);
        if (sErrMsg.length > 0) {
            bPasswordOK = false;
            await MessageBoxRes("GENERAL_WARNING", '<%=GetTranslationData("ACCEPT_NULL_CASE")%>', sErrMsg);
            <% If bReAuthEnabled Then %>
               document.all.info_pending.style.display = 'block';
               document.all.info_done.style.display = 'none';
               document.all.ReAuth.value = "";
               document.all.code.value = "";
               document.all.state.value = "";
            <% End If %>
        }
        else {
            bPasswordOK = true;
        }
    }

async function fn_CheckErrors()
{
    var xmlDoc = this.req.responseXML;
    var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
    if (sErrStr.length > 0)
    {
        bSaveFollowup = false;
        await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("E2B_INCOME_RPTS")%>', sErrStr); 
    }
    else
    {
        bSaveFollowup = true;
    }
}

async function fn_UnLockedReport()
{
    var xmlDoc = this.req.responseXML;
    var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
    if (sErrStr.length > 0)
    {
        await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("INCOME_E2B")%>', sErrStr); 
    }
    return;
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
fn_DisableControl(fn_getElementByName("TXT_user_group"));

function on_ActionItem_change(oDDL)
{
    var TypeAhead_Hdnid = oDDL.id.substr(4);
    var TypeAhead_HdnCtrl = document.getElementById(TypeAhead_Hdnid);
	var oDueIn = fn_getElementByName("due_in");
	var oUserGroup = fn_getElementByName("TXT_user_group");
    if (l_OldAIValue != TypeAhead_HdnCtrl.value)
    {
        if(TypeAhead_HdnCtrl.value > 0)
        {
            oDueIn.style.backgroundColor = ccf_normal_text_color;
            oDueIn.disabled = false;
            oDueIn.readOnly = false;
            oDueIn.select();
            fn_EnableControl(oUserGroup);
        }
        else
        {
            oDueIn.value = "";
            oUserGroup.value = "";
            oDueIn.style.backgroundColor = ccf_readonly_text_color;
            oDueIn.disabled = true;
            oDueIn.readOnly = true;
            fn_DisableControl(oUserGroup);
        }
        l_OldAIValue = TypeAhead_HdnCtrl.value;
    }
}

</script>

<!-- Javascript Functions Ends -->

<script runat="SERVER" language="VBSCRIPT">
Sub ProcessInput()
	Dim oMessage
    Dim lErrNo, sError
	field_id = 25150107
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

<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
