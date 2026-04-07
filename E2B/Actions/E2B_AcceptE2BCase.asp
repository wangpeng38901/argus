<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : AcceptE2BCase.asp
' Description  : Accept E2B Case
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
	<title><%=GetTranslationData("ACCEPT_INITIAL")%></title>
	<!-- Include Stylesheet here -->
	<link rel="stylesheet" href="/css/Relsys.css" />
	<script type="text/javascript" src="/js/Common/CommonReAuthFun.js"></script>
	<!-- Client Library Includes Starts -->
	<!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
	Dim  esm_report_id, AutoNumb
	Dim product_name, report_type_id, country_id, receipt_date, selective_intake, lIsJReport
	Dim Bulk, CustomImport, lError, sError
	Dim bReAuthEnabled,bOIDCReAuthMode
%>
<%	Dim justi_sTitle, justi_oJustificationMsg, justi_oJustification, justi_oJustificationList, justi_row
	Dim field_id, justlen
	Dim selectedIsApplyNewFw
	Dim sDocId

	bReAuthEnabled = IsReAuthEnabled()
    bOIDCReAuthMode = IsOIDCReAuthMode()
	
	ProcessInput()
%>
<!-- Declaration of Page Scope variables Ends -->
<!-- Assign Values to Page Scope variables Starts -->
<!-- Assign Values to Page Scope variables Ends -->
<!-- Page Processing Starts -->
<%
	esm_report_id = GetLong(GetRequest("esm_report_id"), 0)
	receipt_date = GetRequest("receipt_date")
	product_name = GetRequest("product_name")
	report_type_id = GetLong(GetRequest("report_type_id"), 0)
	country_id = GetLong(GetRequest("country_id"), 0)
	selective_intake = GetLong(GetRequest("selectiveintake"), 0)	
	AutoNumb = GetXMLValueDirect (oSession, "NUMBERING_AUTO")
	Bulk = GetLong(GetRequest("Bulk"), 0)	
	sDocId = GetRequest("DocId")
	CustomImport = "1"
	lIsJReport = GetLong(GetRequest("IsJReport"), 0) ' 1 if Authority Id is 4 for the report else 0
	selectedIsApplyNewFw = GetLong(Request("ApplyNewFw"), 0)
%>
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body>
	<%Call BuildHiddenControlDirect("AutoNumb", AutoNumb) %>
	<%Call BuildHiddenControlDirect("esm_report_id", esm_report_id) %>
	<%Call BuildHiddenControlDirect("product_name", product_name) %>
	<%Call BuildHiddenControlDirect("report_type_id", report_type_id) %>
	<%Call BuildHiddenControlDirect("country_id", country_id) %>
	<%Call BuildHiddenControlDirect("receipt_date", receipt_date) %>
	<%Call BuildHiddenControlDirect("SelectiveIntake", selective_intake) %>
	<%Call BuildHiddenControlDirect("CustomImport", CustomImport) %>
	<%Call BuildHiddenControlDirect("IsJReport", lIsJReport) %>
	<%Call BuildHiddenControlDirect("postsave", "") %>
	<table class="table inner-table" style="width: 100%; height: 100%" cellpadding="0"
		cellspacing="0">
		<tr height="25px">
			<td class="section-header-middle">
				<% BuildLocalLabel("ACCEPT_INITIAL").SetStyleSheet("label label-section").Render()%>
			</td>
		</tr>
		<tr valign="top">
			<td class="padding-all">
				<table class="table inner-table" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%">
					<col width="25%" />
					<col />
					<tr height="75px">
						<td colspan="2">
							<% BuildLocalLabel("ACPT_E2B_CASE_INSTRUCTIONS").Render()%>
						</td>
					</tr>
					<tr height="25px">
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
							<% BuildControlDirect(CTL_PASSWORDBOX, "password", "", false, 1, "")._
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

					<tr height="20px">
						<td>
							<% BuildLocalLabel("DATE").Render()%>
						</td>
						<td>
							<%BuildControlDirect(CTL_TEXTBOX_DATETIME, "txtDate", fn_date_from_iso_gmt(TodayNowGMT(), 8, true), true, -1, "").Render()%>
						</td>
					</tr>
					<tr height="20px">
						<td colspan="2">
							<% BuildLocalLabel("NOTES").Render()%>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<%BuildControlDirect(CTL_TEXTAREA, "notes", "", false, 2, "").SetRowCols(8, 65).Style("width:100%").onKeyUp("fn_maxlength(this,190)").Render() %>
						</td>
					</tr>
					<tr height="25px">
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
				<%BuildButton ("btn_OK", "OK", 3).Style("width:50px").OnClick("fn_ClickOK();").Render() %>
				<%BuildButton ("btn_Cancel", "CANCEL", 4).Style("Width:60px").OnClick("fn_ClickCancel();").Render() %>
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
	var selected = 0;
	var sDocId = "<%=sDocId%>";
	var iDiffStatus = 0;
	var g_status = 0;
	var lApplyNewFW = <%=JavaScriptSanitize(selectedIsApplyNewFw) %>;
	var l_delaytime = <%=GetRequestDelayTime() %>;

async function fn_KeyPress() {
		if (window.event.keyCode == 13 && l_attempt < 3) {
			document.all.btn_OK.focus();

			await fn_ClickOK();
			return;
		}
	}

function fn_ClickCancel()
{
	window.close();
}

async function fn_ClickOK()
{
	var s_password, s_notes;
		var sURL, retVal, status;
		s_password = document.all.password.value;
		s_notes = document.all.notes.value;
	
	<% If bReAuthEnabled Then %>
	 if (document.all.ReAuth.value.length == 0) {
			await MessageBoxRes("RE_AUTH_JUSTIFICTN_PROVIDE_PASS");
			return -1;
		}
	<% Else %>
	if (s_password.length == 0) {
			await MessageBoxRes("JUSTIFICTN_PROVIDE_PASS");
			document.all.password.focus();
			return -1;
		}  
	<% End If %>

	    if (s_notes.length == 0) {
			await MessageBoxRes("ACCEPT_CASE_NOTES_REQ");
			document.all.notes.focus();
			return -1;
		}
      var sParams = "password=" + fn_URLEncode(s_password) + "&FormName=Accept ICSR Initial Report";
        <% If bReAuthEnabled And bOIDCReAuthMode Then %>
            sParams = sParams + "&code=" + fn_URLEncode(document.all.code.value) + "&state=" + fn_URLEncode(document.all.state.value);
        <% End If %>

		
		if (l_delaytime > 0) {
			showLoading();
			await sleep(l_delaytime);
		}
		await loadArgusMessage("/CaseForm/Dialogs/CF_AjaxPasswordVerify.asp", fn_password_verify, sParams);
		
		if (!bPasswordOK) {
			document.all.password.focus();
			document.all.password.select();
			l_attempt++;
			if (l_attempt < <%=GetPwdTriesCount()%>) return;
            await MessageBoxRes("E2BACCEPT_PASS_TRY_EXCEED");
            await fn_SessionTimeout();
            window.close();
			return;
		}
	<% If(Bulk = 0) Then %>
		var strURL, case_num = "";
	    if (document.all.AutoNumb.value == 0) {
		    strURL = "/CaseForm/Actions/ManualCaseNum.asp";
		    var strStyle = { dialogHeight: "150", dialogWidth: "300", resizable: false, scrollable: false };
		    case_num = await fn_OpenModalDialog(strURL, window, strStyle);
	    }
		if (case_num != undefined) {
			if (case_num.toLowerCase() == "error")
				await MessageBoxRes("E2BACCEPT_ERROR");
			else {
				showLoading();
				if (document.all.SelectiveIntake.value == 0) {
					strURL = "/e2b/e2bimport/Ajax_E2BDifferenceReport.asp?IsJReport=" + document.all.IsJReport.value + "&report_id=" + document.all.esm_report_id.value + "&e2b_type=1&call_pdf=1" + "&ApplyNewFw=" + lApplyNewFW;
					await loadArgusMessage(strURL, fn_CallBack_DiffData);					
				}
                if (iDiffStatus == 0) {
                    var accept_date = document.getElementById("txtDate").value;
                    strURL = "/E2B/Misc/GetLoadE2BStagingWarning.asp?check_warning=0&custom_import=1&esm_report_id=" + document.all.esm_report_id.value + "&case_num=" + fn_URLEncode(case_num) + "&notes=" + fn_URLEncode(document.all.notes.value);
                    strURL = strURL + "&E2bType=1&DocId=" + sDocId + "&accept_date=" + accept_date + "&IsJReport=" + document.all.IsJReport.value;
					strStyle = { dialogHeight: "560", dialogWidth: "840", resizable: false, scrollable: false };
					var status = await fn_OpenModalDialog(strURL, window, strStyle);
					if (status !== sSessionTimeOutDialogReturn) {					   
						fn_execute(status, 0);						
					}
				}
				else {
					await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + document.all.esm_report_id.value);
				}
			}
			if (status === sSessionTimeOutDialogReturn) {
				setWindowReturnValue(sSessionTimeOutDialogReturn);
				window.close();
			}
			else {
				fn_execute(status, 1);
			}
		}
	
		<% Else %>
			setWindowReturnValue(document.all.notes.value);
		window.close();
	<% End If %>
}

async function fn_UnLockedReport() {
	var xmlDoc = this.req.responseXML;
	var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
	if (sErrStr.length > 0) {
		await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("INCOME_E2B")%>', sErrStr);
		return;
	}
	return;
}

async function fn_CallBack_DiffData() {
	var xmlDoc = this.req.responseXML;
	var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);

	if (sErrStr.length > 0) {
		iDiffStatus = -1;
		await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("INIT_SELECT")%>', sErrStr);
		hideLoading();
	}
	else {
		var asDocId = xmlDoc.getElementsByTagName("GN_REPORT_IDENTIFIER");
		if (asDocId && (asDocId.length > 0))
			sDocId = GetTextContentFromXML(asDocId[0]);
	}
}

	function fn_execute(notes, bClose) {
		var iStatus, Pos, warning, case_numb, saveError;
		
		iStatus = "-1";
		if (notes) {
			Pos = notes.indexOf("~", 1);
			if (Pos > 0) {
				case_numb = " ";
				warning = notes.substring(0, Pos); //either 0 or 1
				if (warning == "NaN")
					warning = "1";
				notes = notes.substring(Pos + 1); //status 102 or 101 + case_number
				Pos = notes.indexOf("~", 1);
				if (Pos > 0) {
					iStatus = notes.substring(0, Pos);
				    notes = notes.substring(Pos + 1);
				    Pos = notes.indexOf("~", 1);
				    if (Pos > 0)
				    {
					    case_numb = notes.substring(0, Pos);
					    saveError = notes.substring(Pos + 1);
					    if ((bClose == 1) && (saveError.length > 0))
					    {
						    document.all.postsave.value = saveError;
					    }
				    }
				    else
				    {
					    case_numb = notes;
				    }
			    }
			    else
				    iStatus = "-1";
			}
		}

		g_status = iStatus;
		if (bClose == 1) {
			setWindowReturnValue(warning + "~" + iStatus + "~" + case_numb + "~" + document.all.notes.value + "~" + document.all.postsave.value);
			window.close();
		}
	}	

	async function fn_password_verify() {
		if (l_delaytime > 0) {
			hideLoading();
		}
		var sErrMsg = fn_GetAjaxErrorMsg(this.req.responseXML);
		if (sErrMsg.length > 0) {
			bPasswordOK = false;
            await MessageBoxRes("GENERAL_WARNING", '<%=GetTranslationData("E2B_ACCEPT_CASE")%>', sErrMsg);
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

	function fn_KeyDownInList(event) {
		var lKey = event.keyCode;
		var sCharacter = String.fromCharCode(lKey);
		var lPickRow = 1;
		var sJustification;

		sCharacter = sCharacter.toUpperCase();
		for (lRow = 1; lRow <= <%= justi_oJustificationList.length %>; lRow++) {
			sJustification = eval("document.all.td" + lRow).innerText;
			sJustification = sJustification.toUpperCase().substring(0, 1);
			if (sCharacter >= sJustification) {
				lPickRow = lRow;
			}
		}
		fn_select(lPickRow);
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
