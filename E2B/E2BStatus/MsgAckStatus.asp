<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : MsgAckStatus.asp
' Description  : 
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
    <title>Message Acknowledgment Status</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
		
	Dim AckType, MessageNumb, ErrorNum, Error, oOutMsg, lGmtOffSet
	Dim icsr_message_numb, icsr_sender_id, icsr_message_date, transmit_ack_code, sender_ack_message_numb, icsr_receiver_id, ack_date
	Dim oRecList, lNumRec
	Dim oRec, CaseNumb, LocalNumb
	
	AckType = GetRequest("AckType")
	MessageNumb = GetLong(GetRequest("MessageNumb"), 0)
	lGmtOffSet = GetFloat(GetRequest("GmtOffSet"), 0)
	If ExecACKStatus() Then
		Set oRecList = oOutMsg.selectnodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")
		icsr_message_numb = GetXMLValueDirect(oOutMsg,"/MESSAGE/TABLE_RPT_E2B/RPT_E2B/RPT_E2B_LMESSAGENUMB")
		icsr_sender_id = GetXMLValueDirect(oOutMsg,"/MESSAGE/TABLE_RPT_E2B/RPT_E2B/RPT_E2B_MESSAGE_SENDER_ID")
		icsr_message_date = GetXMLValueDirect(oOutMsg,"/MESSAGE/TABLE_RPT_E2B/RPT_E2B/RPT_E2B_MESSAGE_DATE")
		transmit_ack_code = GetXMLValueDirect(oOutMsg,"/MESSAGE/TABLE_RPT_E2B/RPT_E2B/RPT_E2B_PARSING_MESSAGE_ERROR")
		sender_ack_message_numb = GetXMLValueDirect(oOutMsg,"/MESSAGE/TABLE_RPT_E2B/RPT_E2B/RPT_E2B_RMESSAGENUMB")
		icsr_receiver_id = GetXMLValueDirect(oOutMsg,"/MESSAGE/TABLE_RPT_E2B/RPT_E2B/RPT_E2B_MESSAGE_RECEIVER_ID")
		ack_date = GetXMLValueDirect(oOutMsg,"/MESSAGE/TABLE_RPT_E2B/RPT_E2B/RPT_E2B_ACK_MESSAGE_DATE")
		lNumRec = oRecList.length
	Else
		Call ExecuteErrorPage(ErrorNum, Error)
	End If
%>
<body>
    <form name="Frm" id="Frm" action="" method="post">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <!-- Outer Box Starts -->
        <table class="table" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%">
            <tr style="height: 25px">
                <td class="section-header-middle">
                    <%BuildLabelDirect("Message Acknowledgement Status").SetStyleSheet("label label-section").Render()%>
                </td>
            </tr>
            <tr style="height: 50px">
                <td valign="top">
                    <table class="table border-blue inner-table" width="100%" cellspacing="2">
                        <col width="23%" />
                        <col width="22%" />
                        <col width="30%" />
                        <col width="25%" />
                        <tr>
                            <td>
                                <%BuildLabelDirect("ICSR Message Number").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "icsr_message_numb", icsr_message_numb, true, 1, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                            </td>
                            <td>
                                <%BuildLabelDirect("Acknowledgment Message # ").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "sender_ack_message_numb", sender_ack_message_numb, True, 2, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%BuildLabelDirect("ICSR Message Date").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "icsr_message_date", icsr_message_date, True, 3, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                            </td>
                            <td>
                                <%BuildLabelDirect("Acknowledgment Message Date ").Render() %>
                            </td>
                            <td>
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "ack_date", ack_date, True, 4, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%BuildLabelDirect("ICSR Message Sender ID").Render() %>
                            </td>
                            <td colspan="3">
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "icsr_message_date", icsr_sender_id, True, 5, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%BuildLabelDirect("ICSR Message Receiver ID ").Render() %>
                            </td>
                            <td colspan="3">
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "icsr_receiver_id", icsr_receiver_id, True, 6, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%BuildLabelDirect("Transmission ACK Code ").Render() %>
                            </td>
                            <td colspan="3">
                                <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "transmit_ack_code", transmit_ack_code, True, 7, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr style="height: 50px">
                <td valign="top">
                    <div id="Divheader" class="table-scroll Scroll_hide" style="height: 100%; border-bottom-width: 0px; overflow-y: scroll;">
                        <table cellpadding="0" cellspacing="0" class="table inner-table" style="width: 100%; height: 100%; border: 0;">
                            <col width="23%" />
                            <col width="23%" />
                            <col width="18%" />
                            <col width="18%" />
                            <col width="18%" />
                            <tr class="tblheader-lightblue">
                                <td class="alc-header" style="padding-left: 5px; padding-top: 5px; padding-bottom: 5px;">
                                    <span class="label label-section" style='color: black; background-color: Transparent; font-size: 8pt'>Report Number </span>
                                </td>
                                <td class="alc-header" style="padding-left: 5px;">
                                    <span class="label label-section" style='color: black; background-color: Transparent; font-size: 8pt'>Authority / Company Number </span>
                                </td>
                                <td class="alc-header" style="padding-left: 5px;">
                                    <span class="label label-section" style='color: black; background-color: Transparent; font-size: 8pt'>Other Number </span>
                                </td>
                                <td class="alc-header" style="padding-left: 5px;">
                                    <span class="label label-section" style='color: black; background-color: Transparent; font-size: 8pt'>Report Status </span>
                                </td>
                                <td class="alc-header" style="padding-left: 5px;">
                                    <span class="label label-section" style='color: black; background-color: Transparent; font-size: 8pt'>Report Type </span>
                                </td>
                            </tr>
                            <tr class="tblheader-lightblue">
                                <td colspan="6" class="padding-all">
                                    <span class="label label-section" style='color: black; background-color: Transparent; font-size: 8pt'>Error Message / Comments </span>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td valign="top" width="100%" style="width: 100%;">
                    <div id="DivResult" class="table-scroll" style="height: 100%; overflow-y: scroll">
                        <table cellpadding="0" cellspacing="0" width="100%" border="0" class="table inner-table">
                            <col width="23%" />
                            <col width="23%" />
                            <col width="18%" />
                            <col width="18%" />
                            <col width="18%" />
                            <% 
						Dim lCount
						lCount = 0  
						For Each oRec in oRecList 
							lCount = lCount + 1
	                        CaseNumb = GetXMLValueDirect(oRec,"RPT_E2B_SAFETYREPORTID")
	                        LocalNumb = GetXMLValueDirect(oRec,"RPT_E2B_AUTHORITYNUMB")
                            if LocalNumb = "" then
                                LocalNumb = GetXMLValueDirect(oRec,"RPT_E2B_COMPANYNUMB")
                            end if
                            %>
                            <tr valign="top">
                                <td class="grd-header" style="padding-left: 5px;" onclick="this.focus();">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(CaseNumb) %>" style="width: 100%;" id="text1<%=lCount%>"
                                        name="text1<%=lCount%>" onkeypress="return false;" onkeydown="return false;">
                                </td>
                                <td class="grd-header" style="padding-left: 5px;">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(LocalNumb) %>" style="width: 100%;" id="text3<%=lCount%>"
                                        name="text3<%=lCount%>" onkeypress="return false;" onkeydown="return false;">
                                </td>
                                <td class="grd-header" style="padding-left: 5px;">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(GetXMLValueDirect(oRec,"RPT_E2B_OTHERNUMB")) %>"
                                        style="width: 100%;" id="text4<%=lCount%>" name="text4<%=lCount%>" onkeypress="return false;"
                                        onkeydown="return false;">
                                </td>
                                <td class="grd-header" style="padding-left: 5px;">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(GetXMLValueDirect(oRec,"RPT_E2B_REPORT_ACK_CODE")) %>"
                                        style="width: 100%;" id="text5<%=lCount%>" name="text5<%=lCount%>" onkeypress="return false;"
                                        onkeydown="return false;">
                                </td>
                                <td class="grd-header" style="padding-left: 5px;">
                                    <input class="textbox-list" value="<%=Fn_Sanitize(GetXMLValueDirect(oRec,"RPT_E2B_FILE_TYPE")) %>"
                                        style="width: 100%;" id="text6<%=lCount%>" name="text6<%=lCount%>" onkeypress="return false;"
                                        onkeydown="return false;">
                                </td>
                            </tr>
                            <tr>
                                <td colspan="6" class="no-padding-left" align="left">
                                    <%BuildControlDirect(CTL_TEXTAREA, "comment", Mid (GetXMLValueDirect(oRec,"RPT_E2B_ERROR_MESSAGE_COMMENT"),1, 32767), true, 2, "").onkeypress("return false;").Style("width:100%;height:225px;").SetMaxLength(32767).Render()%>
                                </td>
                            </tr>
                            <%Next %>
                        </table>
                    </div>
                </td>
            </tr>
            <tr class="tblheader-gray" style="height: 25px">
                <td align="center" valign="middle">
                    <%BuildButtonDirect("btnPrint", "Print", 100).Style("width:60px")_
				.OnClick("fn_PrintList();").Render()%>
                    <%BuildButtonDirect("btncLOSE", "Close", 102).Style("width:60px")_
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
	var sAckType = <%=JavaScriptSanitize(AckType) %>;
	var sMessageNumb = <%=JavaScriptSanitize(MessageNumb)%>;
    var o = document.querySelector('input[name=icsr_message_numb]');
    o.setSelectionRange(0, 3);


	function fn_CheckFields() {
        document.querySelector('input[name=icsr_message_numb]').focus();
        document.querySelector('input[name=icsr_message_numb]').select();
	}
	
    async function fn_PrintList() {
		var strURL;
		var testdate, lGmtOffSet;
		testdate = new Date();
		lGmtOffSet = (-testdate.getTimezoneOffset() / 60);
		strURL = "/E2B/E2BStatus/MsgAckStatusPrint.asp?AckType=" + fn_URLEncode(sAckType) + "&MessageNumb=" + fn_URLEncode(sMessageNumb) + "&GmtOffSet=" + lGmtOffSet;
        var sDialogStyle = { dialogHeight: "415", dialogWidth: "775", resizable: false, scrollable: false };    
       await fn_OpenModalDialog(strURL, window, sDialogStyle);
	}
</script>
<!-- Javascript Functions Ends -->

<!-- Server-Side Functions Starts -->
<script language="vbscript" runat="server">
Function ExecACKStatus()
	Dim oMessage
	ErrorNum = 0
	Error = ""

	Call CreateMessage (oMessage, 300400004)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_ACK_TYPE", AckType)
	Call SetXMLValueDirect (oMessage, "RPT_E2B_LMESSAGENUMB", MessageNumb)
	Call SetXMLValueDirect (oMessage, "GN_PRINT", 0)
	Call SetXMLValueDirect (oMessage, "GN_RPT_GMT_OFFSET", lGmtOffSet)
	
	Set oOutMsg = ServiceRequest(oArgusSvr, oMessage,ErrorNum, Error)
	If ErrorNum <> 0 Then
		ExecACKStatus = FALSE
		Exit Function
	End If
	ExecACKStatus = TRUE		
End Function
</script>
<!-- Server-Side Functions Ends -->

<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
