<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Umar Rehman
' Page         : ReportTransmissionHistory.asp
' Description  : Display the Transmission History
'******************************************************************************
' Revision History
' Date			Author		Description
' 31MAY2006 	Umar       	Original
'******************************************************************************
%>
<!DOCTYPE html>
<html>
<head>
    <!-- Page Title -->
    <title>Transmission History</title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <!-- Client Library Includes Starts -->
    <!-- Client Library Includes Ends -->
</head>
<!-- Declaration of Page Scope variables Starts -->
<%
	Dim reportId, oOutMessage, oTransmissionList, oTransmission, screen, messageId
	Dim casenum,agency, lGmtOffSet
	Dim totalTime
	Dim query, timediff
	Dim d1,d2, hr,min, lError, sError, lCount, sClass
%>
<!-- Declaration of Page Scope variables Ends -->
<!-- Page Processing Starts -->
<%
	reportId = GetRequest("reg_report_id")
	screen = GetRequest("screen")
	messageId = GetRequest("messageId")
	agency = GetRequest("agency")
	casenum = GetRequest("casenum")
	lGmtOffSet = GetRequest("GmtOffSet")

	if (Not(isnull(reportId))) Then
		getTransmissionHistory
	end if
	
	Call SetParameter("HOUR", lGmtOffSet, PARAM_STRING)
	if lError = 0 then
	   Dim sSqlTd
	   if (screen=1) then
	       Call SetParameter("P_REPORT_ID", reportId, PARAM_NUMBER)
		   Call SetParameter("P_MESSAGE_ID", messageId, PARAM_NUMBER)
		   sSqlTd =        " Select max(date_of_trans) + (:HOUR/24) from "
		   sSqlTd = sSqlTd & " (Select erd.date_of_trans "
		   sSqlTd = sSqlTd & " From esm_rpt_detail erd, esm_rpt_hist erh "
		   sSqlTd = sSqlTd & " WHERE erd.parent_id = erh.id AND erh.reg_report_id = :P_REPORT_ID AND erd.type = 1 "
		   sSqlTd = sSqlTd & " UNION ALL "
		   sSqlTd = sSqlTd & " Select erd.date_of_trans "
		   sSqlTd = sSqlTd & " From esm_rpt_detail erd, esm_msg_hist emh "
		   sSqlTd = sSqlTd & " WHERE erd.parent_id = emh.id AND emh.msg_id = :P_MESSAGE_ID AND  erd.type = 2 ) "
		   d1 = ExecuteSQLReturnStr(sSqlTd, lError, sError)
		   if lError = 0 then
			  sSqlTd =          " Select min(erd.date_of_trans) + (:HOUR/24) "
			  sSqlTd = sSqlTd & " From esm_rpt_detail erd, esm_rpt_hist erh "
			  sSqlTd = sSqlTd & " WHERE erd.parent_id = erh.id AND erh.reg_report_id = :P_REPORT_ID AND erd.type = 1 "
			  d2 = ExecuteSQLReturnStr(sSqlTd, lError, sError)
		      if len(d1) > 0 and len(d2) > 0 then
		         timediff = DateDiff("n",CDate(d2),CDate(d1))
	         end if
		      hr = timediff\60
		      min = timediff-(hr*60)
		   end if
	   elseif (screen=2) then
	       Call SetParameter("P_MESSAGE_ID", messageId, PARAM_NUMBER)
		   sSqlTd = " Select max(erd.date_of_trans) + (:HOUR/24) "
		   sSqlTd = sSqlTd & " From esm_rpt_detail erd, esm_msg_hist emh "
		   sSqlTd = sSqlTd & " WHERE erd.parent_id = emh.id  AND emh.msg_id = :P_MESSAGE_ID AND erd.type = 2 "
		   d1 = ExecuteSQLReturnStr(sSqlTd, lError, sError)
		   if lError = 0 then
			  sSqlTd = " Select min(erd.date_of_trans) + (:HOUR/24) "
			  sSqlTd = sSqlTd & " From esm_rpt_detail erd, esm_msg_hist emh "
			  sSqlTd = sSqlTd & " WHERE erd.parent_id = emh.id  AND emh.msg_id = :P_MESSAGE_ID AND erd.type = 2 "
			  d2 = ExecuteSQLReturnStr(sSqlTd, lError, sError)
		      if len(d1) > 0 and len(d2) > 0 then
		         timediff = DateDiff("n",CDate(d2),CDate(d1))
	         end if
   			
		      hr = timediff\60
		      min = timediff-(hr*60)
		   end if
	   end if
   	
	   Set oTransmissionList = oOutMessage.selectNodes("/MESSAGE/TABLE_RPT_E2B/RPT_E2B")
   end if
%>
<%
Sub getTransmissionHistory()
	Dim sSql
	Dim sDataFormat
	If glDisplayLang = cfCMN_LANG_EN Then
      sDataFormat = "DD-MON-YYYY"
    Else
      sDataFormat = "YYYY/MM/DD"
    End If
	
	Call SetParameter("HOUR", lGmtOffSet, PARAM_STRING)
	if (screen = 1) then
	    Call SetParameter("P_REPORT_ID", reportId, PARAM_NUMBER)
		Call SetParameter("P_MESSAGE_ID", messageId, PARAM_NUMBER)
		sSql = " Select decode(erd.Date_Of_Trans ,Null,' ',to_char(erd.Date_Of_trans + (:HOUR/24),'" + sDataFormat +" HH12:MI AM')) date_of_trans, "
		sSql = sSql & " erd.Status_text, erd.seq_num "
		sSql = sSql & " From esm_rpt_detail erd, esm_rpt_hist erh "
		sSql = sSql & " WHERE erd.parent_id = erh.id AND erh.reg_report_id = :P_REPORT_ID AND erd.type = 1 "
		sSql = sSql & " UNION ALL "
		sSql = sSql & " Select decode(erd.Date_Of_Trans ,Null,' ',to_char(erd.Date_Of_trans + (:HOUR/24),'" + sDataFormat +" HH12:MI AM')) date_of_trans, "
		sSql = sSql & " erd.Status_text, erd.seq_num "
		sSql = sSql & " From esm_rpt_detail erd, esm_msg_hist emh "
		sSql = sSql & " WHERE erd.parent_id = emh.id AND emh.msg_id = :P_MESSAGE_ID AND  erd.type = 2"
		sSql = sSql & " ORDER BY seq_num DESC"
	elseif (screen = 2) then
	    Call SetParameter("P_MESSAGE_ID", messageId, PARAM_NUMBER)
		sSql = "Select decode(erd.Date_Of_Trans ,Null,' ',to_char(erd.Date_Of_trans + (:HOUR/24),'DD-MON-YYYY HH12:MI AM')) date_of_trans, "
		sSql = sSql & " Status_text "
		sSql = sSql & " From esm_rpt_detail erd, esm_msg_hist emh "
		sSql = sSql & " WHERE erd.parent_id = emh.id  AND emh.msg_id = :P_MESSAGE_ID AND erd.type = 2 "
		sSql = sSql & " ORDER BY erd.seq_num DESC"
	end if
	
	Set oOutMessage = ExecuteSQL(sSql, "25150018,25150097", lError, sError)
	
End Sub
%>
<!-- Page Processing Ends -->
<!-- Page Display Starts -->
<body onload="initForm()">
    <table class="table" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%">
        <tr height="25px" class="tblheader-lightblue">
            <td class="section-header-middle">
                <% BuildLocalLabel("TRANS_HISTORY").SetStyleSheet("label label-section").Render()%>
            </td>
        </tr>
        <tr height="20px">
            <td>
                <div class="table-scroll-hide" style="width: 100%; overflow-y: scroll;">
                    <table class="grd-table-header" cellspacing="0" cellpadding="2">
                        <col width="20%" />
                        <col />
                        <tr class="tblheader-lightblue">
                            <th class="alc-header">
                                <% BuildLocalLabel("DATE").Style("width:100%").Render()%>
                            </th>
                            <th class="alc-header">
                                <% BuildLocalLabel("WL_CMN_STATUS_DETAILS").Style("width:100%").Render()%>
                            </th>
                        </tr>
                    </table>
                </div>
            </td>
        </tr>
        <tr valign="top">
            <td>
                <div id="criteriaList" class="table-scroll" style="height: 100%; overflow-y: scroll">
                    <table class="grd-table-body" cellspacing="0" cellpadding="2">
                        <col width="20%" />
                        <col />
                        <% 
                        lCount = 0
                        For each oTransmission in oTransmissionlist
                        lCount = lCount + 1
                        if lCount Mod 2 = 0 Then
			                sClass = "row-normal"
			            Else
			                sClass = "row-alternate"
			            End If
                        %>
                        <tr class="<%=sClass %>">
                            <td class="alc-header">
                                <% BuildControlDirect(CTL_TEXTBOX_SIMPLE, "text_date", GetXMLValueDirect(oTransmission,"RPT_E2B_EDI_TRANSMIT_DATE"), true, 1, "").Style("width:100%").SetStyleSheet("textbox-list").Render() %>
                            </td>
                            <td class="alc-header">
                                <% BuildControlDirect(CTL_TEXTBOX_SIMPLE, "text_message", GetXMLValueDirect(oTransmission,"RPT_E2B_STATUS_TEXT"), true, 2, "").Style("width:100%").SetStyleSheet("textbox-list").Render() %>
                            </td>
                        </tr>
                        <% Next %>
                    </table>
                </div>
            </td>
        </tr>
        <tr valign="top" height="15px">
            <td>
                <% BuildLabelDirect("Total Time Elapsed - " & hr & " Hour " & min & " Minutes").Render()%>
            </td>
        </tr>
        <tr class="tblheader-gray" height="25px">
            <td align="center">
                <%BuildButton("btn_Print", "PRINT", 1).Style("width:60px").OnClick("fn_PrintReport();").Render() %>
                <%BuildButton("btn_OK", "OK", 2).Style("width:50px").OnClick("fn_Close();").Render() %>
            </td>
        </tr>
    </table>
</body>
<!-- Page Display Ends -->
<!-- Javascript Functions Starts -->

<script type="text/javascript">
    async function initForm() {
        var lError = "<%=lError %>";

        if (lError != "0") {
            await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("RPT_TRANS_HIST")%>', <%=JavaScriptClean(sError) %>);
            window.close();
        }
    }
    function fn_Close() {
        window.close();
    }
   async function fn_PrintReport() {
        var sParam;
        var strURL;
        sParam = "casenum=" + fn_URLEncode(<%=JavaScriptSanitize(casenum) %>) + "&agency=" + <%=JavaScriptSanitize(agency) %> + "&messageId=" + <%=JavaScriptSanitize(messageId) %> + "&screen=" + <%=JavaScriptSanitize(screen) %> + "&reportId=" + <%=JavaScriptSanitize(reportId) %> + "&messageNum=" + <%=JavaScriptSanitize(messageId) %> + "&GmtOffSet=" + <%= JavaScriptSanitize(lGmtOffSet)%>;
        strURL = "/E2B/Worklist/AjaxE2B_PrintTransmissionHistory.asp"
        await loadArgusMessage(strURL, fn_CallBackReport, sParam);
    }
    async function fn_CallBackReport() {
        var xmlDoc = this.req.responseXML;
        var sError = xmlDoc.getElementsByTagName("ERROR_STRING");
        var asDocId = xmlDoc.getElementsByTagName("GN_REPORT_IDENTIFIER");

        if (asDocId && (asDocId.length > 0))
            fn_ViewDocument(GetTextContentFromXML(asDocId[0]), "");
        else
            await MessageBoxRes("GENERAL_ERROR",'<%=GetTranslationData("RPT_TRANS_HIST")%>', GetTextContentFromXML(sError[0])); 
        window.close();
    }
</script>

<!-- Javascript Functions Ends -->
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
