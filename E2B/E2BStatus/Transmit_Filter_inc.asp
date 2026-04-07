<!-- #INCLUDE VIRTUAL="/Include/CheckLoginLite_inc.asp" -->
<table class="table" cellspacing="0" cellpadding="0" width="100%" id="Filter_Table">
    <!-- Section Header -Filter Header Starts -->
    <tr>
        <td>
            <table class="table" cellspacing="0" cellpadding="0" width="100%" border="0" id="Filter_Header_Table">
                <tr style="height: 25px">
                    <td class="section-header-left"></td>
                    <td class="section-header-middle">
                        <span id="lblFilterSec" class="label label-section">Search Reports</span>
                    </td>
                    <td class="section-header-middle" align="right" valign="middle"></td>
                    <td class="section-header-right"></td>
                </tr>
            </table>
        </td>
    </tr>
    <!-- Section Header -Filter Header Ends -->
    <!-- Section Contents -Filter Information Starts -->
    <tr>
        <td>
            <div id="Filter_Content_Div">
                <table class="table border-blue inner-table" width="100%">
                    <tr>
                        <td class="no-padding-left">
                            <table class="table inner-table" cellspacing="2" cellpadding="2" width="100%">
                                <col width="20%" />
                                <col width="15%" />
                                <col width="5%" />
                                <col width="15%" />
                                <col width="5%" />
                                <col width="20%" />
                                <col width="20%" />
                                <tr>
                                    <td>
                                        <%BuildLabelDirect("Agency/Trading Partner").Render()%>
                                    </td>
                                    <td colspan="5">
                                        <%
                                        BuildListDirectFromXML("AgencyName", Agency, false, 1, oContactList, "LM_REGULATORY_CONTACT_AGENCY_ID", "LM_REGULATORY_CONTACT_AGENCY_NAME")_
                                            .AddTopOption("ALL:-1").Style("width:100%;").OnChange("fn_ChangeAgency(this.value);").Render()
                                        %>                                    
                                    </td>
                                    <td></td>
                                </tr>
                                <tr>
                                    <td>
                                        <%BuildControlDirect(CTL_RADIOBUTTON, "TxRadio", 0, false, 2, ":0" ).onClick("fn_RadioChange()").Render()%>
                                        <%BuildLabelDirect("Transmit Date Range From").SetStyleSheet("label label-no-padding").Render()%>
                                    </td>
                                    <td>
                                        <%BuildControlDirect(CTL_TEXTBOX_DATE, "DateFrom", DateFrom, false, 3, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                                    </td>
                                    <td>
                                        <%BuildLabelDirect("To ").Render() %>
                                    </td>
                                    <td>
                                        <%BuildControlDirect(CTL_TEXTBOX_DATE, "DateTo", DateTo, false, 4, "").onchange("fn_CustomizedDate()").Style("width:100%").Render()%>
                                    </td>
                                    <td>
                                        <%BuildLabelDirect("Range ").Render() %>
                                    </td>
                                    <td>
                                        <%BuildListDirectFromXML("RangeList",RangeList, false, 5, oDateRangeList, "LM_DATE_RANGES_RANGE_ID", "LM_DATE_RANGES_RANGE_NAME").AddBtmOption("Custom Date Range:-1").onchange("fn_ChangeDateRange(this.value)").Style("width:100%").Setstylesheet("ddlist").Render() %>
                                    </td>
                                    <td>&nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <%BuildControlDirect(CTL_RADIOBUTTON, "TxRadio", IIf(RadioBtn = 1, 1, 0), false, 6, ":1").onClick("fn_RadioChange()").Render()%>
                                        <%BuildLabelDirect("Message # Range From").SetStyleSheet("label label-no-padding").Render()%>
                                    </td>
                                    <td>
                                        <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "MessageFrom", MessageFrom, false, 7, "").SetMaxLength(100).Style("width:100%").Render()%>
                                    </td>
                                    <td>
                                        <%BuildLabelDirect("To").Render() %>
                                    </td>
                                    <td>
                                        <%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "MessageTo", MessageTo, false, 8, "").SetMaxLength(100).Style("width:100%").Render()%>
                                    </td>
                                    <td>
                                        <%BuildLabelDirect("Type").Render() %>
                                    </td>
                                    <td>
                                        <select class="ddlist" style="width: 100%;" name="MessageType" onchange="fn_ChangeTypeRange()"
                                            tabindex="9">
                                            <option value="(Any)" <%if TypeList = "(Any)" Then%>selected<%end if%>>(Any)</option>
                                            <option value="MSG" <%if TypeList = "MSG" Then%>selected<%end if%>>MSG</option>
                                            <option value="ACK" <%if TypeList = "ACK" Then%>selected<%end if%>>ACK</option>
                                            <option value="ATT" <%if TypeList = "ATT" Then%>selected<%end if%>>ATT</option>
                                        </select>
                                    </td>
                                    <td>
                                        <%BuildButtonDirect("Btn_OK", "Search", 10).Style("width:60px")_
                                            .OnClick("fn_GetData();").Render()%>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </div>
        </td>
    </tr>
    <!-- Section Contents -Filter Information Ends -->
</table>

<script type="text/javascript">
    function fn_ChangeDateRange(lValue)
    {
	    sStart = document.all.DateFrom.value;
	    sStop  = document.all.DateTo.value;
    	
	    switch (lValue)
	    {
    <%	
	    Set oDateRangeList = oDateRangeMsg.selectNodes ("/MESSAGE/TABLE_LM_DATE_RANGES/LM_DATE_RANGES")
	    For Each oDateRange in oDateRangeList
		    lID = GetXMLValueDirect(oDateRange, "LM_DATE_RANGES_RANGE_ID")
		    sStart = fn_date_from_iso(GetXMLValueDirect(oDateRange, "LM_DATE_RANGES_FROM_DATE"), 8, false)
		    sStop  = fn_date_from_iso(GetXMLValueDirect(oDateRange, "LM_DATE_RANGES_TO_DATE"), 8, false)
    %>		case "<%=lID%>": sStart = "<%=sStart%>"; sStop = "<%=sStop%>"; break;
    <%	Next%>
	    }

	    document.all.DateFrom.value = sStart;
	    document.all.DateTo.value = sStop;
	    document.all.selected_row.value = lValue;
    }

    function fn_CustomizedDate()
    {
	    document.all.selected_row.value = -1; 	//Custom dates
	    document.all.RangeList.value = -1;
    }

    function fn_ChangeAgency(lValue)
    {
	    document.all.Agency.value = lValue;
    }

    function fn_ChangeTypeRange()
    {
	    document.Frm_E2B.TypeList.value = document.Frm_E2B.MessageType.value;
    }

    function fn_GetData()
    {
        showLoading();
        document.Frm_E2B.PrevSortField.value = 1;
	    document.Frm_E2B.CurrentSortField.value = 1;
	    document.Frm_E2B.SortOrder.value = 1;
        document.Frm_E2B.SearchClicked.value = 2;
        document.Frm_E2B.Frm_E2B_CurrentPage.value = 1;
        document.Frm_E2B.action = "E2bTransmitStatus.asp?bSearch=1" + "&RadioBtn=" + document.Frm_E2B.RadioBtn.value + "&DateFrom=" + document.Frm_E2B.DateFrom.value + "&DateTo=" + document.Frm_E2B.DateTo.value + "&rangelist=" + document.Frm_E2B.RangeList.value + "&typelist=" + document.Frm_E2B.TypeList.value + "&MessageFrom=" + fn_URLEncode(document.Frm_E2B.MessageFrom.value) + "&MessageTo=" + fn_URLEncode(document.Frm_E2B.MessageTo.value) + "&OffSet=" + document.Frm_E2B.OffSet.value + "&GmtOffSet=" + document.Frm_E2B.GmtOffSet.value + "&<%=GetRequestKeyValue()%>";
        fn_ValidateSubmitForm(document.Frm_E2B);
    }

    function fn_RadioChange()
    {
    	
	    for(var i = 0; i < document.Frm_E2B.TxRadio.length; i++)
	    {
		    if(document.Frm_E2B.TxRadio[i].checked)
		    {
			    document.Frm_E2B.RadioBtn.value = document.Frm_E2B.TxRadio[i].value;
			    break;
		    }
	    }

	    if (document.Frm_E2B.RadioBtn.value == "-1")
		    document.Frm_E2B.RadioBtn.value = 0;

	    if (document.Frm_E2B.RadioBtn.value == 0)
	    {
	        document.Frm_E2B.TxRadio[0].checked = true;
	        document.Frm_E2B.TxRadio[1].checked = false;
	    }
	    else if (document.Frm_E2B.RadioBtn.value == 1)
	    {
	        document.Frm_E2B.TxRadio[0].checked = false;
	        document.Frm_E2B.TxRadio[1].checked = true;
	    }
	    else
	    {
	        document.Frm_E2B.TxRadio[0].checked = false;
	        document.Frm_E2B.TxRadio[1].checked = false;
	    }
    	
	    if (document.Frm_E2B.TxRadio[0].checked)
	     {
		    document.Frm_E2B.DateFrom.disabled=false;
		    document.Frm_E2B.DateTo.disabled=false;
		    document.Frm_E2B.MessageFrom.disabled=true;
		    document.Frm_E2B.MessageTo.disabled=true;
		    document.Frm_E2B.RangeList.disabled=false;
	     }
	    else if (document.Frm_E2B.TxRadio[1].checked)
	    {
		    document.Frm_E2B.DateFrom.disabled=true;
		    document.Frm_E2B.DateTo.disabled=true;
		    document.Frm_E2B.MessageFrom.disabled=false;
		    document.Frm_E2B.MessageTo.disabled=false;
		    document.Frm_E2B.RangeList.disabled=true;
	    }
	    else
	    {
		    document.Frm_E2B.DateFrom.disabled=true;
		    document.Frm_E2B.DateTo.disabled=true;
		    document.Frm_E2B.MessageFrom.disabled=true;
		    document.Frm_E2B.MessageTo.disabled=true;
		    document.Frm_E2B.RangeList.disabled=true;
	    }
	    fn_ChangeDateRange(document.Frm_E2B.RangeList.value);
    }

    async function fn_OpenAckStatus(temp, msgType)
    {
	    var strURL;		
	    if (msgType == "ATT")
	    {	    
            strURL = "/E2B/E2bStatus/AjaxGetE2BReportAttachment.asp?Attachment_Id=" + temp;        
	        await loadArgusMessage(strURL, fn_view_attachment);	    
	    }
	    else
	    {			
	        strURL = "/E2B/E2bStatus/MsgAckStatus.asp?AckType=" + "TRANSMIT" + "&MessageNumb=" + temp + "&GmtOffSet=" + document.Frm_E2B.GmtOffSet.value;
            var sDialogStyle = { dialogHeight: "480", dialogWidth: "720", resizable: false, scrollable: false };    
            await fn_OpenModalDialog(strURL, window, sDialogStyle);
	    }
    }

    async function fn_view_attachment()
    {
        var strURL;        
        var xmlDoc = this.req.responseXML;    
        var sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
        if (sErrStr.length > 0)
        {
            await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("E2B_ATTACH")%>', sErrStr);
            return;
        }
        var asDocId = xmlDoc.getElementsByTagName("GN_REPORT_IDENTIFIER");        
        if (asDocId && (asDocId.length > 0))
            fn_ViewDocument(GetTextContentFromXML(asDocId[0]), 0);
    }

    async function fn_PrintList()
    {
        var strURL;
        var lCurrentSortField, lPrevSortField, sSortOrder;
        lCurrentSortField = document.Frm_E2B.CurrentSortField.value
        lPrevSortField = document.Frm_E2B.PrevSortField.value
        sSortOrder = document.Frm_E2B.SortOrder.value
	    document.Frm_E2B.Agency.value = document.Frm_E2B.AgencyName.value;
        strURL = "/E2B/E2bStatus/E2bTransmitStatusPrint.asp?Agency=" + document.Frm_E2B.Agency.value + "&RadioBtn=" + document.Frm_E2B.RadioBtn.value + "&DateFrom=" + document.Frm_E2B.DateFrom.value + "&DateTo=" + document.Frm_E2B.DateTo.value + "&typelist=" + document.Frm_E2B.MessageType.value + "&rangelist=" + document.Frm_E2B.RangeList.value + "&MessageFrom=" + fn_URLEncode(document.Frm_E2B.MessageFrom.value) + "&MessageTo=" + fn_URLEncode(document.Frm_E2B.MessageTo.value) + "&OffSet=" + document.Frm_E2B.OffSet.value + "&GmtOffSet=" + document.Frm_E2B.GmtOffSet.value + "&CurrentSortField=" + lCurrentSortField + "&PrevSortField=" + lPrevSortField + "&SortOrder=" + sSortOrder;
        strURL += "&cacheId=" + fn_URLEncode(document.Frm_E2B.UserCacheId.value);
        var sDialogStyle = { dialogHeight: "480", dialogWidth: "720", resizable: false, scrollable: false };    
        await fn_OpenModalDialog(strURL, window, sDialogStyle);
    }
    
    function init_Form()
    {
		var testdate = new Date();
		document.getElementById("GmtOffSet").value = (-testdate.getTimezoneOffset()/60);
		document.getElementById("OffSet").value = testdate.getTimezoneOffset()/60;

        hideLoading();
        fn_RadioChange();
        if (document.Frm_E2B.RadioBtn.value == 0)
        {
            document.Frm_E2B.TxRadio[0].checked = true;
	        document.Frm_E2B.TxRadio[1].checked = false;
        }
        else
        {
            document.Frm_E2B.TxRadio[0].checked = false;
	        document.Frm_E2B.TxRadio[1].checked = true;
        }
    }
</script>
