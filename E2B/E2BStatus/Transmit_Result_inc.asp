<!-- #INCLUDE VIRTUAL="/Include/CheckLoginLite_inc.asp" -->
<!-- Grid Control Starts-->
<tr height="25px">
    <td>
        <% 
        gSecHead_sGridName = sFormName
        gSecHead_lGridRowCount = Rpt_lResultsRowCount
        %>
        <!-- #INCLUDE VIRTUAL="/Common/SectionHeaderGrid_inc.asp" -->
    </td>
</tr>
<!-- Grid Control Ends -->
<!-- Section Column Headers - Open Cases Search Results Starts -->
<tr>
    <td class="no-padding-left">
        <table class="table" cellspacing="0" cellpadding="0" style="width: 100%; height: 100%;">
            <thead>
                <tr style="height: 36px;">
                    <td>
                        <div id="Div1" class="scroll_hide" style="width: 100%; overflow-y: scroll;">
                            <table class="table inner-table" cellspacing="0" cellpadding="2" style="width: 100%;">
                                <col width="7%" />
                                <col width="20%" />
                                <col width="15%" />
                                <col width="20%" />
                                <col width="23%" />
                                <col />
                                <tr class="tblheader-lightblue" style="height: 36px;">
                                    <th class="grd-header">
                                        <table cellpadding='0' cellspacing='0' width="100%">
                                            <tr style="cursor: pointer">
                                                <td width="100%" onclick="fn_Sort(1, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Type").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 1) Then response.Write(trim(sArrowStr))  %>
                                                </td>
                                            </tr>
                                            <tr style="cursor: pointer">
                                                <td width="100%" onclick="fn_Sort(2, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Reports").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 2) Then response.Write(trim(sArrowStr))  %>
                                                </td>
                                            </tr>
                                        </table>
                                    </th>
                                    <th class="grd-header">
                                        <table cellpadding='0' cellspacing='0' width="100%">
                                            <tr style="cursor: pointer">
                                                <td width="100%" onclick="fn_Sort(3, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Trading Partner").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 3) Then response.Write(trim(sArrowStr))  %>
                                                </td>
                                            </tr>
                                            <tr style="cursor: pointer">
                                                <td width="100%" onclick="fn_Sort(12, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Control #").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 12) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                        </table>
                                    </th>
                                    <th style="cursor: pointer" class="grd-header">
                                        <table cellpadding='0' cellspacing='0' width="100%">
                                            <tr width="100%">
                                                <td width="100%" onclick="fn_Sort(4, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Local Msg #").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 4) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td onclick="fn_Sort(5, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Remote Msg #").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 5) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                        </table>
                                    </th>
                                    <th style="cursor: pointer" class="grd-header">
                                        <table cellpadding='0' cellspacing='0' width="100%">
                                            <tr>
                                                <td onclick="fn_Sort(6, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("File Name").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 6) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td onclick="fn_Sort(7, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Transmit to EDI").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 7) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                        </table>
                                    </th>
                                    <th style="cursor: pointer" class="grd-header">
                                        <table cellpadding='0' cellspacing='0' width="100%">
                                            <tr>
                                                <td onclick="fn_Sort(8, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("EDI Tracking ID").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 8) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td onclick="fn_Sort(9, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("EDI Transmit Date").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 9) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                        </table>
                                    </th>
                                    <th style="cursor: pointer" class="grd-header">
                                        <table cellpadding='0' cellspacing='0' width="100%">
                                            <tr>
                                                <td onclick="fn_Sort(10, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("Transmission Status").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 10) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td onclick="fn_Sort(11, <%= lNextSortOrder %>);">
                                                    <%BuildLabelDirect("EDI Receive Receipt").SetStyleSheet("label label-no-padding").Render()%>
                                                    <%If (lCurrentSort = 11) Then response.Write(sArrowStr) %>
                                                </td>
                                            </tr>
                                        </table>
                                    </th>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
            </thead>
            <tbody valign="top">
                <tr>
                    <td class="no-padding-left">
                        <div id="Div2" class="table-scroll" style="height: 100%; overflow-y: scroll;">
                            <table id="" class="table inner-table" style="width: 100%; table-layout: fixed" cellpadding="2" cellspacing="0">
                                <col width="7%" />
                                <col width="20%" />
                                <col width="15%" />
                                <col width="20%" />
                                <col width="23%" />
                                <col width="15%" />
                                <%        
                                If (gSecHead_lGridRowCount <= 0) Then
                                    If Rpt_bPostBack Then
                                        Response.Write("<tr><td class='label' align='center' colspan='6'>No report found<td></tr>")
                                    Else
                                        Response.Write("<tr><td class='label' align='center' colspan='6'>Enter Search Criteria<td></tr>")
                                    End If
                                Else
                                    Response.Write(Rpt_oGridControl.Render(1, Rpt_lPageRowCount))
                                End If
                                %>
                            </table>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
    </td>
</tr>
<!-- Section Column Headers - Open Cases Search Results Starts -->
