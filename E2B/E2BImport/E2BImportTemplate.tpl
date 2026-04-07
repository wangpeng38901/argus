<tr id="data_row_[[ROW_NUMBER]]" name="data_row_[[ROW_NUMBER]]" class="[[ROW_CLASS]]"	onmousedown="fn_highlightCurrentRow(this);">
  <!--  1 -->
  <td valign="top" class="alc-header" style="padding-left:5px;">

    <A href="#" onclick="openE2BViewer('[[RPT_E2B_REPORT_ID]]', '{{RPT_E2B_REPORT_ID}}','{{RPT_E2B_COMPANYNUMB}}');return false;">
      [[RPT_E2B_SAFETYREPORTID]]
    </A>
    <input id="E2BTpl_REPORT_TYPE_[[ROW_NUMBER]]" name="E2BTpl_REPORT_TYPE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_REPORT_TYPE]]"  readonly />
  </td>

  <!-- 2 -->
  <td valign="top" class="alc-header" style="padding-left:5px;">
    <input id="E2BTpl_RAGENCY_NAME_[[ROW_NUMBER]]" name="E2BTpl_RAGENCY_NAME_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_AGENCY_NAME]]"  readonly />
    <input id="E2BTpl_COMPANYNUMB_[[ROW_NUMBER]]" name="E2BTpl_COMPANYNUMB_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_COMPANYNUMB]]"  readonly />
  </td>
  <!-- 3-->

  <td valign="top" class="alc-header" style="padding-left:5px;">
    <TABLE cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td width="80%">
          <input id="E2BTpl_STATUS_[[ROW_NUMBER]]" name="E2BTpl_STATUS_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_STATUS_TEXT]] / [[RPT_E2B_ACK_TYPE]]"  readonly />
          <A href="#" onclick="fn_QuickyOpen([{RPT_E2B_CASE_NUM}]);return false;" style="font-size: 8pt;">
      			 [[RPT_E2B_CASE_NUM]]
    			</A>
        </td>
        <td style="float: right;" valign="center">
          <img src="[[EICON]]" id="ExTpl_lockunlock_[[ROW_NUMBER]]" value="[[EICON]]" 
              onclick="ShowError('[[ROW_NUMBER]]');" style="cursor:pointer" />
        </td>
      </tr>
    </TABLE>
  </td>
  <!-- 4-->
  <td valign="top" align="left" class="alc-header" style="padding-left:5px;">
    <input id="E2BTpl_UNM_[[ROW_NUMBER]]" name="E2BTpl_UNM_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[CFG_USERS_USER_FULLNAME]]"  readonly />
    <input id="E2BTpl_NOTES_[[ROW_NUMBER]]" name="E2BTpl_NOTES_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_ERROR_MESSAGE_COMMENT]]"  readonly />
  </td>
  <!--5-->
  <td valign="top" align="left" class="alc-header" style="padding-left:5px;">
		<input id="E2BTpl_dATE_[[ROW_NUMBER]]" name="E2BTpl_dATE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_MESSAGE_DATE]]"  readonly />

    <input id="E2BTpl_dATE_[[ROW_NUMBER]]" name="E2BTpl_dATE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_IMPORT_DATE_CHAR]]"  readonly />
  </td>

  <td valign="middle" align="left" class="alc-header" style="padding-left:5px;">
    <table width ="100%" cellspacing="0" cellpadding="0" border="0">
      <tr style="width:100%">
        <td style="padding: px 0px 0px 0px" style="width:50%">
          <table cellspacing="0" cellpadding="0" border="0"  width="100%">
            <tr>
              <td valign="middle" width="30%">
              </td>
              <td valign="middle" style="width: 5px">
                <img src="[[ICON_1]]" />
              </td>
              <td valign="middle" width="70%">
                <hr />
              </td>
            </tr>
          </table>

        </td>
        <td style="padding: 0px 0px 0px 0px"  style="width:50%">
          <table cellspacing="0" cellpadding="0" border="0" width="100%">
            <tr>
              <td valign="middle" width="70%">
                <hr />
              </td>
              <td valign="middle" style="width: 5px">
                <img src="[[ICON_2]]" />
              </td>
              <td valign="middle" width="30%">
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </td>

  <input id="HDNERROR_[[ROW_NUMBER]]" name="HDNERROR_[[ROW_NUMBER]]"
		type="hidden" value="[[RPT_E2B_VIEW_WARNINGS]]" />

</tr>
