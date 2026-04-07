<tr id="data_row_[[ROW_NUMBER]]" name="data_row_[[ROW_NUMBER]]" class="[[ROW_CLASS]]" onclick="f_SelectRow([[ROW_NUMBER]]);"  oncontextmenu="fn_Show_Menu(event, [[ROW_NUMBER]]);">
  <td>
    <input type="hidden" id="case_id_[[ROW_NUMBER]]" value="[[CSM_CASE_ID]]" />
    <input type="hidden" id="case_num_[[ROW_NUMBER]]" value="[[CSM_CASE_NUM]]" />
    <input type="hidden" id="lock_status_[[ROW_NUMBER]]" value="[[CSM_STATE_ID]]" />
    <input type="hidden" id="delete_status_[[ROW_NUMBER]]" value="[[CSM_LAST_STATE_ID]]" />
    <input type="hidden" id="close_status_[[ROW_NUMBER]]" value="[[GN_NUMBER1]]" />
    <input type="hidden" id="opened_user_[[ROW_NUMBER]]" value="[[CFG_USERS_USER_FULLNAME]]" />
    <A href="#" onclick="CloseSearchCaseList([{CSM_CASE_NUM}]);">[[CSM_CASE_NUM]]</A>
    <input id="pat_initials_[[ROW_NUMBER]]" name="pat_initials_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[CSPI_PAT_INITIALS]]" readonly />
  </td>
  <td>
    <input id="protocol_num_[[ROW_NUMBER]]" name="protocol_num_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[CSS_PROTOCOL_NUM]]" readonly />
    <br />
    <input id="study_num_[[ROW_NUMBER]]" name="study_num_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[CSS_STUDY_NUM]]" readonly />
  </td>
  <td>
    <input id="init_date_[[ROW_NUMBER]]" name="init_date_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[CSM_INIT_REPT_DATE]]" readonly />
    <br />
    <input id="rpt_country_[[ROW_NUMBER]]" name="rpt_country_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[LM_COUNTRIES_COUNTRY]]" readonly />
  </td>
  <td>
    <input id="prod_name_[[ROW_NUMBER]]" name="prod_name_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[CSPD_PRODUCT_NAME]]" readonly />
    <br />
    <input id="desc_coded_[[ROW_NUMBER]]" name="desc_coded_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[CSE_DESC_CODED]]" readonly />
	</td>
  <td>
    <input id="rpt_type_[[ROW_NUMBER]]" name="rpt_type_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[LM_REPORT_TYPE_REPORT_TYPE]]" readonly />
    <br />
    <input id="rep_first_name_[[ROW_NUMBER]]" name="rep_first_name_[[ROW_NUMBER]]" style="width:100%" class="[[ROW_COLOR]]" value="[[CSRP_FIRST_NAME]]" readonly />
  </td>
</tr>