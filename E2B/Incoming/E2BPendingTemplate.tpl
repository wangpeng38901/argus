<tr id="data_row_[[ROW_NUMBER]]" name="data_row_[[ROW_NUMBER]]" class="[[ROW_CLASS]]" 
    onclick="f_e2bPendingStsRpt_SelectRow(this, [[ROW_NUMBER]]);"
    oncontextmenu="fn_Show_Menu(event,[[ROW_NUMBER]]);">
  
  <!-- 1-->


  <td valign="middle" align="center" class="alc-header" style="padding-left:5px;">
    <table width="100%">
      <tr>
        <td style="float: right;" style="width: 50%">
          <input type="checkbox" id="checkbox_[[ROW_NUMBER]]" name="checkbox_[[ROW_NUMBER]]" />
        </td>
        <td style="width: 50%">
          <input id="E2BTpl_REPORT_ID_[[ROW_NUMBER]]" name="E2BTpl_REPORT_ID_[[ROW_NUMBER]]" type="hidden" value="[[RPT_E2B_REPORT_ID]]" />
		  <input id="E2BTpl_ENCRYPT_REPORT_ID_[[ROW_NUMBER]]" name="E2BTpl_ENCRYPT_REPORT_ID_[[ROW_NUMBER]]" type="hidden" value="{{RPT_E2B_REPORT_ID}}" /> 
          <input id="E2BTpl_HL7P_[[ROW_NUMBER]]" name="E2BTpl_HL7P_[[ROW_NUMBER]]" type="hidden" value="[[CFG_PROFILE_hl7_profile]]" />
          <input id="E2BTpl_APPLY_NEW_FW_[[ROW_NUMBER]]" name="E2BTpl_APPLY_NEW_FW_[[ROW_NUMBER]]" type="hidden" value="[[CFG_PROFILE_hl7_profile]]" />                    
        </td>
      </tr>
    </table>
  </td>
    
  <!-- 2-->
  
  <td valign="top" class="alc-header" style="padding-left:5px;">        
	<input id="E2BTpl_TRADING_PARTNER_[[ROW_NUMBER]]" name="E2BTpl_TRADING_PARTNER_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_AGENCY_NAME]]"  readonly />
	<input id="E2BTpl_REPORTTYPEIFN_[[ROW_NUMBER]]" name="E2BTpl_REPORTTYPEIFN_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_REPORT_TYPE]]"  readonly />
	<input id="E2BTpl_WORLDWIDEUNIQUENO_[[ROW_NUMBER]]" name="E2BTpl_WORLDWIDEUNIQUENO_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_COMPANYNUMB]]"  readonly />
	<input type="hidden" id="E2BTpl_ENCRYPT_WORLDWIDEUNIQUENO_[[ROW_NUMBER]]" name="E2BTpl_ENCRYPT_WORLDWIDEUNIQUENO_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="{{RPT_E2B_COMPANYNUMB}}"  readonly />
	<input id="E2BTpl_SENDERCASENO_[[ROW_NUMBER]]" name="E2BTpl_SENDERCASENO_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_OTHERNUMB]]"  readonly />
  </td>
  
  <!-- 3-->
  
  <td valign="top" class="alc-header" style="padding-left:5px;">
    <input id="E2BTpl_TRANSMISSIONSENTSENT_[[ROW_NUMBER]]" name="E2BTpl_TRANSMISSIONSENTSENT_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_TRANSMIT_DATE]]"  readonly />
    <input id="E2BTpl_EDI_TRANSMIT_DATE_[[ROW_NUMBER]]" name="E2BTpl_EDI_TRANSMIT_DATE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_EDI_TRANSMIT_DATE]]"  readonly />
    <input id="E2BTpl_INTERCHANGE_PROCESSED_DATE_[[ROW_NUMBER]]" name="E2BTpl_INTERCHANGE_PROCESSED_DATE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" 
		value="[[RPT_E2B_MESSAGE_DATE]]"  readonly />
  </td>
  
  <!-- 4-->
  
  <td valign="top" class="alc-header" style="padding-left:5px;">
    
          <input id="E2BTpl_CASE_RECEIPT_DATE_[[ROW_NUMBER]]" name="E2BTpl_CASE_RECEIPT_DATE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" 
			value="[[RPT_E2B_ACK_MESSAGE_DATE]]"  readonly />
          <input id="E2BTpl_COUNTRY_=""[[ROW_NUMBER]]" name="E2BTpl_COUNTRY_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_COUNTRY]]"  readonly />
          <input id="E2BTpl_REPORT_TYPE_[[ROW_NUMBER]]" name="E2BTpl_REPORT_TYPE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_REPORT_ACK_CODE]]"  readonly />
          <input type = "hidden" id="E2BTpl_IMPORTED_CASE_NUM_[[ROW_NUMBER]]" name="E2BTpl_IMPORTED_CASE_NUM_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" 
			value="[[RPT_E2B_CASE_NUM]]"  readonly/>
          <A href="#" onclick="fn_QuickyOpen([{RPT_E2B_CASE_NUM}]);return false;">
                [[RPT_E2B_CASE_NUM]]
          </A>  
       
  </td>
  
  <!-- 5-->
  
  <td valign="top" align="left" class="alc-header" style="padding-left:5px;">
          <input id="E2BTpl_PRODUCT_NAME_[[ROW_NUMBER]]" name="E2BTpl_PRODUCT_NAME_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_MEDICINAL_PRODUCT]]"  readonly />
          <input id="E2BTpl_GENERIC_NAME_[[ROW_NUMBER]]" name="E2BTpl_GENERIC_NAME_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_ACTIVESUBSTANCE_NAME]]"  readonly />
          <input id="E2BTpl_not_used1_[[ROW_NUMBER]]" name="E2BTpl_not_used1_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="" readonly />
          <input id="E2BTpl_WORKFLOW_STATE_NAME_[[ROW_NUMBER]]" name="E2BTpl_WORKFLOW_STATE_NAME_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" 
			value="[[CFG_WORKFLOW_STATES_STATE_NAME]]"  readonly />
  </td>
  
  <!-- 6-->
  
  <td valign="top" align="left" class="alc-header" style="padding-left:5px;">
    
          <input id="E2BTpl_EVENT_PT_[[ROW_NUMBER]]" name="E2BTpl_EVENT_PT_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_REPORTER_TITLE]]"  readonly />
          <input id="E2BTpl_EVENT_LLT_[[ROW_NUMBER]]" name="E2BTpl_EVENT_LLT_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[WHO_DRUG_DICT_SALT_CODE]]"  readonly />
          <input id="E2BTpl_not_used2_[[ROW_NUMBER]]" name="E2BTpl_not_used2_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="" readonly />
          <input id="E2BTpl_SITES_DESC_[[ROW_NUMBER]]" name="E2BTpl_SITES_DESC_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" 
			value="[[LM_SITES_SITE_DESC]]"  readonly />
	</td>
  
  <!-- 7-->
  
  <td valign="top" align="left" class="alc-header" style="padding-left:5px;">

          <input id="E2BTpl_PAT_INI_[[ROW_NUMBER]]" name="E2BTpl_PAT_INI_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_PATIENT_INITIAL]]"  readonly />
          <input id="E2BTpl_STYDYID_PATID_[[ROW_NUMBER]]" name="E2BTpl_STYDYID_PATID_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_SPONSOR_STUDY_NUMB]]"  readonly />
          <input id="E2BTpl_REPORTER_TYPE_[[ROW_NUMBER]]" name="E2BTpl_REPORTER_TYPE_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_REPORTER_POSTCODE]]"  readonly />
          <input id="E2BTpl_REPORTER_[[ROW_NUMBER]]" name="E2BTpl_REPORTER_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;" value="[[RPT_E2B_REPORTER_GIVENAME]]"  readonly />
  </td>
  
  <input id="E2BTpl_REPORTTYPEIFN_ID_[[ROW_NUMBER]]" name="E2BTpl_REPORTTYPEIFN_ID_[[ROW_NUMBER]]"
		type="hidden" value="[[RPT_E2B_REPORTER_FAMILY_NAME]]" />
  <input id="HDNERROR_[[ROW_NUMBER]]" name="HDNERROR_[[ROW_NUMBER]]"
		type="hidden" value="[[RPT_E2B_VIEW_WARNINGS]]" />
  <input id="CASE_STATE_[[ROW_NUMBER]]" name="CASE_STATE_[[ROW_NUMBER]]"
		type="hidden" value="[[CSM_STATE_ID]]" />
  <input id="CASE_CLOSED_STATE_[[ROW_NUMBER]]" name="CASE_CLOSED_STATE_[[ROW_NUMBER]]"
		type="hidden" value="[[GN_NUMBER1]]" />
  <input id="CASE_OPENED_USER_[[ROW_NUMBER]]" name="CASE_OPENED_USER_[[ROW_NUMBER]]"
		type="hidden" value="[[CFG_USERS_USER_FULLNAME]]" />
  <input id="CASE_CURRENT_STATE_[[ROW_NUMBER]]" name="CASE_CURRENT_STATE_[[ROW_NUMBER]]"
		type="hidden" value="[[CSM_LAST_STATE_ID]]" />
  <input id="E2BTpl_SITES_ID_[[ROW_NUMBER]]" name="E2BTpl_SITES_ID_[[ROW_NUMBER]]"
		type="hidden" value="[[LM_SITES_SITE_ID]]" />
  <input id="E2BTpl_WORKFLOW_STATE_ID_[[ROW_NUMBER]]" name="E2BTpl_WORKFLOW_STATE_ID_[[ROW_NUMBER]]"
		type="hidden" value="[[CFG_WORKFLOW_STATES_STATE_ID]]" />
  <input id="CASE_ID_[[ROW_NUMBER]]" name="CASE_ID_[[ROW_NUMBER]]"
		type="hidden" value="[[CSM_CASE_ID]]" />
  <input id="RPT_E2B_AUTHORITY_ID_[[ROW_NUMBER]]" name="RPT_E2B_AUTHORITY_ID_[[ROW_NUMBER]]"
		type="hidden" value="[[RPT_E2B_AUTHORITY_ID]]" />                 
</tr>
