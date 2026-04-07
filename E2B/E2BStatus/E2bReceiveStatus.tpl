<tr id="data_row_[[ROW_NUMBER]]" name="data_row_[[ROW_NUMBER]]" class='[[ROW_CLASS]]' >
    <td valign="middle" align="center" class="alc-header">
    <table cellpadding="0" cellspacing="0" width="30px">
        <tr class='[[ROW_CLASS]]' width='100%'>
          <td valign="middle" align="left" width='100%'>
            <input id="bbfTpl_state_[[ROW_NUMBER]]" name="bbfTpl_STATE_[[ROW_NUMBER]]" class="textbox-worklist" style="
                  background-color: [[ROW_COLOR]];" value="[[RPT_E2B_MESSAGE_TYPE]]"  readonly />
          </td>
        </tr>
        <tr class='[[ROW_CLASS]]' >
          <td valign="middle" align="left">
            <a href="#" onclick="fn_OpenAckStatus('[[RPT_E2B_LMESSAGE_ID]]');">
            <img src="/img/Reports/ReportDetails.gif"  id="Img_[[ROW_NUMBER]]" style="border:none;"  value="[[RPT_E2B_LMESSAGENUMB]]"
                  alt="[[RPT_E2B_LMESSAGENUMB]]" title="[[RPT_E2B_LMESSAGENUMB]]" />
            </a>
          </td>
        </tr>
        
      </table>
    
    </td>
    <!-- 2 -->
    <td valign="middle" class="alc-header">
            <input id="bbfTpl_state_[[ROW_NUMBER]]" name="Tpl_Agency_[[ROW_NUMBER]]" class="textbox-worklist" style="width:100%;
                  background-color: [[ROW_COLOR]];" value="[[RPT_E2B_AGENCY_NAME]]"  readonly />            
            <input id="bbfTpl_state_[[ROW_NUMBER]]" name="bbfTpl_STATE_[[ROW_NUMBER]]" class="textbox-worklist" style="width:100%;
                  background-color: [[ROW_COLOR]];" value="[[RPT_E2B_CONTROL_NO]]"  readonly />
    </td>
    <!-- 3 -->

    <td valign="middle" align="left" class="alc-header">
      <input id="bbfTpl_product_[[ROW_NUMBER]]" name="bbfTpl_product_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;
                background-color: [[ROW_COLOR]];" value="[[RPT_E2B_LMESSAGENUMB]]" readonly />
      <input id="bbfTpl_Diagnosis_[[ROW_NUMBER]]" name="bbfTpl_Diagnosis_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;
            background-color: [[ROW_COLOR]]; " value="[[RPT_E2B_RMESSAGENUMB]]" readonly />
      <!--
      <table cellpadding="0" cellspacing="0" width='100%'>
        <tr class='[[ROW_CLASS]]' width='100%'>
          <td valign="middle" align="left" width='100%'>
            <input id="bbfTpl_product_[[ROW_NUMBER]]" name="bbfTpl_product_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 200;
                background-color: [[ROW_COLOR]];" value="[[RPT_E2B_LMESSAGENUMB]]" readonly />
          </td>
        </tr>
        <tr class='[[ROW_CLASS]]' width='100%'>
          <td valign="middle" align="left" width='100%'>
            <input id="bbfTpl_Diagnosis_[[ROW_NUMBER]]" name="bbfTpl_Diagnosis_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 200;
            background-color: [[ROW_COLOR]]; " value="[[RPT_E2B_RMESSAGENUMB]]" readonly />
          </td>
        </tr>
      </table>
      -->
    </td>
    <!-- 4 -->
    <td valign="middle" align="left" class="alc-header">
      <table cellpadding="0" cellspacing="0">
        <tr class='[[ROW_CLASS]]'>
          <td valign="middle" align="left">
                [[RPT_E2B_NUMB_REPORTS]]   
          </td>
        </tr>
        <tr class='[[ROW_CLASS]]'>
          <td valign="middle" align="left">
            <input id="bbfTpl_flt_[[ROW_NUMBER]]" name="bbfTpl_flt_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 200;
                  background-color: [[ROW_COLOR]];" value="[[RPT_E2B_REJECT_REPORTS]]" readonly />
          </td>
        </tr>
      </table>

    </td>
    <!-- 5 -->
    <td valign="middle" align="left" class="alc-header">
      <input id="fname_[[ROW_NUMBER]]" name="fname_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;
                  background-color: [[ROW_COLOR]];" value="[[RPT_E2B_FILENAME]]" readonly />
      <input id="bbfTpl_flt_[[ROW_NUMBER]]" name="bbfTpl_flt_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 100%;
                    background-color: [[ROW_COLOR]];" value="[[RPT_E2B_RECEIVE_DATE]]" readonly />
      <!--
      <table cellpadding="0" cellspacing="0">
        <tr class='[[ROW_CLASS]]'>
          <td valign="middle" align="left">
            <input id="fname_[[ROW_NUMBER]]" name="fname_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 325;
                  background-color: [[ROW_COLOR]];" value="[[RPT_E2B_FILENAME]]" readonly />
            
          </td>
        </tr>
        <tr class='[[ROW_CLASS]]'>
          <td valign="middle" align="left">
            <input id="bbfTpl_flt_[[ROW_NUMBER]]" name="bbfTpl_flt_[[ROW_NUMBER]]" class="textbox-worklist" style="width: 325;
                    background-color: [[ROW_COLOR]];" value="[[RPT_E2B_RECEIVE_DATE]]" readonly />
          </td>
        </tr>
      </table>
      -->

    </td>
    <!-- 6-->
    <td valign="middle" align="left" class="alc-header">
      <table cellpadding="0" cellspacing="0" width='100%'>
        <tr class='[[ROW_CLASS]]' width='100%'>
          <td valign="middle" align="left" width='100%'>
            <input id="bbfTpl_FOLLOWUP_[[ROW_NUMBER]]" name="bbfTpl_FOLLOWUP_[[ROW_NUMBER]]" class="textbox-worklist" style="width:90%;
                  background-color: [[ROW_COLOR]];" value="[[RPT_E2B_TRANSMIT_STATUS]]" readonly />
          </td>
        </tr>
        <tr class='[[ROW_CLASS]]' width='100%'>
          <td valign="middle" align="left" width='100%'>
            <input id="bbfTpl_FOLLOWUP_[[ROW_NUMBER]]" name="bbfTpl_FOLLOWUP_[[ROW_NUMBER]]" class="textbox-worklist" style="width:90%;
                  background-color: [[ROW_COLOR]];" value="" readonly />
          </td>
        </tr>
      </table>
    </td>
  </tr>

