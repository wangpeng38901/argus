<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
	gModuleID = "Argus.E2BApp"
%>
<%
'******************************************************************************
' Author       : Saurabh K
' Page         : E2b_ViewDifferences.asp
' Description  : 4.2 HF2 2.10.3 View E2B Import Follow-up differences Screen
'                
'******************************************************************************
' Revision History
' Date		        Author		Description
' 15-DEC-2006       Saurabh K   Initial Version
'******************************************************************************
%>
<%
	Const SEQ_NUM_FIELD = 1
	Const IMPORT_CHECKBOX = 2
	Const E2B_CURRENT_FIELD = 3
	Const E2B_DATABASE_FIELD = 4
	Const E2B_PREVIOUS_FIELD = 5
	Const E2B_DIFFCODES_FIELD = 6
	Const E2B_DTD_ELEMENT = 7
	Const E2B_IS_PARENT = 8
	Const E2B_FULL_PATH = 9
%>
<%
	Dim sTradingPartner, sDTDVersion, sCaseNum, sFollowUpNum, oTabIndex, lSelectiveIntake
	Dim sLockStatus, sLockImage, lCaseID, lUserID, lPreviousID, lCurrentID,lIsJReport 
	Dim lError, sError, oDifferenceList, lDiffCount, sCurrentElementPath, lCurrentParentTD, lCurrentLevel
	Dim LockStatus, DeleteStatus, CloseStatus, OpenedUser, esm_initial_report_id, lOpenCasePermit, lLockCasePermit
	Dim sESMReportIdCaseNum, sSelectedCaseNum
	Dim sTitleOfDialog
	Dim oParentList
	Dim selectedIsApplyNewFw
	
	Set oParentList = ExecuteSQL("select parent_element from cfg_esm_parent_mapping", "81750011", lError, sError) 
	E2bDiff_Initialize()
	Set oTabIndex = new CTabIndexIterator
	oTabIndex.SetStartTabIndex 1
%>
<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<title><%=GetTranslationData(sTitleOfDialog)%></title>
	<link href="/css/relsys.css" rel="stylesheet" />
	<base target="_self" />
	<style type="text/css">
		td {
			float: none !important;
		}
	</style>
	<script type="text/javascript">
			var lDiffCount = <%=JavaScriptSanitize(lDiffCount)%>;
			var sRedBgColor = "#FF0000";
			var sGrayBgColor = "#C1BDBD";
			var sYellowBgColor = "#FEFF99";
			var sWhiteBgColor = "#FFFFFF";
			var aElementPaths = new Array();
			var aDiffCodes = new Array();
			var aIsParent = new Array();
			var bAcceptButtonPress = true;
			var CaseLocked = <%=JavaScriptSanitize(LockStatus) %>;
			var CaseDelete = <%=JavaScriptSanitize(DeleteStatus) %>;
			var CaseClosed = <%=JavaScriptSanitize(CloseStatus) %>;
			var OpenedUser = <%=JavaScriptSanitize(OpenedUser) %>;
			var openCasePermit = <%=JavaScriptSanitize(lOpenCasePermit) %>;
			var lockCasePermit = <%=JavaScriptSanitize(lLockCasePermit) %>;
			var selectiveIntake = <%=JavaScriptSanitize(lSelectiveIntake) %>;
			var diffCount = <%=JavaScriptSanitize(lDiffCount) %>;
			var SaveorPrint = 0;
			var lMinReq = 0;
			var AcceptedCase, notes, status, sDiffDocId;
			var titleE2BImport = '<%=GetTranslationData("IMPORT")%>';
			var titleIncomingE2B = '<%=GetTranslationData("INCOME_E2B")%>';
			var sPostSave = "";
			
			// Open the PDF report for the differences
			async function fn_PrintList()
			{
				showLoading();
				var strURL;
				if (SaveorPrint == 0)
				    await fn_SaveUserCheckBox();
				
				strURL = "/Lookup/PassLargeParameters.asp?target=/E2B/Misc/PrintE2BDiffReport.asp&urldata=esm_report_id:<%=Server.URLEncode(lCurrentID)%>;case_id:<%=Server.URLEncode(lCaseID)%>;prev_id:<%=Server.URLEncode(lPreviousID)%>";
				var sDialogStyle = { dialogHeight: "200", dialogWidth: "350", resizable: false, scrollable: false };    
				sDiffDocId = await fn_OpenModalDialog(strURL,window,sDialogStyle);
				 if ((SaveorPrint == 0) && (sDiffDocId.length > 0) && (sDiffDocId != undefined))
					fn_ViewDocument(sDiffDocId, "PDFE2BDIFFWIN");
			}
			
			// On load function
			function fn_init()
            {
                document.getElementsByTagName("body")[0].style.display = "block";
				fn_DiffOptionChanged(2);
				if (selectiveIntake == 1)
                    fn_ChangeDisplay();
			}
			
			// Change display for selective intake
			function fn_ChangeDisplay()
			{
				document.getElementById("lbl_case_num").style.display = "none";
				fn_getElementByName("case_num").style.display = "none";
				document.getElementById("lbl_fup_num").style.display = "none";
				fn_getElementByName("fup_num").style.display = "none";
				fn_getElementByName("lock_state").style.display = "none";
				fn_getElementByName("diff_option").style.display = "none";
				fn_getElementByName("btn_accept").innerHTML = '<%=GetTranslationData("ACCEPT_INITIAL")%>';
				fn_getElementByName("btn_reject").innerHTML = '<%=GetTranslationData("REJECT_INITIAL")%>';
				fn_getElementByName("diff_option").style.display = "none";
			}
			
			// Onchange event for difference options
			function fn_DiffOptionChanged(lCompare)
			{
				if (selectiveIntake == 0)
				{
					showLoading();
					window.setTimeout("fn_DisplayDiff(" + lCompare + ");", 250);
				}
			}
			
			// Display the differences on screen
			function fn_DisplayDiff(lCompare)
			{
				for(var i=0; i < lDiffCount; i++)
					for(var j=1; j <=9; j*=3) 
						fn_SetCellBgColor(j*lCompare, j, i);
				
				hideLoading();
			}
			
			// lType -> 1:Current, 3:Database, 9:Previous
			// lCompare -> 1:None, 2:Current v/s Database, 4: Current v/s Previous, 8:Database v/s Previous
			// Set the background color of a cell depending upon the user selection and differences
			function fn_SetCellBgColor(lCompareType, lType, lIndex)
			{
				if(aIsParent[lIndex] == "1") 
					return;
				
				var sElement = fn_ElementName(lType);
				var sColor = fn_GetColor(lCompareType, lType, lIndex);
				
				document.getElementById("td_" + sElement + "_" + lIndex).style.backgroundColor = sColor;
			}
			
			// lType -> 1:Current, 3:Database, 9:Previous
			// lCompare -> 1:None, 2:Current v/s Database, 4: Current v/s Previous, 8:Database v/s Previous
			// Color Matrix - lCompare v/s lType. W-White, D-Calculate from Diff Code
			//   1      2      4      8
			// 1 W(1)   D(2)   D(4)   W(8)
			// 3 W(3)   D(6)   W(12)  D(24)
			// 9 W(9)   W(18)  D(36)  D(72)
			function fn_GetColor(lCompareType, lType, lIndex)
			{
				if(lCompareType == 2 || lCompareType == 6)  return fn_GetColorDiffCode(1, lIndex);
				if(lCompareType == 4 || lCompareType == 36)  return fn_GetColorDiffCode(3, lIndex); 
				if(lCompareType == 24 || lCompareType == 72)  return fn_GetColorDiffCode(2, lIndex);
				
				return sWhiteBgColor;
			}
			
			// Color based upon code and row index
			// 1. Add, 2. Update, 3.Delete
			function fn_GetColorDiffCode(lCode, lIndex)
			{
				var lCodeValue = aDiffCodes[lIndex].charAt(lCode-1);
				
				if(lCodeValue == '1')
					return sGrayBgColor;
				else if(lCodeValue == '2')
					return sYellowBgColor;
				else if(lCodeValue == '3')
					return sRedBgColor;
					
				return sWhiteBgColor;
			}
			
			// Page Element Name from it's type
			function fn_ElementName(lType)
			{
				var sElement = "curr";
				if(lType==1)
					sElement = "curr";
				else if(lType==3)
					sElement = "database";
				else if(lType==9)
					sElement = "prev";
				
				return sElement;
			}

			function fn_RecAction(lIndex)
			{
				var obj, src;
				obj = fn_getElementByName("action_" + lIndex);
				src = obj.src.toLowerCase();
				if (src.indexOf("undelete.gif") >= 0) 
				{
					obj.src= src.replace("undelete.gif", "delete.gif"); 
					var sTempPath;
					var sElementPath = aElementPaths[lIndex];
					for(var i=lIndex+1; i < lDiffCount; i++)
					{
						sTempPath = aElementPaths[i];
						if(sElementPath == sTempPath) 
							break;
						if((sTempPath.indexOf(sElementPath)==0) && (sElementPath != sTempPath))
						{
							var oControl = fn_ImportRecAction(i);
							if(oControl)
							{
								if (oControl.src.toLowerCase().indexOf("undelete.gif") >= 0)
									oControl.src= oControl.src.toLowerCase().replace("undelete.gif", "delete.gif");
							}
							oControl = fn_ImportCheckBox(i);
							if(oControl)
							{
								if(oControl.checked) {
									oControl.checked = false;
									oControl.setAttribute("rec_delete", "1");
								}
							}
						}
					}
				}
				else
				{
					obj.src= src.replace("delete.gif", "undelete.gif");
					var sTempPath;
					var sElementPath = aElementPaths[lIndex];
					for(var i=lIndex-1; i >= 0; i--) //move up
					{
						sTempPath = aElementPaths[i];
						if((sElementPath.indexOf(sTempPath)==0) && (sElementPath != sTempPath))
						{
							var oControl = fn_ImportRecAction(i);
							if(oControl)
							{
								if (oControl.src.toLowerCase().indexOf("undelete.gif") == -1)
									oControl.src= oControl.src.toLowerCase().replace("delete.gif", "undelete.gif");
							}
						}
					}

					sElementPath = aElementPaths[lIndex];
					for(var i=lIndex+1; i < lDiffCount; i++) //move down
					{
						sTempPath = aElementPaths[i];
						if(sElementPath == sTempPath) 
							break;
						if((sTempPath.indexOf(sElementPath)==0) && (sElementPath != sTempPath))
						{
							var oControl = fn_ImportRecAction(i);
							if(oControl)
							{
								if (oControl.src.toLowerCase().indexOf("undelete.gif") == -1)
									oControl.src= oControl.src.toLowerCase().replace("delete.gif", "undelete.gif");
							}

							oControl = fn_ImportCheckBox(i);
							if(oControl)
							{
								if(!(oControl.checked)) {
									oControl.checked = true;
									oControl.setAttribute("rec_delete", "0");
								}
							}
						}
					}
				}
			}
			
			// Import Check box click event
			function fn_Import(lIndex)
			{
				var oImportCheck = fn_ImportCheckBox(lIndex);
				if(aIsParent[lIndex] == "1") // Parent Element is checked or unchecked, Go all the way down
				{
					var bChecked = oImportCheck.checked;
					var sTempPath;
					var sElementPath = aElementPaths[lIndex];
					for(var i=lIndex+1; i < lDiffCount; i++)
					{
						sTempPath = aElementPaths[i];
						if(sTempPath.indexOf(sElementPath)==0 && sElementPath != sTempPath)
						{
							var oControl = fn_ImportCheckBox(i);
							if(oControl)
								if (oControl.disabled == false)
									oControl.checked = bChecked;
						}
						else
						{
							hideLoading();
							return; // A new node has started return
						 }
					}   
				}
				else
				{
					// Child Element is checked, Go all the way up
					if(oImportCheck.checked)
					{
						var sTempPath;
						var sElementPath = aElementPaths[lIndex];
						for(var i=lIndex-1; i >= 0; i--)
						{
							sTempPath = aElementPaths[i];
							if(sElementPath.indexOf(sTempPath)==0 && sElementPath != sTempPath)
							{
								var oControl = fn_ImportCheckBox(i);
								if(oControl)
								{
									oControl.checked = true;
									sElementPath = sTempPath;
								}
							}
						}
					}
				}
				hideLoading();
				return;
			}

			function fn_GetRowDTDElement(lIndex)
			{
				var oDTDElement = fn_getElementByName("dtd_element_" + lIndex);
				if(!oDTDElement)
					return "";

				var sDTDElement = leftTrim(oDTDElement.outerText);
				var lPos = sDTDElement.indexOf("]");
				if(lPos >= 0)
					sDTDElement = sDTDElement.substring(lPos + 1);
				if (sDTDElement)
					sDTDElement = trim(sDTDElement);

				return sDTDElement.toUpperCase();
			}

			function fn_IsGroupMatch(lIndex, sGroup)
			{
				if (sGroup == "" || sGroup == "ALL")
					return true;

				var sPath = "";
				if (aElementPaths[lIndex])
					sPath = ("" + aElementPaths[lIndex]).toUpperCase();

				var sDTDElement = fn_GetRowDTDElement(lIndex);
				if (sGroup == "TEST")
					return (sPath.indexOf("/TEST/") >= 0) || (sDTDElement == "TEST") || (sDTDElement.indexOf("TEST") == 0);

				if (sGroup == "EVENT")
					return (sPath.indexOf("/REACTION/") >= 0) || (sPath.indexOf("/EVENT/") >= 0) || (sDTDElement == "PRIMARYSOURCEREACTION") || (sDTDElement == "REACTIONMEDDRAPT") || (sDTDElement == "REACTIONMEDDRALLT") || (sDTDElement == "REACTION") || (sDTDElement.indexOf("REACTION") == 0) || (sDTDElement.indexOf("EVENT") == 0);

				if (sGroup == "DRUG")
					return (sPath.indexOf("/DRUG/") >= 0) || (sPath.indexOf("/MEDICINALPRODUCT/") >= 0) || (sDTDElement == "DRUG") || (sDTDElement == "MEDICINALPRODUCT") || (sDTDElement.indexOf("DRUG") == 0) || (sDTDElement.indexOf("MEDICINALPRODUCT") == 0);

				return false;
			}

			function fn_SelectImportRows(bSelect, sGroup)
			{
				showLoading();

				if (sGroup == null || sGroup == undefined)
					sGroup = "ALL";
				else
					sGroup = ("" + sGroup).toUpperCase();

				for (var i = 0; i < lDiffCount; i++)
				{
					var oControl = fn_ImportCheckBox(i);
					if (oControl)
					{
						if (oControl.disabled == false && fn_IsGroupMatch(i, sGroup))
							oControl.checked = bSelect;
					}
				}

				hideLoading();
			}

			// Save the User Options
			async function fn_SaveUserOptions()
			{
				if (await fn_ProceedConfirmation())
				{
					SaveorPrint = 1;
					showLoading();
					setTimeout("fn_SaveUserCheckBox()",500);
				}
			}

			async function fn_ProceedConfirmation()
			{
				var sESMReportIdCaseNum = "";    		
				var selectedCaseNum = "";
				var vAns  = 1;
				sESMReportIdCaseNum = <%=JavaScriptSanitize(sESMReportIdCaseNum)%>;  
				selectedCaseNum = <%=JavaScriptSanitize(sSelectedCaseNum) %>;
				if (sESMReportIdCaseNum !=  0 && sESMReportIdCaseNum !=  "" && selectedCaseNum !=  "" && selectedCaseNum.toLowerCase() != sESMReportIdCaseNum.toLowerCase())
				{    vAns = await MessageBoxRes("E2BINCOME_Q_ACCEPT_FU", "", selectedCaseNum,sESMReportIdCaseNum );   	    
				}
				
				if (vAns == MB_YES )    
					return true;
				else 
					return false;
			}

			async function fn_SaveUserCheckBox()
			{
				var sCheckedPostParams = new Array();
				var sNotCheckedPostParams = new Array();
				var sRecDeletePostParams = new Array();
				var sRecUnDeletePostParams = new Array();
				
				for(var i=0; i<lDiffCount; i++)
				{
					var oControl = fn_ImportCheckBox(i);
					var oCtlAction = fn_ImportRecAction(i);
					if(oControl)
					{
						var lSeqNum = fn_getElementByName("seq_num_" + i).value;
						
						if(oControl.checked)
							sCheckedPostParams.push(lSeqNum);
						else
							sNotCheckedPostParams.push(lSeqNum);
							
						//Save the user options for child tags for which parent tag has been deleted/undeleted
						if(oControl.rec_delete == "1")
							sRecDeletePostParams.push(lSeqNum);
						else if (oControl.rec_delete == "0")
							sRecUnDeletePostParams.push(lSeqNum);
					}
					
					//Save the user options for parent tag which is deleted/undeleted
					if (oCtlAction)
					{
						var lSeqNum = fn_getElementByName("seq_num_" + i).value;
						var src = oCtlAction.src.toLowerCase();
						if (src.indexOf("undelete.gif") >= 0)
						{
							sRecUnDeletePostParams.push(lSeqNum);
						}
						else
						{
							sRecDeletePostParams.push(lSeqNum);
						}
					}
				}
				if (selectiveIntake == 1 && SaveorPrint == 1)
					lMinReq = fn_CheckMinReq();
				
				if (lMinReq == 1)
				{
					lMinReq = 0;
					hideLoading();
					await MessageBoxRes("E2BIMPRT_AS_REQ_NOT_MET"); 
				}
				else
				{
					var sPostParms = "checked=" + sCheckedPostParams.join(",") + "&not_checked=" + sNotCheckedPostParams.join(",") + "&rec_undelete=" + sRecUnDeletePostParams.join(",") + "&rec_delete=" + sRecDeletePostParams.join(",");
					await loadArgusMessage("/E2B/Incoming/E2B_AjaxE2bImportUserOptionsSave.asp", fn_AcceptFollowup, sPostParms);
				}
			}
			
			// Check Box control at position lIndex
			function fn_ImportCheckBox(lIndex)
			{
				var oControl = fn_getElementByName("import_" + lIndex);
				if(oControl && oControl.type.toUpperCase()=="CHECKBOX")
					return oControl;
				
				return null;
			}

			function fn_ImportRecAction(lIndex)
			{
				var oControl = fn_getElementByName("action_" + lIndex);
				if(oControl)
					return oControl;
				return null;
			}

			// Call back method for accepting follow-up
			async function fn_AcceptFollowup()
			{
				var oXML = this.req.responseXML;
				var sError = fn_GetAjaxErrorMsg(oXML);
				
				bAcceptButtonPress = true;
				hideLoading();
				if (sError.length > 0) {             // Check if the returned message has error
					await MessageBoxRes("GENERAL_ERROR", "", sError);
					return;
				}
				if (SaveorPrint == 1)
				{
					await fn_PrintList();
					await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "esm_report_id=" + <%=JavaScriptSanitize(lCurrentID) %>);
				}
			}
			
			function fn_CheckMinReq()
			{
				var lDrugExists = 0, lEventExists = 0, lCOIExists = 0, oDTDElement, lPos;
				for(var i=0; i<lDiffCount; i++)
				{
					var oControl = fn_ImportCheckBox(i);
					if(oControl)
					{
						oDTDElement = fn_getElementByName("dtd_element_" + i);
						sDTDElement = leftTrim(oDTDElement.outerText);
						lPos = 0;
						lPos = sDTDElement.indexOf(']');
						sDTDElement = sDTDElement.substring(lPos + 1);
						if (sDTDElement)
							sDTDElement = trim(sDTDElement);

						if (oControl.checked)
						{
							if (sDTDElement == "MEDICINALPRODUCT")
								lDrugExists = 1;
							else if (sDTDElement == "PRIMARYSOURCEREACTION" || sDTDElement == "REACTIONMEDDRAPT" || sDTDElement == "REACTIONMEDDRALLT")
								lEventExists = 1;
							else if (sDTDElement == "OCCURCOUNTRY" || sDTDElement == "PRIMARYSOURCECOUNTRY" || sDTDElement == "REPORTERCOUNTRYR3")
								lCOIExists = 1;
						}
						else
						{
							if (sDTDElement == "OCCURCOUNTRY" || sDTDElement == "REPORTTYPE" || sDTDElement == "RECEIPTDATE" || sDTDElement == "RECEIPTDATER3") 
								return 1;
						}       
					}
				}
				
				if (lDrugExists == 0 || lEventExists == 0 || lCOIExists == 0) 
					return 1;
				else
					return 0;
			}
			
			function leftTrim(sString) 
			{
				while (sString.substring(0,1) == ' ')
				{
					sString = sString.substring(1, sString.length);
				}
				return sString;
			}
			async function fn_LockedReport()
			{
				var xmlDoc = this.req.responseXML;
				var sErrStr;
				var sLocked_User;
				
				SaveorPrint = 0;
				sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
				if (sErrStr.length > 0)
				{
					await MessageBoxRes("GENERAL_ERROR", titleIncomingE2B, sErrStr); 
					return;
				}
				var asLocked_User = xmlDoc.getElementsByTagName("LOCKED_USER");
				if (asLocked_User)
				{
					 sLocked_User = GetTextContentFromXML(asLocked_User[0]);
					if (sLocked_User.length > 0){
						await MessageBoxRes("GENERAL_INFORMATION", titleIncomingE2B, sLocked_User); }
					else
					{
						if (bAcceptButtonPress)
							await fn_AcceptSingleFollowup();
						else
						{
							await fn_RejectSingleFollowup();
						}
					}
				}
				else
				{
					if (bAcceptButtonPress)
						await fn_AcceptSingleFollowup();
					else
					{
						await fn_RejectSingleFollowup();
					}	
				}
				hideLoading();
				return;
			}
			async function fn_UnLockedReport() {
				var xmlDoc = this.req.responseXML;
				var sErrStr;
				var sLocked_User;
				sErrStr = fn_GetAjaxErrorMsg(xmlDoc);
				if (sErrStr.length > 0) {
					await MessageBoxRes("GENERAL_ERROR", titleIncomingE2B, sErrStr);
					return;
				}
				return;
			}            
			// Call back method for accepting follow-up
			async function fn_AcceptSingleFollowup()
			{
				var strURL, lWarning, Pos, sPostParms;
				var sError = "";
				var selectedCaseID = <%=JavaScriptSanitize(lCaseID) %>;
				var E2bReportID = <%=JavaScriptSanitize(lCurrentID) %>;
				var lIsE2BJReport = <%=JavaScriptSanitize(lIsJReport) %>;
				var lApplyNewFW = <%=JavaScriptSanitize(selectedIsApplyNewFw) %>;
				var justnotes = "";
				sPostSave = "";
				hideLoading();
				if (sError != "") {             // Check if the returned message has error
					await MessageBoxRes("GENERAL_ERROR", "", sError);
					return;
				}
				else
				{
					if (selectiveIntake == 0)// if it is a follow-up difference
					{
						strURL = "/E2B/Actions/E2B_AcceptFollowupE2BCase.asp?E2bType=3&IsJReport=" + lIsE2BJReport + "&esm_report_id=" + E2bReportID + "&case_id=" + selectedCaseID + "&CloseStatus=" + CaseClosed + "&LockStatus=" + CaseLocked + "&ApplyNewFw=" + lApplyNewFW + "&DocId=" + sDiffDocId  + "&esm_initial_report_id=" + <%=JavaScriptSanitize(esm_initial_report_id) %>;
						var sDialogStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };    

						if (OpenedUser != "-99")
							await MessageBoxRes("INCOMERPT_ALREADY_IN_USE", '', OpenedUser);
						else if (CaseDelete == 1)
							await MessageBoxRes("E2B_DELCASE_UPDT_FAILED");
						else if (CaseClosed == 1 && openCasePermit == "0")
							await MessageBoxRes("INCOMERPT_CLOSECASE_RESTRICTION");
						else if (CaseLocked == 1 && lockCasePermit == "0")
							await MessageBoxRes("INCOMERPT_UNLOCKCASE_RESTRICTION");
						else    
							notes = await fn_OpenModalDialog(strURL, window, sDialogStyle,false);
					}
					else
					{
						strURL = "/E2B/Actions/E2B_AcceptE2BCase.asp?E2bType=1&IsJReport="+ lIsE2BJReport +"&esm_report_id=" + E2bReportID + "&receipt_date=-1&product_name=-1&report_type_id=-1&country_id=-1&selectiveintake=1" + "&ApplyNewFw=" + lApplyNewFW + "&DocId=" + sDiffDocId;
						var strStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };
						notes = await fn_OpenModalDialog(strURL, window, strStyle, false);
					}
					if (!notes)
				    {
						await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + <%=JavaScriptSanitize(lCurrentID) %> + "&do_not_clear_inter_tables=1");
				    }
				    else
				    {
					    AcceptedCase = "";
						//"1~101~CaseNum~accepted~postsave"
					    Pos = notes.indexOf("~",1);
					    if (Pos > 0)
					    {
						    lWarning = notes.substring(0,Pos);
						    notes = notes.substring(Pos + 1);
						    Pos = notes.indexOf("~",1);
						    if (Pos > 0)
						    {
							    status = notes.substring(0,Pos);
							    notes = notes.substring(Pos + 1);
							    Pos = notes.indexOf("~",1);
							    if (Pos > 0)
							    {
								    AcceptedCase = notes.substring(0,Pos);
								    notes = notes.substring(Pos + 1)
									Pos = notes.indexOf("~", 1);
									if (Pos > 0) {
										justnotes = notes.substring(0, Pos);
										sPostSave = notes.substring(Pos + 1);
									}
									else {
										justnotes = notes;
									}
									if (justnotes.length < 1)
									    return;
								    if (status <= 0)
									    status = "";
							    }
						    }
						    else
							    status = "";
					    }
					    else
						    status = "";

					    if (status == "")
						    window.close();
					    else
					    {
						    if ((status != "102") || (AcceptedCase == "N/A"))
						    {
						        await MessageBoxRes("E2BVWR_CASE_NO_UPDT");
							    window.close();
						    }
						    else
						    {
							    sPostParms = "esm_report_id=" + <%=JavaScriptSanitize(lCurrentID)%> + "&esm_status=" + status + "&notes=" + fn_URLEncode(justnotes);
							    await loadArgusMessage("/E2B/Misc/AjaxE2bStatusChange.asp", fn_checkerrors, sPostParms);
							
						    }
					    }
				    }
				}
			}
			async function fn_checkerrors()
			{
				var oXML = this.req.responseXML; 
				var sError = fn_GetAjaxErrorMsg(oXML);
				var e2bType;
				if (selectiveIntake == 0)
					e2bType = 3;
				else
					e2bType = 1;
				if (sError.length > 0)              // Check if the returned message has error
				{
					await MessageBoxRes("GENERAL_ERROR", titleE2BImport, sError);
					window.close();
				}
				else
				{   
					if (selectiveIntake == 0)
					{
						sError = <%=JavaScriptSanitize(Replace(GetTranslationData("CASE_UPDATED_SUCCESSFULLY"),"case_num",sCaseNum))%>;
						setWindowReturnValue(2);
					}
					else
					{
						sError = <%=JavaScriptSanitize(GetTranslationData("CASE_ACCEPTED_AS"))%>;
						sError = sError.replace("case_num", AcceptedCase) ;
					}
					if (sPostSave.length > 0)
						sError = sError + "<br>" + sPostSave;
					await MessageBoxRes("GENERAL_INFORMATION",titleIncomingE2B, sError); 
					window.close();
				}
			}

			async function fn_checkreject()
			{
				var oXML = this.req.responseXML; 
				var sError = fn_GetAjaxErrorMsg(oXML);
				if (sError.length > 0)              // Check if the returned message has error
				{
				    await MessageBoxRes("GENERAL_ERROR",titleE2BImport,sError); 
					window.close();
				}
				else
				{
					if (selectiveIntake == 0)
					{
						setWindowReturnValue(2);
						window.close();
					}
					else
					{
						var retVal = "Rejected" + "#||#" + notes;
						setWindowReturnValue(retVal);
						window.close();
					}
				}
			 }
			async function fn_RejectFollowup() 
			{
				bAcceptButtonPress = false;
				await loadArgusMessage("/E2B/E2BImport/Ajax_E2BLockedReport.asp", fn_LockedReport, "esm_report_id=" + <%=JavaScriptSanitize(lCurrentID) %>);
			}
			// Method for rejecting follow-up
			async function fn_RejectSingleFollowup()
			{
				var strURL, notes, sPostParms, results;
				if (selectiveIntake == 0)// if it is a follow-up difference
					strURL = "/E2B/Actions/RejectFollowupE2BCase.asp?fType=1";
				else
					strURL = "/E2B/Actions/RejectE2BCase.asp?fType=1";
				var sDialogStyle = { dialogHeight: "480", dialogWidth: "480", resizable: false, scrollable: false };                    
				notes = await fn_OpenModalDialog(strURL, window, sDialogStyle);

				if (!notes)
				{
					 await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + <%=JavaScriptSanitize(lCurrentID) %> + "&do_not_clear_inter_tables=1");
				}    
				else
				{
					sPostParms = "esm_initial_report_id=" + <%=JavaScriptSanitize(esm_initial_report_id) %> + "&esm_report_id=" + <%=JavaScriptSanitize(lCurrentID)%> + "&esm_status=103" + "&notes=" + fn_URLEncode(notes);
					await loadArgusMessage("/E2B/Misc/AjaxE2bStatusChange.asp", fn_checkreject, sPostParms);
				}
			}
			
			async function fn_ClearInterTableAll()
			{
				var currentID = <%=JavaScriptSanitize(lCurrentID)%>;
				if (currentID > 0)
				{
					await loadArgusMessage("/E2B/Incoming/Ajax_DeleteE2BDiffData.asp", fn_AfterDelete, "esm_report_id=" + currentID);
				}
				window.close();
			}

			async function fn_AfterDelete()
			{
				var currentID = <%=JavaScriptSanitize(lCurrentID)%>;
				var oXML = this.req.responseXML; 
                var sError = fn_GetAjaxErrorMsg(oXML);
				if (sError.length > 0)             // Check if the returned message has error
					await MessageBoxRes("GENERAL_ERROR", '<%=GetTranslationData("IMPORT")%>', sError); 
				await loadArgusMessage("/E2B/E2BImport/Ajax_E2BUnLockedReport.asp", fn_UnLockedReport, "esm_report_id=" + currentID);
			}

			async function fn_ClearTables_Unload()
			{
                var currentID = <%=JavaScriptSanitize(lCurrentID) %>;
				if (currentID > 0)
				{
                    var strURL = "/E2B/Incoming/Ajax_DeleteE2BDiffData.asp?esm_report_id=" + currentID;
                    SendArgusMessage(strURL);

                    strURL = "/E2B/E2BImport/Ajax_E2BUnLockedReport.asp?esm_report_id=" + currentID;
                    SendArgusMessage(strURL);
				}
			}
	</script>
</head>
<body style="height: 100%; display:none;" onload="fn_init();">
	<table style="width: 100%; height: 100%" cellpadding="0" cellspacing="0">
		<tr style="height: 25px">
			<td style="width: 100%" class="section-header-middle">
				<%BuildLocalLabel(sTitleOfDialog).SetStyleSheet("label label-section").Render()%>
			</td>
		</tr>
		<tr style="height: 45px">
			<td style="width: 100%">
				<table style="width: 100%" cellpadding="5" cellspacing="0">
					<col style="width: 15%; padding-left: 5px; padding-top: 2px; padding-right: 5px;" />
					<col style="width: 48%; padding-left: 5px; padding-top: 2px; padding-right: 5px;" />
					<col style="width: 9%; padding-left: 5px; padding-top: 2px; padding-right: 5px;" />
					<col style="width: 28%; padding-left: 5px; padding-top: 2px; padding-right: 5px;" />
					<tr>
						<td>
							<%BuildLocalLabel("TRADING_PARTNER").Style("margin:0px").Render()%>
						</td>
						<td>
							<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "trading_partner", sTradingPartner, true, oTabIndex.NextIndex(), "").Style("width:92%").Render()%>
							<%BuildImage("lock_state", sLockImage).Render()%>
						</td>
						<td>
							<%BuildLocalLabel("DTD_VER").Render()%>
						</td>
						<td>
							<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "dtd_version", sDTDVersion, true, oTabIndex.NextIndex(), "").Style("width:99%").Render()%>
						</td>
					</tr>
					<tr>
						<td>
							<%BuildLocalLabel("ARG_CASE_NUM").SetName("lbl_case_num").Render()%>
						</td>
						<td>
							<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "case_num", sCaseNum, true, oTabIndex.NextIndex(), "").Style("width:100%").Render()%>
						</td>
						<td>
							<%BuildLocalLabel("EXP_RPT_FOLLOW_UP").SetName("lbl_fup_num").Render()%>
						</td>
						<td>
							<%BuildControlDirect(CTL_TEXTBOX_SIMPLE, "fup_num", sFollowUpNum, true, oTabIndex.NextIndex(), "").Style("width:100%").Render()%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="width: 100%; padding: 5px 5px 5px 5px">
				<table style="height: 100%; width: 100%; border: 1px solid gray" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<table style="height: 100%; width: 100%" cellpadding="0" cellspacing="0">
								<tr style="height: 25px">
									<td class="section-header-middle" style="width: 100%;">
										<table style="width: 100%" cellpadding="0" cellspacing="0">
											<tr style="height: 25px">
												<td style="width: 20%">
													<%BuildLabelDirect(GetTranslationData("NUM_ROWS") & " (" & lDiffCount & ")" ).SetStyleSheet("label label-section").Render()%>
												</td>
												<td style="width: 80%; padding-left: 5px">
													<%BuildControlDirect(CTL_DROPDOWNLIST, "diff_option", "-1", false, oTabIndex.NextIndex(), ""). _
														  AddTopOption(GetTranslationData("CURRENT_E2B_VS_CURRENT_CASE_IN_DB") & ":2;" & GetTranslationData("CURRENT_E2B_VS_LAST_IMPORTED_E2B") & ":4;" & _
																	   GetTranslationData("CURRENT_CASE_IN_DB_VS_LAST_IMPORTED_E2B") & ":8").Style("width:300px"). _
																	   OnChange("fn_DiffOptionChanged(diff_option.value);").Render()%>
												</td>
											</tr>
											<tr style="height: 25px">
												<td colspan="2" style="padding-left: 5px">
													<%BuildButton("btn_select_all", "SELECT_ALL", oTabIndex.NextIndex()).Style("width:90px").OnClick("fn_SelectImportRows(true, 'ALL');").Render()%>
													<%BuildButton("btn_deselect_all", "DESELECT_ALL", oTabIndex.NextIndex()).Style("width:90px").OnClick("fn_SelectImportRows(false, 'ALL');").Render()%>
													<%BuildButtonDirect("btn_test_select_all", "TEST " & GetTranslationData("SELECT_ALL"), oTabIndex.NextIndex()).Style("width:115px").OnClick("fn_SelectImportRows(true, 'TEST');").Render()%>
													<%BuildButtonDirect("btn_test_deselect_all", "TEST " & GetTranslationData("DESELECT_ALL"), oTabIndex.NextIndex()).Style("width:115px").OnClick("fn_SelectImportRows(false, 'TEST');").Render()%>
													<%BuildButtonDirect("btn_event_select_all", "Event " & GetTranslationData("SELECT_ALL"), oTabIndex.NextIndex()).Style("width:120px").OnClick("fn_SelectImportRows(true, 'EVENT');").Render()%>
													<%BuildButtonDirect("btn_event_deselect_all", "Event " & GetTranslationData("DESELECT_ALL"), oTabIndex.NextIndex()).Style("width:120px").OnClick("fn_SelectImportRows(false, 'EVENT');").Render()%>
													<%BuildButtonDirect("btn_drug_select_all", "Drug " & GetTranslationData("SELECT_ALL"), oTabIndex.NextIndex()).Style("width:120px").OnClick("fn_SelectImportRows(true, 'DRUG');").Render()%>
													<%BuildButtonDirect("btn_drug_deselect_all", "Drug " & GetTranslationData("DESELECT_ALL"), oTabIndex.NextIndex()).Style("width:120px").OnClick("fn_SelectImportRows(false, 'DRUG');").Render()%>
												</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td style="width: 100%">
										<table style="height: 100%; width: 100%" cellpadding="0" cellspacing="0">
											<thead>
												<tr style="height: 20px" class="tblheader-lightblue">
													<th style="width: 100%">
														<div id="HeaderDiv" class="table-scroll scroll_hide" style="width: 100%; height: 100%; border-width: 0px; overflow-y: scroll">
															<table class="table inner-table" cellpadding="5" cellspacing="0" style="width: 100%; height: 100%;">
																<col style="width: 5%;" />
																<col style="width: 29%" />
																<% if lSelectiveIntake = 0 then %>
																<col style="width: 22%" />
																<col style="width: 22%" />
																<col style="width: 22%" />
																<% else%>
																<col style="width: 66%" />
																<% end if%>
																<tr class="tblheader-lightblue">
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray">
																		<%BuildLocalLabel("SELECT").Render()%>
																	</td>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray">
																		<%BuildLocalLabel("E2B_ELEM").Render()%>
																	</td>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray">
																		<%BuildLocalLabel("CURRENT_E2B").Render()%>
																	</td>
																	<% if lSelectiveIntake = 0 then %>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray">
																		<%BuildLocalLabel("CURRENT_CASE_IN_DB").SetName("current_case").Render()%>
																	</td>
																	<td style="float: left; border-bottom: solid 1px gray">
																		<%BuildLocalLabel("LAST_IMPORTED_E2B").SetName("last_report").Render()%>
																	</td>
																	<% end if%>
																</tr>
															</table>
														</div>
													</th>
												</tr>
											</thead>
											<tbody>
												<tr>
													<td>
														<div id="BodyDiv" class="table-scroll" style="width: 100%; height: 100%; overflow-y: scroll; border-width: 0px; margin-top: 1px;">
															<table id="difftabledata" class="table inner-table" cellpadding="5" cellspacing="0" style="width: 100%; height: 100%;">
																<col style="width: 5%" />
																<col style="width: 29%" />
																<% if lSelectiveIntake = 0 then %>
																<col style="width: 22%" />
																<col style="width: 22%" />
																<col style="width: 22%" />
																<% else%>
																<col style="width: 66%" />
																<% end if%>
																<%
																	Dim lCounter, sElementName, lIsJElement, lParent
																	For lCounter=0 to lDiffCount-1
																		lIsJElement  = GetXMLValueDirect(oDifferenceList(lCounter), "ESM_DIFFERENCE_REPORT_IS_J_ELEMENT")
																		sElementName = GetXMLValueDirect(oDifferenceList(lCounter), "ESM_DIFFERENCE_REPORT_DTD_ELEMENT")
																		lParent = GetXMLValueDirect(oDifferenceList(lCounter), "ESM_DIFFERENCE_REPORT_PARENT")%>
																<% if lIsJReport = 1 and lIsJElement = 1 then %>
																<% if sElementName = "MHLWADMINITEMSICSR" then %>
																<tr style="background-color: #cccccc; height: 25px">
																	<td align="center" style="border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top">&nbsp;
																	</td>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top">&nbsp;
																	</td>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top" id="td4">&nbsp;
																	</td>
																	<% if lSelectiveIntake = 0 then %>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top" id="td5">&nbsp;
																	</td>
																	<td style="float: left; border-bottom: solid 1px gray; vertical-align: top" id="td6">&nbsp;
																	</td>
																	<% end if%>
																</tr>
																<% end if%>
																<% end if%>
																<tr>
																	<td align="center" style="border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top">
																		<%Draw E2B_DIFFCODES_FIELD, lCounter%>
																		<%Draw E2B_IS_PARENT, lCounter%>
																		<%Draw E2B_FULL_PATH, lCounter%>
																		<%Draw SEQ_NUM_FIELD, lCounter%>
																		<%Draw IMPORT_CHECKBOX, lCounter%>
																	</td>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top; <%if lParent=1 then%> font-weight: bold <% end if%>">
																		<%Draw E2B_DTD_ELEMENT, lCounter%><%if lParent=1 then%><img src="/img/TreeView/FolderOpen.gif" alt="Node" title="Node"><%end if%>
																	</td>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top" id="td_curr_<%=lCounter%>">
																		<%Draw E2B_CURRENT_FIELD, lCounter%>
																	</td>
																	<% if lSelectiveIntake = 0 then %>
																	<td style="float: left; border-right: solid 1px gray; border-bottom: solid 1px gray; vertical-align: top" id="td_database_<%=lCounter%>">
																		<%Draw E2B_DATABASE_FIELD, lCounter%>
																	</td>
																	<td style="float: left; border-bottom: solid 1px gray; vertical-align: top" id="td_prev_<%=lCounter%>">
																		<%Draw E2B_PREVIOUS_FIELD, lCounter%>
																	</td>
																	<% end if%>
																</tr>
																<% Response.Flush()
																Next%>
															</table>
														</div>
													</td>
												</tr>
											</tbody>
										</table>
									</td>
								</tr>
								<tr style="height: 25px">
									<td style="width: 100%; border-top: solid 1px gray">
										<table style="width: 100%" cellpadding="0" cellspacing="0">
											<tr style="background-color: #cccccc; height: 25px">
												<td>&nbsp;
												</td>
												<td style="float: right; padding-right: 3px; text-align: right;">
													<%BuildButton("btn_accept", "E2B_ACCEPT_FU", oTabIndex.NextIndex()).Style("width:100px").OnClick("fn_SaveUserOptions()").Render()%>
													<%BuildButton("btn_reject", "E2B_REJECT_FU", oTabIndex.NextIndex()).Style("width:100px").OnClick("fn_RejectFollowup()").Render()%>
													<%BuildButton("btn_print", "PRINT_LIST", oTabIndex.NextIndex()).Style("width:80px").OnClick("fn_PrintList();").Render()%>
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr style="background-color: #cccccc; height: 25px">
			<td style="float: right;" style="padding-right: 5px">
				<%BuildButton("btn_close", "BTN_CLOSE", oTabIndex.NextIndex()).Style("width:70px; float:right; margin-right:5px").OnClick("fn_ClearInterTableAll();").Render()%>
			</td>
		</tr>
	</table>
</body>
</html>
<script type="text/javascript">
    addListener(this, "unload", async function(){await fn_ClearTables_Unload();});
</script>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->

<%
'*******************************************************************
' Author		:   Saurabh K
' Called From	:   This Page
' Parameters	:   NA
' Returns		:   NA
' Description	:   Page Init Function
' Revision History :
' Date                      Author                          Description
' 15-DEC-2006               Saurabh K                       Initial Version
'********************************************************************
	Sub E2bDiff_Initialize
		lSelectiveIntake = 0
		sESMReportIdCaseNum = Request.QueryString("report_case_num")
		sSelectedCaseNum = Request.QueryString("selected_case_num")        
		lCaseID = GetLong(Request.QueryString("case_id"), 0)
		lUserID = Request.QueryString("user_id")
		lPreviousID = GetLong(Request.QueryString("prev_id"), -1)
		lCurrentID = GetLong(Request.QueryString("current_id"), -1)
		LockStatus = GetLong(GetRequest("LockStatus"), -1)
		DeleteStatus = GetLong(GetRequest("DeleteStatus"), -1)
		CloseStatus = GetLong(GetRequest("CloseStatus"), -1)
		OpenedUser = GetRequest("OpenedUser")
		lIsJReport = GetRequest("isJReport")
		selectedIsApplyNewFw = GetLong(Request("ApplyNewFw"), 0)
		esm_initial_report_id = GetLong(GetRequest("esm_initial_report_id"), 0)
		lOpenCasePermit = GetXMLValueDirect(oSession, "CFG_USERS_ALLOW_CLOSING")
		lLockCasePermit = GetXMLValueDirect(oSession, "CFG_USERS_ALLOW_LOCKING")
		sCurrentElementPath = ""
		if lCaseID = 0 Then         
			lSelectiveIntake = 1
		end if
		
		If IsNullOrEmpty(lUserID) Then lUserID = oArgusUser.GetUserId()
		
		E2bDiff_GetData
		If lSelectiveIntake = 0 Then Call E2bDiff_InitImage

		If lSelectiveIntake = 0 Then
			sTitleOfDialog = "E2B_DIFF_REPORT"
		Else
			sTitleOfDialog = "INI_SEL_ACC_RPT"
		End If
	End Sub

'*******************************************************************
' Author		:   Saurabh K
' Called From	:   This Page
' Parameters	:   NA
' Returns		:   NA
' Description	:   Page Init Function
' Revision History :
' Date                      Author                          Description
' 15-DEC-2006               Saurabh K                       Initial Version
'********************************************************************    
	Function E2bDiff_GetData
		
		Dim oCaseMasterList, oMessage, oOutMsg, oCaseMaster
		
		Call CreateMessage (oMessage, 501200027) ' MID_bus_app_e2b_differences
		Call SetXMLValueDirect (oMessage, "CSM_CASE_ID", lCaseID)
		Call SetXMLValueDirect (oMessage, "CFG_USERS_USER_ID", lUserID)
		Call SetXMLValueDirect (oMessage, "RPT_E2B_PREVIOUS_FOLLOWUP", lPreviousID)
		Call SetXMLValueDirect (oMessage, "RPT_E2B_REPORT_ID", lCurrentID)
		Call SetXMLValueDirect (oMessage, "GN_NUMBER1", lSelectiveIntake)
		
		Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lError, sError)
		'oOutMsg.save "c:\temp\diff.xml"
		Set oCaseMasterList = oOutMsg.selectNodes ("/MESSAGE/TABLE_CASE_MASTER/CASE_MASTER")
		If oCaseMasterList.length <> 0 Then
			Set oCaseMaster = oCaseMasterList(0)
			sTradingPartner = GetXMLValueDirect(oCaseMaster, "RPT_E2B_SAFETYREPORTID")
			sDTDVersion = GetXMLValueDirect(oCaseMaster, "CFG_PROFILE_PROFILE_VERSION")
			sCaseNum = GetXMLValueDirect(oCaseMaster, "CSM_CASE_NUM")
			sFollowUpNum = GetXMLValueDirect(oCaseMaster, "GN_RPT_STRING")
			sLockStatus = GetXMLValueDirect(oCaseMaster, "Integer")
		End If
		
		Set oDifferenceList = oOutMsg.selectNodes ("/MESSAGE/TABLE_ESM_DIFFERENCE_REPORT/ESM_DIFFERENCE_REPORT")
		lDiffCount = oDifferenceList.length
		
		E2bDiff_GetData = True
	End Function

'*******************************************************************
' Author		:   Saurabh K
' Called From	:   This Page
' Parameters	:   NA
' Returns		:   NA
' Description	:   Page Init Function
' Revision History :
' Date                      Author                          Description
' 15-DEC-2006               Saurabh K                       Initial Version
'********************************************************************
	Sub E2bDiff_InitImage
		
		If sLockStatus = "0" Then
			sLockImage = "/img/cf/case_open.gif"
		Else 
			If sLockStatus = "1" Then
				sLockImage = "/img/cf/case_locked.gif"
			Else
				sLockImage = "/img/cf/case_closed.gif"
			End If
		End If 
	End Sub

'*******************************************************************
' Author		:   Saurabh K
' Called From	:   This Page
' Parameters	:   Field Type, Index in the Difference NodeList
' Returns		:   NA
' Description	:   Draws the controls on the page
' Revision History :
' Date                      Author                          Description
' 16-DEC-2006               Saurabh K                       Initial Version
'********************************************************************    
	Sub Draw(lType, lIndex)
		If(lType = SEQ_NUM_FIELD) Then
			DrawHidden lType, lIndex
		ElseIf(lType = IMPORT_CHECKBOX) Then
			DrawImportCheckBox lIndex
		ElseIf(lType = E2B_DTD_ELEMENT OR lType = E2B_CURRENT_FIELD OR lType = E2B_DATABASE_FIELD OR lType = E2B_PREVIOUS_FIELD) Then
			DrawElement lType, lIndex
		ElseIf(lType = E2B_FULL_PATH OR lType=E2B_IS_PARENT OR lType=E2B_DIFFCODES_FIELD) Then
			PutInArray lType, lIndex
		End If
	End Sub

'*******************************************************************
' Author		:   Saurabh K
' Called From	:   This Page
' Parameters	:   Field Type, Index in the Difference NodeList
' Returns		:   NA
' Description	:   Draws the hidden controls on the page
' Revision History :
' Date                      Author                          Description
' 16-DEC-2006               Saurabh K                       Initial Version
'********************************************************************        
	Sub DrawHidden(lType, lIndex)
		Dim sName, sValue
		
		If(lType = SEQ_NUM_FIELD) Then
			sName = "seq_num_" & lIndex
			sValue = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_SEQ_NUM")
		End If
		
		Call BuildHiddenControlDirect(sName, sValue)
	End Sub

'*******************************************************************
' Author		:   Saurabh K
' Called From	:   Draw
' Parameters	:   Field Type, Index
' Returns		:   NA
' Description	:   Populates the arrays required for on screen operations
' Revision History :
' Date                      Author                          Description
' 16-DEC-2006               Saurabh K                       Initial Version
'********************************************************************            
	Sub PutInArray(lType, lIndex)
		Dim sParentElement, sDTDElement, lParent
		If(lType = E2B_FULL_PATH) Then
			sParentElement  = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_PARENT_ELEMENT")
			sDTDElement     = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DTD_ELEMENT")
			lParent         = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_PARENT")
			
			If (sParentElement = "MHLWDUMMY") then
				sParentElement = "MHLWADMINITEMSICSR"
			End If
			
			Call CalculateCurrentElementPath (sParentElement,sDTDElement,lParent)
			
			Response.Write("<script type='text/javascript'>aElementPaths[" & lIndex & "] = '" & sCurrentElementPath & "'</script>")
		ElseIf(lType = E2B_DIFFCODES_FIELD) Then
			Dim sDiffCode1, sDiffCode2, sDiffCode3, sDiffCodes
			sDiffCode1 = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DIFF_CODE1")
			If IsNullOrEmpty(sDiffCode1) OR sDiffCode1 = "-1" Then sDiffCode1 = "0"
			sDiffCode2 = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DIFF_CODE2")
			If IsNullOrEmpty(sDiffCode2) OR sDiffCode2 = "-1" Then sDiffCode2 = "0"
			sDiffCode3 = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DIFF_CODE3")
			If IsNullOrEmpty(sDiffCode3) OR sDiffCode3 = "-1" Then sDiffCode3 = "0"
			
			sDiffCodes = sDiffCode1 & sDiffCode2 & sDiffCode3
			Response.Write("<script type='text/javascript'>aDiffCodes[" & lIndex & "] = '" & sDiffCodes & "'</script>")
		ElseIf(lType = E2B_IS_PARENT) Then
			Dim sIsParent
			sIsParent =  GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_PARENT")
			Response.Write("<script type='text/javascript'>aIsParent[" & lIndex & "] = '" & sIsParent & "'</script>")
		End If
	End Sub
	
'*******************************************************************
' Author		:   Saurabh K
' Called From	:   This Page
' Parameters	:   Index in the Difference NodeList
' Returns		:   NA
' Description	:   Draws the import check box on the page
' Revision History :
' Date                      Author                          Description
' 16-DEC-2006               Saurabh K                       Initial Version
'********************************************************************    
	Sub DrawImportCheckBox(lIndex)
		Dim lCheck, sValue, lDiffCode1, lAlwaysImport, lParent, lRecAction, oImage, bDelImage, oNodeList
		lAlwaysImport = 0
		lParent = 0
		lRecAction = -1
		bDelImage = false
		sValue = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DTD_ELEMENT")
		lCheck = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_IMP_CHECK")
		lAlwaysimport = GetXMLValueDirect(oDifferenceList(lIndex), "RPT_E2B_ALWAYS_IMPORT")
		lParent = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_PARENT")
		lRecAction = GetLong(GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_REC_ACTION"), -1)
		
		If lParent = 1 and (lRecAction = 3 or lRecAction = 0) then
			Set oNodeList = oParentList.selectNodes ("/MESSAGE/TABLE_ESM_DIFFERENCE_REPORT/ESM_DIFFERENCE_REPORT[ESM_DIFFERENCE_REPORT_PARENT_ELEMENT='"& sValue & "']")
			if oNodeList.length > 0 then bDelImage = true
		end if

		If(IsNullOrEmpty(lCheck)) Then lCheck = "-1"
		If (IsNullOrEmpty(lAlwaysimport)) Then lAlwaysimport = 0
		If (IsNullOrEmpty(lParent)) Then lParent = 0
		If (IsNullOrEmpty(lRecAction)) Then lRecAction = -1

		If bDelImage=true then
			If lRecAction = 3 then
				BuildImage("action_" & lIndex, "\img\e2b\Delete.gif").Style("width:10px").OnClick("fn_RecAction(" & lIndex & ");").Render()
			Elseif lRecAction = 0 then
				BuildImage("action_" & lIndex, "\img\e2b\UnDelete.gif").Style("width:10px").OnClick("fn_RecAction(" & lIndex & ");").Render()
			End If
		Else
			If(lCheck = "-1") Then
				Call BuildHiddenControlDirect("import_" & lIndex, "-1")
			Else
				If (lSelectiveIntake = 1) and ((sValue = "OCCURCOUNTRY") or (sValue = "REPORTTYPE") or (sValue = "RECEIPTDATE"))   Then
					'in selective intake for mandatory elements make it disabled
					BuildControlDirect(CTL_CHECKBOX, "import_" & lIndex, (lCheck = "1"), true, oTabIndex.NextIndex(), "").OnClick("showLoading(); setTimeout('fn_Import(" & lIndex & ")', 250)").Render()
				Else
					lDiffCode1 = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DIFF_CODE1")
					if (lAlwaysimport = 1 And lParent = 0) Then 'If the element has been marked as Always Import then check box will be always be chekced and read only                                    
						BuildControlDirect(CTL_CHECKBOX, "import_" & lIndex, true, true, oTabIndex.NextIndex(), "").OnClick("showLoading(); setTimeout('fn_Import(" & lIndex & ")', 250)").Render()
					ElseIf lDiffCode1 = "3" Then 'for red items check box will be unchecked and disabled
						BuildControlDirect(CTL_CHECKBOX, "import_" & lIndex, (lRecAction=0), true, oTabIndex.NextIndex(), "").OnClick("showLoading(); setTimeout('fn_Import(" & lIndex & ")', 250)").Render()            
					Else
						BuildControlDirect(CTL_CHECKBOX, "import_" & lIndex, (lCheck = "1"), false, oTabIndex.NextIndex(), "").OnClick("showLoading(); setTimeout('fn_Import(" & lIndex & ")', 250)").Render()
					End If
				End If
			End If        
		End if
		
		
	End Sub

'*******************************************************************
' Author		:   Saurabh K
' Called From	:   This Page
' Parameters	:   Field Type, Index in the Difference NodeList
' Returns		:   NA
' Description	:   Draws the visible elements on the grid on the page
' Revision History :
' Date                      Author                          Description
' 16-DEC-2006               Saurabh K                       Initial Version
'********************************************************************    
	Sub DrawElement(lType, lIndex)
		Dim sName, sValue, sDataElement
		If(lType = E2B_DTD_ELEMENT) Then
			sName = "dtd_element_" & lIndex
			sValue = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DTD_ELEMENT")
			sDataElement = GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DATA_ELEMENT")
			sValue = IIf(IsNullOrEmpty(sDataElement), sValue,  "[" & sDataElement & "] " & sValue)
			sValue = Gap() & Fn_Sanitize(sValue)
		ElseIf(lType = E2B_CURRENT_FIELD) Then
			sName = "e2b_curr_" & lIndex
			sValue = Fn_Sanitize(GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DIFF_VALUE1"))
		ElseIf(lType = E2B_DATABASE_FIELD) Then
			sName = "e2b_database_" & lIndex
			sValue = Fn_Sanitize(GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DIFF_VALUE2"))
		ElseIf(lType = E2B_PREVIOUS_FIELD) Then
			sName = "e2b_prev_" & lIndex
			sValue = Fn_Sanitize(GetXMLValueDirect(oDifferenceList(lIndex), "ESM_DIFFERENCE_REPORT_DIFF_VALUE3"))
		End If
		
		If IsNullOrEmpty(sValue) Then sValue = "&nbsp;"
		BuildLabelDirect(sValue).SetStyleSheet("page-text").SetName(sName).Encode(false).Render()
	End Sub
		
'*******************************************************************
' Author		:   Saurabh K
' Called From	:   DrawHidden
' Parameters	:   Parent DTD Name, Element DTD Name, Parent Flag
' Returns		:   NA
' Description	:   Sets the current path and level in the XML hierarchy
' Revision History :
' Date                      Author                          Description
' 16-DEC-2006               Saurabh K                       Initial Version
'********************************************************************    
	Sub CalculateCurrentElementPath(sParent, sElement, lIsParent)
		
		Dim lPos
		If(sCurrentElementPath = "") Then
			lPos = 0
		Else
			lPos = InStr(sCurrentElementPath, "/" & sParent & "/")
			If lPos <> 0 Then
				sCurrentElementPath = Left(sCurrentElementPath, lPos-1)
				sCurrentElementPath = sCurrentElementPath & "/" & sParent & "/" & sElement & "/"
			End If
		End If
		
		If lPos = 0 Then
			If IsNullOrEmpty(sParent) Then
				sCurrentElementPath = "/" & sElement & "/"
			Else
				sCurrentElementPath = "/" & sParent & "/" & sElement & "/"
			End If
		End If
		lCurrentLevel = UBound(Split(sCurrentElementPath, "/"))-1
	End Sub

'*******************************************************************
' Author		:   Saurabh K
' Called From	:   DrawElement
' Parameters	:   NA
' Returns		:   White Space String
' Description	:   Returns 3*(lCurrentLevel-1) white spaces to show the E2B data hierarchy
' Revision History :
' Date                      Author                          Description
' 16-DEC-2006               Saurabh K                       Initial Version
'********************************************************************    
	Function Gap()
		Dim iCount
		Gap = ""
		For iCount=0 To lCurrentLevel-3
			Gap = Gap & "&nbsp;&nbsp;&nbsp;&nbsp;"
		Next
	End Function
%>

