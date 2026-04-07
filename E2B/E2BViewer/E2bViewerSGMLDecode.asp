<!-- #INCLUDE VIRTUAL="/Nav/DialogHeader_inc.asp" -->
<%
    gModuleID = "Argus.E2BApp"
%>
<% Server.ScriptTimeout = 86400 %>
<%
'******************************************************************************
' Author       : Prakash Singh
' Page         : E2bViewerSGMLDecode.asp
' Description  : this E2B View provides the ability to view the decode version of the element
'******************************************************************************
' Revision History
' Date		Author		Description
' 15jun2006 Prakash     Original
'******************************************************************************
%>
<!DOCTYPE html>
<%    
	Dim sDocId,XMLDisp,strSQL,sFields,i,j,profileStr,EsmReportID,sLastNode, AuthorityId
	Dim maxcount,arraycount,printarray(),strData,sDTDVersion, lReportFormId
	Dim lErrNo, sError
	Dim bSuccess, NRows	
	Dim E2BViewType,sHL7Profile, sProfile1, sProfile2, sRootElement1, sRootElement2, oOutMsg,sProfileRelease
	Dim iPos,IncomingE2b
	Dim DateFormatExtension
    Dim sValidationErrorReportDocId

    sValidationErrorReportDocId = ""    
	sErrorCFGM2Cache = ""
	bSuccess = True
	arraycount = "0"
	sLastNode = ""
	sProfile1 = ""
	sProfile2 = ""
	sRootElement1 = ""
	sRootElement2 = ""
    sProfileRelease=""
    lReportFormId = ""
	
	sDocId = GetString(GetRequest("DocId"), "")
	EsmReportID = GetRequest("EsmReportID")
    AuthorityId = GetLong(GetRequest("AuthorityId"), 0)
    sHL7Profile=GetRequest("HL7Profile")
	sDTDVersion= GetRequest("sDTDVersion")
    E2BViewType = GetLong(GetRequest("E2BViewType"), 0)
    sProfileRelease=GetRequest("sProfileRelease")
    IncomingE2b = GetLong(GetRequest("IncomingE2b"), 0)
    lReportFormId = GetLong(GetRequest("ReportFormId"),0)
    DateFormatExtension = "DATE_EXTENSION"
    Call SetParameter("P_REPORT_ID", EsmReportID, PARAM_STRING)

    If (IncomingE2b = 1 AND sHL7Profile="1" AND AuthorityId = 3 ) Then
		sProfile1 = GetValueFromCMN_PROFILEOnKeySection ("E2B_R3_IMPORT_PROFILE", "DATABASE")
        If LEN(sProfile1) > 0 THEN sProfileRelease = "2.0" End If
    ELSE
        strSQL = "select profile from  SAFETYREPORT where REPORT_ID = :P_REPORT_ID "
        Set oOutMsg = ExecuteSQL (strSQL, "29150004", lErrNo, sError)        
        sProfile1 = GetXMLValueDirect(oOutMsg, "/MESSAGE/TABLE_CFG_RECEIVER/CFG_RECEIVER/CFG_RECEIVER_MESSAGE_PROFILE")   

        strSQL = "select profile from  MHLWADMINITEMSICSR where REPORT_ID = :P_REPORT_ID "
        Set oOutMsg = ExecuteSQL (strSQL, "29150012", lErrNo, sError)
        sProfile2 = GetXMLValueDirect(oOutMsg, "/MESSAGE/TABLE_CFG_RECEIVER/CFG_RECEIVER/CFG_RECEIVER_MESSAGE_PROFILE2")    
    	
        If IsNullOrEmpty(sProfile1) Then
			strSQL = "select CFG_RECEIVER.MESSAGE_PROFILE, CFG_RECEIVER.MESSAGE_PROFILE2 from CFG_RECEIVER, SAFETYREPORT " &_
              "where SAFETYREPORT.REPORT_ID = :P_REPORT_ID AND SAFETYREPORT.AGENCY_ID = CFG_RECEIVER.AGENCY_ID (+) "
        	Set oOutMsg = ExecuteSQL (strSQL, "29150004, 29150012", lErrNo, sError)
        	sProfile1 = GetXMLValueDirect(oOutMsg, "/MESSAGE/TABLE_CFG_RECEIVER/CFG_RECEIVER/CFG_RECEIVER_MESSAGE_PROFILE")
        	sProfile2 = GetXMLValueDirect(oOutMsg, "/MESSAGE/TABLE_CFG_RECEIVER/CFG_RECEIVER/CFG_RECEIVER_MESSAGE_PROFILE2") 
        END IF  
     
    END IF

    If Not IsNullOrEmpty(sProfile1) Then
        Call SetParameter("PROFILE1", sProfile1, PARAM_STRING)
        strSQL = "select root_element from cfg_profile where profile = :PROFILE1"
        sRootElement1 =  ExecuteSQLReturnStr(strSQL, lErrNo, sError)
        If IsNullOrEmpty(sRootElement1) Then
            bSuccess = False
            sErrorCFGM2Cache = "ROOT_ELEMENT_NOT_CONFIG"
        End If
    End If

    If Not IsNullOrEmpty(sProfile2) and bSuccess Then
        Call SetParameter("PROFILE2", sProfile2, PARAM_STRING)
        strSQL = "select root_element from cfg_profile where profile = :PROFILE2"
        sRootElement2 =  ExecuteSQLReturnStr(strSQL, lErrNo, sError)
        If IsNullOrEmpty(sRootElement2) Then
            bSuccess = False
            sErrorCFGM2Cache = "ROOT_ELEMENT_NOT_CONFIG"
        End If
    End If

     
    
    If bSuccess Then
        If E2BViewType = 4 Then 'I - Decodede View
            If UCase(sRootElement1) ="ICHICSR" Then 
                profileStr =sProfile1
            Else
                profileStr =sProfile2
            End If
        Else                  'J - Decoded View      
            If UCase(sRootElement1) ="MHLWADMINITEMSICSR" Then 
                profileStr =sProfile1
            Else
                profileStr =sProfile2
            End If
        End If
    End If

    If IsNullOrEmpty(sDocId) and bSuccess Then
        bSuccess = False
        'If Validation Error Pdf is shown, donot display below error message.  
        sValidationErrorReportDocId = GetRequest("sValidationErrorReportDocId")
        If IsNullOrEmpty(sValidationErrorReportDocId) Then             
           sErrorCFGM2Cache = "ERROR_GEN_FILE"
        End If       
    End If
    If IsNullOrEmpty(profileStr) and bSuccess Then
        bSuccess = False
        sErrorCFGM2Cache = "PROFILE_NOT_EXISTS"
    End If
      
    If bSuccess = True Then
        'Load the XML from DOC ID 
        bSuccess = LoadXMLMessage (XMLDisp, sDocId)
        If (bSuccess = False) Then
            sErrorCFGM2Cache = "ERROR_GEN_FILE"
        End If
    End If

    If (bSuccess = True) Then
        Call SetParameter("PROFILE", profileStr, PARAM_STRING)
        strSQL = "select DTD_ELEMENT,DATA_ELEMENT,dtd_element_title" & IIF (glDisplayLang = cfCMN_LANG_JP,gSUFFIX_JP,"") & ",repeatable, DTD_ELEMENT_TYPE from CFG_E2B where profile = :PROFILE"
        set sFields = ExecuteSQL(strSQL, "25150002,25150004,25150011,25150075,25110110", lErrNo, sError)
        maxcount = sFields.childnodes(0).childnodes(0).childnodes.length
        Traverse XMLDisp,"0","0", sFields, 0 ,sProfileRelease       
    End If
%>
<html>
<head>
    <!-- Page Title -->
    <title></title>
    <!-- Include Stylesheet here -->
    <link rel="stylesheet" href="/css/Relsys.css" />
    <style type="text/css">
        .collapse {
            position: absolute;
            visibility: hidden;
            display: none;
        }

        .expand {
            position: relative;
            visibility: visible;
            display: block;
        }
    </style>
    <script type="text/javascript">
        var l_IncomingE2b = <%=IncomingE2b%>;

        async function window_onload() {
            var s = <%=JavaScriptClean(sErrorCFGM2Cache) %>;
            if (s != "") {
                await MessageBoxRes(s);
                window.parent.close();
            }
        }
        
        function fn_ShowAttachment(attachId)
        {  
            strURL = "/E2B/E2BVIEWER/GETE2BATTACHMENT.ASP?attach_id="+attachId + "&IncomingE2b=" + l_IncomingE2b;
            var sFeatures = 'scrollbars=yes,resizable=yes,width=400,height=160';
            strURL = fn_AddRequestToken(strURL);
            window.open(strURL, "", sFeatures);
            window.close();                
        }
    </script>
</head>
<body class="no-margin" onload="window_onload()">
    <form name="Frm" id="Frm" action="" method="post" class="no-margin">
    	<!-- #INCLUDE VIRTUAL="/Include/CommonForm_inc.asp" -->
        <table class="table border-blue inner-table" cellspacing="0" cellpadding="5" id="TabDisplay" width="100%">
            <tr class="tblheader" style="height: 20px">
                <td class="grd-header" width="30%" background='/img/Common/table_top.gif' class="margin-left-right">
                    <span class="label label-section" style='color: white; background-color: Transparent; font-size: 8pt'>#</span>
                </td>
                <td class="grd-header" width="30%" background='/img/Common/table_top.gif' class="margin-left-right">
                    <span class="label label-section" style='color: white; background-color: Transparent; font-size: 8pt'>
                        <%BuildLocalLabel("FOLDER_TREE").SetStyleSheet("label label-section").Render() %></span>
                </td>
                <td class="grd-header" width="40%" background='/img/Common/table_top.gif'>
                    <span class="label label-section" style='color: white; background-color: Transparent; font-size: 8pt'>
                        <%BuildLocalLabel("DESC").SetStyleSheet("label label-section").Render() %></span>
                </td>
            </tr>
            <%For j = 2 to arraycount%>
            <tr>
                <td class="alc-header">
                    <table class="inner-table" width="100%">
                        <tr>
                            <td width='<% Response.Write ((printarray(0,j)-22)*1.2) & "%"%>'></td>
                            <td class="padding-all" width='<% Response.Write (100 -(printarray(0,j)-22)*1.2) & "%"%>'>
                                <%BuildLabelDirect(printarray(1,j)).Render()%>
                            </td>
                        </tr>
                    </table>
                </td>
                <td class="alc-header">
                    <table cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <td width='<% Response.Write ((printarray(0,j)-10)*0.9) & "%"%>' align="right">
                                <%If printarray(6,j) = 1 Then
				            If printarray(5,j) = 1 Then%>
                                <img src="/img/TreeView/repeatablenodes.bmp" alt="Repeating Node" title="Repeating Node">
                                <%Else%>
                                <img src="/img/TreeView/FolderOpen.gif" alt="Node" title="Node">
                                <%End If
			            End If%>
                            </td>
                            <td class="padding-all" width='<% Response.Write (100 -(printarray(0,j)-10)*0.9) & "%"%>'>
                                <%If len(printarray(2,j)) < 210 then
				            If printarray(6,j) <> 1 then				        
				                BuildLabelDirect(printarray(2,j)).SetStyleSheet("label-nobold").Render()
				            Else
				                BuildLabelDirect(printarray(2,j)).Render()
				            End If
				        Else%>
                                <textarea rows="3" readonly class="textarea text-readonly" id="textarea1" name="textarea1"
                                    style="width: 100%" id="Tree_<%j%>"> <%=strdata %></textarea>
                                <%end if%>
                            </td>
                        </tr>
                    </table>
                </td>
                <%
			strData = printarray(3,j)
            if printarray(8,j) = "10" or printarray(8,j) = "11" then
                Dim nfStr
                nfStr = GetNullFlavorDecodeValue(strData)
                If IsNullOrEmpty(nfStr) Then
                    strData = strData & (dateformatconversion (strData, Null))	
                else
                    strData = strData & " (" & printarray(4,j) & ")"
                end if		
			elseif  printarray(4,j) <> "" then
				strData = strData & " (" & printarray(4,j) & ")"
			elseif instr(ucase(cstr(printarray(7,j-1))),"DATEFORMAT")> 0 then
				strData = strData & ( dateformatconversion (cstr(printarray(3,j)),cstr(printarray(4,j-1))) )		
			elseif Right(ucase(cstr(printarray(7,j))), len(DateFormatExtension)) = DateFormatExtension then
			    strData = strData & ( dateformatconversion (cstr(printarray(3,j)), Null) )		
			end if%>
                <td class="alc-header">
                    <% if ucase(printarray(7,j))="INCLUDEDOCUMENT" or ucase(printarray(7,j))="INCLUDEDDOCUMENT" or  ucase(printarray(7,j))="LITINCLUDEDOCUMENT" or ucase(printarray(7,j))="TEXT" Then 
            
               Dim attach_ID 
               If (IncomingE2b = 1) Then
                   attach_ID = strData
               else
               if InStr(strData, "-")>0 Then
               attach_ID = Mid(strData,1,InStr(strData, "-")-1)   
               end if
               end if
                    %>
                    <a href="#" onclick="fn_ShowAttachment(<%= attach_ID %>)"><%=Fn_Sanitize(strData) %> </a>
                    <% else %>
                    <%if len(strData) > 255 then
                NRows = GetIndex(len(strData))
                NRows = CInt(NRows) * 5
                    %>
                    <textarea rows="<%=NRows%>" readonly class="textarea text-readonly" style="width: 100%"> <%=Fn_Sanitize(strdata)%></textarea>
                    <%else
				if trim(strData) = "" then strData = " "
				BuildLabelDirect(strData).SetStyleSheet("label-nobold").Encode(true).Render()
			end if%>
                    <%end if %>
                &nbsp;
                </td>
            </tr>
            <%Next%>
        </table>
        <table class="table border-blue inner-table" cellspacing="0" cellpadding="0" id="TabPrint"
            style="display: none">
            <tr class="tblheader" style="height: 20px">
                <td class="grd-header" width="30%" background='/img/Common/table_top.gif' class="margin-left-right">
                    <span class="label label-section" style='color: white; background-color: Transparent; font-size: 8pt'>#</span>
                </td>
                <td class="grd-header" width="30%" background='/img/Common/table_top.gif' class="margin-left-right">
                    <span class="label label-section" style='color: white; background-color: Transparent; font-size: 8pt'>
                        <%BuildLocalLabel("FOLDER_TREE").Render() %></span>
                </td>
                <td class="grd-header" width="40%" background='/img/Common/table_top.gif'>
                    <span class="label label-section" style='color: white; background-color: Transparent; font-size: 8pt'>
                        <%BuildLocalLabel("DESC").Render() %>
                    </span>
                </td>
            </tr>
            <%For j = 2 to arraycount%>
            <tr>
                <td class="alc-header">
                    <table class="inner-table" width="100%">
                        <tr>
                            <td width='<% Response.Write ((printarray(0,j)-22)*1.2) & "%"%>'></td>
                            <td class="padding-all" width='<% Response.Write (100 -(printarray(0,j)-22)*1.2) & "%"%>'>
                                <%BuildLabelDirect(printarray(1,j)).Render()%>
                            </td>
                        </tr>
                    </table>
                </td>
                <td class="alc-header">
                    <table width="100%">
                        <tr>
                            <td width='<% Response.Write ((printarray(0,j)-10)*0.9) & "%"%>' align="right">
                                <%if printarray(6,j) = 1 then
				            if printarray(5,j) = 1 then%>
                                <img src="/img/TreeView/repeatablenodes.bmp" alt="Repeating Node" title="Repeating Node">
                                <%else%>
                                <img src="/img/TreeView/FolderOpen.gif" alt="Node" title="Node">
                                <%end if
			            end if%>
                            </td>
                            <td class="padding-all" width='<% Response.Write (100 -(printarray(0,j)-10)*0.9) & "%"%>'>
                                <%if printarray(6,j) <> 1 then				        
				            BuildLabelDirect(printarray(2,j)).SetStyleSheet("label-nobold").Render()
				        else
				            BuildLabelDirect(printarray(2,j)).Render()
				        end if%>
                            </td>
                        </tr>
                    </table>
                </td>
                <%
			strData = printarray(3,j)
			if  printarray(4,j) <> "" then
				strData = strData & " (" & printarray(4,j) & ")"
			elseif instr(ucase(cstr(printarray(7,j-1))),"DATEFORMAT")> 0 then
				strData = strData & ( dateformatconversion (cstr(printarray(3,j)),cstr(printarray(4,j-1))) )			
			elseif Right(ucase(cstr(printarray(7,j))), len(DateFormatExtension)) = DateFormatExtension then
			    strData = strData & ( dateformatconversion (cstr(printarray(3,j)), Null) )		
			end if%>
                <td class="alc-header">
                    <%
				if trim(strData) = "" then strData = " "
				BuildLabelDirect(strData).SetStyleSheet("label-nobold").Render()			
                    %>
                &nbsp;
                </td>
            </tr>
            <%Next%>
        </table>
    </form>
</body>
</html>

<script language="vbscript" runat="server">
'********************************************************************
'  Author      : Akash Dixit
'  Called From :
'  Parameters  : tree-XML,Indentation-numeric (for recursive use),indentationOrg-numeric (indentation of last parent for recursive use),DataElement as string for recursive use.
'  Returns     : sets the printarray so that it can be painted.
'  Description : function is called recursively,and moves to next parent's when the child is not found.

'  Revision History
'  Date			Author		   Description
'  27-Jan-2006 Akash Dixit     Original
'  04-Mar-2006 SR			   No need to lookup for the text whose length is more than 2000 chars (This causes sql problems)
'  11-Aug-2006 PS              IMPLEMENTED 4.2 Standards  
'********************************************************************
Sub Traverse(tree,indentation,indentationOrg,DataElement,dtdelementtype,profilerelease)
    if bSuccess = False then exit sub
    
    Dim i, nodes,indentation1,indentationOr,j,dataElementname,extratd,additional,dataelementTitle,vsql
    Dim oDesiredDataElementNode
    indentation1=indentation
    arraycount = Ccur(arraycount)+1
    ReDim Preserve printarray(9, Ccur(arraycount)+1)
    indentationOr = indentationOrg
    printarray(0,arraycount) = Ccur(indentation)
    
    if sErrorCFGM2Cache <> "" then 'this is set global variable for error set in GetCFGM2XMLFromCache 
        exit sub
    end if 
    
    If (tree.hasChildNodes()) Then
                
        Set oDesiredDataElementNode = DataElement.childnodes(0).childnodes(0).selectSingleNode("RPT_E2B[RPT_E2B_E2B="""& UCase(CStr(tree.nodename)) & """]")       	    
        if not oDesiredDataElementNode is nothing then        
		    if (ucase(tree.nodename) = oDesiredDataElementNode.childnodes(0).text )	then
			    printarray(7,arraycount) = oDesiredDataElementNode.childnodes(0).text
			    printarray(1,arraycount)  =	oDesiredDataElementNode.childnodes(1).text
			    printarray(5,arraycount) = oDesiredDataElementNode.childnodes(3).text
			    printarray(2,arraycount) = oDesiredDataElementNode.childnodes(2).text
			    dtdelementtype = Ccur(oDesiredDataElementNode.childnodes(4).text)
                printarray(8,arraycount) = oDesiredDataElementNode.childnodes(4).text 
		    end if
		else
			printarray(2,arraycount) = UCase(CStr(tree.nodename))
		end if
		
		nodes = tree.childNodes.length

		if Ccur(nodes)>1 then
			printarray(6,arraycount) ="1"
		else
			if tree.childNodes(0).hasChildNodes then
				printarray(6,arraycount) ="1"
			else
				printarray(6,arraycount) ="0"
			end if
		end if

        For i = 0 To Ccur(nodes) - 1
	        indentation1 = Ccur(indentation) +10
	        if sErrorCFGM2Cache <> "" then 'this is set global variable for error set in GetCFGM2XMLFromCache 
                exit sub
            end if 
	        traverse tree.childNodes(i),indentation1,indentationOr,DataElement,dtdelementtype,profilerelease 
        Next
    Else
		arraycount = arraycount-1
		if sLastNode <> printarray(7,arraycount) or tree.Text <> "" then
		    sLastNode = printarray(7,arraycount)
		    printarray(3,arraycount)= tree.Text
		
 		    If sHL7Profile="1" and (Ccur(dtdelementtype) = 7 or Ccur(dtdelementtype) = 10 or Ccur(dtdelementtype) = 11 ) then
                        printarray(4,arraycount) =GetNullFlavorDecodeValue(tree.Text)
		    elseif (Len(tree.Text) < 2000 and Ccur(dtdelementtype) > 1 and Ccur(dtdelementtype) < 8 ) then ' 1 is for Other Type Dtd Elements i.e. whichc do not exist in CFG_M2 Table
			    printarray(4,arraycount) = GetDecodedValue(trim(Ucase(tree.parentnode.tagname)), trim((tree.Text)), dtdelementtype,profilerelease, profileStr )            			
            elseif (dtdelementtype = 1) and (trim(Ucase(tree.parentnode.tagname))="DRUGPHARMADOSEFORMTERMID" or trim(Ucase(tree.parentnode.tagname))="DRUGROUTEOFADMINTERMID" or trim(Ucase(tree.parentnode.tagname))="DRUGPARROUTEOFADMINTERMID") then
                printarray(4,arraycount) = GetEDQMDesc(trim(Ucase(tree.parentnode.tagname)), trim(Ucase(tree.Text)))
		    else
                 If sHL7Profile="1" then
                    printarray(4,arraycount) =GetNullFlavorDecodeValue(tree.Text)
                else
                    printarray(4, arraycount) = ""
                end if
		    end if
        end if		
        indentation1 = indentationOr
    End If
 End Sub
 '********************************************************************
'  Author      : Akash Dixit
'  Called From : Main inline
'  Parameters  : XML file,path of XML file
'  Returns     :
'  Description : loads XML to display

'  Revision History
'  Date			Author		   Description
'  27-Jan-2006 Akash Dixit     Original
'********************************************************************
Function LoadXMLMessage (oXMLDom, docID)
    Dim oMessage, oOutMsg
    Dim lError, sError
    Dim sXMLReport
    Dim oNode, oReport
    Dim safetyreportNode, medicaldeviceinformationNode, managementinformationNode, node, nodes

    On Error Resume Next
    Call CreateMessage (oMessage, 300100294)    ' MID_db_app_get_user_cache
    Call SetXMLValueDirect (oMessage, "USER_CACHE_ID", docID)
    Set oOutMsg = ServiceRequest(oArgusSvr, oMessage, lError, sError) 
      
    If (lError = 0) Then
        Set oNode = oOutMsg.selectSingleNode("/MESSAGE/USER_CACHE_DATA")
		If IsNullOrEmpty(oNode) Then Exit Function
        Set oReport = oNode.childNodes(1)
        If IsNullOrEmpty(oReport) Then Exit Function
		sXMLReport = oReport.xml
        Set oXMLDom = Server.CreateObject ("MSXML2.DOMDocument.6.0")
        oXMLDom.loadxml (sXMLReport)
        if lReportFormId = 47 then
            Set safetyreportNode = oXMLDom.SelectSingleNode("/icsreport/safetyreport")
            Set nodes = oXMLDom.SelectNodes("/icsreport/safetyreport/safetyreportid | /icsreport/safetyreport/management_type | /icsreport/safetyreport/approvalnumber | " &_
                                            "/icsreport/safetyreport/year | /icsreport/safetyreport/reportreferencenumber")
            For each node in nodes
                if not node is nothing or not node is empty then
                    safetyreportNode.removeChild(node)
                end if
            Next

            Set managementinformationNode = oXMLDom.SelectSingleNode("/icsreport/safetyreport/dataroot/managementinformation")
            Set medicaldeviceinformationNode= oXMLDom.SelectSingleNode("/icsreport/safetyreport/dataroot/medicaldeviceinformation")
            For each node in nodes
                if not node is nothing or not node is empty then
                    if node.tagName = "approvalnumber" then
                        medicaldeviceinformationNode.appendChild(node)
                    else
                        managementinformationNode.appendChild(node)
                    end if
                end if
            Next
        end if
        
        LoadXMLMessage = True
    Else
        LoadXMLMessage = False
    End If
End Function

Function GetIndex(iLen)
    Dim i
    i = 1
    Do While (i < 41)
    If iLen <= (500 * i) Then
       GetIndex = i
       Exit Do
    End If
    i = i + 1
    Loop
End Function

Function GetNullFlavorDecodeValue(sNullflavor)
    Dim svalue
    If (glDisplayLang = cfCMN_LANG_JP) Then
        svalue = GetNullFlavorDecodeValueJa(sNullflavor)
    Else
        svalue = GetNullFlavorDecodeValueEn(sNullflavor)
    End If
    GetNullFlavorDecodeValue=svalue
end Function


Function GetNullFlavorDecodeValueEn(sNullflavor)
    Dim svalue
    if sNullflavor="[ASKU]"  then 
	   svalue = "ASKED BUT UNKNOWN"
    ElseIf sNullflavor="[MSK]"   then
	   svalue = "MASKED"
     ElseIf sNullflavor="[NINF]"  then
	   svalue = "NEGATIVE INFINITY"
     ElseIf sNullflavor="[NI]"    then
	   svalue = "NO INFORMATION"
     ElseIf sNullflavor="[NA]"   	then
	   svalue = "NOT APPLICABLE"
     ElseIf sNullflavor="[NASK]"  then
	   svalue = "NOT ASKED"
     ElseIf sNullflavor="[OTH]"   then
	   svalue = "OTHER"
     ElseIf sNullflavor="[PINF]"  then
	        svalue = "POSITIVE INFINITY"
     ElseIf sNullflavor="[UNK]"   then
	   svalue = "UNKNOWN"
     Else
     svalue = ""
    End if
    GetNullFlavorDecodeValueEn=svalue
end Function

Function GetNullFlavorDecodeValueJa(sNullflavor)
    Dim svalue
    if sNullflavor="[ASKU]"  then 
	   svalue = "不明"
    ElseIf sNullflavor="[MSK]"   then
	   svalue = "不明"
     ElseIf sNullflavor="[NINF]"  then
	   svalue = "NEGATIVE INFINITY"
     ElseIf sNullflavor="[NI]"    then
	   svalue = "不明"
     ElseIf sNullflavor="[NA]"   	then
	   svalue = "NOT APPLICABLE"
     ElseIf sNullflavor="[NASK]"  then
	   svalue = "不明"
     ElseIf sNullflavor="[OTH]"   then
	   svalue = "その他"
     ElseIf sNullflavor="[PINF]"  then
	        svalue = "POSITIVE INFINITY"
     ElseIf sNullflavor="[UNK]"   then
	   svalue = "不明"
     Else
     svalue = ""
    End if
    GetNullFlavorDecodeValueJa=svalue
end Function
</script>
<!-- #INCLUDE VIRTUAL="/Nav/DialogFooter_inc.asp" -->
