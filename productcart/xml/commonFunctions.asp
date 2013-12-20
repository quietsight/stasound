
<%
Function state_Change()
	if objXML.readyState=4 then
		if objXML.status=200 then
			pcv_CountCompleted=pcv_CountCompleted+1
		end if
	end if
End Function

Public Function pcf_Log(message)	
	pcv_tmpLogFile=server.MapPath("xml_tools_log.txt")
	pcv_tmpLogFile=pcv_tmpLogFile
    logFilename = pcv_tmpLogFile	
	Dim oFs
	Dim oTextFile
	Set oFs = Server.createobject("Scripting.FileSystemObject")
	Const ioMode = 8
	Set oTextFile = oFs.openTextFile(logFilename, ioMode, True)
	oTextFile.writeLine message
	oTextFile.close
	Set oTextFile = Nothing
	Set oFS = Nothing	
End Function

Function EncodeExtendedAsciiChars(pcv_str)
	pcv_strNew = ""	
	For i = 1 to Len(pcv_str)		
		j = Asc(Mid(pcv_str,i,1))	
		Select Case j
			Case (j>=0 and j<=31): pcv_strNew = pcv_strNew & ""	
			Case (j>=32 and j<=127): pcv_strNew = pcv_strNew & Mid(pcv_str,i,1)
			Case Else: pcv_strNew = pcv_strNew & "&#" & j & ";"
		End Select			
	Next	
	EncodeExtendedAsciiChars = pcv_strNew
End Function

Function HTMLDecode(pcv_str)
    Dim I
    pcv_str = Replace(pcv_str, "&quot;", Chr(34))
    pcv_str = Replace(pcv_str, "&lt;"  , Chr(60))
    pcv_str = Replace(pcv_str, "&gt;"  , Chr(62))
    pcv_str = Replace(pcv_str, "&amp;" , Chr(38))
    pcv_str = Replace(pcv_str, "&nbsp;", Chr(32))
    For I = 1 to 255
        pcv_str = Replace(pcv_str, "&#" & I & ";", Chr(I))
    Next
    HTMLDecode = pcv_str
End Function

Sub InitResponseDocument(Method)
	Set oXML=Server.CreateObject("MSXML2.DOMDocument"&scXML)
	Set oRoot = oXML.createNode(1,Method,"")
	oXML.appendChild(oRoot)
	Set oNode = oXML.createProcessingInstruction("xml", " version='1.0' encoding='"&xmlInit_encoding&"'") 
	oXML.insertBefore oNode, oXML.firstChild
End Sub

Function New_HTMLEncode(tmpValue)
Dim tmpValue1
	tmpValue1=""
	if tmpValue<>"" then
		tmpValue1=Server.HTMLEncode(tmpValue)
	end if
	New_HTMLEncode=tmpValue1
End Function

Function ConvertFromXMLDate(tmpDate)
Dim tmp1,tmp2
	tmp1=tmpDate
	tmp2=split(tmp1,"-")
	tmp1=CDate(tmp2(1) & "/" & tmp2(2) & "/" & tmp2(0))
	ConvertFromXMLDate=tmp1
End Function

Function ConvertToXMLDate(tmpDate)
Dim tmp1,tmp2,tmp3
	tmp1=CDate(tmpDate)
	tmp2=Year(tmp1)
	tmp3=Month(tmp1)
	if tmp3<10 then
		tmp3="0" & tmp3
	end if
	tmp2=tmp2 & "-" & tmp3
	tmp3=Day(tmp1)
	if tmp3<10 then
		tmp3="0" & tmp3
	end if
	tmp2=tmp2 & "-" & tmp3
	ConvertToXMLDate=tmp2
End Function

Sub GetXMLSettings()
Dim query,rs
	call opendb()
	
	cm_LogTurnOn=0
	cm_LogErrors=0
	cm_CaptureRequest=0
	cm_CaptureResponse=0
	cm_EnforceHTTPs=0
	
	query="SELECT pcXMLSet_Log,pcXMLSet_LogErrors,pcXMLSet_CaptureRequest,pcXMLSet_CaptureResponse,pcXMLSet_EnforceHTTPs FROM pcXMLSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		cm_LogTurnOn=rs("pcXMLSet_Log")
		cm_LogErrors=rs("pcXMLSet_LogErrors")
		cm_CaptureRequest=rs("pcXMLSet_CaptureRequest")
		cm_CaptureResponse=rs("pcXMLSet_CaptureResponse")
		cm_EnforceHTTPs=rs("pcXMLSet_EnforceHTTPs")
	end if
	set rs=nothing
	
	call closedb()

End Sub


Sub XMLcreateError(errorCode,errorString)
Dim tmpNode
	if errorCode<>"" or errorString<>"" then
		if statNode="" then
			Set statNode=oRoot.selectSingleNode(cm_requestStatus_name)
			if statNode is Nothing then
				Set statNode=oXML.createNode(1,cm_requestStatus_name,"")
				oRoot.appendChild(statNode)
			end if
			statNode.Text=cm_HaveErrors
		end if
		if eNode="" then
			Set eNode=oRoot.selectSingleNode(cm_errorList_name)
			if eNode is Nothing then
				Set eNode=oXML.createNode(1,cm_errorList_name,"")
				oRoot.appendChild(eNode)
			end if
		end if
		Set tmpNode=oXML.createNode(1,cm_errorCode_name,"")
		tmpNode.Text=errorCode
		eNode.appendChild(tmpNode)
		Set tmpNode=oXML.createNode(1,cm_errorDesc_name,"")
		tmpNode.Text=errorString
		eNode.appendChild(tmpNode)
		xmlHaveErrors=1
	end if
End Sub

Sub CheckHTTPHeaders()
Dim tmpPartner,tmpPartnerIP,IPTurnOn,rs,query,tmpHTTPs
	tmpPartner=Request.ServerVariables("HTTP_XML_AGENT")
	if UCase(tmpPartner)<>"PRODUCTCART XML PARTNER" then
		call XMLcreateError(110,cm_errorStr_110)
		call returnXML()
	end if
	
	IF cm_EnforceHTTPs=1 THEN
		tmpHTTPs=Request.ServerVariables("HTTPS")
		if UCase(tmpHTTPs)="OFF" then
			call XMLcreateError(127,cm_errorStr_127)
			call returnXML()
		end if
	END IF
	
	tmpPartnerIP=Request.ServerVariables("REMOTE_ADDR")
	IPTurnOn=0
	
	call opendb()
	
	query="SELECT pcXIP_id,pcXIP_IPAddr,pcXIP_TurnOn FROM pcXMLIPs;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		IPTurnOn=rs("pcXIP_TurnOn")
	end if
	set rs=nothing
	
	if IPTurnOn="1" then
		query="SELECT pcXIP_id FROM pcXMLIPs WHERE pcXIP_IPAddr LIKE '" & tmpPartnerIP & "';"
		set rs=connTemp.execute(query)
		if rs.eof then
			set rs=nothing
			call closedb()
			call XMLcreateError(126,cm_errorStr_126)
			call returnXML()
		end if
		set rs=nothing
	end if
	
	call closedb()
	
End Sub

Sub CheckValidXMLDocument()
	If iXML.parseError.errorCode <> 0 Then
		call XMLcreateError(101,cm_errorStr_101)
		call returnXML()
	End if
End Sub

Sub CheckValidXMLTag(tmpNode,ReqTag,valueType,tmpValue)
Dim tmp1,tmp2,tmp3,tmpName
	tmpName=tmpNode.NodeName
	tmp1=tmpNode.Text
	If (tmp1="") and (ReqTag=1) Then
		call XMLcreateError(111,cm_errorStr_111 & tmpName & cm_errorStr_111a)
		call returnXML()
	End if
	if trim(tmp1<>"") then
		Select Case valueType
		Case 0: 
			If Not IsNumeric(tmp1) then
				call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107b)
				call returnXML()
			else
				If Fix(tmp1)<>Cdbl(tmp1) then
					call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107b)
					call returnXML()
				end if
			end if
		Case 1: 
			If Not IsNumeric(tmp1) then
				call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107b)
				call returnXML()
			else
				If Fix(tmp1)<>Cdbl(tmp1) then
					call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107b)
					call returnXML()
				else
					If Cdbl(tmp1)<0 then
						call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107b & cm_errorStr_107d)
						call returnXML()
					end if
				end if
			end if
		Case 2:
			If Not IsNumeric(tmp1) then
				call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107c)
				call returnXML()
			end if
		Case 3:
			If Not IsNumeric(tmp1) then
				call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107c)
				call returnXML()
			else
				If Cdbl(tmp1)<0 then
					call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107c & cm_errorStr_107d)
					call returnXML()
				end if
			end if
		Case 4:
			tmp2=split(tmp1,"-")
			If ubound(tmp2)<>2 then
				call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107e)
				call returnXML()
			Else
				tmp3=tmp2(1) & "/" & tmp2(2) & "/" & tmp2(0)
				if not IsDate(tmp3) then
					call XMLcreateError(107,cm_errorStr_107 & tmpName & cm_errorStr_107a & tmp1 & cm_errorStr_107e)
					call returnXML()
				end if
			end if
		Case 5: 'string
		End Select
	
		if tmpValue<>"" then
			if (valueType<4) then
				if cdbl(tmp1)<cdbl(tmpValue) then
					call XMLcreateError(108,cm_errorStr_108 & tmpName & cm_errorStr_108a & tmp1 & cm_errorStr_108b & tmpValue)
					call returnXML()
				end if
			else
				if tmp1<tmpValue then
					call XMLcreateError(108,cm_errorStr_108 & tmpName & cm_errorStr_108a & tmp1 & cm_errorStr_108b & tmpValue)
					call returnXML()
				end if
			end if
		end if
	end if
End Sub

Function CheckExistTag(tagName)
Dim tmpNode
	Set tmpNode=iRoot.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		CheckExistTag=False
	Else
		CheckExistTag=True
	End if
End Function

Function CheckExistTagEx(parentNode,tagName)
Dim tmpNode
	Set tmpNode=parentNode.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		CheckExistTagEx=False
	Else
		CheckExistTagEx=True
	End if
End Function

Sub CheckRequiredXMLTag(tagName)
	Set tmpNode=iRoot.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		call XMLcreateError(102,cm_errorStr_102 & cm_errorStr_102a & tagName)
		call returnXML()
	End if
	If tmpNode.Text="" Then
		call XMLcreateError(102,cm_errorStr_102 & cm_errorStr_102b & tagName & cm_errorStr_102c)
		call returnXML()
	End if
End Sub

Sub CheckRequiredXMLTagEx(parentNode,tagName)
	Set tmpNode=parentNode.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		call XMLcreateError(102,cm_errorStr_102 & cm_errorStr_102a & parentNode.NodeName & "/" & tagName)
		call returnXML()
	End if
	If tmpNode.Text="" Then
		call XMLcreateError(102,cm_errorStr_102 & cm_errorStr_102b & parentNode.NodeName & "/" & tagName & cm_errorStr_102c)
		call returnXML()
	End if
End Sub

Sub CheckCommonRequiredXMLTags()
Dim tmpNodeValue,tmpProductCartResponse_name,query,rs1

	Call CheckRequiredXMLTag(cm_partnerID_name)
	cm_partnerID_ex=1
	cm_partnerID_value=getUserInput(tmpNode.Text,0)
	
	call opendb()
	query="SELECT pcXP_ID,pcXP_ExportAdmin FROM pcXMLPartners WHERE pcXP_PartnerID LIKE '" & cm_partnerID_value & "';"
	set rs1=connTemp.execute(query)
	if not rs1.eof then
		pcv_PartnerID=rs1("pcXP_ID")
		cm_ExportAdmin=rs1("pcXP_ExportAdmin")
		if IsNull(cm_ExportAdmin) OR cm_ExportAdmin="" then
			cm_ExportAdmin=0
		end if
	end if
	set rs1=nothing
	call closedb()

	Select Case iRoot.nodeName
		Case cm_SearchProductsRequest_name,cm_SearchCustomersRequest_name,cm_SearchOrdersRequest_name,cm_GetProductDetailsRequest_name,cm_GetCustomerDetailsRequest_name,cm_GetOrderDetailsRequest_name,cm_NewProductsRequest_name,cm_NewCustomersRequest_name,cm_NewOrdersRequest_name,cm_AddProductRequest_name,cm_AddCustomerRequest_name,cm_UpdateProductRequest_name,cm_UpdateCustomerRequest_name,cm_UndoRequest_name,cm_MarkAsExportedRequest_name:
		Case Else:
			call XMLcreateError(102,cm_errorStr_102d)
			call returnXML()
	End Select
	
	cm_methodName_ex=1
	cm_methodName_value=iRoot.nodeName
	
	Select Case cm_methodName_value
		Case cm_SearchProductsRequest_name:
			tmpProductCartResponse_name=cm_SearchProductsResponse_name
		Case cm_SearchCustomersRequest_name:
			tmpProductCartResponse_name=cm_SearchCustomersResponse_name
		Case cm_SearchOrdersRequest_name:
			tmpProductCartResponse_name=cm_SearchOrdersResponse_name
		Case cm_GetProductDetailsRequest_name:
			tmpProductCartResponse_name=cm_GetProductDetailsResponse_name
		Case cm_GetCustomerDetailsRequest_name:
			tmpProductCartResponse_name=cm_GetCustomerDetailsResponse_name
		Case cm_GetOrderDetailsRequest_name:
			tmpProductCartResponse_name=cm_GetOrderDetailsResponse_name	
		Case cm_NewProductsRequest_name:
			tmpProductCartResponse_name=cm_NewProductsResponse_name
		Case cm_NewCustomersRequest_name:
			tmpProductCartResponse_name=cm_NewCustomersResponse_name
		Case cm_NewOrdersRequest_name:
			tmpProductCartResponse_name=cm_NewOrdersResponse_name
		Case cm_AddProductRequest_name:
			tmpProductCartResponse_name=cm_AddProductResponse_name
		Case cm_AddCustomerRequest_name:
			tmpProductCartResponse_name=cm_AddCustomerResponse_name
		Case cm_UpdateProductRequest_name:
			tmpProductCartResponse_name=cm_UpdateProductResponse_name
		Case cm_UpdateCustomerRequest_name:
			tmpProductCartResponse_name=cm_UpdateCustomerResponse_name
		Case cm_UndoRequest_name:
			tmpProductCartResponse_name=cm_UndoResponse_name
		Case cm_MarkAsExportedRequest_name:
			tmpProductCartResponse_name=cm_MarkAsExportedResponse_name			
	End Select
	
	Set oNode=nothing
	Set oRoot=nothing
	Set oXML=nothing
	
	call InitResponseDocument(tmpProductCartResponse_name)	
	
	Call CheckRequiredXMLTag(cm_partnerPassword_name)
	cm_partnerPassword_ex=1
	cm_partnerPassword_value=getUserInput(tmpNode.Text,0)
	
	Call CheckRequiredXMLTag(cm_partnerKey_name)
	cm_partnerKey_ex=1
	cm_partnerKey_value=getUserInput(tmpNode.Text,0)
	
	if CheckExistTag(cm_callbackURL_name) then
		tmpNodeValue=iRoot.selectSingleNode(cm_callbackURL_name).Text
		if trim(tmpNodeValue)<>"" then
			if (Instr(ucase(tmpNodeValue),"HTTP://")<1) AND (Instr(ucase(tmpNodeValue),"HTTPS://")<1) then
				call XMLcreateError(107,cm_errorStr_107 & cm_callbackURL_name & cm_errorStr_107a & tmpNodeValue & cm_errorStr_107f)
				call returnXML()
			else
				cm_callbackURL_ex=1
				cm_callbackURL_value=tmpNodeValue
			end if
		end if
	end if
	
End Sub

Sub CheckValidPartner()
	Dim rs1,query,tmpPass,tmpPStatus
	on error resume next
	call opendb()
	tmpPass=EnDecrypt(cm_partnerPassword_value, scCrypPass)
	query="SELECT pcXP_ID,pcXP_Status FROM pcXMLPartners WHERE pcXP_PartnerID like '" & cm_partnerID_value & "' AND pcXP_Password like '" & tmpPass & "' AND pcXP_Key like '" & cm_partnerKey_value & "' AND pcXP_Removed=0;"
	set rs1=connTemp.execute(query)
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if rs1.eof then
		set rs1=nothing
		call closedb()
		call XMLcreateError(103,cm_errorStr_103)
		call returnXML()
	else
		pcv_PartnerID=rs1("pcXP_ID")
		tmpPStatus=Clng(rs1("pcXP_Status"))
	end if
	set rs1=nothing
	call closedb()
	
	Select Case tmpPStatus
		Case 0:
			call XMLcreateError(112,cm_errorStr_112)
			call returnXML()
		Case 2:
			call XMLcreateError(113,cm_errorStr_113)
			call returnXML()
		Case 3:
			call XMLcreateError(114,cm_errorStr_114)
			call returnXML()
	End Select
End Sub

Function CreateRequestRecord(tmpPartnerID,requestType,updatedID,backupFile,tmpUndo,resultCount,lastID,undoID)
	Dim query,rs1,backupFileStr
	Dim requestKey,TodayDate,i,myC,ReqExist
	on error resume next

	requestKey=""

	IF cm_LogTurnOn=1 THEN

	call opendb()
	DO
		requestKey=""
		For i=1 to 15
			Randomize
			myC=Fix(2*Rnd)
			Select Case myC
				Case 0: 
					Randomize
					requestKey=requestKey & Cstr(Fix(10*Rnd))
				Case 1: 
					Randomize
					requestKey=requestKey & Chr(Fix(26*Rnd)+65)		
			End Select		
		Next
	
		ReqExist=0
	
		query="SELECT pcXL_ID FROM pcXMLLogs WHERE pcXL_RequestKey LIKE '" & requestKey & "'" 
		set rs1=connTemp.execute(query)
		if Err.number<>0 then
			set rs1=nothing
			call closedb()
			call XMLcreateError(115,cm_errorStr_115)
			call returnXML()
		end if
		if not rs1.eof then
			ReqExist=1
		end if
		set rs1=nothing
	LOOP UNTIL ReqExist=0
	
	if backupFile=1 then
		backupFileStr=requestKey & ".txt"
	else
		backupFileStr=""
	end if
	
	Todaydate=Date()
	if SQL_Format="1" then
		Todaydate=Day(Todaydate)&"/"&Month(Todaydate)&"/"&Year(Todaydate)
	else
		Todaydate=Month(Todaydate)&"/"&Day(Todaydate)&"/"&Year(Todaydate)
	end if	
	if scDB="Access" then
		query="INSERT INTO pcXMLLogs (pcXP_id,pcXL_RequestKey,pcXL_RequestType,pcXL_UpdatedID,pcXL_BackupFile,pcXL_Undo,pcXL_ResultCount,pcXL_Date,pcXL_LastID,pcXL_UndoID) VALUES (" & tmpPartnerID & ",'" & requestKey & "'," & requestType & "," & updatedID & ",'" & backupFileStr & "'," & tmpUndo & "," & resultCount & ",#" & Todaydate & "#," & lastID & "," & undoID & ");"
	else
		query="INSERT INTO pcXMLLogs (pcXP_id,pcXL_RequestKey,pcXL_RequestType,pcXL_UpdatedID,pcXL_BackupFile,pcXL_Undo,pcXL_ResultCount,pcXL_Date,pcXL_LastID,pcXL_UndoID) VALUES (" & tmpPartnerID & ",'" & requestKey & "'," & requestType & "," & updatedID & ",'" & backupFileStr & "'," & tmpUndo & "," & resultCount & ",'" & Todaydate & "'," & lastID & "," & undoID & ");"
	end if	

	set rs1=connTemp.execute(query)
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	set rs1=nothing
	call closedb()
	
	END IF

	cm_requestKey_value=requestKey
	CreateRequestRecord=requestKey
	
End Function

Sub XMLCreateNode(parentNode,tmpNodeName,tmpValue)
Dim attNode
	Set attNode=oXML.createNode(1,tmpNodeName,"")
	if tmpValue<>"" then
		if (tmpValue=-1) and (tmpNodeName<>prdStock_name) then
			tmpValue=1
		end if
		attNode.Text=tmpValue
	end if
	parentNode.appendChild(attNode)
End Sub

Sub CheckUndoRequestTags()
	Dim query,rs
	
	Call CheckRequiredXMLTag(cm_requestKey_name)
	Set strNode=iRoot.selectSingleNode(cm_requestKey_name)
	call CheckValidXMLTag(strNode,1,5,"")
	cm_requestKey_ex=1
	cm_requestKey_value=getUserInput(tmpNode.Text,0)
	
	call opendb()

	query="SELECT pcXMLLogs.pcXL_id,pcXMLLogs.pcXL_RequestType,pcXMLLogs.pcXL_BackupFile,pcXMLLogs.pcXL_Undo FROM pcXMLPartners INNER JOIN pcXMLLogs ON pcXMLPartners.pcXP_ID=pcXMLLogs.pcXP_ID WHERE pcXMLLogs.pcXL_RequestKey LIKE '" & cm_requestKey_value & "' AND pcXMLPartners.pcXP_PartnerID LIKE '" & cm_partnerID_value & "';"
	set rs=connTemp.execute(query)
	
	if rs.eof then
		set rs=nothing
		call closedb()
		call XMLcreateError(116,cm_errorStr_116c & cm_requestKey_value)
		call returnXML()
	else
		xmlRequestID_value=rs("pcXL_id")
		xmlRequestKey_value=cm_requestKey_value
		xmlRequestType_value=rs("pcXL_RequestType")
		xmlBackup_value=rs("pcXL_BackupFile")
		xmlUndo_value=rs("pcXL_Undo")
		set rs=nothing
		call closedb()
	end if
	
	if (cint(xmlRequestType_value)<9) OR (cint(xmlRequestType_value)>12) then
		call XMLcreateError(123,cm_errorStr_123 & cm_requestKey_value & cm_errorStr_123a)
		call returnXML()
	end if
	
	if xmlUndo_value="1" then
		call XMLcreateError(124,cm_errorStr_124 & cm_requestKey_value & cm_errorStr_124a)
		call returnXML()
	end if
	
	if IsNull(xmlBackup_value) or xmlBackup_value="" then
		call XMLcreateError(125,cm_errorStr_125 & cm_requestKey_value & cm_errorStr_125a)
		call returnXML()
	end if
End Sub

Sub RunUndoRequest()
	dim rs, query, TempLine, tmpID, TempStr
	on error resume next

	Set fso = server.CreateObject("Scripting.FileSystemObject")
	findit = Server.MapPath("logs/" & xmlBackup_value)
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	if Err.number>0 then
		Set f=nothing
		Set fso=nothing
		call XMLcreateError(125,cm_errorStr_125 & cm_requestKey_value & cm_errorStr_125a)
		call returnXML()
	end if

	DO WHILE not f.AtEndofStream
		TempLine=f.Readline
	
		IF TempLine<>"" then
			TempStr=split(TempLine,chr(9))
			tmpID=TempStr(1)
			
			Select Case Ucase(TempStr(0))
			Case "DELPRD":
				call opendb()
				query="DELETE FROM Products WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM DProducts WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM pcGC WHERE pcGC_IDProduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM pcProductsOptions WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM options_optionsGroups WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM categories_products WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELCUST":
				call opendb()
				query="DELETE FROM Customers WHERE idcustomer=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM pcCustomerFieldsValues WHERE idcustomer=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM recipients WHERE idcustomer=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELRECI":
				call opendb()
				query="DELETE FROM recipients WHERE idRecipient=" & tmpID & " AND idcustomer=" & TempStr(2) & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELCATPRD":
				call opendb()
				query="DELETE FROM categories_products WHERE idproduct=" & tmpID & " AND idcategory=" & TempStr(2) & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELPRDGRP":
				call opendb()
				query="DELETE FROM pcProductsOptions WHERE idproduct=" & tmpID & " AND idOptionGroup=" & TempStr(2) & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELPRDOPT":
				call opendb()
				query="DELETE FROM options_optionsGroups WHERE idproduct=" & tmpID & " AND idOptionGroup=" & TempStr(2) & " AND idOption=" & TempStr(3) & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "UPDPRD":
				Call XMLUpdRecord("Products","idproduct",tmpID,TempLine)
			Case "UPDRECI":
				Call XMLUpdRecord("Recipients","idRecipient",tmpID,TempLine)
			Case "UPDCUST":
				Call XMLUpdRecord("Customers","idcustomer",tmpID,TempLine)
			Case "UPDPRDGRP":
				Call XMLUpdRecord("pcProductsOptions","pcProdOpt_ID",tmpID,TempLine)
			Case "UPDPRDOPT":
				Call XMLUpdRecord("options_optionsGroups","idoptoptgrp",tmpID,TempLine)
			Case "ADDDP":
				Call XMLAddRecord("DProducts","idproduct",tmpID,TempLine)
			Case "ADDGC":
				Call XMLAddRecord("pcGC","pcGC_IDProduct",tmpID,TempLine)
			End Select
		END IF 'TempLine<>""
	
	LOOP
	
	f.close
	Set fso=nothing
	
	call opendb()
	query="UPDATE pcXMLLogs SET pcXL_Undo=1 WHERE pcXL_id=" & xmlRequestID_value & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closedb()
	
	IF cm_LogTurnOn=1 THEN
		requestKey=CreateRequestRecord(pcv_PartnerID,13,0,0,0,0,0,xmlRequestID_value)
		cm_requestKey_value=requestKey
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	if xmlHaveErrors=0 then
		Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
		tmpNode.Text=cm_SuccessCode
		oRoot.appendChild(tmpNode)
	else
		oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_HalfSuccessCode
	end if

End Sub

Sub XMLAddRecord(tmpTable,tmpIDName,tmpIDvalue,tmpValueStr)
Dim query,rstemp,rs,query2,PreRecord1,PreRecord2,k

	call opendb()
	query2=split(tmpValueStr,chr(9))
	
	query="SELECT * FROM " & tmpTable & ";"
	set rstemp=conntemp.execute(query)
	
	IF not rstemp.eof THEN
	
		query="DELETE FROM " & tmpTable & " WHERE " & tmpIDName & "=" & tmpIDvalue & ";"
		set rs=conntemp.execute(query)
		set rs=nothing
			
		PreRecord1=""
		PreRecord2=""
	
		iCols = rstemp.Fields.Count
	    for k=1 to iCols-1
	    	if query2(k)<>"##" then
	    	if k=1 then
	    		PreRecord1=PreRecord1 & "(" & Rstemp.Fields.Item(k).Name 
	    		PreRecord2=PreRecord2 & "(" & query2(k)
	    	else
	    		PreRecord1=PreRecord1 & "," & Rstemp.Fields.Item(k).Name
	    		PreRecord2=PreRecord2 & "," & query2(k)
	    	end if
	    	end if
	    next
	
	    PreRecord1=PreRecord1 & ")"
	    PreRecord2=PreRecord2 & ")"
	    
		query="INSERT INTO " & tmpTable & " " & PreRecord1 & " VALUES " & PreRecord2 & ";"
		query=replace(query,"DuLTVDu",vbcrlf)
		query=replace(query,"##","' '")
		set rstemp=connTemp.execute(query)
	END IF
	set rstemp=nothing
	call closedb()
	
End Sub

Sub XMLUpdRecord(tmpTable,tmpIDName,tmpIDvalue,tmpValueStr)
Dim query,rstemp,query2,PreRecord1,k

	call opendb()
	query2=split(tmpValueStr,chr(9))
	
	query="SELECT * FROM " & tmpTable & " WHERE " & tmpIDName & "=" & tmpIDvalue & ";"
	set rstemp=conntemp.execute(query)
			
	IF not rstemp.eof THEN
		PreRecord1=""
		iCols = rstemp.Fields.Count
	    for k=1 to iCols-1
	    	if query2(k+1)<>"##" then
	    	if k=1 then
	    		PreRecord1=PreRecord1 & Rstemp.Fields.Item(k).Name & "=" & query2(k+1)
	    	else
	    		PreRecord1=PreRecord1 & "," & Rstemp.Fields.Item(k).Name & "=" & query2(k+1)
	    	end if
	    	end if
	    next
		query="UPDATE " & tmpTable & " SET " & PreRecord1 & " WHERE " & tmpIDName & "=" & tmpIDvalue & ";"
		query=replace(query,"DuLTVDu",vbcrlf)
		set rstemp=connTemp.execute(query)
	END IF
	set rstemp=nothing
	call closedb()
	
End Sub

Sub returnXML()
	Dim SendError,tmpStatus,objXML,tmp1,query,rs,tmp2,tmp3,requestKey,tmpNode
	On Error Resume Next
	IF (cm_LogTurnOn=1) AND (cm_LogErrors=1) AND (oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_HaveErrors) THEN
		if pcv_PartnerID<>"" then
		else
			pcv_PartnerID=0
		end if
		Select Case cm_methodName_value
		Case cm_SearchProductsRequest_name:
			tmp1=0
			tmp2=0
			tmp3=0
		Case cm_SearchCustomersRequest_name:
			tmp1=1
			tmp2=0
			tmp3=0
		Case cm_SearchOrdersRequest_name:
			tmp1=2
			tmp2=0
			tmp3=0
		Case cm_GetProductDetailsRequest_name:
			tmp1=3
			tmp2=prdID_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0
		Case cm_GetCustomerDetailsRequest_name:
			tmp1=4
			tmp2=custID_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0
		Case cm_GetOrderDetailsRequest_name:
			tmp1=5
			tmp2=ordID_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0	
		Case cm_NewProductsRequest_name:
			tmp1=6
			tmp2=0
			tmp3=0
		Case cm_NewCustomersRequest_name:
			tmp1=7
			tmp2=0
			tmp3=0
		Case cm_NewOrdersRequest_name:
			tmp1=8
			tmp2=0
			tmp3=0
		Case cm_AddProductRequest_name:
			tmp1=9
			tmp2=prdID_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0
		Case cm_AddCustomerRequest_name:
			tmp1=10
			tmp2=custID_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0
		Case cm_UpdateProductRequest_name:
			tmp1=11
			tmp2=prdID_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0
		Case cm_UpdateCustomerRequest_name:
			tmp1=12
			tmp2=custID_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0
		Case cm_UndoRequest_name:
			tmp1=13
			tmp2=0
			tmp3=xmlRequestID_value
			if tmp3<>"" then
			else
				tmp3=0
			end if
		Case cm_MarkAsExportedRequest_name:
			tmp1=14
			tmp2=ExportedFlag_value
			if tmp2<>"" then
			else
				tmp2=0
			end if
			tmp3=0
		Case Else:
			tmp1=-1
			tmp2=0
			tmp3=0
		End Select
		requestKey=CreateRequestRecord(pcv_PartnerID,tmp1,tmp2,0,0,0,0,tmp3)
		cm_requestKey_value=requestKey
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF 
	SendError=0
	if cm_callbackURL_ex=1 then
		Set objXML = Server.CreateObject("MSXML2.serverXMLHTTP"&scXML)
		objXML.open "POST", cm_callbackURL_value, false
		objXML.setRequestHeader "CONTENT_TYPE", xmldocument_encoding
		objXML.send(oXML.xml)
		if Err.number<>0 then
			SendError=1
			Err.number=0
			Err.description=""
		end if
		tmpStatus=objXML.Status
		set objXML=nothing
		if (tmpStatus<>200) then
			SendError=1
		end if
	end if
	if cm_callbackURL_ex<>1 OR SendError=1 then
		if SendError=1 then
			call XMLcreateError(109,cm_errorStr_109)
			if oRoot.selectSingleNode(cm_requestStatus_name).Text<>cm_HaveErrors then
				oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_HalfSuccessCode
			end if
		end if
		Response.ContentType = xmldocument_encoding
		response.write oXML.xml
	end if
	
	IF cm_LogTurnOn=1 THEN
		if oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_HalfSuccessCode OR oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_SuccessCode then
			if oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_HalfSuccessCode then
				tmp1="pcXL_Status=2"
			else
				tmp1="pcXL_Status=1"
			end if
			IF cm_CaptureRequest=1 THEN
				if tmp1<>"" then
					tmp1=tmp1 & ","
				end if
				tmp1=tmp1 & "pcXL_RequestXML='" & pcf_SanitizeXML(iXML.xml) & "'"
			END IF
			IF cm_CaptureResponse=1 THEN
				if tmp1<>"" then
					tmp1=tmp1 & ","
				end if
				tmp1=tmp1 & "pcXL_ResponseXML='" & pcf_SanitizeXML(oXML.xml) & "'"
			END IF
			if tmp1<>"" then
				call opendb()
				query="UPDATE pcXMLLogs SET " & tmp1 & " WHERE pcXL_RequestKey like '" & cm_requestKey_value & "';"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			end if
		else
			IF cm_LogErrors=1 AND cm_requestKey_value<>"" THEN
				tmp1="pcXL_Status=0"
				IF cm_CaptureRequest=1 THEN
					if tmp1<>"" then
						tmp1=tmp1 & ","
					end if
					tmp1=tmp1 & "pcXL_RequestXML='" & pcf_SanitizeXML(iXML.xml) & "'"
				END IF
				IF cm_CaptureResponse=1 THEN
					if tmp1<>"" then
						tmp1=tmp1 & ","
					end if
					tmp1=tmp1 & "pcXL_ResponseXML='" & pcf_SanitizeXML(oXML.xml) & "'"
				END IF
				if tmp1<>"" then
					call opendb()
					query="UPDATE pcXMLLogs SET " & tmp1 & " WHERE pcXL_RequestKey like '" & cm_requestKey_value & "';"
					set rs=connTemp.execute(query)
					set rs=nothing
					call closedb()
				end if
			END IF
		end if
	END IF
	
	Set oNode=nothing
	Set oRoot=nothing
	Set oXML=nothing
	response.end
End Sub


Sub CheckMarkAsExportedRequestTags()
	On Error Resume Next
	Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	Call CheckRequiredXMLTag("//" &cm_requests_name & "/" & cm_request_name & "/" & ExportedFlag_name)
	Call CheckRequiredXMLTag("//" &cm_requests_name & "/" & cm_request_name & "/" & ExportedID_name)
	Call CheckRequiredXMLTag("//" &cm_requests_name & "/" & cm_request_name & "/" & ExportedIDType_name)	
	Set rNode=iRoot.selectSingleNode("//" &cm_requests_name & "/" & cm_request_name)	
	Set ChildNodes = rNode.childNodes	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		Select Case tmpNodeName
			Case ExportedFlag_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					ExportedFlag_ex=1
					ExportedFlag_value=tmpNodeValue
				end if
			Case ExportedID_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					ExportedID_ex=1
					ExportedID_value=tmpNodeValue
				end if
			Case ExportedIDType_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					ExportedIDType_ex=1
					ExportedIDType_value=tmpNodeValue
				end if
			Case Else:
				call XMLcreateError(105, tmpNodeName)
				call returnXML()
		End Select
	Next
End Sub



Sub RunMarkAsExportedRequest()
	On Error Resume Next
	Dim query,rs,custNode,i,pcArr,pcv_HaveRecords,attNode,subNode,queryQ,rsQ,tmpExportedFlag	
	tmpExportedFlag=0
	
	IF cm_LogTurnOn=1 THEN
		requestKey=CreateRequestRecord(pcv_PartnerID,14,ExportedID_value,0,0,0,0,0)
		cm_requestKey_value=requestKey
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
	tmpNode.Text=cm_SuccessCode
	oRoot.appendChild(tmpNode)
		
	IF ExportedFlag_value=1 THEN	
		call opendb()
		queryQ="SELECT pcXEL_ExportedID FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=" & ExportedIDType_value & " AND pcXEL_ExportedID=" & ExportedID_value & ";"
		set rsQ=connTemp.execute(queryQ)
		if not rsQ.eof then
			tmpExportedFlag=1
		else
			queryQ="INSERT INTO pcXMLExportLogs (pcXP_ID,pcXEL_ExportedID,pcXEL_IDType) VALUES (" & pcv_PartnerID & "," & ExportedID_value & "," & ExportedIDType_value & ");"
			set rsQ=connTemp.execute(queryQ)
			tmpExportedFlag=1
		end if
		set rsQ=nothing
		call closedb()		
	ELSE  	
		call opendb()
		queryQ="Delete FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=" & ExportedIDType_value & " AND pcXEL_ExportedID=" & ExportedID_value & ";"
		set rsQ=connTemp.execute(queryQ)
		set rsQ=nothing
		'if error.number<>0 then
		'	call XMLcreateError(error.number, err.description)
		'	call returnXML()
		'end if
		call closedb()		
	END IF
	
	Set attNode=oXML.createNode(1,ExportedFlag_name,"")
	attNode.Text=New_HTMLEncode(tmpExportedFlag)
	oRoot.appendChild(attNode)
	
	Set attNode=oXML.createNode(1,ExportedID_name,"")
	attNode.Text=New_HTMLEncode(ExportedID_value)
	oRoot.appendChild(attNode)
	
	Set attNode=oXML.createNode(1,ExportedIDType_name,"")
	attNode.Text=New_HTMLEncode(ExportedIDType_value)
	oRoot.appendChild(attNode)
	
	Set pXML1=Server.CreateObject("MSXML2.DOMDocument"&scXML)
	pXML1.async=false
	pXML1.load(oXML)
	If (pXML1.parseError.errorCode <> 0) Then	
		Set oXML=nothing
		call InitResponseDocument(cm_MarkAsExportedResponse_name)
		call XMLcreateError(pXML1.parseError.errorCode, pXML1.parseError.reason)
		call returnXML()
	End If
	set pXML1 = nothing
	
End Sub



Function UpdateExportFlag(pExportedFlag,pExportedID,pIDType)
	On Error Resume Next						
	Call InitResponseDocument(cm_MarkAsExportedRequest_name)					
	Call XMLCreateNode(oRoot,cm_partnerID_name,tmp_PartnerID)
	Call XMLCreateNode(oRoot,cm_partnerPassword_name,tmp_PartnerPass)
	Call XMLCreateNode(oRoot,cm_partnerKey_name,tmp_PartnerKey)
	Set tmpNode = oXML.createNode(1,cm_requests_name,"")
	oRoot.appendChild(tmpNode)
	Set tmpNode1 = oXML.createNode(1,cm_request_name,"")
	tmpNode.appendChild(tmpNode1)
	Set rNode1 = oXML.createNode(1,ExportedFlag_name,"")
	tmpNode1.appendChild(rNode1)	
	rNode1.Text=pExportedFlag
	Set rNode2 = oXML.createNode(1,ExportedID_name,"")
	tmpNode1.appendChild(rNode2)	
	rNode2.Text=pExportedID
	Set rNode3 = oXML.createNode(1,ExportedIDType_name,"")
	tmpNode1.appendChild(rNode3)
	rNode3.Text=pIDType								
	Set objXML2=Server.CreateObject("MSXML2.serverXMLHTTP"&scXML)
	objXML2.setTimeouts lResolve, lConnect, lSend, lReceive
	objXML2.open "POST",ProductCartXMLServer, True
	objXML2.setRequestHeader "XML-Agent", "ProductCart XML Partner"
	objXML2.send(oXML.xml)
	objXML2.waitForResponse(6)
	Set objXML2=nothing
	Set oRoot=nothing
	Set oXML=nothing							
End Function



Public Function pcf_IsResponseGood()
	On Error Resume Next
	
	If objXML.readyState <> 4  Then
		TransactionReady = objXML.waitForResponse(5)
		If TransactionReady = False Then
			pcf_IsResponseGood=False
			objXML.Abort
			Exit Function
		End If
	End If  	

	If Err.Number <> 0 then
		pcf_IsResponseGood=False
		objXML.Abort
		Exit Function
	Else
		If (objXML.readyState <> 4) Or (objXML.Status <> 200) Then
			pcf_IsResponseGood=False
			objXML.Abort
			Exit Function
		Else
			pcf_IsResponseGood=True
			Exit Function
		End If	
	End If		

	If Err.Number <> 0 then
		pcf_IsResponseGood=False
		objXML.Abort
		Exit Function
	End If
	
	On Error Goto 0
End Function



Public Function pcf_SanitizeXML(ObjXML)
	On Error Resume Next	
	ObjXML=replace(ObjXML,"'","''")
	ObjXML=replace(ObjXML,"""","""""")
	pcf_SanitizeXML=ObjXML
	On Error Goto 0
End Function



Public Function pcf_SanitizeNULL(NULLValue, DefaultValue)
	On Error Resume Next	
	if isNULL(NULLValue)=True then pcf_SanitizeNULL = DefaultValue
	On Error Goto 0
End Function
%>
