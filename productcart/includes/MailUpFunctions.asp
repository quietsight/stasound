<%'MailUp Integration Functions
Dim srvXmlHttp
Dim MU_ErrMsg, MU_XmlDoc

Public Function pcf_IsResponseGood()
	On Error Resume Next
	
	If srvXmlHttp.readyState <> 4  Then
		if not IsNull(MaxRequestTime) then
			TransactionReady = srvXmlHttp.waitForResponse(MaxRequestTime)
		else
			TransactionReady = srvXmlHttp.waitForResponse(5)
		end if
		If TransactionReady = False Then
			pcf_IsResponseGood=False
			srvXmlHttp.Abort
			Exit Function
		End If
	End If  	

	If Err.Number <> 0 then
		pcf_IsResponseGood=False
		srvXmlHttp.Abort
		Exit Function
	Else
		If (srvXmlHttp.readyState <> 4) Then
			pcf_IsResponseGood=False
			srvXmlHttp.Abort
			Exit Function
		Else
			pcf_IsResponseGood=True
			Exit Function
		End If	
	End If		

	If Err.Number <> 0 then
		pcf_IsResponseGood=False
		srvXmlHttp.Abort
		Exit Function
	End If
	
	On Error Goto 0
End Function

Function ConnectServer(tmpURL,tmpMethod,tmpContentType,tmpSOAPHead,tmpData)
Dim rersult1,tmpStatus

on error resume next
Set srvXmlHttp=nothing

IF StopHTTPRequests<>"1" THEN
	Set srvXmlHttp = server.createobject("MSXML2.serverXMLHTTP"&scXML)
	if not IsNull(MaxRequestTime) then
		srvXmlHttp.open tmpMethod, tmpURL, True
	else
		srvXmlHttp.open tmpMethod, tmpURL, False
	end if
	srvXmlHttp.setRequestHeader "Content-Type", tmpContentType
	if tmpSOAPHead<>"" then
	srvXmlHttp.setRequestHeader "SOAPAction", tmpSOAPHead
	end if
	srvXmlHttp.send tmpData
	
	if not IsNull(MaxRequestTime) then
		if pcf_IsResponseGood()=False then
			ConnectServer="TIMEOUT"
			exit function
			StopHTTPRequests=1
		end if
	end if
	
	result1 = srvXmlHttp.responseText
	
	if err.number<>0 then
		err.number=0
		err.description=""
		ConnectServer="ERROR"
		if not IsNull(MaxRequestTime) then
			StopHTTPRequests=1
		end if
		set srvXmlHttp=nothing
	else
		tmpStatus=srvXmlHttp.Status
		set srvXmlHttp=nothing
		if (tmpStatus<>200) then
			if result1<>"" then
				if Instr(result1,"ReturnCode")=0 AND (not IsNumeric(result1)) then
					ConnectServer="ERROR"
					if not IsNull(MaxRequestTime) then
						StopHTTPRequests=1
					end if
				else
					ConnectServer=result1
				end if
			else
				ConnectServer="ERROR"
				if not IsNull(MaxRequestTime) then
					StopHTTPRequests=1
				end if
			end if
		else
			ConnectServer=result1
		end if
	end if
ELSE
	ConnectServer="TIMEOUT"
END IF
End Function

Function FindReturnCode(tmpXML)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"ReturnCode>")>0 then
			tmp1=split(tmpXML,"ReturnCode>")
			tmp2=replace(tmp1(1),"</","")
			FindReturnCode=Clng(tmp2)
		else
			if Instr(tmpXML,"ReturnCode&gt;")>0 then
				tmp1=split(tmpXML,"ReturnCode&gt;")
				tmp2=replace(tmp1(1),"&lt;/","")
				FindReturnCode=Clng(tmp2)
			else
				FindReturnCode=-3333
			end if
		end if
	else
		FindReturnCode=-3333
	end if
End Function

Function FindStatusCode(tmpXML)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"StatusCode>")>0 then
			tmp1=split(tmpXML,"StatusCode>")
			tmp2=replace(tmp1(1),"</","")
			FindStatusCode=Clng(tmp2)
		else
			if Instr(tmpXML,"StatusCode&gt;")>0 then
				tmp1=split(tmpXML,"StatusCode&gt;")
				tmp2=replace(tmp1(1),"&lt;/","")
				FindStatusCode=Clng(tmp2)
			else
				FindStatusCode=-3333
			end if
		end if
	else
		FindStatusCode=-3333
	end if
End Function

Function FindConfirmationSent(tmpXML)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"ConfirmationSent>")>0 then
			tmp1=split(tmpXML,"ConfirmationSent>")
			tmp2=replace(tmp1(1),"</","")
			FindConfirmationSent=tmp2
		else
			if Instr(tmpXML,"ConfirmationSent&gt;")>0 then
				tmp1=split(tmpXML,"ConfirmationSent&gt;")
				tmp2=replace(tmp1(1),"&lt;/","")
				FindConfirmationSent=tmp2
			else
				FindConfirmationSent=-3333
			end if
		end if
	else
		FindConfirmationSent=-3333
	end if
End Function

Function RegisterAcc(APIUser,APIPass,tmpURL)
Dim tmp1,result,tmpCode
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp1=tmp1 & "frontend/WSActivation.aspx?usr=" & APIUser & "&pwd=" & APIPass & "&nl_url=" & tmpURL & "&ws_name=" & Server.URLEncode("WSMailupImport")
	result=ConnectServer(tmp1,"GET","application/x-www-form-urlencoded","","")
	IF result="ERROR" or result="TIMEOUT" THEN
		RegisterAcc="0"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=FindReturnCode(result)
		MU_ErrMsg=""
		if tmpCode="0" then
			RegisterAcc=1
		else
			RegisterAcc="0"
			Select Case tmpCode
				Case "-2": MU_ErrMsg="'ws name' has not been specified"
				Case "-4": MU_ErrMsg="'user name' has not been specified"
				Case "-8": MU_ErrMsg="'password' has not been specified"
				Case "-16": MU_ErrMsg="'nl url' has not been specified"
				Case "-1000": MU_ErrMsg="unrecognized error"
				Case "-1001": MU_ErrMsg="the account is not valid"
				Case "-1002": MU_ErrMsg="the password is not valid"
				Case "-1003": MU_ErrMsg="suspended account"
				Case "-1004": MU_ErrMsg="inactive account"
				Case "-1005": MU_ErrMsg="expired account"
				Case "-1006": MU_ErrMsg="the web service is not enabled"
				Case "-1007": MU_ErrMsg="the web service is not active"
				Case "-1008": MU_ErrMsg="the web service is already active"
				Case "-1009": MU_ErrMsg="web service activation error"
				Case "-1010": MU_ErrMsg="IP registration error"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
End Function

Sub UpdateMUList(MU_XmlDoc)
Dim tmpXML,tmpXML1,iXML,iRoot
Dim rs,query,tmpListID,tmpListGuid,tmpListName

	tmpXML=split(MU_XmlDoc,"mailupMessage&gt;")
	tmpXML1="&lt;mailupMessage&gt;" & tmpXML(1) & "mailupMessage&gt;"
	tmpXML=replace(tmpXML1,"&lt;","<")
	tmpXML=replace(tmpXML,"&gt;",">")
	
	Set iXML=Server.CreateObject("MSXML2.DOMDocument")
	iXML.async=false
	iXML.loadXML(tmpXML)
	
	If iXML.parseError.errorCode=0 Then
		Set iRoot=iXML.documentElement
		Set parentNode=iRoot.selectSingleNode("mailupBody/Lists")
		if parentNode.hasChildNodes then
			query="UPDATE pcMailUpLists SET pcMailUpLists_Removed=1;"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
		Set ChildNodes = parentNode.childNodes
		For Each strNode In ChildNodes
			Set NodeAttrs = strNode.attributes
			tmpListID=0
			tmpListGuid=0
			tmpListName=""
			For Each intAtt in NodeAttrs
				if ucase(intAtt.name)="IDLIST" then
					tmpListID=intAtt.text
				end if
				if ucase(intAtt.name)="LISTGUID" then
					tmpListGuid=intAtt.text
				end if
				if ucase(intAtt.name)="LISTNAME" then
					tmpListName=intAtt.text
					if tmpListName<>"" then
						tmpListName=replace(tmpListName,"'","''")
					end if
				end if
			Next
			if tmpListID<>"0" AND tmpListGuid<>"0" then
				query="SELECT pcMailUpLists_ID FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & tmpListID & " AND pcMailUpLists_ListGuid='" & tmpListGuid & "';"
				set rs=connTemp.execute(query)
				if not rs.eof then
					query="UPDATE pcMailUpLists SET pcMailUpLists_Removed=0 WHERE pcMailUpLists_ListID=" & tmpListID & " AND pcMailUpLists_ListGuid='" & tmpListGuid & "';"
					set rs=connTemp.execute(query)
				else
					query="INSERT INTO pcMailUpLists (pcMailUpLists_ListID,pcMailUpLists_ListGuid,pcMailUpLists_ListName,pcMailUpLists_Active) VALUES (" & tmpListID & ",'" & tmpListGuid & "','" & tmpListName & "',0);"
					set rs=connTemp.execute(query)
				end if
				set rs=nothing
				query="SELECT pcMailUpLists_ID FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & tmpListID & ";"
				set rs=connTemp.execute(query)
				tmpListIdx=rs("pcMailUpLists_ID")
				set rs=nothing
					
				Set GroupsNodes = strNode.childNodes
				For Each GroupsNode in GroupsNodes
					Set SubNodes=GroupsNode.childNodes
					For Each subNode in SubNodes
						Set NodeAttrs = subNode.attributes
						tmpGrpID=0
						tmpGrpName=""
						For Each intAtt in NodeAttrs
							if ucase(intAtt.name)="IDGROUP" then
								tmpGrpID=intAtt.text
							end if
							if ucase(intAtt.name)="GROUPNAME" then
								tmpGrpName=intAtt.text
								if tmpGrpName<>"" then
									tmpGrpName=replace(tmpGrpName,"'","''")
								end if
							end if
						Next
						if tmpGrpID<>"0" then
							query="SELECT pcMailUpGroups_ID FROM pcMailUpGroups WHERE pcMailUpLists_ID=" & tmpListIdx & " AND pcMailUpGroups_GroupID=" & tmpGrpID & ";"
							set rs=connTemp.execute(query)
							if not rs.eof then
								query="UPDATE pcMailUpGroups SET pcMailUpGroups_GroupName='" & tmpGrpName & "' WHERE pcMailUpLists_ID=" & tmpListIdx & " AND pcMailUpGroups_GroupID=" & tmpGrpID & ";"
								set rs=connTemp.execute(query)
							else
								query="INSERT INTO pcMailUpGroups (pcMailUpLists_ID,pcMailUpGroups_GroupID,pcMailUpGroups_GroupName) VALUES (" & tmpListIdx & "," & tmpGrpID & ",'" & tmpGrpName & "');"
								set rs=connTemp.execute(query)
							end if
							set rs=nothing
						end if
					Next
				Next
			end if
		Next
	End if
End Sub

Function GetMUList(APIUser,APIPass,tmpURL)
Dim tmp1,tmpData,result,tmpCode
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp1=tmp1 & "Services/WSMailupImport.asmx"
	tmpData=""
	tmpData=tmpData & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbcrlf
	tmpData=tmpData & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & vbcrlf
	tmpData=tmpData & "<soap:Header>" & vbcrlf
	tmpData=tmpData & "<Authentication xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<User>" & APIUser & "</User>" & vbcrlf
	tmpData=tmpData & "<Password>" & APIPass & "</Password>" & vbcrlf
	tmpData=tmpData & "</Authentication>" & vbcrlf
	tmpData=tmpData & "</soap:Header>" & vbcrlf
	tmpData=tmpData & "<soap:Body>" & vbcrlf
	tmpData=tmpData & "<GetNlLists xmlns=""http://ws.mailupnet.it/"" />" & vbcrlf
	tmpData=tmpData & "</soap:Body>" & vbcrlf
	tmpData=tmpData & "</soap:Envelope>"
	result=ConnectServer(tmp1,"POST","text/xml; charset=utf-8","http://ws.mailupnet.it/GetNlLists",tmpData)
	IF result="ERROR" or result="TIMEOUT" THEN
		GetMUList="0"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=FindReturnCode(result)
		MU_ErrMsg=""
		MU_XmlDoc=""
		if tmpCode="0" then
			GetMUList="1"
			MU_XmlDoc=result
			Call UpdateMUList(MU_XmlDoc)
		else
			GetMUList="0"
			Select Case tmpCode
				Case "-1000": MU_ErrMsg="unrecognized error"
				Case "-1001": MU_ErrMsg="the account is not valid"
				Case "-1002": MU_ErrMsg="the password is not valid"
				Case "-1003": MU_ErrMsg="suspended account"
				Case "-1004": MU_ErrMsg="inactive account"
				Case "-1005": MU_ErrMsg="expired account"
				Case "-1006": MU_ErrMsg="the web service is not enabled"
				Case "-1007": MU_ErrMsg="the web service is not active"
				Case "-1011": MU_ErrMsg="IP is not registered"
				Case "-1012": MU_ErrMsg="IP is registered but access is denied"
				Case "-200": MU_ErrMsg="unrecognized error"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
End Function

Function RegisterSingleUser(CustID,CustEmail,ListID,tmpURL,tmpAuto)
Dim tmp1,result,tmpCode,i,j,k,intCount
Dim tmpField,tmpValue,CustEmailDB
Dim rs,query,dtTodaysDate
	IF tmpAuto="1" THEN
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	query="SELECT [name],lastname,customerCompany,email FROM Customers WHERE idcustomer=" & CustID & ";"
	set rs=connTemp.execute(query)
	tmpField=""
	tmpValue=""
	if not rs.eof then
		if rs("name")<>"" then
			tmpField="campo1"
			tmpValue=Server.URLEncode(rs("name"))
		end if
		if rs("lastname")<>"" then
			if tmpField<>"" then
				tmpField=tmpField & ";"
				tmpValue=tmpValue & ";"
			end if
			tmpField=tmpField & "campo2"
			tmpValue=tmpValue & Server.URLEncode(rs("lastname"))
		end if
		if rs("customerCompany")<>"" then
			if tmpField<>"" then
				tmpField=tmpField & ";"
				tmpValue=tmpValue & ";"
			end if
			tmpField=tmpField & "campo3"
			tmpValue=tmpValue & Server.URLEncode(rs("customerCompany"))
		end if
		CustEmailDB=rs("email")
		if CustEmail<>"" then
			CustEmailDB=CustEmail
		end if
	end if
	set rs=nothing
	
	tmp1=tmp1 & "frontend/xmlSubscribe.aspx?email=" & CustEmailDB & "&list=" & ListID
	if tmpField<>"" then
	tmp1=tmp1 & "&csvFldNames=" & tmpField & "&csvFldValues=" & tmpValue
	end if
	if pcMU_sendconfirm="1" then
		tmp1=tmp1 & "&confirm=true"
	end if
	if pcMU_sendconfirm="0" then
		tmp1=tmp1 & "&confirm=false"
	end if
	result=ConnectServer(tmp1,"GET","application/x-www-form-urlencoded","","")
	IF result="ERROR" OR result="TIMEOUT" THEN
		RegisterSingleUser="0"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=Clng(result)
		MU_ErrMsg=""
		if tmpCode="0" or tmpCode="3" then
			RegisterSingleUser=1
		else
			RegisterSingleUser="0"
			Select Case tmpCode
				Case "1": MU_ErrMsg="Generic error"
				Case "2": MU_ErrMsg="E-mail address not valid"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
	ELSE
		RegisterSingleUser="0"
	END IF
	
	query="SELECT pcMailUpLists_ID FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & ListID & ";"
	set rs=connTemp.execute(query)
	ListIDidx=0
	if not rs.eof then
		ListIDidx=rs("pcMailUpLists_ID")
	end if
	set rs=nothing
	
	query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
	set rs=connTemp.execute(query)
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if not rs.eof then
		if scDB="SQL" then
			query="UPDATE pcMailUpSubs SET idCustomer=" & CustID & ",pcMailUpLists_ID=" & ListIDidx & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=" & Abs(Clng(RegisterSingleUser)-1) & ",pcMailUpSubs_Optout=0 WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
		else
			query="UPDATE pcMailUpSubs SET idCustomer=" & CustID & ",pcMailUpLists_ID=" & ListIDidx & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=" & Abs(Clng(RegisterSingleUser)-1) & ",pcMailUpSubs_Optout=0 WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
		end if
	else
		if scDB="SQL" then
			query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & CustID & "," & ListIDidx & ",'" & dtTodaysDate & "'," & Abs(Clng(RegisterSingleUser)-1) & ",0);"
		else
			query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & CustID & "," & ListIDidx & ",#" & dtTodaysDate & "#," & Abs(Clng(RegisterSingleUser)-1) & ",0);"
		end if
	end if
	set rs=nothing
	set rs=connTemp.execute(query)
	set rs=nothing
End Function

Function UpdUserReg(CustID,CustEmail,ListID,ListGuid,tmpURL,tmpAuto)
Dim tmp1,tmp2,result,tmpCode,i,j,k,intCount
Dim rs,query,dtTodaysDate
	IF tmpAuto="1" THEN
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp2=tmp1
	tmp1=tmp1 & "frontend/xmlChkSubscriber.aspx?listguid=" & ListGuid & "&list=" & ListID & "&email=" & CustEmail
	result=ConnectServer(tmp1,"GET","application/x-www-form-urlencoded","","")
	IF result="ERROR" OR result="TIMEOUT"  THEN
		UpdUserReg="0"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=Clng(result)
		MU_ErrMsg=""
		if tmpCode="1" then
			tmpRes=RegisterSingleUser(CustID,CustEmail,ListID,tmpURL,tmpAuto)
			UpdUserReg=tmpRes
		else
			if tmpCode="3" then
				tmpRes=RegisterSingleUser(CustID,CustEmail,ListID,tmpURL,tmpAuto)
				UpdUserReg=tmpRes
				tmp1=tmp2
				tmp1=tmp1 & "frontend/xmlUpdSubscriber.aspx?listguid=" & ListGuid & "&list=" & ListID & "&email=" & CustEmail
				result=ConnectServer(tmp1,"GET","application/x-www-form-urlencoded","","")
				IF result="ERROR" OR result="TIMEOUT" THEN
					UpdUserReg="0"
					MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
				ELSE
					tmpCode=Clng(result)
					MU_ErrMsg=""
					if tmpCode="0" then
						UpdUserReg=1
					else
						UpdUserReg="0"
						Select Case tmpCode
							Case "1": MU_ErrMsg="Generic error"
							Case Else: MU_ErrMsg="Error Code: " & tmpCode
						End Select
					end if
				END IF
			else
				UpdUserReg=1
			end if
		end if
	END IF
	ELSE
		UpdUserReg="0"
	END IF
	
	query="SELECT pcMailUpLists_ID FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & ListID & ";"
	set rs=connTemp.execute(query)
	ListIDidx=0
	if not rs.eof then
		ListIDidx=rs("pcMailUpLists_ID")
	end if
	set rs=nothing
	
	query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
	set rs=connTemp.execute(query)
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if not rs.eof then
		if scDB="SQL" then
			query="UPDATE pcMailUpSubs SET idCustomer=" & CustID & ",pcMailUpLists_ID=" & ListIDidx & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=" & Abs(Clng(UpdUserReg)-1) & ",pcMailUpSubs_Optout=0 WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
		else
			query="UPDATE pcMailUpSubs SET idCustomer=" & CustID & ",pcMailUpLists_ID=" & ListIDidx & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=" & Abs(Clng(UpdUserReg)-1) & ",pcMailUpSubs_Optout=0 WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
		end if
	else
		if scDB="SQL" then
			query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & CustID & "," & ListIDidx & ",'" & dtTodaysDate & "'," & Abs(Clng(UpdUserReg)-1) & ",0);"
		else
			query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & CustID & "," & ListIDidx & ",#" & dtTodaysDate & "#," & Abs(Clng(UpdUserReg)-1) & ",0);"
		end if
	end if
	set rs=nothing
	set rs=connTemp.execute(query)
	set rs=nothing
End Function

Function CheckUserStatus(CustID,CustEmail,ListID,ListGuid,tmpURL,tmpAuto)
Dim tmp1,tmp2,result,tmpCode,i,j,k,intCount
Dim rs,query
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp2=tmp1
	tmp1=tmp1 & "frontend/xmlChkSubscriber.aspx?listguid=" & ListGuid & "&list=" & ListID & "&email=" & CustEmail
	result=ConnectServer(tmp1,"GET","application/x-www-form-urlencoded","","")
	IF result="ERROR" or result="TIMEOUT" THEN
		CheckUserStatus="-1"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		If not IsNumeric(result) then
			CheckUserStatus="-1"
		Else
			If clng(result)<1 or clng(result)>4 then
				CheckUserStatus="-1"
			else
				CheckUserStatus=Clng(result)
			end if
		End if
	END IF
End Function


Function UnsubUser(CustID,CustEmail,ListID,ListGuid,tmpURL,tmpAuto)
Dim tmp1,result,tmpCode,i,j,k,intCount
Dim rs,query,dtTodaysDate
	IF tmpAuto="1" THEN
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp1=tmp1 & "frontend/xmlUnSubscribe.aspx?ListGuid=" & ListGuid & "&list=" & ListID & "&email=" & CustEmail
	result=ConnectServer(tmp1,"GET","application/x-www-form-urlencoded","","")
	IF result="ERROR" or result="TIMEOUT" THEN
		UnsubUser="0"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=Clng(result)
		MU_ErrMsg=""
		if tmpCode="0" or tmpCode="2" or tmpCode="3" then
			UnsubUser="1"
		else
			UnsubUser="0"
			Select Case tmpCode
				Case "1": MU_ErrMsg="Generic error"
				'Case "3": MU_ErrMsg="User either does not exist in the list or had already been unsubscribed"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
	ELSE
		UnsubUser=0
	END IF
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	
	query="SELECT pcMailUpLists_ID FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & ListID & ";"
	set rs=connTemp.execute(query)
	ListIDidx=0
	if not rs.eof then
		ListIDidx=rs("pcMailUpLists_ID")
	end if
	set rs=nothing
	
	if UnsubUser="1" then
		query="DELETE FROM pcMailUpSubs WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
	else
		if scDB="SQL" then
			query="UPDATE pcMailUpSubs SET idCustomer=" & CustID & ",pcMailUpLists_ID=" & ListIDidx & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=1,pcMailUpSubs_Optout=1 WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
		else
			query="UPDATE pcMailUpSubs SET idCustomer=" & CustID & ",pcMailUpLists_ID=" & ListIDidx & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=1,pcMailUpSubs_Optout=1 WHERE idCustomer=" & CustID & " AND pcMailUpLists_ID=" & ListIDidx & ";"
		end if
	end if
	set rs=connTemp.execute(query)
	set rs=nothing
End Function

Sub GetProIDsList(MU_XmlDoc)
Dim tmpXML,tmpXML1,iXML,iRoot
Dim parentNode,ChildNodes,strNode,GroupsNodes,subNode

	tmpXML=split(MU_XmlDoc,"mailupMessage&gt;")
	tmpXML1="&lt;mailupMessage&gt;" & tmpXML(1) & "mailupMessage&gt;"
	tmpXML=replace(tmpXML1,"&lt;","<")
	tmpXML=replace(tmpXML,"&gt;",">")
	
	Set iXML=Server.CreateObject("MSXML2.DOMDocument")
	iXML.async=false
	iXML.loadXML(tmpXML)
	
	If iXML.parseError.errorCode=0 Then
		Set iRoot=iXML.documentElement
		Set parentNode=iRoot.selectSingleNode("mailupBody/processes")
		Set ChildNodes = parentNode.childNodes
		session("CP_MU_ReturnIDs")=""
		session("CP_MU_ReturnList")=""
		session("CP_MU_ReturnCode")=""
		For Each strNode In ChildNodes
			Set GroupsNodes = strNode.childNodes
			For Each subNode in GroupsNodes
				if ucase(subNode.nodeName)="PROCESSID" then
					if session("CP_MU_ReturnIDs")<>"" then
						session("CP_MU_ReturnIDs")=session("CP_MU_ReturnIDs") & ";"
					end if
					session("CP_MU_ReturnIDs")=session("CP_MU_ReturnIDs") & subNode.text
				end if
				if ucase(subNode.nodeName)="LISTID" then
					if session("CP_MU_ReturnList")<>"" then
						session("CP_MU_ReturnList")=session("CP_MU_ReturnList") & ";"
					end if
					session("CP_MU_ReturnList")=session("CP_MU_ReturnList") & subNode.text
				end if
				if ucase(subNode.nodeName)="RETURNCODE" then
					if session("CP_MU_ReturnCode")<>"" then
						session("CP_MU_ReturnCode")=session("CP_MU_ReturnCode") & ";"
					end if
					session("CP_MU_ReturnCode")=session("CP_MU_ReturnCode") & subNode.text
				end if
			Next
		Next
	End if
End Sub

Function MUImport(APIUser,APIPass,tmpURL,ListID,ListGuid,xmlDoc,IDGroups,OptIn,OptOut,SendConfirm)
Dim tmp1,tmpData,result,tmpCode
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp1=tmp1 & "Services/WSMailupImport.asmx"
	tmpData=""
	tmpData=tmpData & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbcrlf
	tmpData=tmpData & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & vbcrlf
	tmpData=tmpData & "<soap:Header>" & vbcrlf
	tmpData=tmpData & "<Authentication xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<User>" & APIUser & "</User>" & vbcrlf
	tmpData=tmpData & "<Password>" & APIPass & "</Password>" & vbcrlf
	tmpData=tmpData & "</Authentication>" & vbcrlf
	tmpData=tmpData & "</soap:Header>" & vbcrlf
	tmpData=tmpData & "<soap:Body>" & vbcrlf
	tmpData=tmpData & "<StartImportProcesses xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<listsIDs>" & ListID & "</listsIDs>" & vbcrlf
	tmpData=tmpData & "<listsGUIDs>" & ListGuid & "</listsGUIDs>" & vbcrlf
	tmpData=tmpData & "<xmlDoc>" & Server.HTMLEncode(xmlDoc) & "</xmlDoc>" & vbcrlf
	tmpData=tmpData & "<groupsIDs>" & IDGroups & "</groupsIDs>" & vbcrlf
	tmpData=tmpData & "<importType>3</importType>" & vbcrlf
	tmpData=tmpData & "<mobileInputType>1</mobileInputType>" & vbcrlf
	tmpData=tmpData & "<asPending>0</asPending>" & vbcrlf
	tmpData=tmpData & "<ConfirmEmail>" & SendConfirm & "</ConfirmEmail>" & vbcrlf
	tmpData=tmpData & "<asOptOut>" & OptOut & "</asOptOut>" & vbcrlf
	tmpData=tmpData & "<forceOptIn>" & OptIn & "</forceOptIn>" & vbcrlf
	tmpData=tmpData & "<replaceGroups>0</replaceGroups>" & vbcrlf
	tmpData=tmpData & "<idConfirmNL>0</idConfirmNL>" & vbcrlf
	tmpData=tmpData & "</StartImportProcesses>" & vbcrlf
	tmpData=tmpData & "</soap:Body>" & vbcrlf
	tmpData=tmpData & "</soap:Envelope>"
	result=ConnectServer(tmp1,"POST","text/xml; charset=utf-8","http://ws.mailupnet.it/StartImportProcesses",tmpData)
	IF result="ERROR" or result="TIMEOUT" THEN
		MUImport="0"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=FindReturnCode(result)
		MU_XmlDoc=""
		if tmpCode="0" then
			MUImport="1"
			call GetProIDsList(result)
		else
			MUImport="0"
			Select Case tmpCode
				Case "-450": MU_ErrMsg="the listIDs, listGUIDs and groupsIDs do not have the same number of elements. The number of each of them must match up."
				Case "-400": MU_ErrMsg="unrecognized error"
				Case "-401": MU_ErrMsg="xmlDoc is empty"
				Case "-402": MU_ErrMsg="convert xml to csv failed"
				Case "-403": MU_ErrMsg="create new import process failed"
				Case "-100": MU_ErrMsg="unrecognized error"
				Case "-101": MU_ErrMsg="verification failed"
				Case "-102": MU_ErrMsg="List Guid format is not valid"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
End Function

Function MUCreateGroup(APIUser,APIPass,tmpURL,ListID,ListGuid,NewGroupName)
Dim tmp1,tmpData,result,tmpCode
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp1=tmp1 & "Services/WSMailupImport.asmx"
	tmpData=""
	tmpData=tmpData & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbcrlf
	tmpData=tmpData & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & vbcrlf
	tmpData=tmpData & "<soap:Header>" & vbcrlf
	tmpData=tmpData & "<Authentication xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<User>" & APIUser & "</User>" & vbcrlf
	tmpData=tmpData & "<Password>" & APIPass & "</Password>" & vbcrlf
	tmpData=tmpData & "</Authentication>" & vbcrlf
	tmpData=tmpData & "</soap:Header>" & vbcrlf
	tmpData=tmpData & "<soap:Body>" & vbcrlf
	tmpData=tmpData & "<CreateGroup xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<idList>" & ListID & "</idList>" & vbcrlf
	tmpData=tmpData & "<listGUID>" & ListGuid & "</listGUID>" & vbcrlf
	tmpData=tmpData & "<newGroupName>" & NewGroupName & "</newGroupName>" & vbcrlf
	tmpData=tmpData & "</CreateGroup>" & vbcrlf
	tmpData=tmpData & "</soap:Body>" & vbcrlf
	tmpData=tmpData & "</soap:Envelope>"
	result=ConnectServer(tmp1,"POST","text/xml; charset=utf-8","http://ws.mailupnet.it/CreateGroup",tmpData)
	IF result="ERROR" or result="TIMEOUT" THEN
		MUCreateGroup="0"
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=FindReturnCode(result)
		MU_XmlDoc=""
		if tmpCode>"0" then
			MUCreateGroup=tmpCode
		else
			MUCreateGroup="0"
			Select Case tmpCode
				Case "-300": MU_ErrMsg="unrecognized error"
				Case "-301": MU_ErrMsg="the list has not been specified"
				Case "-302": MU_ErrMsg="the group name has not been specified"
				Case "-303": MU_ErrMsg="the group already exists"
				Case "-100": MU_ErrMsg="unrecognized error"
				Case "-101": MU_ErrMsg="verification failed"
				Case "-102": MU_ErrMsg="List Guid format is not valid"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
End Function

Function MUGetIMStatus(APIUser,APIPass,tmpURL,ListID,ListGuid,IDProcess)
Dim tmp1,tmpData,result,tmpCode,tmpStatus,tmpStatusCode,tmpConfirm,tmpConfirmCode
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp1=tmp1 & "Services/WSMailupImport.asmx"
	tmpData=""
	tmpData=tmpData & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbcrlf
	tmpData=tmpData & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & vbcrlf
	tmpData=tmpData & "<soap:Header>" & vbcrlf
	tmpData=tmpData & "<Authentication xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<User>" & APIUser & "</User>" & vbcrlf
	tmpData=tmpData & "<Password>" & APIPass & "</Password>" & vbcrlf
	tmpData=tmpData & "</Authentication>" & vbcrlf
	tmpData=tmpData & "</soap:Header>" & vbcrlf
	tmpData=tmpData & "<soap:Body>" & vbcrlf
	tmpData=tmpData & "<GetProcessDetails xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<idList>" & ListID & "</idList>" & vbcrlf
	tmpData=tmpData & "<listGUID>" & ListGuid & "</listGUID>" & vbcrlf
	tmpData=tmpData & "<idProcess>" & IDProcess & "</idProcess>" & vbcrlf
	tmpData=tmpData & "</GetProcessDetails>" & vbcrlf
	tmpData=tmpData & "</soap:Body>" & vbcrlf
	tmpData=tmpData & "</soap:Envelope>" & vbcrlf
	result=ConnectServer(tmp1,"POST","text/xml; charset=utf-8","http://ws.mailupnet.it/GetProcessDetails",tmpData)
	IF result="ERROR" or result="TIMEOUT" THEN
		MUGetIMStatus=""
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=FindReturnCode(result)
		MU_XmlDoc=""
		if tmpCode="0" then
			MUGetIMStatus=""
			tmpStatusCode=FindStatusCode(result)
			tmpStatus=""
			Select Case tmpStatusCode
				Case "1": tmpStatus="Not yet started"
				Case "2": tmpStatus="Running"
				Case "3": tmpStatus="Completed"
				Case Else: tmpStatus="Unknown"
			End Select
			tmpConfirmCode=FindConfirmationSent(result)
			tmpConfim=""
			Select Case UCase(tmpConfirmCode)
				Case "FALSE","0": tmpConfim="Not sent"
				Case "TRUE","1","-1": tmpConfim="Sent"
				Case Else: tmpConfim="Unknown"
			End Select
			MUGetIMStatus="<td>" & tmpStatus & "</td><td>" & tmpConfim & "</td>"
		else
			MUGetIMStatus=""
			Select Case tmpCode
				Case "-500": MU_ErrMsg="unrecognized error"
				Case "-501": MU_ErrMsg="idProcess not found"
				Case "-100": MU_ErrMsg="unrecognized error"
				Case "-101": MU_ErrMsg="verification failed"
				Case "-102": MU_ErrMsg="List Guid format is not valid"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
End Function

Function MUStartProcess(APIUser,APIPass,tmpURL,ListID,ListGuid,IDProcess)
Dim tmp1,tmpData,result,tmpCode,tmpStatus,tmpStatusCode,tmpConfirm,tmpConfirmCode
	if Ucase(Left(tmpURL,7))<>"HTTP://" then
		tmp1="http://" & tmpURL
	end if
	if Right(tmp1,1)<>"/" then
		tmp1=tmp1 & "/"
	end if
	tmp1=tmp1 & "Services/WSMailupImport.asmx"
	tmpData=""
	tmpData=tmpData & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbcrlf
	tmpData=tmpData & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & vbcrlf
	tmpData=tmpData & "<soap:Header>" & vbcrlf
	tmpData=tmpData & "<Authentication xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<User>" & APIUser & "</User>" & vbcrlf
	tmpData=tmpData & "<Password>" & APIPass & "</Password>" & vbcrlf
	tmpData=tmpData & "</Authentication>" & vbcrlf
	tmpData=tmpData & "</soap:Header>" & vbcrlf
	tmpData=tmpData & "<soap:Body>" & vbcrlf
	tmpData=tmpData & "<StartProcess xmlns=""http://ws.mailupnet.it/"">" & vbcrlf
	tmpData=tmpData & "<idList>" & ListID & "</idList>" & vbcrlf
	tmpData=tmpData & "<listGUID>" & ListGuid & "</listGUID>" & vbcrlf
	tmpData=tmpData & "<idProcess>" & IDProcess & "</idProcess>" & vbcrlf
	tmpData=tmpData & "</StartProcess>" & vbcrlf
	tmpData=tmpData & "</soap:Body>" & vbcrlf
	tmpData=tmpData & "</soap:Envelope>" & vbcrlf
	result=ConnectServer(tmp1,"POST","text/xml; charset=utf-8","http://ws.mailupnet.it/StartProcess",tmpData)
	IF result="ERROR" or result="TIMEOUT" THEN
		MUStartProcess=0
		MU_ErrMsg="Cannot connect to newsletter management system (MailUp)"
	ELSE
		tmpCode=FindReturnCode(result)
		MU_XmlDoc=""
		if tmpCode="0" then
			MUStartProcess=1
		else
			MUStartProcess=0
			Select Case tmpCode
				Case "-600": MU_ErrMsg="unrecognized error"
				Case "-601": MU_ErrMsg="an import process is already running for the list"
				Case "-602": MU_ErrMsg="an import process is already running for a different list"
				Case "-603": MU_ErrMsg="error checking process status"
				Case "-604": MU_ErrMsg="error starting the process job"
				Case "-100": MU_ErrMsg="unrecognized error"
				Case "-101": MU_ErrMsg="verification failed"
				Case "-102": MU_ErrMsg="List Guid format is not valid"
				Case Else: MU_ErrMsg="Error Code: " & tmpCode
			End Select
		end if
	END IF
End Function
%>