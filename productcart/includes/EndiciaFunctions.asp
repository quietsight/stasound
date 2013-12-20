<%'Endicia Integration Functions
Dim srvXmlHttp
Dim EDC_ErrMsg, EDC_SuccessMsg, EDC_XmlDoc

EDC_ErrMsg=""
EDC_SuccessMsg=""

tmpscXML=".3.0"
private const DeveloperTest="N" '// "Y" or "N"
private const EDCURL="https://LabelServer.Endicia.com/LabelService/EwsLabelService.asmx/"
'private const EDCURL="https://www.envmgr.com/LabelService/EwsLabelService.asmx/"
private const EDCURLSpc="https://www.endicia.com/ELS/ELSServices.cfc?wsdl"
private const EDCPartnerID="leip"
private const EDCGPostage="GetPostageLabelXML"
private const EDCBPostage="BuyPostageXML"
private const EDCCPass="ChangePassPhraseXML"
private const EDCCal="CalculatePostageRateXML"
private const EDCAStatus="GetAccountStatusXML"
private const EDCSignUp="UserSignup"
private const EDCRefund="RefundRequest"
private const EDCExpDays=10

Dim MaxRequestTime,StopHTTPRequests
Dim EDCUserID,EDCPassP,EDCAutoRefill,EDCTriggerAmount,EDCLogTrans,EDCReg,EDCTestMode,EDCBuyAmount,EDCABalance,EDCReXML,EDCFillAmount,EDCAutoRmv
Dim EDCTrackingNum,EDCPIC,EDCCustomsNum,EDCTransID,EDCFPostage,EDCRBalance,EDCSPostage,EDCFees,EDCFeesDetails,EDCID,EDCCalMode,EDCLabelFile,EDCIsPIC

'maximum seconds for each HTTP request time
MaxRequestTime=10

StopHTTPRequests=0

Sub GetEDCSettings()
Dim query,rsQ

call opendb()

query="SELECT pcES_UserID,pcES_PassP,pcES_AutoRefill,pcES_TriggerAmount,pcES_FillAmount,pcES_LogTrans,pcES_Reg,pcES_TestMode,pcES_AutoRmvLogs FROM pcEDCSettings;"
set rsQ=connTemp.execute(query)

EDCUserID=0
EDCReg=0
EDCTestMode=1
EDCAutoRefill=0
EDCTriggerAmount=0
EDCFillAmount=0
EDCLogTrans=0
EDCAutoRmv=0
EDCPassP=""

if not rsQ.eof then
	EDCUserID=rsQ("pcES_UserID")
	EDCPassP=rsQ("pcES_PassP")
	if EDCPassP<>"" then
		EDCPassP=enDeCrypt(EDCPassP, scCrypPass)
	end if
	EDCAutoRefill=rsQ("pcES_AutoRefill")
	EDCTriggerAmount=rsQ("pcES_TriggerAmount")
	EDCFillAmount=rsQ("pcES_FillAmount")
	EDCLogTrans=rsQ("pcES_LogTrans")
	EDCReg=rsQ("pcES_Reg")
	EDCTestMode=rsQ("pcES_TestMode")
	EDCAutoRmv=rsQ("pcES_AutoRmvLogs")
end if
set rsQ=nothing

End Sub

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

Function XMLReplace(tmpData)
Dim tmp1
	tmp1=tmpData
	tmp1=replace(tmp1,"&","&amp;")
	tmp1=replace(tmp1,"<","&lt;")
	tmp1=replace(tmp1,">","&gt;")
	tmp1=replace(tmp1,"""","&quot;")
	tmp1=replace(tmp1,"'","&apos;")
	XMLReplace=tmp1
End Function

Function XMLNumber(tmpData)
Dim tmp1
	tmp1=tmpData
	if scDecSign="," then
		tmp1=replace(tmp1,".","")
		tmp1=replace(tmp1,",",".")
	else
		tmp1=replace(tmp1,",","")
	end if
	XMLNumber=tmp1
End Function

Function ConnectServer(tmpURL,tmpMethod,tmpContentType,tmpSOAPHead,tmpData)
Dim rersult1,tmpStatus

on error resume next
Set srvXmlHttp=nothing

IF StopHTTPRequests<>"1" THEN
	Set srvXmlHttp = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
	if not IsNull(MaxRequestTime) then
		srvXmlHttp.open tmpMethod, tmpURL, True
	else
		srvXmlHttp.open tmpMethod, tmpURL, False
	end if
	if tmpContentType="" then
		srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	else
		srvXmlHttp.setRequestHeader "Content-Type", tmpContentType
	end if
	srvXmlHttp.setRequestHeader "CharSet", "UTF-8"
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
				if Instr(result1,"Response>")=0 then
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

Function FindStatusCode(tmpXML)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"<Status>")>0 then
			tmp1=split(tmpXML,"<Status>")
			tmp2=split(tmp1(1),"</Status>")
			FindStatusCode=Clng(tmp2(0))
		else
			if Instr(tmpXML,"&lt;Status&gt;")>0 then
				tmp1=split(tmpXML,"&lt;Status&gt;")
				tmp2=split(tmp1(1),"&lt;/Status&gt;")
				FindStatusCode=Clng(tmp2(0))
			else
				FindStatusCode=0
			end if
		end if
	else
		FindStatusCode=0
	end if
End Function

Function FindXMLValue(tmpXML,tmpName)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"<" & tmpName & ">")>0 then
			tmp1=split(tmpXML,"<" & tmpName & ">")
			tmp2=split(tmp1(1),"</" & tmpName & ">")
			FindXMLValue=tmp2(0)
		else
			if Instr(tmpXML,"&lt;" & tmpName & "&gt;")>0 then
				tmp1=split(tmpXML,"&lt;" & tmpName & "&gt;")
				tmp2=split(tmp1(1),"&lt;/" & tmpName & "&gt;")
				FindXMLValue=tmp2(0)
			else
				FindXMLValue=""
			end if
		end if
	else
		FindXMLValue=""
	end if
End Function

Function FindErrMsg(tmpXML)
Dim tmp1,tmp2
	if tmpXML<>"" then
		if Instr(tmpXML,"<ErrorMessage>")>0 then
			tmp1=split(tmpXML,"<ErrorMessage>")
			tmp2=split(tmp1(1),"</ErrorMessage>")
			FindErrMsg=tmp2(0)
		else
			if Instr(tmpXML,"&lt;ErrorMessage&gt;")>0 then
				tmp1=split(tmpXML,"&lt;ErrorMessage&gt;")
				tmp2=split(tmp1(1),"&lt;/ErrorMessage&gt;")
				FindErrMsg=tmp2(0)
			else
				FindErrMsg=""
			end if
		end if
	else
		FindErrMsg=""
	end if
End Function

Function BuyPostage(tmpValue)
Dim tmp1,tmpXML,result,tmpCode,query,rs

	EDCBuyAmount=0

	if (tmpValue="") OR (Not IsNumeric(tmpValue)) then
		EDC_ErrMsg="Please enter correct numbers!"
		BuyPostage=0
		Exit Function
	end if
	
	if (CDbl(tmpValue)<10) OR (CDbl(tmpValue)>99999.99) then
		EDC_ErrMsg="Please enter correct numbers in range: 10.00 - 99999.99"
		BuyPostage=0
		Exit Function
	end if
	
	EDC_ErrMsg=""
	EDC_SuccessMsg=""
	
	tmp1=EDCURL & "BuyPostageXML"
	tmpXML="<RecreditRequest>"
	tmpXML=tmpXML & "<RequesterID>" & EDCPartnerID & "</RequesterID>"
	tmpXML=tmpXML & "<RequestID>PCEDC12345678</RequestID>"
	tmpXML=tmpXML & "<CertifiedIntermediary>"
	tmpXML=tmpXML & "<AccountID>" & EDCUserID & "</AccountID>"
	tmpXML=tmpXML & "<PassPhrase>" & EDCPassP & "</PassPhrase>"
	tmpXML=tmpXML & "</CertifiedIntermediary>"
	tmpXML=tmpXML & "<RecreditAmount>" & tmpValue & "</RecreditAmount>"
	tmpXML=tmpXML & "</RecreditRequest>"
	tmpXML="recreditRequestXML=" & tmpXML
	result=ConnectServer(tmp1,"POST","","",tmpXML)
	IF result="ERROR" or result="TIMEOUT" THEN
		BuyPostage=0
		tmpCode=0
		EDC_ErrMsg="Cannot connect to Endicia Label Server"
	ELSE
		tmpCode=FindStatusCode(result)
		if tmpCode="0" then
			BuyPostage=1
			tmpCode=1
			EDCBuyAmount=tmpValue
			EDC_SuccessMsg="Your Endicia account was successfully refilled!"
		else
			BuyPostage=0
			tmpCode=0
			EDC_ErrMsg=FindErrMsg(result)
		end if
	END IF
	Call SaveTrans(tmpXML,result,tmpCode,2)
End Function

Function AutoRefill()
Dim tmpCode1
EDC_ErrMsg=""
EDC_SuccessMsg=""
IF (EDCAutoRefill="1") AND (EDCFillAmount>"0") THEN
	tmpCode1=GetAccountStatus()
	if tmpCode1=0 then
		AutoRefill=0
		exit function
	else
		if Cdbl(EDCABalance)<cdbl(EDCTriggerAmount) then
			tmpCode1=BuyPostage(EDCFillAmount)
			if tmpCode1=0 then
				AutoRefill=0
				exit function
			else
				AutoRefill=1
				EDC_SuccessMsg="Your Endicia account was successfully refilled!"
			end if
		else
			AutoRefill=0
			exit function
		end if
	end if
ELSE
	AutoRefill=0
	exit function
END IF

End Function

Function GetAccountStatus()
Dim tmp1,tmpXML,result,tmpCode,query,rs

	EDCABalance=0
	
	EDCReXML=""

	EDC_ErrMsg=""
	EDC_SuccessMsg=""
	
	tmp1=EDCURL & "GetAccountStatusXML"
	tmpXML="<AccountStatusRequest>"
	tmpXML=tmpXML & "<RequesterID>" & EDCPartnerID & "</RequesterID>"
	tmpXML=tmpXML & "<RequestID>PCEDC12345678</RequestID>"
	tmpXML=tmpXML & "<CertifiedIntermediary>"
	tmpXML=tmpXML & "<AccountID>" & EDCUserID & "</AccountID>"
	tmpXML=tmpXML & "<PassPhrase>" & EDCPassP & "</PassPhrase>"
	tmpXML=tmpXML & "</CertifiedIntermediary>"
	tmpXML=tmpXML & "</AccountStatusRequest>"
	tmpXML="accountStatusRequestXML=" & tmpXML
	result=ConnectServer(tmp1,"POST","","",tmpXML)
	IF result="ERROR" or result="TIMEOUT" THEN
		GetAccountStatus=0
		tmpCode=0
		EDC_ErrMsg="Cannot connect to Endicia Label Server"
	ELSE
		tmpCode=FindStatusCode(result)
		if tmpCode="0" then
			GetAccountStatus=1
			tmpCode=1
			EDCReXML=result
			EDCABalance=FindXMLValue(EDCReXML,"PostageBalance")
		else
			GetAccountStatus=0
			tmpCode=0
			EDC_ErrMsg=FindErrMsg(result)
		end if
	END IF
	Call SaveTrans(tmpXML,result,tmpCode,5)
End Function


Function ChangePassP(APIUser,APIPassP,APINewPassP)
Dim tmp1,tmpXML,result,tmpCode,query,rs
	if (EDCUserID>"0") then
		if (Clng(APIUser)<>Clng(EDCUserID)) OR (EDCPassP<>APIPassP) then
			EDC_ErrMsg="Please enter correct your current UserID and Pass Phrase"
			ChangePassP=0
			Exit Function
		end if
	end if
	
	if trim(APINewPassP)="" then
		EDC_ErrMsg="Please enter your new Pass Phrase"
		ChangePassP=0
		Exit Function
	end if
	
	EDC_ErrMsg=""
	EDC_SuccessMsg=""
	
	tmp1=EDCURL & "ChangePassPhraseXML"
	tmpXML="<ChangePassPhraseRequest>"
	tmpXML=tmpXML & "<RequesterID>" & EDCPartnerID & "</RequesterID>"
	tmpXML=tmpXML & "<RequestID>PCEDC12345678</RequestID>"
	tmpXML=tmpXML & "<CertifiedIntermediary>"
	tmpXML=tmpXML & "<AccountID>" & APIUser & "</AccountID>"
	tmpXML=tmpXML & "<PassPhrase>" & APIPassP & "</PassPhrase>"
	tmpXML=tmpXML & "</CertifiedIntermediary>"
	tmpXML=tmpXML & "<NewPassPhrase>" & APINewPassP & "</NewPassPhrase>"
	tmpXML=tmpXML & "</ChangePassPhraseRequest>"
	tmpXML="changePassPhraseRequestXML=" & tmpXML
	result=ConnectServer(tmp1,"POST","","",tmpXML)
	IF result="ERROR" or result="TIMEOUT" THEN
		ChangePassP=0
		tmpCode=0
		EDC_ErrMsg="Cannot connect to Endicia Label Server"
	ELSE
		tmpCode=FindStatusCode(result)
		if tmpCode="0" then
			ChangePassP=1
			tmpCode=1
			call opendb()
			APINewPassP=enDeCrypt(APINewPassP, scCrypPass)
			if (EDCUserID>"0") then
			query="UPDATE pcEDCSettings SET pcES_UserID=" & APIUser & ",pcES_PassP='" & APINewPassP & "' WHERE pcES_UserID=" & APIUser & ";"
			else
			query="UPDATE pcEDCSettings SET pcES_UserID=" & APIUser & ",pcES_PassP='" & APINewPassP & "' WHERE pcES_Reg=1;"
			end if
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
			EDC_SuccessMsg="Your Pass Phrase was updated successfully!"
		else
			ChangePassP=0
			tmpCode=0
			EDC_ErrMsg=FindErrMsg(result)
		end if
	END IF
	Call SaveTrans(tmpXML,result,tmpCode,3)
End Function

Sub SaveTrans(RequestXML,ResponseXML,tmpSuccess,MethodCode)
Dim query,rsQ,objXMLDoc,objStream

	dim dtTodaysDate
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & Time()
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & Time()
	end if

	Select Case MethodCode
	
	Case 1:
		strFileName=""
		EDCTrackingNum=""
		EDCPIC=""
		EDCCustomsNum=""
		EDCTransID=""
		EDCFPostage=0
		EDCRBalance=0
		EDCSPostage=0
		EDCFees=0
		EDCFeesDetails=""
		EDCLabelFile=""
		EDCIsPIC=0
		IF tmpSuccess="1" THEN
			IF EDCCalMode<>"yes" THEN
			Base64LabelImage=FindXMLValue(ResponseXML,"Base64LabelImage")
			if Base64LabelImage="" then
				Set iXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
				iXML.async=false
				iXML.loadXML(ResponseXML)
				If iXML.parseError.errorCode=0 Then
					Set iRoot=iXML.documentElement
					Set parentNode=iRoot.selectSingleNode("Label")
					Set ChildNodes = parentNode.childNodes
					For Each strNode In ChildNodes
						if ucase(strNode.nodeName)="IMAGE" then
							Base64LabelImage=Base64LabelImage & strNode.text
						end if
					Next
				end if
				Set iXML=nothing
			end if
			if Base64LabelImage<>"" then
				GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"">"&Base64LabelImage&"</Base64Data>"
				set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
				objXMLDoc.async = False
				objXMLDoc.validateOnParse = False
							
				If objXMLDoc.loadXML(GraphicXML) Then
					Set objStream = Server.CreateObject("ADODB.Stream")
					objStream.Type = 1
					objStream.Open
					objStream.Write objXMLDoc.selectSingleNode("/Base64Data").nodeTypedValue
					strFileName = year(Date()) & month(Date()) & day(Date()) & hour(Time()) & Minute(Time()) & Second(Time()) & Session("pcAdmincustomerRefNo1") & Session("pcAdminOrderID") & "." & session("pcEDCLabelFormat")
					objStream.SaveToFile server.mappath("USPSLabels\" & strFileName), 2
					if err.number<>0 then
						err.number=0
						err.description=""
					else
					EDCLabelFile=strFileName
					end if
					objStream.Close()
					Set objStream = Nothing
				end if
				set objXMLDoc = Nothing
			end if
			END IF
			
			EDCTrackingNum=FindXMLValue(ResponseXML,"TrackingNumber")
			EDCPIC=FindXMLValue(ResponseXML,"PIC")
			if EDCPIC<>"" then
				EDCIsPIC="1"
			end if
			EDCCustomsNum=FindXMLValue(ResponseXML,"CustomsNumber")
			EDCTransID=FindXMLValue(ResponseXML,"TransactionID")
			EDCFPostage=FindXMLValue(ResponseXML,"FinalPostage")
			EDCRBalance=FindXMLValue(ResponseXML,"PostageBalance")
			
			
			Set iXML=Server.CreateObject("MSXML2.DOMDocument"&tmpscXML)
			iXML.async=false
			iXML.loadXML(ResponseXML)
			If iXML.parseError.errorCode=0 Then
				on error resume next
				Set iRoot=iXML.documentElement
				IF EDCCalMode="yes" THEN	
					Set parentNode=iRoot.selectSingleNode("PostagePrice")
					Set sNodeAttributes = parentNode.attributes
					For Each strAtt In sNodeAttributes
						if ucase(strAtt.name)="TOTALAMOUNT" then
							EDCFPostage=strAtt.value
							exit for
						end if
					Next
				END IF
				Set parentNode=iRoot.selectSingleNode("PostagePrice/Postage")
				Set sNodeAttributes = parentNode.attributes
				For Each strAtt In sNodeAttributes
					if ucase(strAtt.name)="TOTALAMOUNT" then
						EDCSPostage=strAtt.value
						exit for
					end if
				Next
				Set parentNode=iRoot.selectSingleNode("PostagePrice/Fees")
				Set sNodeAttributes = parentNode.attributes
				For Each strAtt In sNodeAttributes
					if ucase(strAtt.name)="TOTALAMOUNT" then
						EDCFees=strAtt.value
						exit for
					end if
				Next
				Set parentNode=iRoot.selectSingleNode("PostagePrice/Fees")
				Set ChildNodes = parentNode.childNodes
				For Each strNode In ChildNodes
					if strNode.text<>"0" then
						EDCFeesDetails=EDCFeesDetails & strNode.nodeName & "|!|" & scCurSign & money(strNode.text) & "|~|"
					end if
				Next
			end if
			Set iXML=nothing
			
		END IF 'Successfully
		
		IF EDCCalMode<>"yes" THEN		
		if scDB="SQL" then
			query="INSERT INTO pcEDCTrans (IDOrder,pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg,pcET_LabelFile,pcET_PicNum,pcET_CustomsNum,pcET_TransID,pcET_Postage,pcET_subPostage,pcET_Fees,pcET_FeesDetails) VALUES (" & Session("pcAdminOrderID") & ",'" & dtTodaysDate & "',1," & tmpSuccess & ",'" & EDC_ErrMsg & "','" & EDCLabelFile & "','" & EDCPIC & "','" & EDCCustomsNum & "','" & EDCTransID & "'," & EDCFPostage & "," & EDCSPostage & "," & EDCFees & ",'" & EDCFeesDetails & "');"
		else
			query="INSERT INTO pcEDCTrans (IDOrder,pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg,pcET_LabelFile,pcET_PicNum,pcET_CustomsNum,pcET_TransID,pcET_Postage,pcET_subPostage,pcET_Fees,pcET_FeesDetails) VALUES (" & Session("pcAdminOrderID") & ",#" & dtTodaysDate & "#,1," & tmpSuccess & ",'" & EDC_ErrMsg & "','" & EDCLabelFile & "','" & EDCPIC & "','" & EDCCustomsNum & "','" & EDCTransID & "'," & EDCFPostage & "," & EDCSPostage & "," & EDCFees & ",'" & EDCFeesDetails & "');"
		end if
		END IF
	
	Case 2:
		if scDB="SQL" then
			query="INSERT INTO pcEDCTrans (pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg,pcET_Postage) VALUES ('" & dtTodaysDate & "',2," & tmpSuccess & ",'" & EDC_ErrMsg & "'," & EDCBuyAmount & ");"
		else
			query="INSERT INTO pcEDCTrans (pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg,pcET_Postage) VALUES (#" & dtTodaysDate & "#,2," & tmpSuccess & ",'" & EDC_ErrMsg & "'," & EDCBuyAmount & ");"
		end if
		EDCABalance=FindXMLValue(ResponseXML,"PostageBalance")
	
	Case 3:
		if scDB="SQL" then
			query="INSERT INTO pcEDCTrans (pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg) VALUES ('" & dtTodaysDate & "',3," & tmpSuccess & ",'" & EDC_ErrMsg & "');"
		else
			query="INSERT INTO pcEDCTrans (pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg) VALUES (#" & dtTodaysDate & "#,3," & tmpSuccess & ",'" & EDC_ErrMsg & "');"
		end if
		
	Case 5:
		if scDB="SQL" then
			query="INSERT INTO pcEDCTrans (pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg,pcET_Postage) VALUES ('" & dtTodaysDate & "',5," & tmpSuccess & ",'" & EDC_ErrMsg & "'," & EDCABalance & ");"
		else
			query="INSERT INTO pcEDCTrans (pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg,pcET_Postage) VALUES (#" & dtTodaysDate & "#,5," & tmpSuccess & ",'" & EDC_ErrMsg & "'," & EDCABalance & ");"
		end if
	Case 7:
		if pcv_IsPIC="1" then
			tmpStr1="pcET_PicNum"
		else
			tmpStr1="pcET_CustomsNum"
		end if
		if scDB="SQL" then
			query="INSERT INTO pcEDCTrans (" & tmpStr1 & ",IDOrder,pcPackageInfo_ID,pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg) VALUES ('" & pcv_TrackNum & "'," & pcv_IDOrder & "," & PackID & ",'" & dtTodaysDate & "',7," & tmpSuccess & ",'" & EDC_ErrMsg & "');"
		else
			query="INSERT INTO pcEDCTrans (" & tmpStr1 & ",IDOrder,pcPackageInfo_ID,pcET_TransDate,pcET_Method,pcET_Success,pcET_ErrMsg) VALUES ('" & pcv_TrackNum & "'," & pcv_IDOrder & "," & PackID & ",#" & dtTodaysDate & "#,7," & tmpSuccess & ",'" & EDC_ErrMsg & "');"
		end if
	End Select

	IF EDCCalMode<>"yes" THEN
	call opendb()
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	
	query="SELECT pcET_ID FROM pcEDCTrans ORDER BY pcET_ID DESC;"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		EDCID=rsQ("pcET_ID")
	end if
	set rsQ=nothing
	
	If MethodCode="7" then
		query="UPDATE pcEDCTrans SET pcET_RefundID=" & EDCID & " WHERE (pcET_Method=1) AND ((pcET_PicNum like '" & pcv_TrackNum & "') OR (pcET_CustomsNum like '" & pcv_TrackNum & "'));"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	End if
	
	if EDCLogTrans="1" then
		query="INSERT INTO pcEDCLogs (pcET_ID,pcELog_Request,pcELog_Response) VALUES (" & EDCID & ",'" & replace(RequestXML,"'","''") & "','" & replace(ResponseXML,"'","''") & "');"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	end if
	END IF
	EDCCalMode="no"

End Sub
%>