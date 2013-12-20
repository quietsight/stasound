<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="../includes/GoogleCheckout_APIFunctions.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Handler.asp"-->
<!--#include file="inc_GenDownloadInfo.asp"-->
<% on error resume next
dim conntemp, query, rs, qry_ID

'// Define objects used to create and send Google Checkout Order Processing API requests
Dim xmlRequest
Dim xmlResponse
Dim attrGoogleOrderNumber
Dim elemAmount
Dim elemReason
Dim elemComment
Dim elemCarrier
Dim elemTrackingNumber
Dim elemMessage
Dim elemSendEmail
Dim elemMerchantOrderNumber
Dim transmitResponse

qry_ID=request("qry_ID")

call opendb()
'find out original status first
query="SELECT orderstatus,paymentCode,pcOrd_Payer, pcOrd_GoogleIDOrder FROM orders WHERE idOrder="& qry_ID
set rs=server.CreateObject("ADODB.RecordSet")
Set rs=conntemp.execute(query)
porigstatus=rs("orderstatus")
paymentCode=rs("paymentCode")
pcOrd_Payer=rs("pcOrd_Payer")
pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder")
set rs=nothing
call closedb()

'****************************************************
' START - EIG processing codes
'****************************************************
IF request("SubmitEIG1")<>"" THEN

	pcv_strAdminPrefix="1"

	Public Function pcf_EIGChars(pgwTransId)
		pgwTransId=replace(pgwTransId,chr(0),"")
		pgwTransId=replace(pgwTransId,chr(13),"")
		pgwTransId=replace(pgwTransId,chr(10),"")
		pgwTransId=replace(pgwTransId,chr(34),"")
		pcf_EIGChars=trim(pgwTransId)
	End Function

	Public Function deformatNVP(nvpstr)
		On Error Resume Next

		Dim AndSplitedArray, EqualtoSplitedArray, Index1, Index2, NextIndex
		Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
		AndSplitedArray = Split(nvpstr, "&", -1, 1)
		NextIndex=0
		For Index1 = 0 To UBound(AndSplitedArray)
			EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
			For Index2 = 0 To UBound(EqualtoSplitedArray)
				NextIndex=Index2+1
				NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
				'response.Write(URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex)) & "<br />")
				Index2=Index2+1
			Next
		Next
		Set deformatNVP = NvpCollection

	End Function

	Function URLDecode(str)
		On Error Resume Next

		str = Replace(str, "+", " ")
		For i = 1 To Len(str)
		sT = Mid(str, i, 1)
			If sT = "%" Then
				sR = sR & Chr(CLng("&H" & Mid(str, i+1, 2)))
				i = i+2
			Else
				sR = sR & sT
			End If
		Next
		URLDecode = sR
	End Function

	call opendb()

	'// Load Settings
	query="SELECT pcPay_EIG_Type, pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key, pcPay_EIG_Curcode, pcPay_EIG_CVV, pcPay_EIG_TestMode, pcPay_EIG_SaveCards, pcPay_EIG_UseVault FROM pcPay_EIG WHERE pcPay_EIG_ID=1"
	set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if NOT rs.eof then
		x_Username=rs("pcPay_EIG_Username")
		x_Username=enDeCrypt(x_Username, pcs_GetSecureKey)
		x_Password=rs("pcPay_EIG_Password")
		x_Password=enDeCrypt(x_Password, pcs_GetSecureKey)
		x_Key=rs("pcPay_EIG_Key")
		x_Key=enDeCrypt(x_Key, pcs_GetSecureKey)
		x_CVV=rs("pcPay_EIG_CVV")
		x_Type=rs("pcPay_EIG_Type")
		x_TypeArray=Split(x_Type,"||")
		x_TransType=x_TypeArray(0)
		x_Curcode=rs("pcPay_EIG_Curcode")
		x_TestMode=rs("pcPay_EIG_TestMode")
		x_SaveCards=rs("pcPay_EIG_SaveCards")
		x_UseVault=rs("pcPay_EIG_UseVault")
	end if
	set rs=nothing

	pcgwTransId=request("EIGTransID")
	pcgwAmount=-1

	ActionType=Ucase(trim(request("SubmitEIG1")))

	'// Contact EIG
	strTest = ""
	strTest = strTest & "username=" & x_Username
	strTest = strTest & "&password=" & x_Password
	strTest = strTest & "&type=refund"
	strTest = strTest & "&transactionid=" & pcgwTransId
	'strTest = strTest & "&amount=" & ?

	'response.Write(strTest)
	'response.End()

	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.open "POST", "https://secure.networkmerchants.com/api/transact.php", false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send strTest
	strStatus = xml.Status
	strRetVal = xml.responseText
	Set xml = Nothing

	'response.Write(strRetVal)
	'response.End()

	'// Check for success
	Set resArray = deformatNVP(strRetVal)
	ack = resArray("response")
	ackDesc = resArray("responsetext")

	If ack="1" Then

		if ActionType="REFUND" then

			'// Update EIG Status - Refunded
			query="UPDATE Orders SET pcOrd_PaymentStatus=6, gwTransId='"& pcf_EIGChars(pcgwTransId)&"' WHERE idorder=" & qry_ID & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'redirect back
			call closedb()
			response.redirect "Orddetails.asp?id="&qry_ID&"&ActiveTab=1"
		end if

	Else

		call closedb()
		response.redirect "Orddetails.asp?id="&qry_ID&"&ActiveTab=1&msg2=" & Server.URLEncode(ackDesc)

	End If

END IF
'****************************************************
' END - EIG processing codes
'****************************************************


'****************************************************
' START - PayPal processing codes
'****************************************************
PayPalProcessOrder=0
PayPalCancelOrder=0
call opendb()
IF request("SubmitPayPal1")<>"" THEN
	pcv_strAdminPrefix="1"

	Public Function pcf_PayPalChars(pgwTransId)
		pgwTransId=replace(pgwTransId,chr(0),"")
		pgwTransId=replace(pgwTransId,chr(13),"")
		pgwTransId=replace(pgwTransId,chr(10),"")
		pgwTransId=replace(pgwTransId,chr(34),"")
		pcf_PayPalChars=trim(pgwTransId)
	End Function

	set objPayPalClass = New pcPayPalClass

	'// Query our PayPal Settings
	objPayPalClass.pcs_SetAllVariables()

	pcgwTransId=request("PayPalTransID")
	pcgwTransParentId=request("PayPalTransParentID")
	pcgwAmount=-1

	ActionType=Ucase(trim(request("SubmitPayPal1")))

	if ActionType="CAPTURE" OR ActionType="REAUTHORIZE" then

		'// We need to use the Order Total Amount because maybe it is different with previous transaction amount
		query="SELECT CurrencyCode FROM pcPay_PayPal_Authorize WHERE authcode like '" & pcgwTransId & "';"
		set rs=server.CreateObject("ADODB.RecordSet")
		Set rs=conntemp.execute(query)
		if not rs.eof then
			pcgwCurrencyCode=rs("CurrencyCode")
		else
			pcgwCurrencyCode="USD"
		end if
		set rs=nothing
		if pcgwCurrencyCode="" then
			pcgwCurrencyCode="USD"
		end if

		query="SELECT total FROM orders WHERE idorder=" & qry_ID & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		Set rs=conntemp.execute(query)
		if not rs.eof then
			pcgwAmount=rs("total")
		end if
		set rs=nothing

	end if
	'response.Write(pcgwAmount)
	'response.End()

	'// Add the required NVP’s
	nvpstr="" '// clear
	if ActionType="CAPTURE" OR ActionType="REAUTHORIZE" then
		objPayPalClass.AddNVP "AUTHORIZATIONID", pcgwTransId
	elseif ActionType="VOID" then
		objPayPalClass.AddNVP "AUTHORIZATIONID", pcgwTransParentId
	else
		objPayPalClass.AddNVP "TRANSACTIONID", pcgwTransId
		objPayPalClass.AddNVP "REFUNDTYPE", "Full"
	end if
	if ActionType="CAPTURE" OR ActionType="REAUTHORIZE" then
		pcgwAmount=money(pcgwAmount)
		pcgwAmount=pcf_CurrencyField(pcgwAmount)
		objPayPalClass.AddNVP "AMT", pcgwAmount
		objPayPalClass.AddNVP "CURRENCYCODE", pcgwCurrencyCode
	end if
	if ActionType="CAPTURE" THEN
		objPayPalClass.AddNVP "COMPLETETYPE", "Complete"
	end if

	'// Post to PayPal by calling .hash_call
	PayPalFuncName=""
	Select Case ActionType
		Case "CAPTURE": PayPalFuncName="DoCapture"
		Case "REAUTHORIZE": PayPalFuncName="DoReauthorization"
		Case "REFUND": PayPalFuncName="RefundTransaction"
		Case "VOID": PayPalFuncName="DoVoid"
	End Select


	'response.Write(PayPalFuncName & ": " & nvpstr)
	'response.End()

	Set resArray = objPayPalClass.hash_call(PayPalFuncName,nvpstr)

	'// Set Response
	Set Session("nvpResArray")=resArray

	'// Check for success
	ack = UCase(resArray("ACK"))

	'// Check for code errors
	if err.number <> 0 then
		'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
		Dim pcv_PayPalErrMessage
		%><!--#include file="../includes/pcPayPalErrors.asp"--><%
	end if

	If ack="SUCCESS" Then
		if ActionType="CAPTURE" then

			'// CAPTURED
			pgwTransId=resArray("TRANSACTIONID")

			'// Update pcPay_PayPal_Authorize to captured
			query="UPDATE pcPay_PayPal_Authorize SET captured=1 WHERE authcode='"& pcgwTransId &"';"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'// Update Payment Status - PAID
			query="UPDATE Orders SET pcOrd_PaymentStatus=2, gwTransId='"& pgwTransId &"' WHERE idorder=" & qry_ID & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'// Process order if it isn't processed or shipped
			if (porigstatus<>"3") AND (porigstatus<>"4") AND (porigstatus<>"7") AND (porigstatus<>"8") then
				PayPalProcessOrder=1
			else
				'redirect back - customer wanted their orders process at checkout time - not correct to do this with Authorize only, but possible.
				call closedb()
				response.redirect "Orddetails.asp?id="&qry_ID
			end if
		end if

		if ActionType="REAUTHORIZE" then

			'// RE-AUTHORIZED
			pgwTransId=resArray("AUTHORIZATIONID")

			'// Update Authorized Date on the table pcPay_PayPal_Authorize
			pTodaysDate=Date()
			if SQL_Format="1" then
				pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
			else
				pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
			end if
			if scDB="Access" then
				tmpStr="#"& pTodaysDate &"#"
			else
				tmpStr="'"& pTodaysDate &"'"
			end if
			query="UPDATE pcPay_PayPal_Authorize SET AuthorizedDate=" & tmpStr & " WHERE authcode like '" & pcgwTransId & "';"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'// Update Payment Status and Order Status - Authorized and Pending
			query="UPDATE Orders SET pcOrd_PaymentStatus=1, orderStatus=2, gwTransId='"&pcf_PayPalChars(pgwTransId)&"' WHERE idorder=" & qry_ID & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'redirect back
			call closedb()
			response.redirect "Orddetails.asp?id="&qry_ID&"&ActiveTab=1"
		end if

		if ActionType="REFUND" then

			'// RE-AUTHORIZED
			pgwTransId=resArray("REFUNDTRANSACTIONID")

			'// Update Payment Status - Refunded
			query="UPDATE Orders SET pcOrd_PaymentStatus=6, gwTransId='"&pcf_PayPalChars(pgwTransId)&"' WHERE idorder=" & qry_ID & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'redirect back
			call closedb()
			response.redirect "Orddetails.asp?id="&qry_ID&"&ActiveTab=1"
		end if

		if ActionType="VOID" then

			'// VOIDED
			pgwTransId=resArray("AUTHORIZATIONID")

			'// Remove pending PayPal transaction from the table pcPay_PayPal_Authorize
			query="DELETE FROM pcPay_PayPal_Authorize WHERE authcode like '" & pcgwTransId & "';"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'// Update Payment Status - VOIDED
			query="UPDATE Orders SET pcOrd_PaymentStatus=8, gwTransId='"&pcf_PayPalChars(pgwTransId)&"' WHERE idorder=" & qry_ID & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			set rs=nothing

			'// Cancel order if it isn't cancelled
			if (porigstatus<>"5") then
				PayPalCancelOrder=1
			end if

		end if


	else
		'// append the user friendly errors to API errors
		objPayPalClass.GenerateErrorReport()

		if ActionType="CAPTURE" then
			'// Update Payment Status - PENDING
			'query="UPDATE Orders SET pcOrd_PaymentStatus=0 WHERE idorder=" & qry_ID & ";"
			'set rs=server.CreateObject("ADODB.RecordSet")
			'Set rs=conntemp.execute(query)
			'set rs=nothing
		end if

		'redirect back
		call closedb()
		response.redirect "Orddetails.asp?id="&qry_ID&"&ActiveTab=1&msg1=" & Server.URLEncode(pcv_PayPalErrMessage)
	end if
END IF
call closedb()
'****************************************************
' END- PayPal processing codes
'****************************************************


if request.querystring("Submit1")<>"" then
	'update customer information and return to order information
	fname=replace(request.querystring("name"),"'","''")
	lastname=replace(request.querystring("lastName"),"'","''")
	company=replace(request.querystring("customerCompany"),"'","''")
	address=replace(request.querystring("address"),"'","''")
	address2=replace(request.querystring("address2"),"'","''")
	city=replace(request.querystring("city"),"'","''")
	stateCode=request.querystring("stateCode")
	pstate=request.QueryString("state")
	zip=request.querystring("zip")
	country=request.querystring("CountryCode")
	phone=request.querystring("phone")
	email=request.querystring("email")
	idcustomer=request.querystring("idcustomer")
	if (country<>"US" AND country<>"CA") AND pstate<>"" then
		stateCode=""
	end if
	call opendb()
	'insert updated information into Customer Table
	query="UPDATE customers SET [name]='"&fname&"',lastName='"&lastname&"',email='"&email&"',customerCompany='"&company&"', address='"&address&"', address2='"&address2&"', city='"&city&"', zip='"&zip&"', stateCode='"&stateCode&"', state='"&pstate&"', CountryCode='"&country&"', phone='"&phone&"' WHERE idCustomer="& idcustomer
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	'insert updated information into Order Table
	query="UPDATE orders SET address='"&address&"', address2='"&address2&"', city='"&city&"', zip='"&zip&"', stateCode='"&stateCode&"', state='"&pstate&"', CountryCode='"&country&"' WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	'Start Special Customer Fields
	pcv_IDCustomer=idcustomer
	if session("cp_nc_custfieldsExists")="YES" then
		pcArr=session("cp_nc_custfields")
		For k=0 to ubound(pcArr,2)
			tmp_cf=""
			tmp_cf=request("custfield_" & pcArr(0,k))
			if not IsNull(tmp_cf) then
				tmp_cf=replace(tmp_cf,"'","''")
			end if
			pcArr(3,k)=tmp_cf
		Next

		pcv_IDCustomer=idcustomer

		For k=0 to ubound(pcArr,2)
			query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
			set rs=connTemp.execute(query)
			if not rs.eof then
				query="UPDATE pcCustomerFieldsValues SET pcCFV_Value='" & pcArr(3,k) & "' WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
			else
				query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & pcv_IDCustomer & "," & pcArr(0,k) & ",'" & pcArr(3,k) & "');"
			end if
			set rs=connTemp.execute(query)
			set rs=nothing
		Next

		session("cp_nc_custfields")=""
	end if
	'End of Special Customer Fields

	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID&"&ActiveTab=6"
end if

'Update Affiliate Earned commissions
IF request.querystring("Submit12")<>"" THEN
	' Admin Comments
	adminComments=replace(request("adminCommentsAffliate"),"'","''")

	'New Commissions
	Dim commByPercentAmount
	commByPercentAmount = Request("optByPercentAmount")
	NewComm             = request("comm1")

	if( commByPercentAmount = "" ) then
		commByPercentAmount = "1"
	end if

	if( commByPercentAmount = "1" ) then
		PrdSales=request("PrdSales")
		NewPay=PrdSales*NewComm/100
	else
		NewPay=NewComm
	end if

	call opendb()

	'insert updated information into Order Table
	query="UPDATE orders SET affiliatePay="&NewPay&" WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
END IF
'End of Update Affiliate Earned commissions



if request.querystring("Submit2")<>"" then
	'update payment information and return to order information
	if request.QueryString("custcardtype")="1" then
		call opendb()
		ccCnt=request.QueryString("ccCnt")
		for i=1 to ccCnt
			pIdCCOrder=request.QueryString("CCID"&i)
			pFormValue=request.QueryString("CCO"&pIdCCOrder)
			query="UPDATE customcardOrders SET strFormValue='"&replace(pFormValue,"'","''")&"' WHERE idCCOrder="&pIdCCOrder&";"
			Set rs=Server.CreateObject("ADODB.Recordset")
			Set rs=conntemp.execute(query)
		next
		call closedb()
	end if
	ccp=request.querystring("ccp")
	if ccp="Y" then
		call opendb()
		'update credit card info
		CCType=request.querystring("CCT")
		cardNumber=request.querystring("CCNum")

		pcv_SecurityPass = pcs_GetSecureKey
		pcv_SecurityKeyID = pcs_GetKeyID

		cardNumber=enDeCrypt(cardNumber, pcv_SecurityPass)
		cardComments=replace(request.querystring("CCcomments"),"'","''")
		if SQL_Format="1" then
			expiration="1/" & request.querystring("CCexpM") & "/" & request.querystring("CCexpY")
		else
			expiration=request.querystring("CCexpM") & "/1/" & request.querystring("CCexpY")
		end if
		' validates expiration
		if DateDiff("d", Month(Now)&"/"&Year(now), request.querystring("CCexpM")&"/"&request.querystring("CCexpY"))<=-1 then
			 conntemp.Close
			 response.redirect "msgb.asp?message="&Server.UrlEncode(dictLanguage.Item(session("language")&"_paymntb_o_6") )
		end if
		'update card
		if scDB="SQL" then
			query="UPDATE creditcards SET cardType='"&CCType&"',cardNumber='"&cardNumber&"', seqcode='na', expiration='"&expiration&"',comments='" &cardComments& "', pcSecurityKeyID = "&pcv_SecurityKeyID&" WHERE idOrder="& qry_ID
		else
			query="UPDATE creditcards SET cardType='"&CCType&"',cardNumber='"&cardNumber&"', seqcode='na',expiration=#"&expiration&"#,comments='" &cardComments& "', pcSecurityKeyID = "&pcv_SecurityKeyID&" WHERE idOrder="& qry_ID
		end if
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)

		call closedb()
	end if

	'redirect back
	response.redirect "Orddetails.asp?id="&qry_ID&"&activetab=5"
end if

if request.querystring("Submit3")<>"" then
	call opendb()
	'update shipping information and return to order information
	shippingFullName=replace(request.querystring("shippingFullName"),"'","''")
	shippingCompany=replace(request.querystring("shippingCompany"),"'","''")
	shippingAddress=replace(request.querystring("shippingAddress"),"'","''")
	shippingAddress2=replace(request.querystring("shippingAddress2"),"'","''")
	shippingcity=replace(request.querystring("shippingcity"),"'","''")
	shippingStateCode=request.querystring("shippingStatecode")
	shippingState=request.QueryString("ShippingState")
	shippingZip=request.querystring("shippingZip")
	shippingPhone=request.querystring("shippingPhone")
	shippingEmail=request.querystring("shippingEmail")
	pOrdPackageNum=request.querystring("ordPackageNum")
	shippingcountry=request.querystring("shippingCountryCode")
	idcustomer=request.querystring("idcustomer")
	if (shippingcountry<>"US" AND shippingcountry<>"CA") AND shippingState<>"" then
		shippingStateCode=""
	end if
	'// Update "no separate shipping address" flag
	pcShowShipAddr=1

	'Update shipping address information in the Orders table
	query="UPDATE orders SET shippingCompany='"&shippingCompany&"', shippingAddress='"&shippingAddress&"', shippingAddress2='"&shippingAddress2&"', shippingcity='"&shippingcity&"', shippingZip='"&shippingZip&"', pcOrd_shippingPhone='"&shippingPhone&"', pcOrd_shippingEmail='"&shippingEmail&"', shippingState='"&shippingstate&"', shippingStateCode='"&shippingStateCode&"', shippingCountryCode='"&shippingcountry&"',shippingFullName='"&shippingFullName&"',ordPackageNum="&pOrdPackageNum&",pcOrd_ShowShipAddr=1 WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID&"&activetab=6"
end if

IF (PayPalProcessOrder=1) OR (request.querystring("Submit4")<>"") OR (request.querystring("Submit4A")<>"") OR (request.querystring("Submit4B")<>"") THEN
	pIdOrder=qry_ID
	call opendb()

	pCheckEmail="YES"

	'Start SDBA
	pcv_AdmComments=""
	pcv_AdmComments=request("AdmComments1")
	if pcv_AdmComments<>"" then
		pcv_AdmComments=replace(pcv_AdmComments,"'","''")
	end if
	pcv_SubmitType=0
	pcv_CustomerReceived=0
	if request.querystring("Submit4A")<>"" then
		pcv_SubmitType=1
		pcv_AdmComments=request("AdmComments1A")
		if pcv_AdmComments<>"" then
			pcv_AdmComments=replace(pcv_AdmComments,"'","''")
		end if
		pcv_CustomerReceived=0
	else
		if request.querystring("Submit4B")<>"" then
			pcv_SubmitType=2
			pcv_AdmComments=request("AdmComments0")
			if pcv_AdmComments<>"" then
				pcv_AdmComments=replace(pcv_AdmComments,"'","''")
			end if
			pcv_CustomerReceived=1
		end if
	end if
	'End SDBA

	if PayPalProcessOrder=1 then
		pcv_CustomerReceived=0
		pcv_AdmComments=""
		pcv_SubmitType=3
	end if

	'Start Process Order and Send Notification E-mails%>
	<!--#include file="inc_ProcessOrder.asp"-->
	<%'End of Process Order and Send Notification E-mails

	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
END IF


IF (request.querystring("Submit4Google")<>"" OR request.querystring("Submit5Google")<>"" OR request.querystring("Submit6Google")<>"" OR request.querystring("Submit7Google")<>"" OR request.querystring("Submit8Google")<>"" OR request.querystring("Submit9Google")<>"") THEN

	pIdOrder=qry_ID
	call opendb()

	pcv_strGoogleMethod = request("GoogleMethod")
	pcv_strGoogleMethod2 = request("GoogleMethod2")
	pcv_strGoogleMethod3 = request("GoogleMethod3")
	pcv_strGoogleMethod4 = request("GoogleMethod4")

	if request.querystring("Submit7Google")<>"" then
		pcv_strGoogleMethod = pcv_strGoogleMethod2
	end if
	if request.querystring("Submit8Google")<>"" then
		pcv_strGoogleMethod = pcv_strGoogleMethod3
	end if
	if request.querystring("Submit9Google")<>"" then
		pcv_strGoogleMethod = pcv_strGoogleMethod4
	end if

	pcv_strReason = pcf_ReverseHTML(request("strReason"))
	pcv_strComment = pcf_ReverseHTML(request("strComment"))
	if pcv_strReason = "" then pcv_strReason = "Merchant Cancelled."
	if pcv_strComment = "" then pcv_strComment = " "

	pcv_strRefundReason = pcf_ReverseHTML(request("strRefundReason"))
	pcv_strRefundComment = pcf_ReverseHTML(request("strRefundComment"))
	if pcv_strRefundReason = "" then pcv_strRefundReason = "Merchant Issued Refund."
	if pcv_strRefundComment = "" then pcv_strRefundComment = " "

	pcv_strBuyerMessage = pcf_ReverseHTML(request("strBuyerMessage"))
	pcv_strShipper = request("strShipper")
	'response.write pcv_strGoogleMethod
	'response.end

	'// Run the Order Management Code
	'Start Google include %>
	<!--#include file="../includes/GoogleCheckout_OrderManagement.asp"-->
	<% 'End Google include


	'redirect back
	call closedb()
	if pcv_strGoogleMethod = "message" then
		response.redirect "Orddetails.asp?id="&qry_ID&"&msg=Your message has been sent to the buyer."
	else
		response.redirect "Orddetails.asp?id="&qry_ID&"&msg=1"
	end if

END IF

Public Function pcf_ReverseHTML(HTMLcode)
	HTMLcode=replace(HTMLcode,"&quot;","""")
	HTMLcode=replace(HTMLcode,"&amp;","&")
	HTMLcode=replace(HTMLcode,"&lt; ","<")
	HTMLcode=replace(HTMLcode,"&gt;",">")
	HTMLcode=replace(HTMLcode,"&copy;","©")
	HTMLcode=replace(HTMLcode,"&reg;","®")
	HTMLcode=replace(HTMLcode,"&#8482;","™")
	HTMLcode=replace(HTMLcode,"&#8220;","“")
	HTMLcode=replace(HTMLcode,"&#8221;","”")
	HTMLcode=replace(HTMLcode,"&#8216;","‘")
	HTMLcode=replace(HTMLcode,"&#8217;","’")
	HTMLcode=replace(HTMLcode,"&#8218;","‚")
	pcf_ReverseHTML=HTMLcode
End Function

'// Resend Order Shipped Email
if (request.querystring("Submit5")<>"") OR (request.querystring("Submit5A")<>"") then

	call opendb()

	'Start SDBA
	if request.querystring("Submit5A")<>"" then
		pcv_SubmitType=1
		pcv_AdmComments=request("AdmComments3A")
		if pcv_AdmComments<>"" then
			pcv_AdmComments=replace(pcv_AdmComments,"'","''")
		end if
	else
		pcv_SubmitType=0
		pcv_AdmComments=request("AdmComments3")
		if pcv_AdmComments<>"" then
			pcv_AdmComments=replace(pcv_AdmComments,"'","''")
		end if
	end if
	query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=3;"
	set rstemp=connTemp.execute(query)
	query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments) VALUES (" & qry_ID & ",3,'" & pcv_AdmComments & "');"
	set rstemp=connTemp.execute(query)
	'End SDBA

	'// Update orderstatus to 4(shipped) and input form variables
	IF pcv_SubmitType=0 THEN
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Pre v3 Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		shipDate=request.querystring("shipDate")
		shipVia=replace(request.querystring("shipVia"),"'","''")
		TrackingNum=request.querystring("TrackingNum")
		if shipDate<>"" then
			if scDateFrmt="DD/MM/YY" then
				err.number=0
				tempShpDt=shipDate
				shpDtArr=split(shipDate,"/")
				shipDate=(shpDtArr(1)&"/"&shpDtArr(0)&"/"&shpDtArr(2))
				if SQL_Format="1" then
					shipDate=tempShpDt
				end if
				if err.number<>0 then
					shipDate=Date()
					if SQL_Format="1" then
						shipDate=Day(shipDate)&"/"&Month(shipDate)&"/"&Year(shipDate)
					else
						shipDate=Month(shipDate)&"/"&Day(shipDate)&"/"&Year(shipDate)
					end if
				end if
			end if
		end if
		if shipDate="" then
			shipDate=Date()
			if SQL_Format="1" then
				shipDate=Day(shipDate)&"/"&Month(shipDate)&"/"&Year(shipDate)
			else
				if scDateFrmt="DD/MM/YY" then
					shipDate=Day(shipDate)&"/"&Month(shipDate)&"/"&Year(shipDate)
				else
					shipDate=Month(shipDate)&"/"&Day(shipDate)&"/"&Year(shipDate)
				end if
			end if
		end if
		if scDB="Access" then
			query="UPDATE orders SET orderstatus=4, shipDate=#"&shipDate&"#,shipVia='"&shipVia&"', TrackingNum='"&TrackingNum&"' WHERE idOrder="& qry_ID
		else
			query="UPDATE orders SET orderstatus=4, shipDate='"&shipDate&"',shipVia='"&shipVia&"', TrackingNum='"&TrackingNum&"' WHERE idOrder="& qry_ID
		end if
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Pre v3 Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	ELSE
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: v3 Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Check Package Shipped
		query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & qry_ID & ";"
		set rs=connTemp.execute(query)
		if rs.eof then
			pcv_SubmitType=0
		end if
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: v3 Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	END IF

	IF pcv_SubmitType=0 THEN
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Pre v3 Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="Select orders.idcustomer, orders.orderdate, orders.shipDate, orders.CountryCode, orders.shippingCountryCode, Orders.pcOrd_ShippingEmail FROM orders WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		pIdCustomer=rs("idcustomer")
		pOrderDate=rs("orderdate")
		todaydate=showDateFrmt(rs("shipDate"))
		' Get country code to determine FedEx tracking URL
		pCountryCode=rs("CountryCode")
		pshippingCountryCode=rs("shippingCountryCode")
		if pshippingCountryCode <> "" then
				strFedExCountryCode=pshippingCountryCode
			else
				strFedExCountryCode=pCountryCode
		end if
		' End get country code to determine FedEx tracking URL
		pShippingEmail=rs("pcOrd_ShippingEmail")

		query="Select name,lastname,email,customercompany FROM customers WHERE idcustomer="& pIdCustomer
		Set rsCust=Server.CreateObject("ADODB.Recordset")
		Set rsCust=conntemp.execute(query)

		' compile emails
		customerShippedEmail=Cstr("")

		'Customized message from store owner
		personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
		personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
		personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
		personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
		personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",rsCust("name")&" "&rsCust("lastname"))
		personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(qry_ID)))
		personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pOrderDate))
		personalmessage=replace(personalmessage,"//","/")
		personalmessage=replace(personalmessage,"http:/","http://")
		personalmessage=replace(personalmessage,"https:/","https://")
		If scShippedEmail<>"" Then
			customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf
			if pcv_AdmComments<>"" then
				customerShippedEmail=customerShippedEmail & replace(pcv_AdmComments,"''","'") & vbcrlf
			end if
			customerShippedEmail=customerShippedEmail & vbcrlf
		end if
		if shipVia <> "" then
			customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_15") &replace(shipVia,"'","''")& vbCrLf
		end if
		customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_16") &ShowDateFrmt(shipDate)& vbCrLf
		if TrackingNum <> "" then
			customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_17") &TrackingNum& vbCrLf
			' Start tracking URL, if any
				if instr(ucase(shipVia),"UPS") then
					customerShippedEmail=customerShippedEmail & scStoreURL & "/" & scPcFolder & "/pc/custUPSTracking.asp?itracknumber=" & TrackingNum & vbCrLf & vbCrLf
					customerShippedEmail=replace(customerShippedEmail,"//","/")
					customerShippedEmail=replace(customerShippedEmail,"http:/","http://")
					else
						if instr(ucase(shipVia),"FEDEX") then
							if ucase(strFedExCountryCode)="US" then
								customerShippedEmail=customerShippedEmail & "http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=" & TrackingNum & vbCrLf & vbCrLf
								else
								customerShippedEmail=customerShippedEmail & "http://www.fedex.com/Tracking?cntry_code=" & strFedExCountryCode & vbCrLf & vbCrLf
							end if
						end if
				end if
			' End tracking URL, if any
		else
			customerShippedEmail=customerShippedEmail & vbCrLf & vbCrLf
		end if
		CustomerShippedEmail=replace(CustomerShippedEmail,"//","/")
		CustomerShippedEmail=replace(CustomerShippedEmail,"http:/","http://")
		CustomerShippedEmail=replace(CustomerShippedEmail,"https:/","https://")
		CustomerShippedEmail=replace(CustomerShippedEmail,"''",chr(39))

		pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_7")
		pEmail=rsCust("email")

		if request.QueryString("sendEmailShip")="YES" then
			call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
			'//Send email to shipping email if it is different and exist
			if trim(pShippingEmail)<>"" AND trim(pShippingEmail)<>trim(pEmail) then
				call sendmail (scCompanyName, scEmail, pShippingEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
			end if
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Pre v3 Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	ELSE
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: v3 Shipping Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & qry_ID & ";"
		set rs=connTemp.execute(query)
		PackCount=0
		if not rs.eof then
			pcPackArr=rs.getRows()
			set rs=nothing
			PackCount=ubound(pcPackArr,2)
			For ipa=0 to PackCount
				pcv_PackageID=pcPackArr(0,ipa)
				pcv_SendCust="1"
				pcv_SendAdmin="0"
				pcv_ResendShip="1"
				if clng(ipa)=clng(PackCount) then
					pcv_LastShip="1"
				end if

				If pcv_LastShip="1" Then
					'// Perform a Google Action
					pcv_strGoogleMethod = "mark" ' // Marks the order shipped at Google
					%> <!--#include file="../includes/GoogleCheckout_OrderManagement.asp"--> <%
				End If
				%>
				<!--#include file="../pc/inc_PartShipEmail.asp"-->
				<%
			Next
		end if
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: v3 Shipping Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	END IF
	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
end if

'Start SDBA
if request.querystring("SubmitA3")<>"" then
	call opendb()
	pcv_DropShipperID=request("pcv_DS_ID")
	pcv_IsSupplier=request("pcv_DS_IsSupplier")
	pcv_AdmComments=request("pcv_DS_Comments")
	if pcv_AdmComments<>"" then
		pcv_AdmComments=replace(pcv_AdmComments,"'","''")
	end if

	query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=4 AND pcDropShipper_ID=" & pcv_DropShipperID & " AND pcACom_IsSupplier=" & pcv_IsSupplier & ";"
	set rstemp=connTemp.execute(query)
	query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier) VALUES (" & qry_ID & ",4,'" & pcv_AdmComments & "'," & pcv_DropShipperID & "," & pcv_IsSupplier & ");"
	set rstemp=connTemp.execute(query)
	if pcv_AdmComments<>"" then
		pcv_AdmComments=replace(pcv_AdmComments,"''","'")
	end if
	pcv_ManualSend=1%>
	<!--#include file="../pc/inc_DropShipperNotificationEmail.asp"-->
	<%
	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID&"&activetab=2"
end if
'End SDBA

'Start SDBA
if request.querystring("SubmitA4")<>"" then
	call opendb()
	pcv_DropShipperID=request("pcv_DS_ID")
	pcv_IsSupplier=request("pcv_DS_IsSupplier")
	pcv_AdmComments=request("pcv_DS_Comments")
	if pcv_AdmComments<>"" then
		pcv_AdmComments=replace(pcv_AdmComments,"'","''")
	end if
	pcv_PackageID=request("pcv_DS_IDPackage")
	pcv_SendCust="1"
	pcv_SendAdmin="0"

	query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID & ";"
	set rstemp=connTemp.execute(query)
	query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier,pcPackageInfo_ID) VALUES (" & qry_ID & ",2,'" & pcv_AdmComments & "'," & pcv_DropShipperID & "," & pcv_IsSupplier & "," & pcv_PackageID & ");"
	set rstemp=connTemp.execute(query)
	%>
	<!--#include file="../pc/inc_PartShipEmail.asp"-->
	<%
	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
end if
'End SDBA

if request.querystring("Submit9")<>"" then
	call opendb()
	'update orderstatus to 4(shipped) and input form variables
	shipDate=request.querystring("shipDate")
	shipVia=replace(request.querystring("shipVia"),"'","''")
	TrackingNum=request.querystring("TrackingNum")
	if shipDate<>"" then
		if scDateFrmt="DD/MM/YY" then
			err.number=0
			tempShpDt=shipDate
			shpDtArr=split(shipDate,"/")
			shipDate=(shpDtArr(1)&"/"&shpDtArr(0)&"/"&shpDtArr(2))
			if SQL_Format="1" then
				shipDate=tempShpDt
			end if
			if err.number<>0 then
				shipDate=Date()
				if SQL_Format="1" then
					shipDate=Day(shipDate)&"/"&Month(shipDate)&"/"&Year(shipDate)
				else
					shipDate=Month(shipDate)&"/"&Day(shipDate)&"/"&Year(shipDate)
				end if
			end if
		end if
	end if
	if shipDate="" then
		shipDate=Date()
		if SQL_Format="1" then
			shipDate=Day(shipDate)&"/"&Month(shipDate)&"/"&Year(shipDate)
		else
			shipDate=Month(shipDate)&"/"&Day(shipDate)&"/"&Year(shipDate)
		end if
	end if
	if scDB="Access" then
		query="UPDATE orders SET orderstatus=4, shipDate=#"&shipDate&"#,shipVia='"&shipVia&"', TrackingNum='"&TrackingNum&"' WHERE idOrder="& qry_ID
	else
		query="UPDATE orders SET orderstatus=4, shipDate='"&shipDate&"',shipVia='"&shipVia&"', TrackingNum='"&TrackingNum&"' WHERE idOrder="& qry_ID
	end if
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	query="Select orders.idcustomer, orders.orderdate, orders.CountryCode, orders.shippingCountryCode, Orders.pcOrd_ShippingEmail FROM orders WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	pIdCustomer=rs("idcustomer")
	pOrderDate=rs("orderdate")
	' Get country code to determine FedEx tracking URL
	pCountryCode=rs("CountryCode")
	pshippingCountryCode=rs("shippingCountryCode")
	if pshippingCountryCode <> "" then
			strFedExCountryCode=pshippingCountryCode
		else
			strFedExCountryCode=pCountryCode
	end if
	' End get country code to determine FedEx tracking URL
	pShippingEmail=rs("pcOrd_ShippingEmail")

	query="Select name,lastname,email,customercompany FROM customers WHERE idcustomer="& pIdCustomer
	Set rsCust=Server.CreateObject("ADODB.Recordset")
	Set rsCust=conntemp.execute(query)

	'send email to customer
	customerShippedEmail=Cstr("")

	'Customized message from store owner
	todaydate=showDateFrmt(now())
	personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
	personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
	personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
	personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
	personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",rsCust("name")&" "&rsCust("lastname"))
	personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(qry_ID)))
	personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pOrderDate))
	personalmessage=replace(personalmessage,"//","/")
	personalmessage=replace(personalmessage,"http:/","http://")
	personalmessage=replace(personalmessage,"https:/","https://")
	If scShippedEmail<>"" Then
		customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf & vbCrLf
	end if
	If shipVia <> "" Then
	customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_15") &replace(shipVia,"'","''")& vbCrLf
	End if
	customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_16") &ShowDateFrmt(shipDate)& vbCrLf
	If TrackingNum <> "" Then
	customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_17") &TrackingNum& vbCrLf & vbCrLf
	' Start tracking URL, if any
		if instr(ucase(shipVia),"UPS") then
			customerShippedEmail=customerShippedEmail & scStoreURL & "/" & scPcFolder & "/pc/custUPSTracking.asp?itracknumber=" & TrackingNum & vbCrLf & vbCrLf
			customerShippedEmail=replace(customerShippedEmail,"//","/")
			customerShippedEmail=replace(customerShippedEmail,"http:/","http://")
			else
				if instr(ucase(shipVia),"FEDEX") then
					if ucase(strFedExCountryCode)="US" then
						customerShippedEmail=customerShippedEmail & "http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=" & TrackingNum & vbCrLf & vbCrLf
						else
						customerShippedEmail=customerShippedEmail & "http://www.fedex.com/Tracking?cntry_code=" & strFedExCountryCode & vbCrLf & vbCrLf
					end if
				end if
		end if
	' End tracking URL, if any
	else
	customerShippedEmail=customerShippedEmail & vbCrLf & vbCrLf
	end if
	CustomerShippedEmail=replace(CustomerShippedEmail,"//","/")
	CustomerShippedEmail=replace(CustomerShippedEmail,"http:/","http://")
	CustomerShippedEmail=replace(CustomerShippedEmail,"https:/","https://")
	CustomerShippedEmail=replace(CustomerShippedEmail,"''",chr(39))

	pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_7")
	'PAYPALEXP EMAIL
	pEmail=rsCust("email")

	if request.QueryString("sendEmailShip2")="YES" then
		call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
		'//Send email to shipping email if it is different and exist
		if trim(pShippingEmail)<>"" AND trim(pShippingEmail)<>trim(pEmail) then
			call sendmail (scCompanyName, scEmail, pShippingEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
		end if
	end if
	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
end if

if PayPalCancelOrder=1 OR request.querystring("Submit6")<>"" OR request.querystring("Submit10")<>"" OR request.querystring("Submit10A")<>"" then
	call opendb()
	'Start SDBA
	if request.querystring("Submit10A")<>"" then
		pcv_SubmitType=1
		pcv_AdmComments=request("AdmComments5A")
		if pcv_AdmComments<>"" then
			pcv_AdmComments=replace(pcv_AdmComments,"'","''")
		end if
	else
		pcv_SubmitType=0
		pcv_AdmComments=request("AdmComments5")
		if pcv_AdmComments<>"" then
			pcv_AdmComments=replace(pcv_AdmComments,"'","''")
		end if
	end if
	query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=5;"
	set rstemp=connTemp.execute(query)
	query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments) VALUES (" & qry_ID & ",5,'" & pcv_AdmComments & "');"
	set rstemp=connTemp.execute(query)
	'End SDBA

	'find out original status first
	query="SELECT orderstatus,paymentCode FROM orders WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	porigstatus=rs("orderstatus")
	pPaymentCode=rs("paymentCode")

	'update orderstatus to 5(canceled) and input form variables
	query="UPDATE orders SET orderstatus=5 WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	if pcv_SubmitType=0 then
	Set rs=conntemp.execute(query)
	end if
	'update startDate
	StartDate=CDate("1/1/2000")
	if scDB="Access" then
		query="UPDATE DPRequests SET StartDate=#"&StartDate&"# WHERE idOrder="& qry_ID
	else
		query="UPDATE DPRequests SET StartDate='"&StartDate&"' WHERE idOrder="& qry_ID
	end if
	Set rs=Server.CreateObject("ADODB.Recordset")
	if pcv_SubmitType=0 then
	Set rs=conntemp.execute(query)
	end if

	'GGG Add-on start

		query="UPDATE pcGCOrdered SET pcGO_Status=0 WHERE pcGO_idOrder="& qry_ID
		Set rs=conntemp.execute(query)

		query="Select total,pcOrd_GcCode,pcOrd_GcUsed,pcOrd_GCDetails,pcOrd_GCAmount from orders WHERE idOrder="& qry_ID
		Set rs=conntemp.execute(query)

		ototal=rs("total")
		GCDetails=rs("pcOrd_GCDetails")
		GCAmount=rs("pcOrd_GCAmount")
		if GCAmount="" OR IsNull(GCAmount) then
			GCAmount=0
		end if
		pGiftCode=rs("pcOrd_GcCode")
		pGiftUsed=rs("pcOrd_GcUsed")

		if GCDetails<>"" then
			ototal=cdbl(ototal)+cdbl(GCAmount)
			query="update orders set total=" & ototal & ",pcOrd_GcCode='',pcOrd_GcUsed=0,pcOrd_GCDetails='',pcOrd_GCAmount=0 WHERE idOrder="& qry_ID
			Set rs=conntemp.execute(query)

			GCArry=split(GCDetails,"|g|")
			intArryCnt=ubound(GCArry)

			for k=0 to intArryCnt

			if GCArry(k)<>"" then
				GCInfo = split(GCArry(k),"|s|")
				if GCInfo(2)="" OR IsNull(GCInfo(2)) then
					GCInfo(2)=0
				end if
				pGiftCode=GCInfo(0)
				pGiftUsed=GCInfo(2)

				query="select pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
				Set rs=conntemp.execute(query)
				if not rs.eof then
					pGCAmount=rs("pcGO_Amount")
					if pGCAmount="" OR IsNull(pGCAmount) then
						pGCAmount=0
					end if
					pGCStatus=rs("pcGO_Status")
					if pGCStatus="" OR IsNull(pGCStatus) then
						pGCStatus=0
					end if

					pGCAmount=cdbl(pGCAmount)+cdbl(pGiftUsed)

					if pGCAmount>"0" then
						pGCStatus=1
					end if

					query="update pcGCOrdered set pcGO_Amount=" & pGCAmount & ",pcGO_Status=" & pGCStatus & " WHERE pcGO_GcCode='" & pGiftCode & "'"
					Set rs=conntemp.execute(query)
				end if
			end if

			Next
		end if 'Have Gift Code

		'Increase Remaining Products of Gift Registry
		query="Select pcPO_EPID,quantity from ProductsOrdered WHERE idOrder="& qry_ID
		Set rs=conntemp.execute(query)
		do while not rs.eof
			geID=rs("pcPO_EPID")
			if geID<>"" then
			else
			geID="0"
			end if
			gQty=rs("quantity")
			if gQty<>"" then
			else
			gQty="0"
			end if
			if geID<>"0" then
			query="Update pcEvProducts set pcEP_HQty=pcEP_HQty-" & gQty & " WHERE pcEP_ID="& geID
			Set rs1=conntemp.execute(query)
			end if
			rs.MoveNext
		loop

		set rs=nothing
		set rs1=nothing

	'GGG Add-on end

	'set any Auth or PFP orders to captured
	select case pPaymentCode
		case "PFLink", "PFPro", "PFPRO", "PFLINK"
			query="UPDATE pfporders SET captured=1 WHERE idOrder="& qry_ID
			Set rs=Server.CreateObject("ADODB.Recordset")
			if pcv_SubmitType=0 then
			Set rs=conntemp.execute(query)
			end if
		case "Authorize"
			query="UPDATE authorders SET captured=1 WHERE idOrder="& qry_ID
			Set rs=Server.CreateObject("ADODB.Recordset")
			if pcv_SubmitType=0 then
			Set rs=conntemp.execute(query)
			end if
	end select
	'update reward pts.
	query="Select * FROM orders WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	pIdCustomer=rs("idcustomer")

		piRewardPoints=rs("iRewardPoints")
		piRewardRefId=rs("iRewardRefId")
		piRewardPointsRef=rs("iRewardPointsRef")
		piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
		'take away points from refferer if any points were awarded. if order was processed
		If porigstatus<>"2" then
			If piRewardRefId>0 AND piRewardPointsRef>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & piRewardRefId
				set rsCust=conntemp.execute(query)
				iAccrued=rsCust("iRewardPointsAccrued") - piRewardPointsRef
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & piRewardRefId
				if pcv_SubmitType=0 then
					set rsCust=conntemp.Execute(query)
				end if
			end if
			'take away accrued points from customer if any points were accrued
			If piRewardPointsCustAccrued>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
				set rsCust=conntemp.execute(query)
				iAccrued=rsCust("iRewardPointsAccrued") - piRewardPointsCustAccrued
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & pIdCustomer
				if pcv_SubmitType=0 then
					set rsCust=conntemp.Execute(query)
				end if
			end if
			If piRewardPoints>0 then
				query="SELECT iRewardPointsUsed, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
				set rsCust=conntemp.execute(query)
				iUsed=rsCust("iRewardPointsUsed") - piRewardPoints
				query="UPDATE customers SET iRewardPointsUsed=" & iUsed & " WHERE idCustomer=" & pIdCustomer
				if pcv_SubmitType=0 then
					set rsCust=conntemp.Execute(query)
				end if
			end if
		end if

	query="Select name,lastname,email,customercompany FROM customers WHERE idcustomer="& pIdCustomer
	Set rsCust=Server.CreateObject("ADODB.Recordset")
	Set rsCust=conntemp.execute(query)
	if porigstatus="3" or porigstatus="2" then
		query="SELECT idproduct,quantity,idconfigSession FROM ProductsOrdered WHERE ProductsOrdered.idOrder="& qry_ID
		set rsOrderDetails=conntemp.execute(query)
		Do While Not rsOrderDetails.EOF
			pidproduct=rsOrderDetails("idproduct")
			pqty=rsOrderDetails("quantity")
			idconfig=rsOrderDetails("idconfigSession")
			'check if stock is ignored or not
			query="SELECT noStock FROM products WHERE idProduct="&pIdProduct
			set rsStockObj=server.CreateObject("ADODB.RecordSet")
			set rsStockObj=conntemp.execute(query)
			pNoStock=rsStockObj("noStock")
			set rsStockObj=nothing
			'---------------
			' increase stock
			'---------------
			query="SELECT stock, sales, description FROM products WHERE idProduct="&pidproduct
			set rsStockObj=server.CreateObject("ADODB.RecordSet")
			set rsStockObj=conntemp.execute(query)
			if err.number <> 0 then
				call closedb()
				response.redirect "techErr.asp?error=95"&Server.Urlencode("Error in processOrder. Error: "&err.description)
			end if
			if pNoStock=0 then
				query="UPDATE products SET stock=stock+"&pqty&" WHERE idProduct="&pidproduct
				if pcv_SubmitType=0 then
				set rsStockObj=conntemp.execute(query)
				end if
				if err.number <> 0 then
					call closedb()
					response.redirect "techErr.asp?error=110"&Server.Urlencode("Error in processOrder. Error: "&err.description)
				end if
				'Update BTO Items & Additional Charges stock and sales
				IF (idconfig<>"") and (idconfig<>"0") then
				query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
				set rs1=conntemp.execute(query)
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")

					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock+" &QtyArr(k)*pqty&",sales=sales-" &QtyArr(k)*pqty&" WHERE idProduct=" &PrdArr(k)
							if pcv_SubmitType=0 then
							set rs1=conntemp.execute(query)
							end if
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")

					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock+" &pqty&",sales=sales-" &pqty&" WHERE idProduct=" &CPrdArr(k)
							if pcv_SubmitType=0 then
							set rs1=conntemp.execute(query)
							end if
						end if
					next
				end if
			END IF
			'End Update BTO Items & Additional Charges

			end if
			set rsStockObj=nothing
			'--------------
			' end increase stock
			'--------------
			'--------------
			'update sales
			'--------------
			query="UPDATE products SET sales=sales-"&pqty&" WHERE idProduct="&pidproduct
			if pcv_SubmitType=0 then
			set rsSalesObj=conntemp.execute(query)
			end if
			if err.number <> 0 then
				call closedb()
				response.redirect "techErr.asp?error=110"&Server.Urlencode("Error in processOrder. Error: "&err.description)
			end if
			set rsSalesObj=nothing
			'--------------
			' end update sales
			'--------------
			set rsStockObj=nothing
			rsOrderDetails.MoveNext
		loop
	end if

	'send email to customer
	customerCancelledEmail=Cstr("")

	'Customized message from store owner
	If scCancelledEmail<>"" Then
		todaydate=showDateFrmt(now())
		personalmessage=replace(scCancelledEmail,"<br>", vbCrlf)
		personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
		personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
		personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
		personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",rsCust("name")&" "&rsCust("lastname"))
		personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(qry_ID)))
		personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(rs("orderDate")))
		personalmessage=replace(personalmessage,"//","/")
		personalmessage=replace(personalmessage,"http:/","http://")
		personalmessage=replace(personalmessage,"https:/","https://")
		customerCancelledEmail=customerCancelledEmail & vbCrLf & personalmessage & vbCrLf
		if pcv_AdmComments<>"" then
			customerCancelledEmail=customerCancelledEmail & vbCrLf & replace(pcv_AdmComments,"''","'") & vbCrLf
		end if
		customerCancelledEmail=replace(customerCancelledEmail,"''",chr(39))
		pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_8")
		'PAYPALEXP EMAIL
		pEmail=rsCust("email")
		if request.querystring("Submit6")<>"" AND request.QueryString("sendEmailCanc")="YES" then
			call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerCancelledEmail, "&quot;", chr(34)))
		end if
		if (request.querystring("Submit10")<>"" AND request.QueryString("sendEmailCanc2")="YES") OR (request("Submit10A")<>"") then
			call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerCancelledEmail, "&quot;", chr(34)))
		end if
	end if

	'Start SDBA
	if pcv_SubmitType=0 then%>
	<!--#include file="inc_DropShipperCancelOrderEmail.asp"-->
	<%end if
	'END SDBA

	' Order summary starts here ...
	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
end if

if request.querystring("Submit7")<>"" then
	call opendb()
	'update orderstatus and input form variables
	query="UPDATE orders SET orderstatus="&request.querystring("resetstat")&" WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	pcv_intOrdReset=request.querystring("resetstat")
	if (Cint(pcv_intOrdReset)>4) AND (Cint(pcv_intOrdReset)<>7)  AND (Cint(pcv_intOrdReset)<>8) then
		'query="UPDATE orders SET DPs=0 WHERE idOrder="& qry_ID
		'Set rs=Server.CreateObject("ADODB.Recordset")
		'Set rs=conntemp.execute(query)
		StartDate=CDate("1/1/2000")
		if scDB="Access" then
			query="UPDATE DPRequests SET StartDate=#"&StartDate&"# WHERE idOrder="& qry_ID
		else
			query="UPDATE DPRequests SET StartDate='"&StartDate&"' WHERE idOrder="& qry_ID
		end if
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
	end if

	if Cint(pcv_intOrdReset)=2 then

	query="UPDATE orders SET DPs=0 WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	query="Delete From DPLicenses WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	query="Delete From DPRequests WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	'GGG Add-on start

	query="UPDATE orders SET pcOrd_GCs=0 WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	query="Delete From pcGCOrdered WHERE pcGO_idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	'Decrease Remaining Products of Gift Registry
		query="Select pcPO_EPID,quantity from ProductsOrdered WHERE idOrder="& qry_ID
		Set rs=conntemp.execute(query)
		do while not rs.eof
			geID=rs("pcPO_EPID")
			if geID<>"" then
			else
			geID="0"
			end if
			gQty=rs("quantity")
			if gQty<>"" then
			else
			gQty="0"
			end if
			if geID<>"0" then
			query="Update pcEvProducts set pcEP_HQty=pcEP_HQty+" & gQty & " WHERE pcEP_ID="& geID
			Set rs1=conntemp.execute(query)
			end if
			rs.MoveNext
		loop

	set rs=nothing
	set rs1=nothing

	'GGG Add-on end

	end if

	if (Cint(pcv_intOrdReset)=3) or (Cint(pcv_intOrdReset)=4)  or (Cint(pcv_intOrdReset)=7) or (Cint(pcv_intOrdReset)=8) then
	query="select idproduct,idconfigSession from ProductsOrdered WHERE idOrder="& qry_ID
	set rs=connTemp.execute(query)
	DPOrder="0"
	do while not rs.eof
		pIdProduct=rs("idproduct")
		tmpidConfig=rs("idconfigSession")
		query="select downloadable from products where idproduct=" & pIdProduct
		set rstemp=connTemp.execute(query)
		if not rstemp.eof then
			pdownloadable=rstemp("downloadable")
			if (pdownloadable<>"") and (pdownloadable="1") then
				DPOrder="1"
			end if
		end if
		set rstemp=nothing
		'Find downloadable items in BTO configuration
		if tmpidConfig<>"" AND tmpidConfig>"0" then
			query="SELECT stringProducts,stringQuantity,stringCProducts FROM configSessions WHERE idconfigSession=" & tmpidConfig & ";"
			set rs1=connTemp.execute(query)
			if not rs1.eof then
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")

					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="SELECT idproduct FROM Products WHERE idProduct=" & PrdArr(k) & " AND Downloadable=1;"
							set rs1=conntemp.execute(query)
							if not rs1.eof then
								DPOrder="1"
							end if
							set rs1=nothing
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")
					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="SELECT idproduct FROM Products WHERE idProduct=" & CPrdArr(k) & " AND Downloadable=1;"
							set rs1=conntemp.execute(query)
							if not rs1.eof then
								DPOrder="1"
							end if
							set rs1=nothing
						end if
					next
				end if
			end if
			set rs1=nothing
		end if
	rs.moveNext
	loop
	set rs=nothing

	query="UPDATE orders SET DPs=" & DPOrder & " WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	StartDate=Date()
	if scDB="Access" then
	query="UPDATE DPRequests SET StartDate=#"&StartDate&"# WHERE idOrder="& qry_ID
	else
	query="UPDATE DPRequests SET StartDate='"&StartDate&"' WHERE idOrder="& qry_ID
	end if
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	'GGG Add-on start

	query="select idproduct from ProductsOrdered WHERE idOrder="& qry_ID
	set rstemp=connTemp.execute(query)
	pGCs="0"
	do while not rstemp.eof
		query="select pcprod_GC from products where idproduct=" & rstemp("idproduct")
		set rs=connTemp.execute(query)
		if not rs.eof then
			pGC=rs("pcprod_GC")
			if (pGC<>"") and (pGC="1") then
				pGCs="1"
			end if
		end if
		rstemp.moveNext
	loop
	set rstemp=nothing

	query="UPDATE orders SET pcOrd_GCs=" & pGCs & " WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	query="UPDATE pcGCOrdered SET pcGO_Status=1 WHERE pcGO_idOrder="& qry_ID
	Set rs=conntemp.execute(query)

	'Decrease Remaining Products of Gift Registry
		query="Select pcPO_EPID,quantity from ProductsOrdered WHERE idOrder="& qry_ID
		Set rs=conntemp.execute(query)
		do while not rs.eof
			geID=rs("pcPO_EPID")
			if geID<>"" then
			else
			geID="0"
			end if
			gQty=rs("quantity")
			if gQty<>"" then
			else
			gQty="0"
			end if
			if geID<>"0" then
			query="Update pcEvProducts set pcEP_HQty=pcEP_HQty+" & gQty & " WHERE pcEP_ID="& geID
			Set rs1=conntemp.execute(query)
			end if
			rs.MoveNext
		loop

	set rs=nothing
	set rs1=nothing

	'GGG Add-on end

	end if

	if porigstatus="5" then
		query="SELECT idproduct,quantity,idconfigSession FROM ProductsOrdered WHERE ProductsOrdered.idOrder="& qry_ID
		set rsOrderDetails=conntemp.execute(query)
		Do While Not rsOrderDetails.EOF
			pidproduct=rsOrderDetails("idproduct")
			pqty=rsOrderDetails("quantity")
			idconfig=rsOrderDetails("idconfigSession")
			'check if stock is ignored or not
			query="SELECT noStock FROM products WHERE idProduct="&pIdProduct
			set rsStockObj=server.CreateObject("ADODB.RecordSet")
			set rsStockObj=conntemp.execute(query)
			pNoStock=rsStockObj("noStock")
			set rsStockObj=nothing
			'---------------
			' decrease stock
			'---------------
			query="SELECT stock, sales, description FROM products WHERE idProduct="&pidproduct
			set rsStockObj=server.CreateObject("ADODB.RecordSet")
			set rsStockObj=conntemp.execute(query)
			if err.number <> 0 then
				call closedb()
				response.redirect "techErr.asp?error=95"&Server.Urlencode("Error in processOrder. Error: "&err.description)
			end if
			if pNoStock=0 then
				query="UPDATE products SET stock=stock-"&pqty&" WHERE idProduct="&pidproduct
				set rsStockObj=conntemp.execute(query)
				if err.number <> 0 then
					call closedb()
					response.redirect "techErr.asp?error=110"&Server.Urlencode("Error in processOrder. Error: "&err.description)
				end if
				'Update BTO Items & Additional Charges stock and sales
				IF (idconfig<>"") and (idconfig<>"0") then
				query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
				set rs1=conntemp.execute(query)
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")

					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &QtyArr(k)*pqty&",sales=sales+" &QtyArr(k)*pqty&" WHERE idProduct=" &PrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")

					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &pqty&",sales=sales+" &pqty&" WHERE idProduct=" &CPrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
			END IF
			'End Update BTO Items & Additional Charges

			end if
			set rsStockObj=nothing
			'--------------
			' end decrease stock
			'--------------
			'--------------
			'update sales
			'--------------
			query="UPDATE products SET sales=sales+"&pqty&" WHERE idProduct="&pidproduct
			set rsSalesObj=conntemp.execute(query)
			if err.number <> 0 then
				call closedb()
				response.redirect "techErr.asp?error=110"&Server.Urlencode("Error in processOrder. Error: "&err.description)
			end if
			set rsSalesObj=nothing
			'--------------
			' end update sales
			'--------------
			set rsStockObj=nothing
			rsOrderDetails.MoveNext
		loop
	end if

	'Cancel Order
	if pcv_intOrdReset="5" then
		query="SELECT idproduct,quantity,idconfigSession FROM ProductsOrdered WHERE ProductsOrdered.idOrder="& qry_ID
		set rsOrderDetails=conntemp.execute(query)
		Do While Not rsOrderDetails.EOF
			pidproduct=rsOrderDetails("idproduct")
			pqty=rsOrderDetails("quantity")
			idconfig=rsOrderDetails("idconfigSession")
			'check if stock is ignored or not
			query="SELECT noStock FROM products WHERE idProduct="&pIdProduct
			set rsStockObj=server.CreateObject("ADODB.RecordSet")
			set rsStockObj=conntemp.execute(query)
			pNoStock=rsStockObj("noStock")
			set rsStockObj=nothing
			'---------------
			' increase stock
			'---------------
			query="SELECT stock, sales, description FROM products WHERE idProduct="&pidproduct
			set rsStockObj=server.CreateObject("ADODB.RecordSet")
			set rsStockObj=conntemp.execute(query)
			if err.number <> 0 then
				call closedb()
				response.redirect "techErr.asp?error=95"&Server.Urlencode("Error in processOrder. Error: "&err.description)
			end if
			if pNoStock=0 then
				query="UPDATE products SET stock=stock+"&pqty&" WHERE idProduct="&pidproduct
				set rsStockObj=conntemp.execute(query)
				if err.number <> 0 then
					call closedb()
					response.redirect "techErr.asp?error=110"&Server.Urlencode("Error in processOrder. Error: "&err.description)
				end if
				'Update BTO Items & Additional Charges stock and sales
				IF (idconfig<>"") and (idconfig<>"0") then
				query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
				set rs1=conntemp.execute(query)
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")

					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock+" &QtyArr(k)*pqty&",sales=sales-" &QtyArr(k)*pqty&" WHERE idProduct=" &PrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")

					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock+" &pqty&",sales=sales-" &pqty&" WHERE idProduct=" &CPrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
			END IF
			'End Update BTO Items & Additional Charges

			end if
			set rsStockObj=nothing
			'--------------
			' end increase stock
			'--------------
			'--------------
			'update sales
			'--------------
			query="UPDATE products SET sales=sales-"&pqty&" WHERE idProduct="&pidproduct
			set rsSalesObj=conntemp.execute(query)
			if err.number <> 0 then
				call closedb()
				response.redirect "techErr.asp?error=110"&Server.Urlencode("Error in processOrder. Error: "&err.description)
			end if
			set rsSalesObj=nothing
			'--------------
			' end update sales
			'--------------
			set rsStockObj=nothing
			rsOrderDetails.MoveNext
		loop

		'Start SDBA
		if pcv_SubmitType=0 then%>
		<!--#include file="inc_DropShipperCancelOrderEmail.asp"-->
		<%end if
		'END SDBA

	end if

	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
end if

if request.querystring("Submit8")<>"" then
	call opendb()
	'update orderstatus to 6(Return) and input form variables
	returnDate=request.querystring("returnDate")
	if returnDate<>"" then
		if SQL_Format="1" then
			returnDate=Day(returnDate)&"/"&Month(returnDate)&"/"&Year(returnDate)
		else
			returnDate=Month(returnDate)&"/"&Day(returnDate)&"/"&Year(returnDate)
		end if
	end if
	if returnDate="" then
		returnDate=Date()
		if SQL_Format="1" then
			returnDate=Day(returnDate)&"/"&Month(returnDate)&"/"&Year(returnDate)
		else
			returnDate=Month(returnDate)&"/"&Day(returnDate)&"/"&Year(returnDate)
		end if
	end if
	returnReason=replace(request.querystring("returnReason"),"'","''")
	if returnReason="" then
		returnReason="None given."
	end if
	if scDB="Access" then
		query="UPDATE orders SET orderstatus=6, returnDate=#"&returnDate&"#,returnReason='"&returnReason&"' WHERE idOrder="& qry_ID
	else
		query="UPDATE orders SET orderstatus=6, returnDate='"&returnDate&"',returnReason='"&returnReason&"' WHERE idOrder="& qry_ID
	end if
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	query="SELECT ProductsOrdered.idproduct, quantity FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& qry_ID
	set rsOrderDetails=conntemp.execute(query)
	Do While Not rsOrderDetails.EOF
		pqty=rsOrderDetails("quantity")
		pidproduct=rsOrderDetails("idproduct")
		'--------------
		'update sales
		'--------------
		query="UPDATE products SET sales=sales-"&pqty&" WHERE idProduct="&pidproduct
		set rsSalesObj=conntemp.execute(query)
		if err.number <> 0 then
			call closedb()
			response.redirect "techErr.asp?error=110"&Server.Urlencode("Error in processOrder. Error: "&err.description)
		end if
		set rsSalesObj=nothing
		'--------------
		' end update sales
		'--------------
		rsOrderDetails.MoveNext
	loop

		query="Select * FROM orders WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		pIdCustomer=rs("idcustomer")
		piRewardPoints=rs("iRewardPoints")
		piRewardRefId=rs("iRewardRefId")
		piRewardPointsRef=rs("iRewardPointsRef")
		piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
		'take away points from refferer if any points were awarded.
		If piRewardRefId>0 AND piRewardPointsRef>0 then
			query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & piRewardRefId
			set rsCust=conntemp.execute(query)
			iAccrued=rsCust("iRewardPointsAccrued") - piRewardPointsRef
			query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & piRewardRefId
			set rsCust=conntemp.Execute(query)
		end if
		'take away accrued points from customer if any points were accrued
		If piRewardPointsCustAccrued>0 then
			query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
			set rsCust=conntemp.execute(query)
			iAccrued=rsCust("iRewardPointsAccrued") - piRewardPointsCustAccrued
			query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & pIdCustomer
			set rsCust=conntemp.Execute(query)
		end if
		'give points back if they were used on purchase
		If piRewardPoints>0 then
			query="SELECT iRewardPointsUsed, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
			set rsCust=conntemp.execute(query)
			iUsed=rsCust("iRewardPointsUsed") - piRewardPoints
			query="UPDATE customers SET iRewardPointsUsed=" & iUsed & " WHERE idCustomer=" & pIdCustomer
			set rsCust=conntemp.Execute(query)
		end if

	call closedb()
	'redirect back
	response.redirect "Orddetails.asp?id="&qry_ID
end if

if request.querystring("Submit11")<>"" then

	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open scDSN
	pord_OrderName=replace(request("ord_OrderName"),"'","''")
	DF1=request("DF1")
	TF1=request("TF1")
	session("admin_DF1")=DF1
	if scDateFrmt="DD/MM/YY" then
		DFArry=split(DF1,"/")
		tDF=(DFArry(1)&"/"&DFArry(0)&"/"&DFArry(2))
		session("admin_DF1")=tDF
	end if
	session("admin_TF1")=TF1
	if session("admin_DF1")<>"" then
		If Not IsDate(session("admin_DF1")) then
			response.redirect "Orddetails.asp?id="&qry_ID&"&r=0&msg=Invalid " & DF1Label & " value."
		end if
		If CDate(session("admin_DF1"))<Date() then
			response.redirect "Orddetails.asp?id="&qry_ID&"&r=0&msg=Invalid " & DF1Label & " value."
		end if
	end if

	if session("admin_TF1")<>"" then
		If Not IsDate(session("admin_TF1")) then
			response.redirect "Orddetails.asp?id="&qry_ID&"&r=0&msg=Invalid " & TF1Label & " value."
		end if
	end if

	If DTCheck="1" then
		if (session("admin_DF1")="") and (session("admin_TF1")="") then
		response.redirect "Orddetails.asp?id="&qry_ID&"&r=0&msg=You must deliver the date/time."
		end if
		if session("admin_DF1")<>"" then
			DF2=CDate(session("admin_DF1"))
			if DF2-Date()<=0 then
				response.redirect "Orddetails.asp?id="&qry_ID&"&r=0&msg=Your Date/Time must be greater than 24 hours from the current date/time."
			else
				if (DF2-Date()=1) then
					if session("admin_TF1")<>"" then
						TF2=CDate(session("admin_TF1"))
						if TF2<time() then
							response.redirect "Orddetails.asp?id="&qry_ID&"&r=0&msg=Your Date/Time must be greater than 24 hours from the current date/time."
						end if
					end if
				end if
			end if
		else
			if session("admin_TF1")<>"" then
				TF2=CDate(session("admin_TF1"))
				if TF2-time()<24 then
					response.redirect "Orddetails.asp?id="&qry_ID&"&r=0&msg=Your Date/Time must be greater than 24 hours from the current date/time."
				end if
			end if
		end if
	end if

	ord_DeliveryDate=session("admin_DF1") & " " & session("admin_TF1")
	ord_DeliveryDate=trim(ord_DeliveryDate)

	if not isDate(ord_DeliveryDate) then
		ord_DeliveryDate="1/1/1900"
	end if

	ord_DeliveryDate=CDate(ord_DeliveryDate)

	query="update orders set ord_OrderName='" & pord_OrderName & "',ord_DeliveryDate="
	if scDB="SQL" then
		query=query & "'" & ord_DeliveryDate  & "'"
	else
		query=query & "#" & ord_DeliveryDate  & "#"
	end if
	query=query & " where idOrder=" & qry_ID

	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conn.execute(query)

	conn.Close
	response.redirect "Orddetails.asp?id="&qry_ID
end if

'// START: PRE-V3 Shipping Managment
if (request.querystring("Submit14")<>"") then
	call opendb()
	'update orderstatus to 4(shipped) and input form variables
	shipDate=request.querystring("shipDate")
	shipVia=replace(request.querystring("shipVia"),"'","''")
	TrackingNum=request.querystring("TrackingNum")
	if shipDate<>"" then
		if scDateFrmt="DD/MM/YY" then
			err.number=0
			tempShpDt=shipDate
			shpDtArr=split(shipDate,"/")
			shipDate=(shpDtArr(1)&"/"&shpDtArr(0)&"/"&shpDtArr(2))
			if SQL_Format="1" then
				shipDate=tempShpDt
			end if
			if err.number<>0 then
				shipDate=Date()
				if SQL_Format="1" then
					shipDate=Day(shipDate)&"/"&Month(shipDate)&"/"&Year(shipDate)
				else
					shipDate=Month(shipDate)&"/"&Day(shipDate)&"/"&Year(shipDate)
				end if
			end if
		end if
	end if
	if shipDate="" then
		shipDate=Date()
		if SQL_Format="1" then
			shipDate=Day(shipDate)&"/"&Month(shipDate)&"/"&Year(shipDate)
		else
			shipDate=Month(shipDate)&"/"&Day(shipDate)&"/"&Year(shipDate)
		end if
	end if
	if scDB="Access" then
		query="UPDATE orders SET orderstatus=4, shipDate=#"&shipDate&"#,shipVia='"&shipVia&"', TrackingNum='"&TrackingNum&"' WHERE idOrder="& qry_ID
	else
		query="UPDATE orders SET orderstatus=4, shipDate='"&shipDate&"',shipVia='"&shipVia&"', TrackingNum='"&TrackingNum&"' WHERE idOrder="& qry_ID
	end if
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)

	query="Select orders.idcustomer, orders.orderdate, orders.CountryCode, orders.shippingCountryCode, Orders.pcOrd_ShippingEmail FROM orders WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	pIdCustomer=rs("idcustomer")
	pOrderDate=rs("orderdate")
	' Get country code to determine FedEx tracking URL
	pCountryCode=rs("CountryCode")
	pshippingCountryCode=rs("shippingCountryCode")
	if pshippingCountryCode <> "" then
			strFedExCountryCode=pshippingCountryCode
		else
			strFedExCountryCode=pCountryCode
	end if
	' End get country code to determine FedEx tracking URL
	pShippingEmail=rs("pcOrd_ShippingEmail")

	query="Select name,lastname,email,customercompany FROM customers WHERE idcustomer="& pIdCustomer
	Set rsCust=Server.CreateObject("ADODB.Recordset")
	Set rsCust=conntemp.execute(query)

	' compile emails
	customerShippedEmail=Cstr("")

	'Customized message from store owner
	todaysdate = showDateFrmt(now())
	personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
	personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
	personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
	personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
	personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",rsCust("name")&" "&rsCust("lastname"))
	personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(qry_ID)))
	personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pOrderDate))
	personalmessage=replace(personalmessage,"//","/")
	personalmessage=replace(personalmessage,"http:/","http://")
	personalmessage=replace(personalmessage,"https:/","https://")
	If scShippedEmail<>"" Then
		customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf
		if pcv_AdmComments<>"" then
			customerShippedEmail=customerShippedEmail & replace(pcv_AdmComments,"''","'") & vbcrlf
		end if
		customerShippedEmail=customerShippedEmail & vbcrlf
	end if
	if shipVia <> "" then
	customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_15") &replace(shipVia,"'","''")& vbCrLf
	end if
	customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_16") &ShowDateFrmt(shipDate)& vbCrLf
	if TrackingNum <> "" then
	customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_17") &TrackingNum& vbCrLf
	' Start tracking URL, if any
		if instr(ucase(shipVia),"UPS") then
			customerShippedEmail=customerShippedEmail & scStoreURL & "/" & scPcFolder & "/pc/custUPSTracking.asp?itracknumber=" & TrackingNum & vbCrLf & vbCrLf
			customerShippedEmail=replace(customerShippedEmail,"//","/")
			customerShippedEmail=replace(customerShippedEmail,"http:/","http://")
			else
				if instr(ucase(shipVia),"FEDEX") then
					if ucase(strFedExCountryCode)="US" then
						customerShippedEmail=customerShippedEmail & "http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=" & TrackingNum & vbCrLf & vbCrLf
						else
						customerShippedEmail=customerShippedEmail & "http://www.fedex.com/Tracking?cntry_code=" & strFedExCountryCode & vbCrLf & vbCrLf
					end if
				end if
		end if
	' End tracking URL, if any
	else
	customerShippedEmail=customerShippedEmail & vbCrLf & vbCrLf
	end if
	CustomerShippedEmail=replace(CustomerShippedEmail,"//","/")
	CustomerShippedEmail=replace(CustomerShippedEmail,"http:/","http://")
	CustomerShippedEmail=replace(CustomerShippedEmail,"https:/","https://")
	CustomerShippedEmail=replace(CustomerShippedEmail,"''",chr(39))

	pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_7")
	pEmail=rsCust("email")

	if request.QueryString("sendEmailShip14")="YES" then
		call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
		'//Send email to shipping email if it is different and exist
		if trim(pShippingEmail)<>"" AND trim(pShippingEmail)<>trim(pEmail) then
			call sendmail (scCompanyName, scEmail, pShippingEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
		end if
	end if
	'redirect back
	call closedb()
	response.redirect "Orddetails.asp?id="&qry_ID
end if
'// END: PRE-V3 Shipping Managment



'Start SDBA
if (request.querystring("SubmitA1")<>"") then
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open scDSN
	pcv_PaymentStatus=request("pcv_PaymentStatus")
	if pcv_PaymentStatus="" then
		pcv_PaymentStatus=0
	end if
	query="UPDATE Orders SET pcOrd_PaymentStatus=" & pcv_PaymentStatus & " WHERE idorder=" & qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conn.execute(query)
	conn.Close
	response.redirect "Orddetails.asp?id="&qry_ID
end if
'End SDBA

if (request.querystring("SubmitA2")<>"") then
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open scDSN
	pcv_OrderStatus=request("pcv_OrderStatus")
	if pcv_OrderStatus="" then
		pcv_OrderStatus=0
	end if
	if pcv_OrderStatus<>0 then
	query="UPDATE Orders SET orderStatus=" & pcv_OrderStatus & " WHERE idorder=" & qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conn.execute(query)
	end if
	conn.Close
	response.redirect "Orddetails.asp?id="&qry_ID
end if
'End SDBA

if request.querystring("submitReSendGCRec")<>"" OR request.querystring("submitReSendGCRecA")<>"" then
	call opendb()
	GC_ReName=getUserInput(request("GC_RecName"),0)
	GC_ReEmail=getUserInput(request("GC_RecEmail"),0)
	if request.querystring("submitReSendGCRecA")<>"" then
		query="SELECT pcOrd_GcReMsg FROM orders WHERE idOrder=" & qry_ID & ";"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		if not rs.eof then
			GC_ReMsg=rs("pcOrd_GcReMsg")
		else
			GC_ReMsg=""
		end if
		Set rs=nothing
		query="UPDATE orders SET pcOrd_GcReName='" & GC_ReName & "',pcOrd_GcReEmail='" & GC_ReEmail & "',pcOrd_GcReMsg='" & GC_ReMsg & "' WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		Set rs=nothing
	else
		GC_ReMsg=getUserInput(request("GC_RecMsg"),0)
		query="UPDATE orders SET pcOrd_GcReName='" & GC_ReName & "',pcOrd_GcReEmail='" & GC_ReEmail & "' WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		Set rs=nothing
	end if

	ReciEmail=""

	query="select idproduct from ProductsOrdered WHERE idOrder="& qry_ID
	pidorder=qry_ID
	set rs11=connTemp.execute(query)
	do while not rs11.eof
		query="select products.Description,pcGCOrdered.pcGO_GcCode,pcGc.pcGc_EOnly from Products,pcGc,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGC.pcGc_IDProduct=products.idproduct and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& qry_ID
		set rs=connTemp.execute(query)

		if not rs.eof then
			pIdproduct=rs11("idproduct")
			pName=rs("Description")
			pCode=rs("pcGO_GcCode")
			pEOnly=rs("pcGc_EOnly")

				query="select pcGO_Amount,pcGO_GcCode,pcGO_ExpDate from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & pidorder
				set rs19=connTemp.execute(query)

				do while not rs19.eof
				pAmount=rs19("pcGO_Amount")
				if pAmount<>"" then
				else
				pAmount="0"
				end if

				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_68") & scCurSign & money(pAmount) & vbcrlf

				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_69") & rs19("pcGO_GcCode") & vbcrlf
				pExpDate=rs19("pcGO_ExpDate")

				if year(pExpDate)="1900" then
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_45b") & vbcrlf
				else
				if scDateFrmt="DD/MM/YY" then
				pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
				else
				pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
				end if
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_70") & pExpDate & vbcrlf
				end if
				if pEOnly="1" then
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_71") & vbcrlf
				end if
				ReciEmail=ReciEmail & vbcrlf
				rs19.movenext
				loop

		end if
	rs11.MoveNext
	loop
	set rs11=nothing

	query="SELECT customers.name,customers.lastname,customers.email,orders.pcOrd_GcReName,orders.pcOrd_GcReEmail,orders.pcOrd_GcReMsg FROM customers INNER JOIN Orders ON customers.idcustomer=orders.idcustomer WHERE idOrder="& qry_ID
	set rs11=connTemp.execute(query)

	if not rs11.eof then
		pCustomerFullName=rs11("name") & " " & rs11("lastname")
		pCustomerFullNamePlusEmail=pCustomerFullName & " (" & rs11("email") & ")"
		GcReName=rs11("pcOrd_GcReName")
		GcReEmail=rs11("pcOrd_GcReEmail")
		if request.querystring("submitReSendGCRecA")<>"" then
			GcReMsg=rs11("pcOrd_GcReMsg")
		else
			GcReMsg=GC_ReMsg
		end if

		if GcReEmail<>"" then
			if GcReName<>"" then
			else
				GcReName=GcReEmail
			end if
			ReciEmail1=replace(dictLanguage.Item(Session("language")&"_sendMail_66"),"<recipient name>",GcReName)
			ReciEmail2=replace(dictLanguage.Item(Session("language")&"_sendMail_67"),"<customer name>",pCustomerFullNamePlusEmail)
			if GcReMsg<>"" then
				ReciEmail3=replace(dictLanguage.Item(Session("language")&"_sendMail_72"),"<customer name>",pCustomerFullNamePlusEmail) & vbcrlf & GcReMsg & vbcrlf
			else
				ReciEmail3=""
			end if
			ReciEmail=ReciEmail1 & vbcrlf & vbcrlf & ReciEmail2 & vbcrlf & vbcrlf & ReciEmail & ReciEmail3
			ReciEmail=ReciEmail & vbcrlf & scCompanyName & vbCrLf & scStoreURL & vbcrlf & vbCrLf
			call sendmail (scCompanyName, scEmail, GcReEmail,pCustomerFullName & dictLanguage.Item(Session("language")&"_sendMail_73"), replace(ReciEmail, "&quot;", chr(34)))
		end if
	end if
	set rs11=nothing

	call closedb()
	'redirect back
	response.redirect "Orddetails.asp?id="&qry_ID
end if

'===== functions=====
Public Function FixedField(ByVal Width, ByVal Justify, ByVal Text)
	Select Case True
		Case Width < Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Left(Text, Width)
				Case Justify="R"
					FixedField=Right(Text, Width)
				Case Else
			End Select

		Case Width=Len(Text)
			FixedField=Text

		Case Width > Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Text & String(Width - Len(Text), " ")
				Case Justify="R"
					FixedField=String(Width - Len(Text), " ") & Text
				Case Else
			End Select

	End Select

End Function
%>