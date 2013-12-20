<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>

<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<!-- #include file="Centinel_Config.asp"-->
<% 'Gateway specific files %>
<%
Dev_Testmode = 2
'SB S
msg=getUserInput(request.querystring("message"),0)
msg=replace(msg, "&lt;BR&gt;", "<BR>")
msg=replace(msg, "&lt;br&gt;", "<br>")
msg=replace(msg, "&lt;b&gt;", "<b>")
msg=replace(msg, "&lt;/b&gt;", "</b>")
msg=replace(msg, "&lt;/font&gt;", "</font>")
msg=replace(msg, "&lt;a href", "<a href")
msg=replace(msg, "&gt;Back&lt;/a&gt;", ">Back</a>")
msg=replace(msg, "&lt;font", "<font")
msg=replace(msg, "&gt;<b>Error&nbsp;</b>:", "><b>Error&nbsp;</b>:")
msg=replace(msg, "&gt;&lt;img src=", "><img src=")
msg=replace(msg, "&gt;&lt;/a&gt;", "></a>")
msg=replace(msg, "&gt;<b>", "><b>")
msg=replace(msg, "&lt;/a&gt;", "</a>")
msg=replace(msg, "&gt;View Cart", ">View Cart")
msg=replace(msg, "&gt;Continue", ">Continue")
msg=replace(msg, "&lt;u>", "<u>")
msg=replace(msg, "&lt;/u>", "</u>")
msg=replace(msg, "&lt;ul&gt;", "<ul>")
msg=replace(msg, "&lt;/ul&gt;", "</ul>")
msg=replace(msg, "&lt;li&gt;", "<li>")
msg=replace(msg, "&lt;/li&gt;", "</li>")
msg=replace(msg, "&gt;", ">") 
msg=replace(msg, "&lt;", "<") 
'SB E
%>
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td><img src="images/checkout_bar_step5.gif" alt=""></td>
	</tr>
	<tr>
		<td class="pcSpacer"></td>
	</tr>
	<tr>
		<td>
		<% if session("GWOrderDone")="YES" then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://")
			session("GWOrderDone")=""
			response.redirect tempURL
		end if

		dim connTemp, rs

		session("redirectPage")="gwAuthorizeAIM.asp" %>

		<% Dim pcCustIpAddress
		pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

		dim tempURL
		If scSSL="" OR scSSL="0" Then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://") 
		Else
			tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://")
		End If
		
		' Get Order ID
		if session("GWOrderId")="" then
			session("GWOrderId")=request("idOrder")
		end if
		
		pcGatewayDataIdOrder=session("GWOrderID")
		%>
		<!--#include file="pcGateWayData.asp"-->
		<% session("idCustomer")=pcIdCustomer
		pcv_IncreaseCustID=(scCustPre + int(pcIdCustomer)) %>

		<%
		call opendb()

		query="SELECT x_Type,x_Login,x_Password,x_Curcode,x_AIMType,x_CVV,x_testmode,x_secureSource FROM authorizeNet Where id=1"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		x_Type=rs("x_Type")
		x_Login=rs("x_Login")
		'decrypt
		x_Login=enDeCrypt(x_Login, scCrypPass)
		x_Password=rs("x_Password")
		'decrypt
		x_Password=enDeCrypt(x_Password, scCrypPass)
		x_Curcode=rs("x_Curcode")
		x_AIMType=rs("x_AIMType")
		x_CVV=rs("x_CVV")
		x_testmode=rs("x_testmode")
		x_secureSource=rs("x_secureSource")
		x_TypeArray=Split(x_Type,"||")
		x_TransType=x_TypeArray(0)
		set rs=nothing
		
		If Request.Form("PaymentGWCentinel")="Go" OR request.QueryString("Centinel")<>"" Then %>
			<%
			'SB S
			'// By pass AIM if the immediate order value is 0
			If pcBillingTotal<0 Then
				pcBillingTotal=0
			End If
			
			If (pcIsSubscription) AND (pcBillingTotal=0) Then

				session("reqCardNumber")=getUserInput(request.Form("cardNumber"),16)
				session("reqExpMonth")=getUserInput(request.Form("expMonth"),0)
				session("reqExpYear")=getUserInput(request.Form("expYear"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("CVV"),4)					
				pExpiration=getUserInput(request("expMonth"),0) & "/01/" & getUserInput(request("expYear"),0)				
				'// Validates expiration
			    if DateDiff("d", Month(Now)&"/01/"&Year(now),pExpiration)<=-1 then
				 	response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_6")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
			    end if
		       	'// Validate card
			    if not IsCreditCard(session("reqCardNumber"), request.form("x_Card_Type")) then
					response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_5")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
			    end if 
				session("GWAuthCode")	= "AUTH-ARB"
				session("GWTransId")	= "0"

				Response.Redirect("gwReturn.asp?s=true&gw=AIM&GWError=1")
				Response.End

			End If
			'SB E

			if Request.Form("PaymentGWCentinel")="Go" then
				session("reqCardNumber")=getUserInput(request.Form("cardNumber"),16)
				session("reqExpMonth")=getUserInput(request.Form("expMonth"),0)
				session("reqExpYear")=getUserInput(request.Form("expYear"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("CVV"),4)
				
				If x_CVV="1" Then
					if not isnumeric(session("reqCVV")) or len(session("reqCVV")) < 3 or len(session("reqCVV")) > 4 then
				  response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_7")&dictLanguage.Item(Session("language")&"_paymntb_c_4")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
					end If
				End if
			end if

			'//Check to see if Centinel is active for this gateway:
			pcPay_CentByPass=0
			If request.QueryString("centinel")="Y" Then
				pcPay_CentByPass=1
			End If

			dim pcPay_CardType
			pcPay_CardType="YES"

			If determineCardType(session("reqCardNumber"))<>"MASTERCARD" AND determineCardType(session("reqCardNumber"))<>"VISA" then
				pcPay_CentByPass=1
				pcPay_CardType="NO"
			end if
			%>
			<%
			If pcPay_Cent_Active=1 AND pcPay_CentByPass=0 Then %>
				<!-- #include file="Centinel_ThinClient.asp"-->
				<%
				'==========================================================================================
				'= Cardinal Commerce (http://www.cardinalcommerce.com)
				'= process the Lookup message
				'==========================================================================================

				dim strTransactionType
				dim centinelRequest, centinelResponse

				strTransactionType = request("txn_type")

				dtTempExpYear=Cstr(session("reqExpYear"))
				if len(dtTempExpYear)=2 then
					dtTempExpYear="20"&dtTempExpYear
				end if
				GWOrderTotal=replace(money(pcBillingTotal),".","")
				GWOrderTotal=replace(GWOrderTotal,",","")

				set centinelRequest = Server.CreateObject("Scripting.Dictionary")
				centinelRequest.Add "Version", Cstr(MessageVersion)
				centinelRequest.Add "MsgType", "cmpi_lookup"
				centinelRequest.Add "ProcessorId", Cstr(ProcessorId)
				centinelRequest.Add "MerchantId", Cstr(MerchantId)
				centinelRequest.Add "TransactionPwd", Cstr(TransactionPwd)
				centinelRequest.Add "TransactionType", Cstr("C")
				centinelRequest.Add "TransactionMode", Cstr("S")
				centinelRequest.Add "OrderNumber", Cstr(session("GWOrderID"))
				centinelRequest.Add "OrderDescription", Cstr(replace(scCompanyName,",","-") & " Order: " & session("GWOrderID"))
				centinelRequest.Add "Amount", Cstr(GWOrderTotal)
				'centinelRequest.Add "RawAmount", Cstr(GWOrderTotal)
				centinelRequest.Add "CurrencyCode", Cstr("840")
				centinelRequest.Add "UserAgent", Cstr(Request.ServerVariables("HTTP_USER_AGENT"))
				centinelRequest.Add "BrowserHeader", Cstr(Request.ServerVariables("HTTP_ACCEPT"))
				centinelRequest.Add "IPAddress", Cstr(pcCustIpAddress)
				centinelRequest.Add "Item_Name_1", Cstr("Web Order")
				centinelRequest.Add "Item_Desc_1", Cstr("Online Product")
				centinelRequest.Add "Item_Price_1", Cstr(GWOrderTotal)
				centinelRequest.Add "Item_Quantity_1", Cstr(1)
				centinelRequest.Add "CardNumber", Cstr(session("reqCardNumber"))
				centinelRequest.Add "CardExpMonth", Cstr(session("reqExpMonth"))
				centinelRequest.Add "CardExpYear", Cstr(dtTempExpYear)

				'=====================================================================================
				' Send the XML Msg to the MAPS Server
				' SendHTTP will send the cmpi_lookup message to the MAPS Server (requires fully qualified URL)
				' The Response is the CentinelResponse Object
				'=====================================================================================

				set centinelResponse = sendMsg(centinelRequest, Cstr(TRANSACTIONURL), CLng(ResolveTimeout), CLng(ConnectTimeout), CLng(SendTimeout), CLng(ReceiveTimeout))

				Session("Centinel_Enrolled") = centinelResponse.item("Enrolled")
				Session("Centinel_ErrorNo") = centinelResponse.item("ErrorNo")
				Session("Centinel_ErrorDesc") = centinelResponse.item("ErrorDesc")
				Session("Centinel_TransactionId") = centinelResponse.item("TransactionId")
				Session("Centinel_OrderId") = centinelResponse.item("OrderId")
				Session("Centinel_ACSURL") = centinelResponse.item("ACSUrl")
				Session("Centinel_Payload") = centinelResponse.item("Payload")
				Session("Centinel_TermURL") = TermURL
				Session("Centinel_TransactionType") = strTransactionType
				'==========================================================================================
				' Determine how to proceed with the transaction - Handle Business Logic
				'==========================================================================================
				If Session("Centinel_Enrolled") = "N" OR Session("Centinel_Enrolled") = "U" Then
					'RESPONSE.WRITE request.QueryString("centinel")
					'RESPONSE.END
				Else
					If Session("Centinel_ErrorNo") = "0" AND Session("Centinel_Enrolled") = "Y" Then
	
						'==========================================================================================
						' Cardholder is enrolled, proceed to redirect page.
						'==========================================================================================
						Response.redirect ("Centinel_Frame.asp")
					Else
						'==================================================================================
						' No redirect to ACS
						'==================================================================================
	
						Session("Centinel_Message") = "Your transaction errored prior to completion. Please provide another form of payment."
						Session("Centinel_Message") = Session("Centinel_Message") + "ErrorNo(s) [" + Session("Centinel_ErrorNo") + "]<br>"
						Session("Centinel_Message") = Session("Centinel_Message") + "ErrorDesc(s) [" + Session("Centinel_ErrorDesc") + "]<br>"
	
						If CStr(DEBUGOUTPUT) = "True" Then
							Response.redirect ("gwAuthorizeAIM.asp?centinel=results")
						Else
							Response.redirect ("msg.asp")
						End If
					End If
			End If
				Set centinelResponse = nothing
				Set centinelRequest = nothing
			End If

			if pcBillingTotal > 0 then
				If request.QueryString("centinel")="Y"	Then
					'//Centinel Results

					if len(Session("Centinel_Message")) > 0 then %>
						<br/><br/>
							<font color="red"><b>Sample Message : <%=Session("Centinel_Message")%></b></font>
						<br/><br/>
					<% end if %>

					<table id="results">
					<tr>
						<td>Enrolled : </td>
						<td><%=Session("Centinel_Enrolled")%></td>
					</tr>
					<tr>
						<td>PAResStatus :</td>
						<td><%=Session("Centinel_PAResStatus")%></td>
					</tr>
					<tr>
						<td>SignatureVerification : </td>
						<td><%=Session("Centinel_SignatureVerification")%></td>
					</tr>
					<tr>
						<td>Transaction Id :</td>
						<td><%=Session("Centinel_TransactionId")%></td>
					</tr>
					<tr>
						<td>Order Id :</td>
						<td><%=Session("Centinel_OrderId")%></td>
					</tr>
					<tr>
						<td>Error No :</td>
						<td><%=Session("Centinel_ErrorNo")%></td>
					</tr>
					<tr>
						<td>Error Desc : </td>
						<td><%=Session("Centinel_ErrorDesc")%></td>
					</tr>
					<tr>
						<td>Reason Code :</td>
						<td><%=Session("Centinel_ReasonCode#")%></td>
					</tr>
					<tr>
						<td>Reason Desc : </td>
						<td><%=Session("Centinel_ReasonDesc#")%></td>
					</tr>
				</table><br><br>

				<% End If

				If scSSL="" OR scSSL="0" Then
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
					tempURL=replace(tempURL,"https:/","https://")
					tempURL=replace(tempURL,"http:/","http://")
				Else
					tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
					tempURL=replace(tempURL,"https:/","https://")
					tempURL=replace(tempURL,"http:/","http://")
				End If

				Dim objXMLHTTP, xml

				'Send the request to the Authorize.NET processor.
				stext="x_Version=3.1"
			stext=stext & "&x_Delim_Data=True"
			stext=stext & "&x_Delim_char=,"
			If x_testmode="1" Then
				stext=stext & "&x_Test_Request=True"
			Else
					stext=stext & "&x_Test_Request=False"
				End If
				stext=stext & "&x_relay_response=FALSE"
				stext=stext & "&x_Login=" & x_Login
				stext=stext & "&x_Tran_Key=" & x_Password
				stext=stext & "&x_method=CC"
				if x_testmode="1" then
					stext=stext & "&x_Amount=1"
					stext=stext & "&x_Card_Num=42222222222222222"
					stext=stext & "&x_Exp_Date=1214"
				else
					stext=stext & "&x_Amount=" & pcBillingTotal
					stext=stext & "&x_Card_Num=" & session("reqCardNumber")
					stext=stext & "&x_Exp_Date=" & session("reqExpMonth")&session("reqExpYear")
				end if
				If x_CVV="1" Then
					stext=stext & "&x_Card_Code=" & session("reqCVV")
				End If
			stext=stext & "&x_customer_ip=" & pcCustIpAddress
			stext=stext & "&x_Type=" & x_TransType
			stext=stext & "&x_Currency_Code=" & x_Curcode
			stext=stext & "&x_Description=" & replace(scCompanyName,",","-") & " Order: " & session("GWOrderID")
			stext=stext & "&x_Invoice_Num=" & session("GWOrderID")
			stext=stext & "&x_Cust_ID=" & pcv_IncreaseCustID
			stext=stext & "&x_first_name=" & pcBillingFirstName
			stext=stext & "&x_last_name=" & pcBillingLastName
			stext=stext & "&x_company=" & replace(pcBillingCompany,",","||")
			stext=stext & "&x_address=" & replace(pcBillingAddress,",","||")
			stext=stext & "&x_city=" & pcBillingCity
			stext=stext & "&x_state=" & pcBillingState
			stext=stext & "&x_zip=" & pcBillingPostalCode
			stext=stext & "&x_country=" & pcBillingCountryCode
			stext=stext & "&x_phone=" & pcBillingPhone
			stext=stext & "&x_email=" & pcCustomerEmail
			if pcShippingFullName<>"" then
				pcShippingNameArry=split(pcShippingFullName, " ")
				if ubound(pcShippingNameArry)>0 then
					pcShippingFirstName=pcShippingNameArry(0)
					if ubound(pcShippingNameArry)>1 then
						 tmpShipFirstName = pcShippingFirstName&" "
						 pcShippingLastName = replace(pcShippingFullName,tmpShipFirstName,"")
					else
						pcShippingLastName=pcShippingNameArry(1)
					end if
				else
					pcShippingFirstName=pcShippingFullName
					pcShippingLastName=pcShippingFullName
				end if
			else
				pcShippingFirstName=pcBillingFirstName
				pcShippingLastName=pcBillingLastName
			end if
			stext=stext & "&x_Ship_To_First_Name=" & pcShippingFirstName '?
			stext=stext & "&x_Ship_To_Last_Name=" & pcShippingLastName '?
				stext=stext & "&x_ship_to_company=" & replace(pcShippingCompany,",","||") '?
			stext=stext & "&x_Ship_To_Address=" & replace(pcShippingAddress,",","||")
			stext=stext & "&x_Ship_To_City=" & pcShippingCity
			stext=stext & "&x_Ship_To_State=" & pcShippingState
			stext=stext & "&x_Ship_To_Zip=" & pcShippingPostalCode
			stext=stext & "&x_Ship_To_Country=" & pcShippingCountryCode
	
			If pcPay_Cent_Active=1 AND pcPay_CentByPass=1 AND pcPay_CardType="YES" Then
				stext=stext & "&x_authentication_indicator=" & Session("Centinel_ECI")
				stext=stext & "&x_cardholder_authentication_value=" & Session("Centinel_CAVV")
			End If

			'Send the transaction info as part of the querystring
			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			'SB S
			if Dev_Testmode = 1 Then
				xml.open "POST", "https://test.authorize.net/gateway/transact.dll?"& stext & "", false
			else
				xml.open "POST", "https://secure.authorize.net/gateway/transact.dll?"& stext & "", false
			end if
			'SB E
			
			xml.send ""
				strStatus = xml.Status

				'store the response
				strRetVal = xml.responseText
				Set xml = Nothing

			'//Save Centinel Details
			call opendb()
			if Session("Centinel_Enrolled")&""<>"" then
				pCentIdOrderF = (clng(session("GWOrderID"))-clng(scpre))
				query = "INSERT INTO pcPay_Centinel_Orders ([pcPay_CentOrd_OrderID],[pcPay_CentOrd_Enrolled],[pcPay_CentOrd_ErrorNo],[pcPay_CentOrd_ErrorDesc],[pcPay_CentOrd_PAResStatus],[pcPay_CentOrd_SignatureVerification],[pcPay_CentOrd_EciFlag],[pcPay_CentOrd_Cavv],[pcPay_CentOrd_rErrorNo],[pcPay_CentOrd_rErrorDesc]) VALUES ("&pCentIdOrderF&",'"&Session("Centinel_Enrolled")&"','"&Session("Centinel_ErrorNo")&"','"&Session("Centinel_ErrorDesc")&"','"&Session("Centinel_PAResStatus")&"','"&Session("Centinel_SignatureVerification")&"','"&Session("Centinel_ECI")&"','"&Session("Centinel_CAVV")&"','"&Session("Centinel_ReasonCode")&"','"&Session("Centinel_ReasonDesc")&"')"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
			end if

			'/////////////////////////////////////////////////////
			'// Create Log of response and save in includes
			'/////////////////////////////////////////////////////
			dim authLogging
			authLogging=1 'Change to 1 to log

			if authLogging=1 then

				if PPD="1" then
					pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/AUTHLOG.txt")
				else
					pcStrFileName=Server.Mappath ("../includes/AUTHLOG.txt")
				end if

				dim strFileName
				dim fs
				dim OutputFile

				'Specify directory and file to store silent post information
				strFileName = pcStrFileName
				Set fs = CreateObject("Scripting.FileSystemObject")
				Set OutputFile = fs.OpenTextFile (strFileName, 8, True)

				OutputFile.WriteLine now()
				OutputFile.WriteLine "===================================="
				OutputFile.WriteLine "Request from ProductCart: " & stext
				OutputFile.WriteBlankLines(1)


				OutputFile.WriteLine "Response from Authorize.Net: " & strRetVal
				OutputFile.WriteBlankLines(2)

				OutputFile.Close
			end if
			'/////////////////////////////////////////////////////
			'// End - Create Log of response and save in includes
			'/////////////////////////////////////////////////////

				strArrayVal = split(strRetVal, ",", -1)
				session("x_response_code") = strArrayVal(0)
				session("x_response_subcode") = strArrayVal(1)
				session("x_response_reason_code") = strArrayVal(2)
				session("x_response_reason_text") = strArrayVal(3)
			session("GWAuthCode")								= strArrayVal(4)    '6 digit approval code
				session("AVSCode") = strArrayVal(5)
			session("GWTransId")								= strArrayVal(6)    'transaction id
			session("x_invoice_num")						= strArrayVal(7)
			session("x_description")						= strArrayVal(8)
			session("x_amount")									= strArrayVal(9)
			session("x_method")									= strArrayVal(10)
			session("x_type")										= strArrayVal(11)
			session("x_cust_id")								= strArrayVal(12)
			session("x_first_name")							= strArrayVal(13)
			session("x_last_name")							= strArrayVal(14)
				session("x_company") = replace(strArrayVal(15),"||",",")
			session("x_address")								= replace(strArrayVal(16),"||",",")
			session("x_city")										= strArrayVal(17)
			session("x_state")									= strArrayVal(18)
			session("x_zip") 										= strArrayVal(19)
			session("x_country")								= strArrayVal(20)
			session("x_phone")									= strArrayVal(21)
			session("x_fax")										= strArrayVal(22)
			session("x_email")									= strArrayVal(23)
			session("x_ship_to_first_name")			= strArrayVal(24)
			session("x_ship_to_last_name")			= strArrayVal(25)
				session("x_ship_to_company") = replace(strArrayVal(26),"||",",")
				session("x_ship_to_address ") = replace(strArrayVal(27),"||",",")
				session("x_ship_to_city") = strArrayVal(28)
				session("x_ship_to_state") = strArrayVal(29)
				session("x_ship_to_zip") = strArrayVal(30)
				session("x_ship_to_country") = strArrayVal(31)
				session("x_tax_value") = strArrayVal(32)
				session("x_duty_value") = strArrayVal(33)
				session("x_freight_value") = strArrayVal(34)
				session("x_tax_exempt") = strArrayVal(35)
				session("x_po_number") = strArrayVal(36)
				session("x_md5hash") = strArrayVal(37)
				session("CVV2Code") = strArrayVal(38)
				session("x_cav_response") = strArrayVal(39)

				'===========================================================================================
				' Remove Centinel Session Contents Once the Transaction Is Complete
				'===========================================================================================
				Session.Contents.Remove("Centinel_ACSURL")
				Session.Contents.Remove("Centinel_CAVV")
				Session.Contents.Remove("Centinel_ECI")
				Session.Contents.Remove("Centinel_Enrolled")
				Session.Contents.Remove("Centinel_ErrorDesc")
				Session.Contents.Remove("Centinel_ErrorNo")
				Session.Contents.Remove("Centinel_Message")
				Session.Contents.Remove("Centinel_OrderId")
				Session.Contents.Remove("Centinel_PAResStatus")
				Session.Contents.Remove("Centinel_PIType")
				Session.Contents.Remove("Centinel_Payload")
				Session.Contents.Remove("Centinel_SignatureVerification")
				Session.Contents.Remove("Centinel_TermURL")
				Session.Contents.Remove("Centinel_TransactionId")
				Session.Contents.Remove("Centinel_TransactionType")
				Session.Contents.Remove("Centinel_XID")

				'Check the ErrorCode to make sure that the component was able to talk to the authorization network
				If (strStatus <> 200) Then
					Response.Write "An error occurred during processing. Please try again later."
					Response.end
				Else
					'save and update order
					If session("x_response_code") = 1 Then
						call opendb()

						pcv_SecurityPass = pcs_GetSecureKey
						pcv_SecurityKeyID = pcs_GetKeyID

					dim pCardNumber, pCardNumber2
					pCardNumber=session("reqCardNumber")
					pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)
					
					'save info in authOrders if AUTH_CAPTURE
					pcTrueOrdnum=(int(session("x_invoice_num"))-scpre)
					If x_TransType="AUTH_ONLY" Then
						
						query="INSERT INTO authorders (idOrder, amount, paymentmethod, transtype, authcode, ccnum, ccexp, idCustomer, fname, lname, address, zip, captured, pcSecurityKeyID) VALUES ("&pcTrueOrdnum&", "&session("x_amount")&", 'CC', '"&x_TransType&"', '"&session("GWAuthCode")&"', '"&pCardNumber2&"', '"&session("reqExpMonth")&session("reqExpYear")&"', "&session("idCustomer")&", '"&replace(session("x_first_name"),"'","''")&"', '"&replace(session("x_last_name"),"'","''")&"', '"&replace(session("x_address"),"'","''")&"', '"&session("x_zip")&"', 0, "&pcv_SecurityKeyID&");"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
					
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						
					End If
					
					call closedb()
						
					set rs=nothing
					
					'clear all card sessions before redirect
					'SB S
					if not pcIsSubscription then
						session("reqCardNumber")=""
						session("reqExpMonth")=""
						session("reqExpYear")=""
						session("reqCVV")=""
					End if
					'SB E
	
					session("x_response_code")=""
						session("x_response_subcode")=""
						session("x_response_reason_code")=""
						session("x_response_reason_text")=""
						session("x_avs_code")=""
						session("x_description")=""
						session("x_amount")=""
						session("x_cust_id")=""
						session("x_first_name")=""
						session("x_last_name")=""
						session("x_company")=""
						session("x_address")=""
						session("x_city")=""
					session("x_state")=""
					session("x_zip")=""
					session("x_country")=""
					session("x_phone")=""
					session("x_fax")=""
					session("x_email")=""
					session("x_ship_to_first_name")=""
					session("x_ship_to_last_name")=""
					session("x_ship_to_company")=""
					session("x_ship_to_address")=""
					session("x_ship_to_city")=""
					session("x_ship_to_state")=""
					session("x_ship_to_zip")=""
					session("x_ship_to_country")=""
	
					Response.redirect "gwReturn.asp?s=true&gw=AIM"
			
					ElseIf session("x_response_code")<>1 Then
						If session("x_response_code") = "3" AND session("x_response_subcode") = "11" Then
							session("x_response_reason_text") = "A duplicate transaction has been submitted.<br>A transaction with identical amount and credit card information was submitted two minutes prior. Click the button below to return to your payment page and wait at least 2 minutes before submitting your payment information again."
						End If
						response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;"&session("x_response_code")&"</b>: "& session("x_response_reason_text")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&session("x_amount")&"""><img src="""&rslayout("back")&""" border=0></a>")
						response.end
					End If
				End If
			end if '// if pcBillingTotal > 0 then
			'End If 'Redirected back after Centinel approval or form submit
		Else
				call opendb()

				query="SELECT x_Type,x_CVV FROM authorizeNet Where id=1;"
				set rs=Server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				x_Type=rs("x_Type")
				x_CVV=rs("x_CVV")
				M="0"
				V="0"
				A="0"
				D="0"
				%>	
	
				<% If x_CVV="1" Then 
					response.write "<form method=""POST"" action=""gwAuthorizeAIM.asp"" name=""form1"" class=""pcForms"">"
				Else %>
					<form action="gwAuthorizeAIM.asp" method="POST" name="form1" class="pcForms">
				<% End If %>

				<input type="hidden" name="PaymentGWCentinel" value="Go">
					<table class="pcShowContent">
			
					<% if Msg<>"" then %>
						<tr valign="top"> 
							<td colspan="2">
								<div class="pcErrorMessage"><%=Msg%></div>
							</td>
						</tr>
					<% end if %>
					<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></p></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_1")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingFirstName&" "&pcBillingLastName%></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingAddress%></p></td>
					</tr>
					<% if pcBillingAddress2<>"" then %>
					<tr>
						<td colspan="2"><p><%=pcBillingAddress2%></p></td>
					</tr>
					<% end if %>
					<tr>
						<td colspan="2"><p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<% if x_testmode="1" then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td><p>Card Type:</p></td>
						<td>
							<select name="x_Card_Type">
							<% 	x_TypeArray=Split(x_Type,"||")
							If ubound(x_TypeArray)=1 Then
								x_Type2=x_TypeArray(1)
								cardTypeArray=split(x_Type2,", ")
								i=ubound(cardTypeArray)
								cardCnt=0
								do until cardCnt=i+1
									cardVar=cardTypeArray(cardCnt)
									select case cardVar
										case "V"
											response.write "<option value=""V"" selected>Visa</option>"
											cardCnt=cardCnt+1
										case "M" 
											response.write "<option value=""M"">MasterCard</option>"
											cardCnt=cardCnt+1
										case "A"
											response.write "<option value=""A"">American Express</option>"
											cardCnt=cardCnt+1
										case "D"
											response.write "<option value=""D"">Discover</option>"
											cardCnt=cardCnt+1
									end select
								loop
							End If %>
							</select>
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td> 
							<input type="text" name="CardNumber" value="">
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="expMonth">
								<option value="01">1</option>
								<option value="02">2</option>
								<option value="03">3</option>
								<option value="04">4</option>
								<option value="05">5</option>
								<option value="06">6</option>
								<option value="07">7</option>
								<option value="08">8</option>
								<option value="09">9</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="12">12</option>
							</select>
							<% dtCurYear=Year(date()) %>
							&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%> 
							<select name="expYear">
								<option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,2)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,2)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,2)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,2)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,2)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,2)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,2)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,2)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,2)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,2)%>"><%=dtCurYear+10%></option>
							</select>
						</td>
					</tr>
					<% If x_CVV="1" Then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
								<input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4">
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% End If %>

					<%
					'SB S
					if pcIsSubscription Then
					%>
					<tr>
						<td><p><%=scSBLang7%></p></td>
						<td><p><%= money((pcBillingTotal + pcBillingSubScriptionTotal))%></p></td>
					</tr>
                    <% Else %>
					<tr>
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><p><%= money(pcBillingTotal)%></p></td>
					</tr>
					<%
					End if
					'SB E
					%>

					<%'SB S
					If pcIsSubscription Then %>
					<tr>
						<td><p><%=scSBLang8%></p></td>
						<td><p><!--#include file="inc_sb_widget.asp"--></p></td>
					</tr>
					<% End If
					'SB E %>
					
					<% if pcPay_Cent_Active=1 then %>
						<script LANGUAGE="JavaScript">
						function popUp(url) {
							popupWin=window.open(url,"win",'toolbar=0,location=0,directories=0,status=1,menubar=1,scrollbars=1,width=570,height=450');
							self.name = "mainWin"; }
						</script>
							
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						 <tr> 
							<td colspan="2">
							<p><a href='javascript:popUp("pcPay_Cent_mcsc.asp")'><img src='images/pc_mcsc.gif' alt="MasterCard SecureCode - Learn More" border='0' /></a>&nbsp;&nbsp;<a href='javascript:popUp("pcPay_Cent_vbv.asp")'><img src='images/pc_vbv.gif' alt="Verified by Visa - Learn More" border='0'/></a></p>
							<p>Your card may be eligible or enrolled in Verified by Visa&#8482; or MasterCard&reg; SecureCode&#8482; payer authentication programs. After clicking the 'Continue' button, your Card Issuer may prompt you for your payer authentication password to complete your purchase.</p>
								<p>&nbsp;</p>
							</td>
						</tr>
					<% end if %>
					<%
                    'SB S
                    If pcIsSubscription AND scSBaymentPageText <>"" Then %>
                    	<tr>
                        	<td colspan="2" class="pcSpacer"></td>
                        </tr>
                        <tr>
                            <td><p><%=scSBLang9%></p></td>
                            <td><p><%=scSBaymentPageText%></p></td>
                        </tr>
					<% End If %>
                    <% If pcIsSubscription AND pcv_intIsTrial AND scSBPaymentPageTrialText <> "" Then %>
                        <tr>
                        	<td colspan="2" class="pcSpacer"></td>
                        </tr>
                        <tr>
                        	<td><p><%=scSBLang10%></p></td>
                            <td><p><%=scSBPaymentPageTrialText%></p></td>
                        </tr>
                    <% 
                    End if 
                    'SB E
                    %>					
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td colspan="2" align="center">
							<!--#include file="inc_gatewayButtons.asp"-->
						</td>
					</tr>
				</table>
			</form>
		<% end if %>
		</td>
	</tr>
</table>
</div>
<% '// Functions
 function IsCreditCard(ByRef anCardNumber, ByRef asCardType)
	Dim lsNumber		' Credit card number stripped of all spaces, dashes, etc.
	Dim lsChar			' an individual character
	Dim lnTotal			' Sum of all calculations
	Dim lnDigit			' A digit found within a credit card number
	Dim lnPosition		' identifies a character position In a String
	Dim lnSum			' Sum of calculations For a specific Set
		
	' Default result is False
	IsCreditCard = False
    			
	' ====
	' Strip all characters that are Not numbers.
	' ====
		
	' Loop through Each character inthe card number submited
	For lnPosition = 1 To Len(anCardNumber)
		' Grab the current character
		lsChar = Mid(anCardNumber, lnPosition, 1)
		' if the character is a number, append it To our new number
		if validNum(lsChar) Then lsNumber = lsNumber & lsChar
		
	Next ' lnPosition
		
	' ====
	' The credit card number must be between 13 and 16 digits.
	' ====
	' if the length of the number is less Then 13 digits, then Exit the routine
	if Len(lsNumber) < 13 Then Exit function

	' if the length of the number is more Then 16 digits, then Exit the routine
	if Len(lsNumber) > 16 Then Exit function
    			
	' Choose action based on Type of card
	Select Case LCase(asCardType)
		' VISA
		Case "visa", "v", "V"
			' if first digit Not 4, Exit function
			if Not Left(lsNumber, 1) = "4" Then Exit function
		' American Express
		Case "american express", "americanexpress", "american", "ax", "A"
			' if first 2 digits Not 37, Exit function
			if Not Left(lsNumber, 2) = "37" AND Not Left(lsNumber, 2) = "34" Then Exit function
		' Mastercard
		Case "mastercard", "master card", "master", "M"
			' if first digit Not 5, Exit function
			if Not Left(lsNumber, 1) = "5" Then Exit function
		' Discover
		Case "discover", "discovercard", "discover card", "D"
			' if first digit Not 6, Exit function
			if Not Left(lsNumber, 1) = "6" Then Exit function
			
		Case Else
	End Select ' LCase(asCardType)
    			
	' ====
	' if the credit card number is less Then 16 digits add zeros
	' To the beginning to make it 16 digits.
	' ====
	' Continue Loop While the length of the number is less Then 16 digits
	While Not Len(lsNumber) = 16
			
		' Insert 0 To the beginning of the number
		lsNumber = "0" & lsNumber
		
	Wend ' Not Len(lsNumber) = 16
		
	' ====
	' Multiply Each digit of the credit card number by the corresponding digit of
	' the mask, and sum the results together.
	' ====
		
	' Loop through Each digit
	For lnPosition = 1 To 16
    				
		' Parse a digit from a specified position In the number
		lnDigit = Mid(lsNumber, lnPosition, 1)
			
		' Determine if we multiply by:
		'	1 (Even)
		'	2 (Odd)
		' based On the position that we are reading the digit from
		lnMultiplier = 1 + (lnPosition Mod 2)
			
		' Calculate the sum by multiplying the digit and the Multiplier
		lnSum = lnDigit * lnMultiplier
			
		' (Single digits roll over To remain single. We manually have to Do this.)
		' if the Sum is 10 or more, subtract 9
		if lnSum > 9 Then lnSum = lnSum - 9
			
		' Add the sum To the total of all sums
		lnTotal = lnTotal + lnSum
    			
	Next ' lnPosition
		
	' ====
	' Once all the results are summed divide
	' by 10, if there is no remainder Then the credit card number is valid.
	' ====
	IsCreditCard = ((lnTotal Mod 10) = 0)

End function ' IsCreditCard
Function determineCardType(Card_Number)

	Dim cardType

	cardType = "UNKNOWN"   ' VISA, MASTERCARD, JCB, AMEX, UNKNOWN

	If (Len(Card_Number) = "16" AND Left(Card_Number, 1) = "4") Then
		cardType = "VISA"
	ElseIf (Len(Card_Number) = "13" AND Left(Card_Number, 1) = "5") Then
		cardType = "MASTERCARD"
	ElseIf (Len(Card_Number) = "16" AND Left(Card_Number, 1) = "5") Then
		cardType = "MASTERCARD"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 4) = "2131") Then
		cardType = "JCB"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 4) = "1800") Then
		cardType = "JCB"
	ElseIf (Len(Card_Number) = "16" AND Left(Card_Number, 1) = "3") Then
		cardType = "JCB"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 2) = "34") Then
		cardType = "AMEX"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 2) = "37") Then
		cardType = "AMEX"
	End If

	determineCardType = cardType

End Function
%>
<!--#include file="footer.asp"-->