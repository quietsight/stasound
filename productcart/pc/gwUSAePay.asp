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
<% 'Gateway specific files %>
<%
'If gateway has sandbox or developer testing mode use this flag
Dev_Testmode = 0

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

		session("redirectPage")="gwUSAePay.asp" %>

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

		query="SELECT pcPay_Uep_SourceKey,pcPay_Uep_TransType,pcPay_Uep_TestMode FROM pcPay_USAePay WHERE pcPay_Uep_Id=1;"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		pcPay_Uep_SourceKey=rs("pcPay_Uep_SourceKey")
		pcPay_Uep_SourceKey=enDeCrypt(pcPay_Uep_SourceKey, scCrypPass)
		pcPay_Uep_TransType=rs("pcPay_Uep_TransType")
		pcPay_Uep_TestMode=rs("pcPay_Uep_TestMode")

		set rs=nothing

		If Request.Form("PaymentGWCentinel")="Go" OR request.QueryString("Centinel")<>"" Then %>
			<%
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
			<!-- #include file="Centinel_Config.asp"-->
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
						Response.redirect ("gwUSAePay.asp?centinel=results")
					Else
						Response.redirect ("msg.asp")
					End If


				End If


				Set centinelResponse = nothing
				Set centinelRequest = nothing

			End If

			if pcBillingTotal > 0 then
				If scSSL="" OR scSSL="0" Then
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
					tempURL=replace(tempURL,"https:/","https://")
					tempURL=replace(tempURL,"http:/","http://")
				Else
					tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
					tempURL=replace(tempURL,"https:/","https://")
					tempURL=replace(tempURL,"http:/","http://")
				End If
				%>
				<%
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

				If pcShippingAddress&""="" Then
					pcShippingFirstName = pcBillingFirstName
					pcShippingLastName = pcBillingLastName
					pcShippingCompany = pcBillingCompany
					pcShippingAddress = pcBillingAddress
					pcShippingAddress2 = pcBillingAddress2
					pcShippingCity = pcBillingCity
					pcShippingState = pcBillingState
					pcShippingPostalCode = pcBillingPostalCode
					pcShippingCountryCode = pcBillingCountryCode
					pcShippingPhone = pcBillingPhone
				End If

				Dim objXMLHTTP, xml

				'Send the request to the Authorize.NET processor.
				stext="UMcommand=cc:sale"
				stext=stext & "&UMkey=" & pcPay_Uep_SourceKey
				stext=stext & "&UMamount=" & pcBillingTotal
				stext=stext & "&UMcard=" & session("reqCardNumber")
				stext=stext & "&UMexpir=" & session("reqExpMonth")&session("reqExpYear")
				stext=stext & "&UMinvoice=" & session("GWOrderID")
				stext=stext & "&UMorderid=" & session("GWOrderID")
				stext=stext & "&UMcvv2=" & session("reqCVV")
				stext=stext & "&UMname=" & pcBillingFirstName &" "& pcBillingLastName
				stext=stext & "&UMstreet=" & replace(pcBillingAddress,",","||")
				stext=stext & "&UMzip=" & pcBillingPostalCode
				'stext=stext & "&UMhash=" & "m/1234/"&strHash&"/n"
				stext=stext & "&UMbillfname=" & pcBillingFirstName
				stext=stext & "&UMbilllname=" & pcBillingLastName
				stext=stext & "&UMbillcompany=" & pcBillingCompany
				stext=stext & "&UMbillstreet=" & pcBillingAddress
				stext=stext & "&UMbillstreet2=" & pcBillingAddress2
				stext=stext & "&UMbillcity=" & pcBillingCity
				stext=stext & "&UMbillstate=" & pcBillingState
				stext=stext & "&UMbillzip=" & pcBillingPostalCode
				stext=stext & "&UMbillcountry=" & pcBillingCountryCode
				stext=stext & "&UMbillphone=" & pcBillingPhone
				stext=stext & "&UMshipfname=" & pcShippingFirstName
				stext=stext & "&UMshiplname=" & pcShippingLastName
				stext=stext & "&UMshipcompany=" & pcShippingCompany
				stext=stext & "&UMshipstreet=" & pcShippingAddress
				stext=stext & "&UMshipsreet2=" & pcShippingAddress2
				stext=stext & "&UMshipcity=" & pcShippingCity
				stext=stext & "&UMshipstate=" & pcShippingState
				stext=stext & "&UMshipzip=" & pcShippingPostalCode
				stext=stext & "&UMshipcountry=" & pcShippingCountryCode
				stext=stext & "&UMshipphone=" & pcShippingPhone
				stext=stext & "&UMemail=" & pcCustomerEmail
				stext=stext & "&UMip=" & Cstr(pcCustIpAddress)

				'GET CENTINEL DETAILS FOR GATEWAY - ALTER THIS
				If pcPay_Cent_Active=1 AND pcPay_CentByPass=1 AND pcPay_CardType="YES" Then
					'stext=stext & "&x_authentication_indicator=" & Session("Centinel_ECI")
					'stext=stext & "&x_cardholder_authentication_value=" & Session("Centinel_CAVV")
				End If

				'Send the transaction info as part of the querystring
				set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)

				if Dev_Testmode = 1 Then
					xml.open "POST", "https://sandbox.usaepay.com/gate?"& stext & "", false
				else
					xml.open "POST", "https://www.usaepay.com/gate?"& stext & "", false
				end if

				xml.send ""
				strStatus = xml.Status

				'store the response
				strRetVal = xml.responseText
				Set xml = Nothing

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
					' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
					ResponseArray = Split(strRetVal, "&")
					Set objDictResponse = server.createobject("Scripting.Dictionary")
					For each ResponseItem in ResponseArray
						NameValue = Split(ResponseItem, "=")
						objDictResponse.Add NameValue(0), NameValue(1)
					Next

					' Parse the response into local vars
					UMstatus = objDictResponse.Item("UMstatus") 'Status of the transaction. The possible values are: Approved, Declined, Verification and Error.
					UMauthCode = objDictResponse.Item("UMauthCode") 'Authorization number.
					UMauthAmount = objDictResponse.Item("UMauthAmount") 'Amount authorized. May be less than amount requested if UMallowPartialAuth=true
					UMrefNum = objDictResponse.Item("UMrefNum") 'Transaction reference number
					UMbatch = objDictResponse.Item("UMbatch") 'Batch reference number. This will only be returned for sale and auth commands. Warning: The batch number returned is for the batch that was open when the transaction was initiated. It is possible that the batch was closed while the transaction was processing. In this case the transaction will get queued for the next batch to open.
					UMavsResult = objDictResponse.Item("UMavsResult") 'AVS result in readable format
					UMavsResultCode = objDictResponse.Item("UMavsResultCode") 'AVS result code.
					UMcvv2Result = objDictResponse.Item("UMcvv2Result") 'CVV2 result in readable format.
					UMcvv2ResultCode = objDictResponse.Item("UMcvv2ResultCode") 'CVV2 result code.
					UMvpasResultCode = objDictResponse.Item("UMvpasResultCode") 'Verified by Visa (VPAS) or Mastercard SecureCode (UCAF) result code.
					UMresult = objDictResponse.Item("UMresult") 'Transaction result - A, D, E, or V (for Verification)
					UMerror = objDictResponse.Item("UMerror") 'Error description if UMstatus is Declined or Error.
					UMerrorcode = objDictResponse.Item("UMerrorcode") 'A numerical error code.
					UMacsurl = objDictResponse.Item("UMacsurl") 'Verification URL to send cardholder to. Sent when UMstatus is verification (cardholder authentication is required).
					UMpayload = objDictResponse.Item("UMpayload") 'Authentication data to pass to verification url. Sent when UMstatus is verification (cardholder authentication is required).
					UMisDuplicate = objDictResponse.Item("UMisDuplicate") 'Indicates whether this transaction is a folded duplicate or not. 'Y' means that this transaction was flagged as duplicate and has already been processed. The details returned are from the original transaction. Send UMignoreDuplicate to override the duplicate folding.
					UMconvertedAmount = objDictResponse.Item("UMconvertedAmount") 'Amount converted to merchant's currency, when using a multi-currency processor.
					UMconvertedAmountCurrency = objDictResponse.Item("UMconvertedAmountCurrency") 'Merchant's currency, when using a multi-currency processor.
					UMconversionRate = objDictResponse.Item("UMconversionRate") 'Conversion rate used to convert currency, when using a multi-currency processor.
					UMcustnum = objDictResponse.Item("UMcustnum") 'Customer reference number assigned by gateway. Returned only if UMaddcustomer=yes.
					UMresponseHash = objDictResponse.Item("UMresponseHash") 'Response verification hash. Only present if response hash was requested in the UMhash. (See Source Pin Code section for further details)
					UMprocRefNum = objDictResponse.Item("UMprocRefNum") 'Transaction Reference number provided by backend processor (platform), blank if not available)
					UMcardLevelResult = objDictResponse.Item("UMcardLevelResult") 'Card level results (for Visa cards only), blank if no results provided
					UMbatchRefNum = objDictResponse.Item("UMbatchRefNum")
					UMcustReceiptResult = objDictResponse.Item("UMcustReceiptResult")
					UMfiller = objDictResponse.Item("UMfiller")

					'response.write "UMstatus :"& UMstatus &"<BR>"
					'response.write "UMauthCode :"& UMauthCode &"<BR>"
					'response.write "UMauthAmount :"& UMauthAmount &"<BR>"
					'response.write "UMrefNum :"& UMrefNum &"<BR>"
					'response.write "UMbatch :"& UMbatch &"<BR>"
					'response.write "UMavsResult :"& UMavsResult &"<BR>"
					'response.write "UMavsResultCode :"& UMavsResultCode &"<BR>"
					'response.write "UMcvv2Result :"& UMcvv2Result &"<BR>"
					'response.write "UMcvv2ResultCode :"& UMcvv2ResultCode &"<BR>"
					'response.write "UMvpasResultCode :"& UMvpasResultCode &"<BR>"
					'response.write "UMresult :"& UMresult &"<BR>"
					'response.write "UMerror :"& UMerror &"<BR>"
					'response.write "UMerrorcode :"& UMerrorcode &"<BR>"
					'response.write "UMacsurl :"& UMacsurl &"<BR>"
					'response.write "UMpayload :"& UMpayload &"<BR>"
					'response.write "UMisDuplicate :"& UMisDuplicate &"<BR>"
					'response.write "UMconvertedAmount :"& UMconvertedAmount &"<BR>"
					'response.write "UMconvertedAmountCurrency :"& UMconvertedAmountCurrency &"<BR>"
					'response.write "UMconversionRate :"& UMconversionRate &"<BR>"
					'response.write "UMcustnum :"& UMcustnum &"<BR>"
					'response.write "UMresponseHash :"& UMresponseHash &"<BR>"
					'response.write "UMprocRefNum :"& UMprocRefNum &"<BR>"
					'response.write "UMcardLevelResult :"& UMcardLevelResult &"<BR>"
					'response.write "UMbatchRefNum :"& UMbatchRefNum &"<BR>"
					'response.write "UMcustReceiptResult :"& UMcustReceiptResult &"<BR>"
					'response.write "UMfiller :"& UMfiller &"<BR>"
					'response.end
					'save and update order
					'Approved, Declined, Verification and Error
					If UMstatus = "Approved" Then
						call opendb()
						session("GWTransId") = UMrefNum
						session("GWAuthCode") = UMauthCode
						pcv_SecurityPass = pcs_GetSecureKey
						pcv_SecurityKeyID = pcs_GetKeyID

						dim pCardNumber, pCardNumber2
						pCardNumber=session("reqCardNumber")
						pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)

				'save info in pcPay_Uep_Orders if AUTH_CAPTURE
						pcTrueOrdnum=(cLng(session("GWOrderID"))-cLng(scpre))
						if pcPay_Uep_TransType<>"1" then

							query="INSERT INTO pcPay_USAePay_Orders (idOrder, amount,paymentmethod,transtype,RefNum,ccCard,ccExp,idCustomer,fname,lname,address,zip,captured,pcSecurityKeyID) VALUES ("&pcTrueOrdnum&", "&pcBillingTotal&",'CC','"&pcPay_Uep_TransType&"','"&session("GWTransId")&"','"&pCardNumber2&"','"&session("reqExpMonth")&session("reqExpYear")&"',"&session("idcustomer")&",'"&replace(pcBillingFirstName,"'","''")&"','"&replace(pcbillingLastName,"'","''")&"','"&replace(pcBillingAddress,"'","''")&"','"&pcBillingPostalCode&"',0,"&pcv_SecurityKeyID&");"
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

						session("GWTransType")=pcPay_Uep_TransType

						Response.redirect "gwReturn.asp?s=true&gw=UEP"
					Else
						if UMerror&""="" Then
							UMerror="There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."
						end if
						response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "&URLDecode(UMerror)&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idOrder="&session("GWOrderID")&"""><img src="""&rslayout("back")&"""></a>")
						response.end
					End If
					'////
				End If
			end if '//// if pcBillingTotal > 0 then
			'//// End If 'Redirected back after Centinel approval or form submit
		Else
			call opendb()

			%>

			<form action="gwUSAePay.asp" method="POST" name="form1" class="pcForms">
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
					<tr>
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><p><%= money(pcBillingTotal)%></p></td>
					</tr>
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

Function URLDecode(str)
	str = Replace(str, "+", " ")
	For i = 1 To Len(str)
		sT = Mid(str, i, 1)
		If sT = "%" Then
			If i+2 < Len(str) Then
				sR = sR & _
					Chr(CLng("&H" & Mid(str, i+1, 2)))
				i = i+2
			End If
		Else
			sR = sR & sT
		End If
	Next
	URLDecode = sR
End Function

%>
<!--#include file="footer.asp"-->