<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
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
'SB-S
If session("SB_SkipPayment")="1" then
	Response.redirect "gwReturn.asp?s=true&gw=EIG"
End if
'SB-E

Dim conntemp

'SB S
Dev_Testmode = 1
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
' Determine BACK button url
If scSSL="1" And scIntSSLPage="1" Then
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
Else
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
%>
<div id="pcMain">
<div id="PleaseWaitDialog" title="" style="display:none">
	<div id="PleaseWaitMsg" class="ui-main"></div>
</div>
<script type="text/javascript">
	$(document).ready(function() {
		//*Please Wait Dialog
		$("#PleaseWaitDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 250,
			minHeight: 50,
			modal: true
		});
	});
</script>
<table class="pcMainTable">
	<tr>
		<td class="pcSpacer"></td>
	</tr>
	<tr>
		<td>
		<%
		Dim pcCustIpAddress
		pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

		' Get Order ID
		if session("GWOrderId")="" then
			session("GWOrderId")=getUserInput(request("idOrder"),0)
		end if

		pcGatewayDataIdOrder=session("GWOrderID")
		%>
		<!--#include file="pcGateWayData.asp"-->
		<% session("idCustomer")=pcIdCustomer
		pcv_IncreaseCustID=(scCustPre + int(pcIdCustomer))
		pcTrueOrdnum=(int(session("GWOrderId"))-scpre) %>
		<%
		'// LOAD SSETTINGS
		call opendb()
		query="SELECT pcPay_EIG_Type, pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key, pcPay_EIG_Curcode, pcPay_EIG_CVV, pcPay_EIG_TestMode, pcPay_EIG_SaveCards, pcPay_EIG_UseVault FROM pcPay_EIG WHERE pcPay_EIG_ID=1"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		x_Username=rs("pcPay_EIG_Username")
		x_Username=enDeCrypt(x_Username, scCrypPass)
		x_Password=rs("pcPay_EIG_Password")
		x_Password=enDeCrypt(x_Password, scCrypPass)
		x_Key=rs("pcPay_EIG_Key")
		x_Key=enDeCrypt(x_Key, scCrypPass)
		x_CVV=rs("pcPay_EIG_CVV")
		x_Type=rs("pcPay_EIG_Type")
		x_TypeArray=Split(x_Type,"||")
		x_TransType=x_TypeArray(0)
		x_Curcode=rs("pcPay_EIG_Curcode")
		x_TestMode=rs("pcPay_EIG_TestMode")
		x_SaveCards=rs("pcPay_EIG_SaveCards")
		x_UseVault=rs("pcPay_EIG_UseVault")
		set rs=nothing
		call closedb()



		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// START:  PROCESS RESULTS WHEN SUBSCRIPTION
		'/////////////////////////////////////////////////////////////////////////////////////////////
		If Request.Form("PaymentGWEIG")="Go" Then %>

			<% session("redirectPage")="gwEIGateway.asp" %>

			<%
			dim tempReturnURL
			If scSSL="" OR scSSL="0" Then
				tempReturnURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
				tempReturnURL=replace(tempReturnURL,"https:/","https://")
				tempReturnURL=replace(tempReturnURL,"http:/","http://")
			Else
				tempReturnURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
				tempReturnURL=replace(tempReturnURL,"https:/","https://")
				tempReturnURL=replace(tempReturnURL,"http:/","http://")
			End If

			'SB S
			'// By pass EIG if the immediate order value is 0
			If pcBillingTotal<0 Then
				pcBillingTotal=0
			End If
			If (pcIsSubscription) AND (pcBillingTotal=0) Then

				session("reqCardNumber")=getUserInput(request.Form("billing-cc-number"),16)
				session("reqExpMonth")=getUserInput(request.Form("billing-cc-exp1"),0)
				session("reqExpYear")=getUserInput(request.Form("billing-cc-exp2"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("billing-cvv"),4)
				pExpiration=getUserInput(request("billing-cc-exp1"),0) & "/01/" & getUserInput(request("billing-cc-exp2"),0)

				'// Validates expiration
				if DateDiff("d", Month(Now)&"/01/"&Year(now),pExpiration)<=-1 then
					response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_6")&"<br><br><a href="""&tempReturnURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
				end if

				'// Validate card
				if not IsCreditCard(session("billing-cc-number"), request.form("x_Card_Type")) then
					'response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_5")&"<br><br><a href="""&tempReturnURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
				end if

				session("GWAuthCode")	= "AUTH-ARB"
				session("GWTransId")	= "0"

				Response.Redirect("gwReturn.asp?s=true&gw=EIG&GWError=1")
				Response.End

			Else

				'// Normal Payment, Let Pass
				session("reqCardNumber")=getUserInput(request.Form("billing-cc-number"),16)
				session("reqExpMonth")=getUserInput(request.Form("billing-cc-exp1"),0)
				session("reqExpYear")=getUserInput(request.Form("billing-cc-exp2"),0)
				session("reqCardType")=getUserInput(request.Form("x_Card_Type"),0)
				session("reqCVV")=getUserInput(request.Form("billing-cvv"),4)

			End if
			'SB E
			%>


			<%
			if pcBillingTotal > 0 then

				If x_CVV="1" Then
					if not isnumeric(session("reqCVV")) or len(session("reqCVV")) < 3 or len(session("reqCVV")) > 4 Then
						response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_7")&dictLanguage.Item(Session("language")&"_paymntb_c_4")&"<br><br><a href="""&tempReturnURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
					End If
				End if

				Dim objXMLHTTP, xml

				'// Send the request to the Authorize.NET processor.
				stext=""
				stext=stext & "username=" & x_Username
				stext=stext & "&password=" & x_Password
				If x_TransType="AUTH_ONLY" Then
					stext=stext & "&type=auth"
				Else
					stext=stext & "&type=sale"
				End If
				stext=stext & "&amount=" & pcBillingTotal
				stext=stext & "&ccnumber=" & session("reqCardNumber")
				stext=stext & "&ccexp=" & session("reqExpMonth")&session("reqExpYear")
				If x_CVV="1" Then
					stext=stext & "&cvv=" & session("reqCVV")
				End If

				'// Send the transaction info as part of the querystring
				set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
				xml.open "POST", "https://secure.networkmerchants.com/api/transact.php", false
				xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				xml.send stext
				strStatus = xml.Status

				'// Store the response
				strRetVal = xml.responseText
				Set xml = Nothing

				'// Check for success
				Set resArray = deformatNVP(strRetVal)
				ack = resArray("response")
				ackDesc = resArray("responsetext")

				If ack="1" Then

					call opendb()

					pcv_SecurityPass = scCrypPass
					pcv_SecurityKeyID = pcs_GetKeyID

					dim pCardNumber, pCardNumber2
					pCardNumber=session("reqCardNumber")
					pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)

					'// Save Batch Processing Record
					If x_TransType="AUTH_ONLY" Then

						session("GWAuthCode") = resArray("authcode")
						session("GWTransId") = resArray("transactionid")

						query="INSERT INTO pcPay_EIG_Authorize (idOrder, amount, vaultToken, paymentmethod, transtype, authcode, ccnum, ccexp, cctype, idCustomer, fname, lname, address, zip, captured, trans_id, pcSecurityKeyID) VALUES ("& pcTrueOrdnum &", "& pcBillingTotal &", '', 'CC', '"& x_TransType &"', '"& session("GWAuthCode") &"', '"& pCardNumber2 &"', '"& session("reqExpMonth")&session("reqExpYear") &"', '"& session("reqCardType") &"', "& session("idCustomer") &", '"&replace(pcBillingFirstName,"'","''")&"', '"&replace(pcBillingLastName,"'","''")&"', '"&replace(pcBillingAddress,"'","''")&"', '"& pcBillingPostalCode &"', 0, '"& session("GWTransId") &"', "& pcv_SecurityKeyID &");"

						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						set rs=nothing

						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if

					End If

					call closedb()

					set rs=nothing


					'// Clear all card sessions before redirect
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
					'session("x_method")=""
					'session("x_type")=""
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

					If pcIsSubscription Then
						Response.redirect "gwReturn.asp?s=true&gw=EIG"
					Else
						Response.redirect "gwReturn.asp?s=false&gw=EIG"
					End If


				Else '// If ack="1" Then

					response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& ackDesc &"<br><br><a href="""&tempReturnURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")

				End If '// If ack="1" Then

			End if '// if pcBillingTotal > 0 then


		End If
		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// END:  PROCESS RESULTS WHEN SUBSCRIPTION
		'/////////////////////////////////////////////////////////////////////////////////////////////




		'// CHECK FOR TOKEN
		Dim TokenID
		TokenID=getUserInput(request("token-id"),0)
		IF len(TokenID)>0 THEN

		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// START:  PROCESS RESULTS WHEN TOKEN EXISTS
		'/////////////////////////////////////////////////////////////////////////////////////////////

			'// COMPLETE ACTION
			strTest = ""
			strTest = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
			strTest = strTest & "<complete-action>"
			strTest = strTest & "<api-key>" & x_Key & "</api-key>"
			strTest = strTest & "<token-id>" & TokenID & "</token-id>"
			strTest = strTest & "</complete-action>"

			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			xml.open "POST", "https://secure.nmi.com/api/v2/three-step", false
			xml.setRequestHeader "Content-Type", "text/xml"
			xml.send strTest
			strStatus = xml.Status
			strRetVal = xml.responseText
			Set xml = Nothing

			strResult = pcf_GetNode(strRetVal, "result", "*")
			strResultText = pcf_GetNode(strRetVal, "result-text", "*")
			strTransactionID = pcf_GetNode(strRetVal, "transaction-id", "*")
			strResultCode = pcf_GetNode(strRetVal, "result-code", "*")
			strAuthorizationCode = pcf_GetNode(strRetVal, "authorization-code", "*")
			pcv_strCustomerVaultID = pcf_GetNode(strRetVal, "customer-vault-id", "*")

			'response.Write(strResult & ".<br />")
			'response.Write(strResultText & ".<br />")
			'response.Write(strTransactionID & ".<br />")
			'response.Write(strResultCode & ".<br />")
			'response.Write(authorization-code & ".<br />")
			'response.Write(pcv_strCustomerVaultID & ".<br />")
			'response.End

			'// PROCESS RESULTS
			'// 1 = Transaction Approved
			'// 2 = Transaction Declined
			'// 3 = Error in transaction data or system error
			If strResult="1" Then


				If (x_TransType="AUTH_ONLY") OR (x_SaveCards="1" AND Session("SF_IsSaved")="true") Then

					call opendb()

					Dim pcv_CardNum, pcv_CardType, pcv_CardExp


					If (len(Session("CustomerVaultID"))=0) Then '// A. New Card was used.

						If (x_UseVault=1) OR (x_SaveCards="1" AND Session("SF_IsSaved")="true") Then '// Vault Storage Enabled - OR - Customer Opt-In (Grab from Secure Vault)

							strTest = ""
							strTest = strTest & "username=" & x_Username
							strTest = strTest & "&password=" & x_Password
							strTest = strTest & "&transaction_id=" & strTransactionID

							set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
							xml.open "POST", "https://secure.networkmerchants.com/api/query.php", false
							xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
							xml.send strTest
							strStatus = xml.Status
							strRetVal = xml.responseText
							Set xml = Nothing

							pcv_CardNum = pcf_GetNode(strRetVal, "cc_number", "//nm_response/transaction")
							pcv_CardType = Session("CardType")
							pcv_CardExp = pcf_GetNode(strRetVal, "cc_exp", "//nm_response/transaction")
							If len(pcv_strCustomerVaultID)=0 Then
								pcv_strCustomerVaultID = pcf_GetNode(strRetVal, "customerid", "//nm_response/transaction")
							End If
							If len(pcv_strCustomerVaultID)>0 Then
								Session("CustomerVaultID")=pcv_strCustomerVaultID
							End If
							pcv_strCustomerVaultID2=enDeCrypt(Session("CustomerVaultID"), scCrypPass)

							'// Save Vault Record
							query="SELECT idOrder FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_CardNum='"& pcv_CardNum &"' AND pcPay_EIG_Vault_CardExp='"& pcv_CardExp &"'"
							set rs=Server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)
							if rs.eof then

								If Session("SF_IsSaved")="true" Then
									pcv_tmpIsSaved = 1
								Else
									pcv_tmpIsSaved = 0
								End If
								query="INSERT INTO pcPay_EIG_Vault (idOrder, idCustomer, IsSaved, pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardType, pcPay_EIG_Vault_CardExp, pcPay_EIG_Vault_Token) VALUES ("&pcTrueOrdnum&", "&session("idCustomer")&", "&pcv_tmpIsSaved&", '"&pcv_CardNum&"', '"& pcv_CardType &"', '"& pcv_CardExp &"', '"& pcv_strCustomerVaultID2 &"');"
								set rs2=server.CreateObject("ADODB.RecordSet")
								set rs2=connTemp.execute(query)
								set rs2=nothing

							end if
							set rs=nothing

						Else '// Vault Storage Disabled (Grab from Session)

							pcv_CardNum = Session("CardNum")
							pcv_CardType = Session("CardType")
							pcv_CardExp = Session("CardExp")
							pcv_strCustomerVaultID2="" '// No vault record

						End If

					Else '// B. Saved Card was used.

						query="SELECT pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardType, pcPay_EIG_Vault_CardExp FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_ID="& Session("VaultID") &""
						set rs=Server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if NOT rs.eof then
							pcv_CardNum = rs("pcPay_EIG_Vault_CardNum")
							pcv_CardType = rs("pcPay_EIG_Vault_CardType")
							pcv_CardExp = rs("pcPay_EIG_Vault_CardExp")
						end if
						set rs=nothing
						pcv_strCustomerVaultID2=enDeCrypt(Session("CustomerVaultID"), scCrypPass)

					End If


					'// Save Batch Processing Record
					If x_TransType="AUTH_ONLY" Then

						pcv_CardNum=enDeCrypt(pcv_CardNum, scCrypPass)

						query="INSERT INTO pcPay_EIG_Authorize (idOrder, amount, vaultToken, paymentmethod, transtype, authcode, ccnum, ccexp, cctype, idCustomer, fname, lname, address, zip, captured, trans_id, pcSecurityKeyID) VALUES ("& pcTrueOrdnum &", "& pcBillingTotal &", '"& pcv_strCustomerVaultID2 &"', 'CC', '"& x_TransType &"', '"& strAuthorizationCode &"', '"& pcv_CardNum &"', '"& pcv_CardExp &"', '"& pcv_CardType &"', "& session("idCustomer") &", '"&replace(pcBillingFirstName,"'","''")&"', '"&replace(pcBillingLastName,"'","''")&"', '"&replace(pcBillingAddress,"'","''")&"', '"& pcBillingPostalCode &"', 0, '"& strTransactionID &"', "& pcs_GetKeyID &");"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						set rs=nothing

					End If

					call closedb()

				End If

				Session("CardType") = ""
				Session("CardNum") = ""
				Session("CardExp") = ""
				Session("CustomerVaultID") = ""
				Session("VaultID") = ""
				Session("SF_IsSaved") = ""
				session("GWAuthCode") = strAuthorizationCode
				session("GWTransId") = strTransactionID
				session("GWTransType") = x_TransType

				Response.redirect "gwReturn.asp?s=true&gw=EIG"

			Else '// If strResult="1" Then

				response.Redirect("gwEIGateway.asp?Error=" & strResultText)

			End If '// If strResult="1" Then
			response.End()

		'/////////////////////////////////////////////////////////////////////////////////////////////
		'// END:  PROCESS RESULTS WHEN TOKEN EXISTS
		'/////////////////////////////////////////////////////////////////////////////////////////////

		END IF

		If len(strError)=0 Then
			strError=getUserInput(Request("Error"),0)
		End If
		%>

			<table class="pcShowContent">

					<% if Msg<>"" then %>
						<tr valign="top">
							<td colspan="2">
								<div class="pcErrorMessage"><%=Msg%></div>
							</td>
						</tr>
					<% end if %>

					<% if strError<>"" then %>
						<tr valign="top">
							<td colspan="2">
								<div class="pcErrorMessage"><%=strError%></div>
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
					<%
					call openDB()
					Dim pcSavedCardsCount
					pcSavedCardsCount=0
					query="SELECT pcPay_EIG_Vault_ID, pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardExp FROM pcPay_EIG_Vault WHERE idCustomer="& Session("idCustomer") &" AND IsSaved=1"
					set rs=Server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					if NOT rs.eof then
						pcArray_SavedCards = rs.GetRows()
						pcSavedCardsCount = Ubound(pcArray_SavedCards)
					end if
					set rs=nothing
					call closeDB()
					%>
					<% If pcSavedCardsCount>0 AND (NOT pcIsSubscription) Then %>
					<tr>
						<td>

							<%
							'// Saved Credit Card
							%>
							<form method="POST" name="form-saved-card" id="form-saved-card" class="pcForms">
								<table class="pcShowContent">
									<tr>
										<td width="30%"><p><strong><%=dictLanguage.Item(Session("language")&"_EIG_7")%></strong></p></td>
										<td>
											<select name="VaultID" id="VaultID">
											<%
											For SavedCardsCounter=0 to ubound(pcArray_SavedCards,2)
												%>
												<option value="<%=pcArray_SavedCards(0,SavedCardsCounter)%>"><%=pcArray_SavedCards(1,SavedCardsCounter)%> (<%=dictLanguage.Item(Session("language")&"_GateWay_8")%> <%=pcArray_SavedCards(2,SavedCardsCounter)%>)</option>
												<%
											Next
											%>
											</select>&nbsp;&nbsp;<a href="CustviewPayment.asp" target="_blank"><%=dictLanguage.Item(Session("language")&"_EIG_9")%></a>
										</td>
									</tr>
									<tr>
										<td colspan="2" class="pcSpacer"></td>
									</tr>
									<tr>
										<td>
										<td align="left">
										<a href="<%=tempURL%>"><img src="<%=rslayout("back")%>"></a>
											 &nbsp;<img src="<%=rslayout("pcLO_placeOrder")%>" name="submit-saved-card" id="submit-saved-card" style="border:none; cursor:pointer">
											<script type="text/javascript">
												$(document).ready(function() {
													$('#submit-saved-card').click(function() {
														$("#PleaseWaitMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_EIG_19"))%>');
														$("#PleaseWaitDialog").dialog('open');
														$(".ui-dialog-titlebar").css({'display' : 'none'});
														$("#PleaseWaitDialog").css({'min-height' : '50px'});
														var tmpdata="";
														tmpdata=$('#VaultID').val();
														$.ajax(
															   {
																type: "GET",
																url: "gwEIGatewayURL.asp",
																data: "VaultID=" + tmpdata,
																timeout: 45000,
																success: function(data, textStatus){
																	if (data.indexOf("OK||")>=0) {
																		var tmpArr=data.split("||")
																		$("#form-saved-card").attr("action", tmpArr[1]);
																		$("#form-saved-card").submit();
																		return true;
																	} else {
																		window.location.href = 'gwEIGateway.asp?Error=' + data;
																		return false;
																	}
																}
														});
													});
												});
											</script>

										</td>
									</tr>
									<tr>
										<td colspan="2" class="pcSpacer" style="height: 40px; vertical-align: middle;"><hr></td>
									</tr>
									<tr>
										<td colspan="2"><strong><%=dictLanguage.Item(Session("language")&"_EIG_18")%></strong></td>
									</tr>
								</table>
							</form>

						</td>
					</tr>
					<% End If %>
					<tr>
						<td>

							<%
							'// New Credit Card
							%>
							<% If NOT pcIsSubscription Then %>
								<form method="POST" name="form-new-card" id="form-new-card" class="pcForms">
							<% Else %>
								<% 'SB S %>
								<form method="POST" name="form-new-card" id="form-new-card" class="pcForms" action="gwEIGateway.asp">
								<input type="hidden" name="PaymentGWEIG" value="Go">
								<% 'SB E %>
							<% End If %>
								<table class="pcShowContent">
									<tr>
										<td><p>Card Type:</p></td>
										<td>
											<select name="x_Card_Type" id="x_Card_Type">
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
										<td width="30%"><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
										<td>
											<input type="text" name="billing-cc-number" id="billing-cc-number" value="" size="30">
										</td>
									</tr>
									<tr>
										<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
										<td><input type="hidden" name="billing-cc-exp" id="billing-cc-exp" value="">
											<!--(<%=dictLanguage.Item(Session("language")&"_EIG_3")%>/<%=dictLanguage.Item(Session("language")&"_EIG_4")%>) -->
											<% dtCurYear=Year(date()) %>
											<%=dictLanguage.Item(Session("language")&"_GateWay_9")%>
											<select name="billing-cc-exp1" id="billing-cc-exp1">
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
											&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%>
											<select name="billing-cc-exp2" id="billing-cc-exp2">
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
												<input name="billing-cvv" type="text" id="billing-cvv" value="" size="4" maxlength="4">
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

									<% If x_SaveCards="1" AND (NOT pcIsSubscription) Then %>
									<tr>
										<td><p><%=dictLanguage.Item(Session("language")&"_EIG_1")%> <img src="images/pcv4_st_icon_info.png" width="20" height="20" alt="<%=dictLanguage.Item(Session("language")&"_EIG_1")%>" title="<%=dictLanguage.Item(Session("language")&"_EIG_2")%>">
										</p></td>
										<td>
                                            <input name="x_SaveCards" id="x_SaveCards" type="checkbox" value="x_SaveCards" <%if Session("SF_IsSaved")="true" then response.Write("checked")%> class="clearBorder"/>

                                        </td>
									</tr>


									<% End If %>
									<tr>
										<td colspan="2" class="pcSpacer"><hr></td>
									</tr>
									<tr>
										<td>
										<td align="left">
											<% 'SB S
											 If (pcIsSubscription) Then %>

												<!--#include file="inc_gatewayButtons.asp"-->

											<% 'SB E
											Else %>

												<a href="<%=tempURL%>"><img src="<%=rslayout("back")%>"></a>  &nbsp;
												<img src="<%=rslayout("pcLO_placeOrder")%>" name="submit-new-card" id="submit-new-card" style="border:none; cursor:pointer">

												<script type="text/javascript">
													$(document).ready(function() {
														$('#submit-new-card').click(function() {
															$("#PleaseWaitMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_EIG_19"))%>');
															$("#PleaseWaitDialog").dialog('open');
															$(".ui-dialog-titlebar").css({'display' : 'none'});
															$("#PleaseWaitDialog").css({'min-height' : '50px'});
															var cardType = $('#x_Card_Type').val();
															<% If (x_UseVault<>1) AND (x_TransType="AUTH_ONLY") Then %>
																var CardNum = $('#billing-cc-number').val();
															<% Else %>
																var CardNum = '';
															<% End If %>
															var exp1 = $('#billing-cc-exp1').val();
															var exp2 = $('#billing-cc-exp2').val();
															var exp_date = exp1 + exp2;
															$('#billing-cc-exp').val(exp_date);
															var tmpdata="";
															if ($('#x_SaveCards').attr('checked')==true) {
																tmpdata=true
															} else {
																tmpdata=false
															}
															$.ajax(
																   {
																	type: "GET",
																	url: "gwEIGatewayURL.asp",
																	data: "IsSaved=" + tmpdata + '&CardType=' + cardType + '&CardNum=' + CardNum + '&ExpDate=' + exp_date,
																	timeout: 45000,
																	success: function(data, textStatus){
																		if (data.indexOf("OK||")>=0) {
																			var tmpArr=data.split("||")
																			$("#form-new-card").attr("action", tmpArr[1]);
																			$("#form-new-card").submit();
																			return true;
																		} else {
																			window.location.href = 'gwEIGateway.asp?Error=' + data;
																			return false;
																		}
																	}
															});
														});
													});
												</script>

											<% End If %>
											<%
                                            '// Secure Form Fields
                                            %>
                                            <input type="hidden" name="billing-first-name" value="<%=pcf_FixXML(pcBillingFirstName)%>">
                                            <input type="hidden" name="billing-last-name" value="<%=pcf_FixXML(pcBillingLastName)%>">
                                            <input type="hidden" name="billing-address1" value="<%=pcf_FixXML(pcBillingAddress)%>">
                                            <input type="hidden" name="billing-address2" value="<%=pcf_FixXML(pcBillingAddress2)%>">
                                            <input type="hidden" name="billing-city" value="<%=pcf_FixXML(pcBillingCity)%>">
                                            <input type="hidden" name="billing-state" value="<%=pcf_FixXML(pcBillingState)%>">
                                            <input type="hidden" name="billing-postal" value="<%=pcf_FixXML(pcBillingPostalCode)%>">
                                            <input type="hidden" name="billing-country" value="<%=pcf_FixXML(pcBillingCountryCode)%>">
                                            <input type="hidden" name="billing-phone" value="<%=pcf_FixXML(pcBillingPhone)%>">
                                            <input type="hidden" name="billing-fax" value="<%=pcf_FixXML(pcShippingFax)%>">
                                            <input type="hidden" name="billing-email" value="<%=pcf_FixXML(pcCustomerEmail)%>">
                                            <input type="hidden" name="billing-company" value="<%=pcf_FixXML(pcBillingCompany)%>">
                                            
                                            <% If len(pcShippingAddress)>0 Then %>
                                            <input type="hidden" name="shipping-address1" value="<%=pcf_FixXML(pcShippingAddress)%>">
                                            <input type="hidden" name="shipping-address2" value="<%=pcf_FixXML(pcShippingAddress2)%>">
                                            <input type="hidden" name="shipping-city" value="<%=pcf_FixXML(pcShippingCity)%>">
                                            <input type="hidden" name="shipping-state" value="<%=pcf_FixXML(pcShippingState)%>">
                                            <input type="hidden" name="shipping-postal" value="<%=pcf_FixXML(pcShippingPostalCode)%>">
                                            <input type="hidden" name="shipping-country" value="<%=pcf_FixXML(pcShippingCountryCode)%>">
                                            <input type="hidden" name="shipping-phone" value="<%=pcf_FixXML(pcShippingPhone)%>">
                                            <input type="hidden" name="shipping-fax" value="<%=pcf_FixXML(pcShippingFax)%>">
                                            <input type="hidden" name="shipping-email" value="<%=pcf_FixXML(pcShippingEmail)%>">
                                            <input type="hidden" name="shipping-company" value="<%=pcf_FixXML(pcShippingCompany)%>">
                                            <% End If %>
										</td>
									</tr>
								</table>
							</form>

						</td>
					</tr>
				</table>

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
%>
<!--#include file="footer.asp"-->
<%
Function pcf_GetNode(responseXML, nodeName, nodeParent)
	Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)
	myXmlDoc.loadXml(responseXML)
	Set Nodes = myXmlDoc.selectnodes(nodeParent)
	For Each Node In Nodes
		pcf_GetNode = pcf_CheckNode(Node,nodeName,"")
	Next
	Set Node = Nothing
	Set Nodes = Nothing
	Set myXmlDoc = Nothing
End Function

Function pcf_CheckNode(Node,tagName,default)
	Dim tmpNode
	Set tmpNode=Node.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		pcf_CheckNode=default
	Else
		pcf_CheckNode=Node.selectSingleNode(tagName).text
	End if
End Function

Function pcf_FixXML(str)
	str=replace(str, "&","and")
	pcf_FixXML=str
End Function

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
%>