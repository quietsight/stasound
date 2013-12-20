<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/pcGenXMLClass.asp"-->
<!--#include file="header.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwTransFirstXP.asp"

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
		
'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if

'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->

<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT pcPay_TransFirstXP.pcPay_TXPGatewayID, pcPay_TransFirstXP.pcPay_TXPRegistrationKey, pcPay_TransFirstXP.pcPay_TXPTransType, pcPay_TransFirstXP.pcPay_TXPTestMode, pcPay_TransFirstXP.pcPay_TXPReqCardCode, pcPay_TransFirstXP.pcPay_TXPCardTypes FROM pcPay_TransFirstXP WHERE (((pcPay_TransFirstXP.pcPay_TransFirstXPID)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect"techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
TXPGatewayID=rs("pcPay_TXPGatewayID")
	TXPGatewayID=enDeCrypt(TXPGatewayID, scCrypPass)
 
TXPRegistrationKey=rs("pcPay_TXPRegistrationKey") 
	TXPRegistrationKey=enDeCrypt(TXPRegistrationKey, scCrypPass)

TXPTransType=rs("pcPay_TXPTransType") 
TXPTestMode=rs("pcPay_TXPTestMode") 
TXPReqCardCode=rs("pcPay_TXPReqCardCode") 
TXPCardTypes=rs("pcPay_TXPCardTypes") 
if TXPReqCardCode<>1 then
	TXPReqCardCode=0
end if

'TXPGatewayID ="7777777750"
'TXPRegistrationKey ="JBDBDNQ8BRRGAKLZ"
'TXPTransType ="1" 'Sale"0" 'Auth Only
'TXPReqCardCode="1"
set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then
	pcBillingTotal = (pcBillingTotal*100)
	if len(pcBillingTotal)<3 then
		pcBillingTotal = "0"&pcBillingTotal
	end if
	txpCardNumber = GetUserInput(Request.Form("CardNumber"),0)
	txpExpMonth = GetUserInput(Request.Form("expMonth"),0)
	txpExpYear = GetUserInput(Request.Form("expYear"),0)
	txpExpDate = txpExpYear & txpExpMonth
	txpSecCode = GetUserInput(Request.Form("cvm"),0)
	
	TXP_PostData=""
	TXP_PostData=TXP_PostData&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
	TXP_PostData=TXP_PostData&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"">"&vbcrlf
	TXP_PostData=TXP_PostData&"<soapenv:Body>"&vbcrlf
	TXP_PostData=TXP_PostData&"<v1:SendTranRequest xmlns:v1=""http://postilion/realtime/merchantframework/xsd/v1/"">"&vbcrlf
		TXP_PostData=TXP_PostData&"<v1:merc>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:id>" & TXPGatewayID &"</v1:id>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:regKey>" & TXPRegistrationKey &"</v1:regKey>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:inType>1</v1:inType>"&vbcrlf
		TXP_PostData=TXP_PostData&"</v1:merc>"&vbcrlf
		TXP_PostData=TXP_PostData&"<v1:tranCode>"&TXPTransType&"</v1:tranCode>"&vbcrlf
		TXP_PostData=TXP_PostData&"<v1:card>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:pan>" & txpCardNumber &"</v1:pan>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:sec>" & txpSecCode &"</v1:sec>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:xprDt>" & txpExpDate &"</v1:xprDt>"&vbcrlf
		TXP_PostData=TXP_PostData&"</v1:card>"&vbcrlf
		TXP_PostData=TXP_PostData&"<v1:contact>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:id>" & session("GWOrderId") &"</v1:id>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:addrLn1>" & pcBillingAddress &"</v1:addrLn1>"&vbcrlf
			'TXP_PostData=TXP_PostData&"<v1:addrLn2>" & pcBillingAddress &"</v1:addrLn2>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:city>" & pcBillingCity &"</v1:city>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:state>" & left(pcShippingStateCode,2) &"</v1:state>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:zipCode>" & pcShippingPostalCode &"</v1:zipCode>"&vbcrlf
			TXP_PostData=TXP_PostData&"<v1:ctry>" & pcShippingCountryCode &"</v1:ctry>"&vbcrlf
		TXP_PostData=TXP_PostData&"</v1:contact>"&vbcrlf
		TXP_PostData=TXP_PostData&"<v1:reqAmt>"&pcBillingTotal&"</v1:reqAmt>"&vbcrlf
	TXP_PostData=TXP_PostData&"</v1:SendTranRequest>"&vbcrlf
	TXP_PostData=TXP_PostData&"</soapenv:Body>"&vbcrlf
	TXP_PostData=TXP_PostData&"</soapenv:Envelope>"&vbcrlf
	
	'response.write TXP_PostData &"<HR>"
	TXPURL ="https://ws.cert.processnow.com/portal/merchantframework/MerchantWebServices-v1"
	dim resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	
	resolveTimeout	= 5000
	connectTimeout	= 5000
	sendTimeout		= 5000
	receiveTimeout	= 10000

	'Send the transaction info as part of the querystring
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout

	xml.open"POST", TXPURL &"", false
	xml.setRequestHeader"Content-Type","text/xml;charset=UTF-8"
	xml.setRequestHeader"Host","ws.cert.processnow.com"
	xml.setRequestHeader"Content-Length", len(TXP_PostData)

	
	xml.send(TXP_PostData)

	strRetVal = xml.responseText

	Dim objGenXMLXmlDoc, objGenXMLStream 
	Dim objGenXMLClass, objOutputXMLDoc, srvGenXMLXmlHttp

	'// SET THE GenXML OBJECT
	set objGenXMLClass = New pcGenXMLClass
	
	call objGenXMLClass.LoadXMLResults(strRetVal)
	objOutputXMLDoc.loadXML strRetVal

	
	 rspCode = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:rspCode")
	 If TXPTransType ="0" Then
		 avsRslt = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:authRsp/ns2:avsRslt")
		 tranId = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:authRsp/ns2:tranId")
		 valCode = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:authRsp/ns2:valCode")
	 End If
	 aci = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:authRsp/ns2:aci")
	 swchKey = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:tranData/ns2:swchKey")
	 tranNr = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:tranData/ns2:tranNr")
	 dtTm = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:tranData/ns2:dtTm")
	 amt = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:tranData/ns2:amt")
	 stan = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:tranData/ns2:stan")
	 auth = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:tranData/ns2:auth")
	 cardType = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:cardType")
	 mapCaid = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:mapCaid")
	 accountType = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:additionalAmount/ns2:accountType")
	 amountType = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:additionalAmount/ns2:amountType")
	 currencyCode = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:additionalAmount/ns2:currencyCode")
	 amountSign = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:additionalAmount/ns2:amountSign")
	 amount = objGenXMLClass.ReadResponseNode("//ns2:SendTranResponse","ns2:additionalAmount/ns2:amount")
 
	 If rspCode ="00" then
		session("GWAuthCode")=auth
		session("GWTransId")=tranNr
		session("GWTransType")=TXPTransType
		response.redirect"gwReturn.asp?s=true&gw=TransFirst"
	Else
		select case rspCode
			case "01"
				responseError ="Refer to card issuer"
			case "02"
				responseError ="Refer to card issuer, special condition"
			case "03"
				responseError ="Invalid merchant"
			case "04"
				responseError ="Pick-up card"
			case "05"
				responseError ="Do not honor"
			case "06"
				responseError ="Error"
			case "07"
				responseError ="Pick-up card, special condition"
			case "08"
				responseError ="Honor with identification"
			case "10"
				responseError ="Approved, partial"
			case "11"
				responseError ="VIP Approval"
			case "12"
				responseError ="Invalid transaction"
			case "13"
				responseError ="Invalid amount"
			case "14"
				responseError ="Invalid card number"
			case "15"
				responseError ="No such issuer"
			case "17"
				responseError ="Customer cancellation"
			case "19"
				responseError ="Re-enter transaction"
			case "21"
				responseError ="No action taken"
			case "25"
				responseError ="Unable to locate record"
			case "28"
				responseError ="File update file locked"
			case "30"
				responseError ="Format error"
			case "32"
				responseError ="Completed partially - Valid for MasterCard Reversal Requests Only - Used in a reversal message to indicate the reversal request is for an amount less than the original transaction."
			case "39"
				responseError ="No credit account"
			case "41"
				responseError ="Lost card, pick-up"
			case "43"
				responseError ="Stolen card, pick-up" 
			case "51"
				responseError ="Not sufficient funds"
			case "52"
				responseError ="No checking account"
			case "53"
				responseError ="No savings account"
			case "54"
				responseError ="Expired card"
			case "55"
				responseError ="Incorrect PIN"
			case "57"
				responseError ="Transaction not permitted to cardholder"
			case "58"
				responseError ="Transaction not permitted on terminal"
			case "59"
				responseError ="Suspected fraud"
			case "61"
				responseError ="Exceeds withdrawal limit"
			case "62"
				responseError ="Restricted card"
			case "63"
				responseError ="Security violation"
			case "65"
				responseError ="Exceeds withdrawal frequency"
			case "68"
				responseError ="Response received too late"
			case "69"
				responseError ="Advice received too late"
			case "70"
				responseError ="Reserved for future use"
			case "75"
				responseError ="PIN tries exceeded" 
			case "76"
				responseError ="Reversal: Unable to locate previous message (no match on Retrieval Reference Number)." 
			case "77"
				responseError ="Previous message located for a repeat or reversal, but repeat or reversal data is inconsistent with original message."
			case "78"
				responseError ="Invalid/non-existent account ? Decline (MasterCard specific)"
			case "79"
				responseError ="Already reversed (by Switch)"
			case "80"
				responseError ="No financial Impact (Reserved for declined debit)" 
			case "81"
				responseError ="PIN cryptographic error found by the Visa security module during PIN decryption."
			case "82"
				responseError ="Incorrect CVV"
			case "83"
				responseError ="Unable to verify PIN"
			case "84"
				responseError ="Invalid Authorization Life Cycle ? Decline (MasterCard) or Duplicate Transaction Detected (Visa)"
			case "85"
				responseError ="No reason to decline a request for Account Number Verification or Address Verification"
			case "86"
				responseError ="Cannot verify PIN"
			case "91"
				responseError ="Issuer or switch inoperative"
			case "92"
				responseError ="Destination Routing error"
			case "93"
				responseError ="Violation of law"
			case "94"
				responseError ="Duplicate Transmission (Integrated Debit and MasterCard)"
			case "96"
				responseError ="System malfunction"
			case "B1"
				responseError ="Surcharge amount not permitted on Visa cards or EBT Food Stamps" 
			case "B2"
				responseError ="Surcharge amount not supported by debit network issuer"
			case "N0"
				responseError ="Force STIP"
			case "N3"
				responseError ="Cash service not available"
			case "N4"
				responseError ="Cash request exceeds Issuer limit"
			case "N5"
				responseError ="Ineligible for re-submission"
			case "N7"
				responseError ="Decline for CVV2 failure"
			case "N8"
				responseError ="Transaction amount exceeds preauthorized approval amount"
			case "P0"
				responseError ="Approved; PVID code is missing, invalid, or has expired"
			case "P1"
				responseError ="Declined; PVID code is missing, invalid, or has expired"
			case "P2"
				responseError ="Invalid biller Information"
			case "R0"
				responseError ="The transaction was declined or returned, because the cardholder requested that payment of a specific recurring or installment payment transaction be stopped."
			case "R1"
				responseError ="The transaction was declined or returned, because the cardholder requested that payment of all recurring or installment payment transactions for a specific merchant account be stopped."
			case "Q1"
				responseError ="Card Authentication failed"
			case "XA"
				responseError ="Forward to Issuer"
			case "XD"
				responseError ="Forward to Issuer"
			case else
				responseError ="Unknown Error"
		end select
	
'		response.write TXP_PostData &"<HR>"
'		response.write strRetVal
'		response.write"<HR>"
'		response.end
		response.redirect"msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:"&responseError &"<br><br><a href="""&tempURL&"?psslurl=gwTransFirstXP.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		Response.end	
	End if
end if 
%>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<img src="images/checkout_bar_step5.gif" alt="">
			</td>
		</tr>
		<tr>
			<td>
				<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="PaymentSubmitted" value="Go">
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
						<td colspan="2"><p><%=pcBillingFirstName&""&pcBillingLastName%></p></td>
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
						<td colspan="2"><p><%=pcBillingCity&","&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% if pcPay_Cys_TestMode="0" then %>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
					<% end if %>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></td> 
						<td>
							<select name="cctype">
							<% cardTypeArray=split(TXPCardTypes,", ")
							i=ubound(cardTypeArray)
							cardCnt=0
							do until cardCnt=i+1
								'response.write cardTypeArray(cardCnt)
								if cardTypeArray(cardCnt)="VISA" then %>
									<option value="VISA" selected>Visa</option>
								<% end if 
								if cardTypeArray(cardCnt)="MAST" then %>
									<option value="MAST">MasterCard</option>
								<% end if 
								if cardTypeArray(cardCnt)="AMER" then %>
									<option value="AMER">American Express</option>
								<% end if 
								if cardTypeArray(cardCnt)="DISC" then %>
									<option value="DISC">Discover</option>
								<% end if 
								cardCnt=cardCnt+1
							loop
							%>
						</select>
					</td>
				</tr>
				<tr> 
					<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
					<td> 
						<input type="text" name="cardnumber" value="">
					</td>
				</tr>
				<tr> 
					<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
					<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
						<select name="expmonth">
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
						<select name="expyear">
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
					<% if TXPReqCardCode="1" then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
								<input name="cvm" type="text" size="4" maxlength="4">
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% end If %>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					
					<tr> 
						<td colspan="2" align="center">
							<!--#include file="inc_gatewayButtons.asp"-->
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->