<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwParaData.asp"

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
	query="SELECT pcPay_ParaData_TransType, pcPay_ParaData_Key, pcPay_ParaData_TestMode, pcPay_ParaData_AVS, pcPay_ParaData_CVC FROM pcPay_ParaData WHERE pcPay_ParaData_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_ParaData_TransType=rs("pcPay_ParaData_TransType") ' auth or sale
pcPay_ParaData_Key=rs("pcPay_ParaData_Key") ' private key
pcPay_ParaData_TestMode=rs("pcPay_ParaData_TestMode")  ' test mode or live mode
pcPay_ParaData_AVS=rs("pcPay_ParaData_AVS") ' avs "on" or "off"
pcv_CVV=rs("pcPay_ParaData_CVC") ' cvc "on" or "off"

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then
	
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	'// Istantiate the COM client
	set eClient = Server.CreateObject("Paygateway.EClient.1")

	'// Initialize
	reqFieldsOK = true

	'// Set the required fields
	'
	' card number  
	If reqFieldsOK Then
		retVal = eClient.SetCreditCardNumber(request.Form("CardNumber"))
		if (retVal = 0) then
			DeclinedString="Invalid credit card number"
			reqFieldsOK = false
		end if
	End If

	' expiration year 
	If reqFieldsOK Then
		retVal = eClient.SetExpireYear(request.Form("expYear"))
		if (retVal = 0) then
			DeclinedString="Invalid expiration year"
			reqFieldsOK = false
		end if
	End If

	' expiration date
	If reqFieldsOK Then
		retVal = eClient.SetExpireMonth(request.Form("expMonth"))
		if (retVal = 0) then
			DeclinedString="Invalid expiry month"
			reqFieldsOK = false
		end if
	End IF
	
	' amount 
	If reqFieldsOK Then
		retVal = eClient.SetChargeTotal(pcBillingTotal)
		if (retVal = 0) then
			DeclinedString="Invalid charge total"
			reqFieldsOK = false
		end if
	End If
	
	'  ChargeType is one of SALE, AUTH, CAPTURE, VOID, CREDIT
	If reqFieldsOK Then
		retVal = eClient.SetChargeType(pcPay_ParaData_TransType)
		if (retVal = 0) then
			DeclinedString="Invalid charge type"
			reqFieldsOK = false
		end if
	End IF
	
	' 7) Set The Optional Fields
	IF reqFieldsOK THEN
		' Order fields
		retVal = eClient.SetOrderId(session("GWOrderId"))
	
		' Credit card fields
		if pcv_CVV = 1 then
			retVal = eClient.SetCreditCardVerificationNumber(request.Form("CVV"))
		end if
	
		' Billing, Shipping and Customer information
		' Billing
		retVal = eClient.SetBillAddressOne(pcBillingAddress)
		retVal = eClient.SetBillAddressTwo(pcBillingAddress2)
		retVal = eClient.SetBillCity(pcBillingCity)
		retVal = eClient.SetShipStateOrProvince(pcBillingState)
		retVal = eClient.SetBillCompany(pcBillingCompany)
		retVal = eClient.SetBillCountryCode(pcBillingCountryCode)
		retVal = eClient.SetBillEmail(pcCustomerEmail)
		retVal = eClient.SetBillFirstName(pcBillingFirstName)
		retVal = eClient.SetBillLastName(pcBillingLastName)
		retVal = eClient.SetBillPhone(pcBillingPhone)
		retVal = eClient.SetBillPostalCode(pcBillingPostalCode)
		' Customer and order info
		retVal = eClient.SetCustomerIpAddress(pcCustIpAddress)		
		retVal = eClient.SetOrderCustomerId(session("idCustomer"))
		retVal = eClient.SetOrderDescription(pOrderDetails)	 '//!!
		retVal = eClient.SetOrderUserId(session("idCustomer"))	
		retVal = eClient.SetPurchaseOrderNumber(session("idCustomer"))
		retVal = eClient.SetBillNote(pOrderComments) '//!!
		retVal = eClient.SetTaxAmount(pOrderTax) '//!!
		' Shipping info
		retVal = eClient.SetShipAddressOne(pcShippingAddress)
		retVal = eClient.SetShipAddressTwo(pcShippingAddress2)
		retVal = eClient.SetShipCity(pcShippingCity)
		retVal = eClient.SetShipStateOrProvince(pcShippingState)
		retVal = eClient.SetShipPostalCode(pcShippingPostalCode)
		retVal = eClient.SetShipCountryCode(pcShippingCountryCode)
		retVal = eClient.SetShipCompany(pcShippingCompany)
		retVal = eClient.SetShipFirstName(pcShippingFirstName)
		retVal = eClient.SetShipLastName(pcShippingLastName)
		retVal = eClient.SetShippingCharge(pshipmentCharge) '//!!
		retVal = eClient.SetShipEmail(pcCustomerEmail)
		retVal = eClient.SetShipPhone(pcShippingPhone)
		' Application name
		retVal = eClient.SetCartridgeType("ProductCart v3")

	END IF

	' 8) Process The Transaction
	IF reqFieldsOK THEN
		' Are we in test mode?
		if pcPay_ParaData_TestMode = 1 then
			pcPay_ParaData_Key = "195325FCC230184964CAB3A8D93EEB31888C42C714E39CBBB2E541884485D04B"
		end if

		retVal = eClient.DoTransaction("trans_key", pcPay_ParaData_Key)
		noErrors = CheckAndDisplayError(retVal)
		If noErrors Then
			eClient.GetOrderId orderId
			eClient.GetResponseCode responseCode
			eClient.GetResponseCodeText responseCodeText
			eClient.GetBankApprovalCode bankApprovalCode
			eClient.GetBankTransactionId bankTransactionId
			eClient.GetIsoCode isoCode
			eClient.GetBatchId batchId

			eClient.GetReferenceId responseReferenceId
			eClient.GetTimeStamp timeStamp
			eClient.GetTimeStampFormatted timeStampFormatted
			' Order fields
			eClient.GetOrderId orderId			
			eClient.GetOrderUserId orderUserId
			eClient.GetInvoiceNumber invoiceNumber
			'eClient.GetTaxExempt taxExempt
			eClient.GetShippingChargeStr shippingCharge
			eClient.GetTransactionConditionCode tcc
			eClient.GetBuyerCode buyerCode
			eClient.GetOrderDescription orderDescription
			' Billing fields
			eClient.GetBillFirstName billFirstName			
			eClient.GetBillLastName  billLastName			
			eClient.GetBillAddressOne billAddressOne
			eClient.GetBillAddressTwo billAddressTwo
			eClient.GetBillCity billCity
			eClient.GetBillStateOrProvince billStateOrProvince
			eClient.GetBillCountryCode billCountryCode
			eClient.GetBillPostalCode billPostalCode
			eClient.GetBillPhone billPhone			
			eClient.GetBillEmail billEmail
			eClient.GetBillNote billNote
	
			' Testing code
			'response.write responseCodeText
			'response.write responseCode
			'response.End()
				
			'ex: "Test transaction response: Transaction successful.1 "
			If responseCode=1 Then
				Result="APPROVED"
			Else
				Result="ERROR"
			End If
		End If
	ELSE

	DeclinedString = "The transaction was not performed. " & DeclinedString

	END IF
	eClient.CleanUp
	set eClient = nothing
 
	' 9) Handle Errors "ERROR" or "APPROVED" or "DECLINED" or "MISSING"
	'    GIVE MESSAGE AND SHOW FORM AGAIN
	If Result="ERROR" Then
		DeclinedString="The transaction was declined by the payment processor for the following reason(s): " & responseCodeText
	End If

	' 10) Handle the transaction approved
	' redirect to gwReturn.asp with proper values
	If Result="APPROVED" Then
		session("GWAuthCode")=bankApprovalCode
		session("GWTransId")=bankTransactionId
		session("GWTransType")=pcPay_ParaData_TransType
		call closedb()
			
		'Redirect to complete order
		response.redirect "gwReturn.asp?s=true&gw=ParaData"
	End If
	
	Msg=DeclinedString

'*************************************************************************************
' END
'*************************************************************************************
end if 

'### Check for errors ###
function CheckAndDisplayError(retVal)
	if (retVal = 0)  then
		eClient.GetErrorString errorString
		Response.Write(errorString)
		CheckAndDisplayError = false
	else
		CheckAndDisplayError = true
	end if
end function

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
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% if pcPay_ParaData_TestMode = 1 then %>
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
								<option value="<%=right(dtCurYear,4)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,4)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,4)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,4)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,4)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,4)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,4)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,4)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,4)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,4)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,4)%>"><%=dtCurYear+10%></option>
							</select>
						</td>
					</tr>
					<% If pcv_CVV="1" Then %>
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