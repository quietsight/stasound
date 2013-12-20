<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/pcPayPalUKClass.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="header.asp"-->
<% response.Buffer=true %>
<script language="javascript">
	function generateCC(){
		var cc_number = new Array(16);
		var cc_len = 16;
		var start = 0;
		var rand_number = Math.random();
		
		switch(document.PaymentForm.creditCardType.value)
				{
			case "Visa":
				cc_number[start++] = 4;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Discover":
				cc_number[start++] = 6;
				cc_number[start++] = 0;
				cc_number[start++] = 1;
				cc_number[start++] = 1;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "MasterCard":
				cc_number[start++] = 5;
				cc_number[start++] = Math.floor(Math.random() * 5) + 1;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Amex":
				cc_number[start++] = 3;
				cc_number[start++] = Math.round(Math.random()) ? 7 : 4 ;
				cc_len = 15;
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Maestro":
				cc_number[start++] = 5;
				cc_len = 16;
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
			case "Solo":
				cc_number[start++] = 6;
				cc_len = 16;
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
				}
				
				for (var i = start; i < (cc_len - 1); i++) {
			cc_number[i] = Math.floor(Math.random() * 10);
				}
		
		var sum = 0;
		for (var j = 0; j < (cc_len - 1); j++) {
			var digit = cc_number[j];
			if ((j & 1) == (cc_len & 1)) digit *= 2;
			if (digit > 9) digit -= 9;
			sum += digit;
		}
		
		var check_digit = new Array(0, 9, 8, 7, 6, 5, 4, 3, 2, 1);
		cc_number[cc_len - 1] = check_digit[sum % 10];
		
		document.PaymentForm.CardNumber.value = "";
		for (var k = 0; k < cc_len; k++) {
			document.PaymentForm.CardNumber.value += cc_number[k];
		}
	}
	function generateCC2(){
		switch(document.PaymentForm.creditCardType.value)
				{
			case "Visa":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Discover":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "MasterCard":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Amex":
				document.getElementById("UKOptions").style.display = 'none';
				document.getElementById("USOptions").style.display = '';
				break;
			case "Maestro":
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
			case "Solo":
				document.getElementById("UKOptions").style.display = '';
				document.getElementById("USOptions").style.display = 'none';
				break;
				}
	}
</script>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<img src="images/checkout_bar_step5.gif" alt="">
			</td>
		</tr>
		<tr>
			<td>
<% 
if session("GWOrderDone")="YES" then
	'tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
	'tempURL=replace(tempURL,"https:/","https://")
	'tempURL=replace(tempURL,"http:/","http://") 
	'session("GWOrderDone")=""
	'response.redirect tempURL
end if
		
dim query, rs, conntemp
%>
		
<!-- #Include File="pcPay_Cent_XMLFunctions.asp"-->
<!-- #Include File="pcPay_Cent_Utility.asp"-->

<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalWPP"

'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass

'//Set redirect page to the current file name
session("redirectPage")="gwPayPalUK.asp"

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

'// Open Db
call opendb()

'///////////////////////////////////////////////////////////////////////////////
'// START: GET DATA FROM DB
 '///////////////////////////////////////////////////////////////////////////////

'// Declare Local Variables at once
'>>> pcPay_PayPal_TransType
'>>> PaymentAction
'>>> pcPay_PayPal_Username
'>>> pcPay_PayPal_Password
'>>> pcPay_PayPal_Sandbox
'>>> pcPay_PayPal_Method
'>>> pcPay_PayPal_Signature
objPayPalClass.pcs_SetAllVariables()
objPayPalClass.pcs_SetShipAddress((int(session("GWOrderId"))-scpre))

'///////////////////////////////////////////////////////////////////////////////
'// END: GET DATA FROM DB
'///////////////////////////////////////////////////////////////////////////////


pIdOrder2=(int(session("GWOrderID"))-scpre)

query="SELECT details, comments, taxAmount, shipmentDetails FROM orders WHERE idOrder="&pIdOrder2
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pOrderDetails = rs("details")
pOrderDetails = replace(pOrderDetails,"Amount: ||"," $")
pOrderComments = rs("comments")
pOrderTax = round(rs("taxAmount"),2)
pshipmentDetails = rs("shipmentDetails")
set rs=nothing

'get shipping details...
shipping=split(pshipmentDetails,",")
if ubound(shipping)>1 then
	if NOT isNumeric(trim(shipping(2))) then
		pshipmentCharge="0"
	else
		pshipmentCharge=trim(shipping(2))
		if ubound(shipping)=>3 then
			serviceHandlingFee=trim(shipping(3))
			if NOT isNumeric(serviceHandlingFee) then
				serviceHandlingFee=0
			end if
		else
			serviceHandlingFee=0
		end if
		pshipmentCharge = round(pshipmentCharge,2) + round(serviceHandlingFee,2)
	end if
else
	pshipmentCharge="0"
end if

'// Close Db
call closedb()

IF (Request.ServerVariables("Content_Length") > 0 AND request("PaymentSubmitted")="Go") OR request.QueryString("centinel")<>"" then

	%>
    <!--#include file="pcCentinelInclude.asp"-->
    <%
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' (2) HANDLE POST BACK FORM DATA  
	'		> Important billing info  
	'	
	CardNumber=request.Form("CardNumber")
	expYear=request.Form("expYear")
	expMonth=request.Form("expMonth")
	expYear2=request.Form("expYear2")
	expMonth2=request.Form("expMonth2")
	startYear=request.Form("startYear")
	startMonth=request.Form("startMonth")
	CVV=request.Form("CVV")
	CC_TYPE=request.Form("creditCardType")
	ISSUENUMBER=request.Form("ISSUENUMBER")	
	if startYear<>"" then
		startYear=right(startYear,2)
	end if
	if expYear<>"" then
		expYear=right(expYear,2)
	end if
	if expYear2<>"" then
		expYear2=right(expYear2,2)
	end if
	
	' (2a) Check the integrity of the data
	'		> Do we have everything that we need?
	'
	reqFieldsOK = true
	
	' ####  card number  
	If reqFieldsOK Then
		retVal = CardNumber
		if (retVal = "") then
			DeclinedString="Invalid credit card number"
			reqFieldsOK = false
		end if
	End If
	
	' ####  valid card number
	if not IsCreditCard(CardNumber) AND (CC_TYPE<>"Solo" AND CC_TYPE<>"Maestro") then
			DeclinedString="You have not entered a valid credit card number"
			reqFieldsOK = false      
	end if
	
	' ####  expiration year 
	If reqFieldsOK Then
		retVal = expYear
		if (retVal = "") then
			DeclinedString="Invalid expiration year"
			reqFieldsOK = false
		end if
	End If
	
	' ####  CVV
	if pcPay_PayPal_CVV=1 then
		If reqFieldsOK Then
			retVal = CVV
			if (retVal = "") then
				DeclinedString="Missing CVV Security Code"
				reqFieldsOK = false
			end if
		End IF
	End If
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	If reqFieldsOK Then  ' start data integrity check conditional submission
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
		'///////////////////////////////////////////////////////////////////////////////
		'// START: Direct Payment Method
		'///////////////////////////////////////////////////////////////////////////////

		
		'***********************************************************************
		'// Start: Posting Details to PayPal
		'***********************************************************************

		'---------------------------------------------------------------------------
		' Construct the parameter string that describes the PayPal payment the varialbes 
		' were set in the web form, and the resulting string is stored in nvpstr
		'
		' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
		'---------------------------------------------------------------------------
		nvpstr="" '// clear 

		objPayPalClass.AddNVP "CLIENTIP", pcCustIpAddress
		objPayPalClass.AddNVP "AMT", pcf_CurrencyField(money(pcBillingTotal))
		objPayPalClass.AddNVP "ACCT", CardNumber
		if CVV<>"" then
			objPayPalClass.AddNVP "CVV2", CVV
		end if	
		objPayPalClass.AddNVP "FIRSTNAME", pcBillingFirstName
		objPayPalClass.AddNVP "LASTNAME", pcBillingLastName
		objPayPalClass.AddNVP "STREET", pcBillingAddress
		objPayPalClass.AddNVP "CITY", pcBillingCity
		objPayPalClass.AddNVP "STATE", pcBillingState
		objPayPalClass.AddNVP "ZIP", pcBillingPostalCode
		objPayPalClass.AddNVP "COUNTRY", pcBillingCountryCode		
		objPayPalClass.AddNVP "CURRENCY", pcPay_PayPal_Currency
		objPayPalClass.AddNVP "BUTTONSOURCE", "ProductCart_Cart_DP_US"
		objPayPalClass.AddNVP "INVNUM", session("GWOrderId")
		if CC_TYPE="Solo" OR CC_TYPE="Maestro" then			
			if startMonth<>"" AND startYear<>"" then
				objPayPalClass.AddNVP "CARDSTART", startMonth & startYear
			end if				
			if ISSUENUMBER<>"" then 
				objPayPalClass.AddNVP "CARDISSUE", ISSUENUMBER
			end if		
			objPayPalClass.AddNVP "EXPDATE", expMonth2 & expYear2		
			objPayPalClass.AddNVP "ACCTTYPE", "MasterCard" '// fix paypal bug			
		else
			objPayPalClass.AddNVP "ACCTTYPE", CC_TYPE
			objPayPalClass.AddNVP "EXPDATE", expMonth & expYear			 		
		end if		
		objPayPalClass.AddNVP "TRXTYPE", PaymentAction
		objPayPalClass.AddNVP "TENDER", "C"	
		
		'// Centinel		
		If pcPay_Cent_Active=1 AND pcPay_CentByPass=1 AND pcPay_CardType="YES" Then
			objPayPalClass.AddNVP "ECI3DS", Session("Centinel_ECI")
			objPayPalClass.AddNVP "CAVV", Session("Centinel_CAVV")
			objPayPalClass.AddNVP "AUTHSTATUS3DS", Session("Centinel_PAResStatus") '// PAResStatus
			objPayPalClass.AddNVP "MPIVENDOR3DS", Session("Centinel_Enrolled") '// Enrolled
			objPayPalClass.AddNVP "XID", Session("Centinel_XID") '// Xid
			'// VERSION=59.0
		End If

		'***********************************************************************
		'// Start: Address Override
		'***********************************************************************	
		if pcv_strShippingStateCode="" OR isNULL(pcv_strShippingStateCode)=True then
			pcv_strShippingStateCode=pcv_strShippingProvince
		end if
		if pcv_strShippingStateCode<>"" AND isNULL(pcv_strShippingStateCode)=False then
			objPayPalClass.AddNVP "SHIPTONAME", pcv_strShippingFullName
			objPayPalClass.AddNVP "SHIPTOSTREET", pcv_strShippingAddress
			objPayPalClass.AddNVP "SHIPTOCITY", pcv_strShippingCity
			objPayPalClass.AddNVP "SHIPTOSTATE", pcv_strShippingStateCode
			objPayPalClass.AddNVP "SHIPTOZIP", pcv_strShippingPostalCode
			objPayPalClass.AddNVP "SHIPTOCOUNTRY", pcv_strShippingCountryCode
		end if
		'***********************************************************************
		'// End: Address Override
		'***********************************************************************	



		'--------------------------------------------------------------------------- 
		' Make the call to PayPal to set the Express Checkout token
		' If the API call succeded, then redirect the buyer to PayPal
		' to begin to authorize payment.  If an error occurred, show the
		' resulting errors
		'---------------------------------------------------------------------------
		Set resArray = objPayPalClass.hash_call("doDirectPayment",nvpstr)
		Set Session("nvpResArray")=resArray
		ack = UCase(resArray("RESULT"))
		
		if err.number <> 0 then	
			'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
			Dim pcv_PayPalErrMessage
			%><!--#include file="../includes/pcPayPalErrors.asp"--><%										
		end if
		
		'clear Centinel sessions
		Session("Centinel_Enrolled")=""
		Session("Centinel_ErrorNo")=""
		Session("Centinel_ErrorDesc")=""
		Session("Centinel_PAResStatus")=""
		Session("Centinel_SignatureVerification")=""
		Session("Centinel_ECI")=""
		Session("Centinel_XID")=""
		Session("Centinel_CAVV")=""
		Session("Centinel_ErrorNo")=""
		Session("Centinel_ErrorDesc")=""
		Session("Centinel_TransactionId")=""
		Session("Centinel_ACSURL")=""
		Session("Centinel_PAYLOAD")=""
		
		If ack=0 Then

			session("GWTransId")=resArray("CORRELATIONID")
			session("AVSCode")=resArray("AVSADDR")
			session("CVV2Code")=resArray("CVV2MATCH")
			session("GWAuthCode")=""
			session("GWTransType")=pcPay_PayPal_TransType

			if session("GWTransId") <> "" then			
			
				'// Save info in pcPay_PayPal_Authorize if "A"			
				If PaymentAction="A" Then
					call opendb()
					Dim pTodaysDate
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
					query="INSERT INTO pcPay_PayPal_Authorize (idOrder, amount, paymentmethod, transtype, authcode, idCustomer, captured, AuthorizedDate, CurrencyCode) VALUES ("&pcGatewayDataIdOrder&", "&pcBillingTotal&", 'PayPalWP', '"&paymentAction&"', '"&session("GWTransId")&"', "&pcIdCustomer&", 0," & tmpStr & ", '"&pcPay_PayPal_Currency&"');"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)				
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					call closedb()
				End If		
			
				response.redirect "gwReturn.asp?s=true&gw=PayPalWP"
				
			else			
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Start: Error Reporting
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				'// Combine User Friendly Errors from "pcPay_PayPal_Errors.asp"
				'// with Code errors from string "Declined String".
				'// Return a formatted error report as the string "pcv_PayPalErrMessage".
				objPayPalClass.GenerateErrorReport()
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' End: Error Reporting
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
				
			end if
		
		'// Unsuccessful Express Checkout / Transaction Not Complete
		Else	
		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start: Error Reporting
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			'// Combine User Friendly Errors from "pcPay_PayPal_Errors.asp"
			'// with Code errors from string "DeclinedString".
			'// Return a formatted error report as the string "pcv_PayPalErrMessage".
			objPayPalClass.GenerateErrorReport()
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' End: Error Reporting
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
		
		End If

	Else '// If reqFieldsOK Then
	
		pcv_PayPalErrMessage = DeclinedString
	
	End If ' end data integrity check

	'*************************************************************************************
	' END
	'*************************************************************************************
end if 
%>

				<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
				<input type="hidden" name="PaymentSubmitted" value="Go">
					<table class="pcShowContent">

					<% if pcv_PayPalErrMessage <> "" then %>
						<tr> 
							<td colspan="2">
								<div class="pcErrorMessage">
									The transaction was not performed for the following reasons: 
									<%=pcv_PayPalErrMessage%>
								</div>
							</td>
						</tr>
					<% end if %>
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
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td nowrap="nowrap" width="20%"><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
						<td width="80%">

							<% if pcPay_PayPal_Method = "sandbox" then 
                                response.write "<select name=""creditCardType"" onChange=""javascript:generateCC(); return false;"">"
                            else %>
                                <select name="creditCardType" onChange="javascript:generateCC2(); return false;">
                            <% end if %>	
                                <% 	
								cardTypeArray=split(pcPay_PayPal_CardTypes,", ")
								i=ubound(cardTypeArray)
								cardCnt=0
								do until cardCnt=i+1
									cardVar=cardTypeArray(cardCnt)
									select case cardVar
										case "V"
                                                response.write "<option value=""Visa"" selected>Visa</option>"
                                                cardCnt=cardCnt+1
										case "M" 
											response.write "<option value=""MasterCard"">MasterCard</option>"
											cardCnt=cardCnt+1
										case "A"
											response.write "<option value=""Amex"">American Express</option>"
											cardCnt=cardCnt+1
										case "D"
											response.write "<option value=""Discover"">Discover</option>"
											cardCnt=cardCnt+1
									end select
								loop
                                %>
								<% If PaymentAction="A" AND pcPay_PayPal_Currency="GBP" Then %>
								<option value="Maestro" <%if CC_TYPE="Maestro" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_20")%></option>
								<option value="Solo" <%if CC_TYPE="Solo" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_21")%></option>
								<% End If %>
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
						<td colspan="2">
							
							<%
							'// Maestro/ Solo Cards
							%>
							<table id="UKOptions" style="display:none">
								<tr> 
									<td nowrap="nowrap" width="92px"><%=dictLanguage.Item(Session("language")&"_GateWay_13")%></td>
									<td align="left">  
										<input name="ISSUENUMBER" type="text" id="ISSUENUMBER" value="" size="2" maxlength="2">
									</td>
								</tr>
								<tr> 
									<td><%=dictLanguage.Item(Session("language")&"_GateWay_14")%></td>
									<td align="left"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
										<select name="startMonth">
											<option value="" selected></option>
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
										<select name="startYear">
											<option value="" selected></option>
											<option value="<%=(dtCurYear-10)%>"><%=dtCurYear-10%></option>
											<option value="<%=(dtCurYear-9)%>"><%=dtCurYear-9%></option>
											<option value="<%=(dtCurYear-8)%>"><%=dtCurYear-8%></option>
											<option value="<%=(dtCurYear-7)%>"><%=dtCurYear-7%></option>
											<option value="<%=(dtCurYear-6)%>"><%=dtCurYear-6%></option>
											<option value="<%=(dtCurYear-5)%>"><%=dtCurYear-5%></option>
											<option value="<%=(dtCurYear-4)%>"><%=dtCurYear-4%></option>
											<option value="<%=(dtCurYear-3)%>"><%=dtCurYear-3%></option>
											<option value="<%=(dtCurYear-2)%>"><%=dtCurYear-2%></option>
											<option value="<%=(dtCurYear-1)%>"><%=dtCurYear-1%></option>											
											<option value="<%=(dtCurYear)%>"><%=dtCurYear%></option>
										</select>
									</td>
								</tr>
								<tr> 
									<td><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></td>
									<td align="left"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
										<select name="expMonth2">
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
										<select name="expYear2">
											<option value="<%=(dtCurYear)%>" selected><%=dtCurYear%></option>
											<option value="<%=(dtCurYear+1)%>"><%=dtCurYear+1%></option>
											<option value="<%=(dtCurYear+2)%>"><%=dtCurYear+2%></option>
											<option value="<%=(dtCurYear+3)%>"><%=dtCurYear+3%></option>
											<option value="<%=(dtCurYear+4)%>"><%=dtCurYear+4%></option>
											<option value="<%=(dtCurYear+5)%>"><%=dtCurYear+5%></option>
											<option value="<%=(dtCurYear+6)%>"><%=dtCurYear+6%></option>
											<option value="<%=(dtCurYear+7)%>"><%=dtCurYear+7%></option>
											<option value="<%=(dtCurYear+8)%>"><%=dtCurYear+8%></option>
											<option value="<%=(dtCurYear+9)%>"><%=dtCurYear+9%></option>
											<option value="<%=(dtCurYear+10)%>"><%=dtCurYear+10%></option>
										</select>
										<div class="pcSmallText"><%=dictLanguage.Item(Session("language")&"_GateWay_15")%></div>
									</td>
								</tr>
							</table>
							
							<%
							'// Visa/ MasterCard/ Discover/ AMEX
							%>
							<table id="USOptions" style="display:''">
								<tr> 
									<td nowrap="nowrap" width="30%"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></td>
									<td align="left"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
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
											<option value="<%=(dtCurYear)%>" selected><%=dtCurYear%></option>
											<option value="<%=(dtCurYear+1)%>"><%=dtCurYear+1%></option>
											<option value="<%=(dtCurYear+2)%>"><%=dtCurYear+2%></option>
											<option value="<%=(dtCurYear+3)%>"><%=dtCurYear+3%></option>
											<option value="<%=(dtCurYear+4)%>"><%=dtCurYear+4%></option>
											<option value="<%=(dtCurYear+5)%>"><%=dtCurYear+5%></option>
											<option value="<%=(dtCurYear+6)%>"><%=dtCurYear+6%></option>
											<option value="<%=(dtCurYear+7)%>"><%=dtCurYear+7%></option>
											<option value="<%=(dtCurYear+8)%>"><%=dtCurYear+8%></option>
											<option value="<%=(dtCurYear+9)%>"><%=dtCurYear+9%></option>
											<option value="<%=(dtCurYear+10)%>"><%=dtCurYear+10%></option>
										</select>
									</td>
								</tr>
								
							</table>
						
						</td>
					</tr>
					<% if pcPay_PayPal_CVV=1 then %>
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
						<td><p>&nbsp;&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					<% 
					'////////////////////////////////////////////////////////////////////
					'// Start: Cardinal Commerce
					'////////////////////////////////////////////////////////////////////
					pcPay_Cent_Active=1
					if pcPay_Cent_Active=1 then %>
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
					<% end if 
					'////////////////////////////////////////////////////////////////////
					'// End: Cardinal Commerce
					'////////////////////////////////////////////////////////////////////
					%>
					<tr> 
						<td colspan="2" align="center">
							<!--#include file="inc_gatewayButtons.asp"-->
						</td>
					</tr>
				</table>
			</form>
			<script language="javascript">
				<% if pcPay_PayPal_Method = "sandbox" then %>
					generateCC();
				<% else %>
					generateCC2();
				<% end if %>	
			</script>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->
<%
'*************************************************************************************
' FUNCTIONS
' START
'
'*************************************************************************************
function IsCreditCard(ByRef anCardNumber)
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
		if IsNumeric(lsChar) Then lsNumber = lsNumber & lsChar
		
	Next ' lnPosition
		
	' ====
	' The credit card number must be between 13 and 16 digits.
	' ====
	' if the length of the number is less Then 13 digits, then Exit the routine
	if Len(lsNumber) < 13 Then Exit function
		
	' if the length of the number is more Then 16 digits, then Exit the routine
	if Len(lsNumber) > 16 Then Exit function
    			    			
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

'*************************************************************************************
' FUNCTIONS
' END
'*************************************************************************************
%>