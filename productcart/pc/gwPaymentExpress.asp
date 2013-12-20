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
session("redirectPage")="gwPaymentExpress.asp"

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
query="SELECT pcPay_PaymentExpress_TestUsername, pcPay_PaymentExpress_TransType, pcPay_PaymentExpress_Username, pcPay_PaymentExpress_Password, pcPay_PaymentExpress_TestMode, pcPay_PaymentExpress_Cvc2, pcPay_PaymentExpress_ReceiptEmail FROM pcPay_PaymentExpress WHERE pcPay_PaymentExpress_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_PaymentExpress_TransType=rs("pcPay_PaymentExpress_TransType") ' auth or sale
pcPay_PaymentExpress_Username=rs("pcPay_PaymentExpress_Username") ' Username
pcPay_PaymentExpress_Password=rs("pcPay_PaymentExpress_Password") ' Password
pcPay_PaymentExpress_TestMode=rs("pcPay_PaymentExpress_TestMode")  ' test mode or live mode
pcv_CVV=rs("pcPay_PaymentExpress_Cvc2") ' cvc "on" or "off"
pcPay_PaymentExpress_ReceiptEmail=rs("pcPay_PaymentExpress_ReceiptEmail") ' additional receipt email
pcPay_PaymentExpress_TestUsername=rs("pcPay_PaymentExpress_TestUsername") ' test username

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	pcv_CardNumber=request.Form("CardNumber")
	pcv_ExpYear=request.Form("expYear")
	pcv_ExpMonth=request.Form("expMonth")
	pcv_CVN=request.Form("CVV")
	
	' validates expiration
	if DateDiff("d", Month(Now)&"/"&Year(now), pcv_ExpMonth&"/20"&pcv_ExpYear)<=-1 then
		response.redirect "msg.asp?message=67"       
	end if

	if not IsCreditCard(pcv_CardNumber) then
		response.redirect "msg.asp?message=68"       
	end if
	
	'// Check the integrity of the data
	reqFieldsOK = true
	
	'/  card number  
	If reqFieldsOK Then
		retVal = pcv_CardNumber
		if (retVal = "") then
			DeclinedString="Invalid credit card number"
			reqFieldsOK = false
		end if
	End If
	
	'/  expiration year 
	If reqFieldsOK Then
		retVal = pcv_ExpYear
		if (retVal = "") then
			DeclinedString="Invalid expiration year"
			reqFieldsOK = false
		end if
	End If
	
	'/ expiration date
	If reqFieldsOK Then
		retVal = pcv_ExpMonth
		if (retVal = "") then
			DeclinedString="Invalid expiry month"
			reqFieldsOK = false
		end if
	End IF
		
	'// pcv_CVN
	If pcv_CVV = 1 Then
		If reqFieldsOK Then
			retVal = pcv_CVN
			if (retVal = "") then
				DeclinedString="Missing Security Code"
				reqFieldsOK = false
			end if
		End IF
	End If
		
	'// ChargeType is one of SALE, AUTH
	If reqFieldsOK Then
		retVal = pcPay_PaymentExpress_TransType
		if (retVal = "") then
			DeclinedString="Invalid charge type"
			reqFieldsOK = false
		end if
	End IF
	
	If reqFieldsOK Then  ' start data integrity check
		'// GENERATE AN XML REQUEST
		sXmlAction = sXmlAction & "<Txn>"	
		' Testmode username or live mode username
		If pcPay_PaymentExpress_TestMode = 1 Then
			sXmlAction = sXmlAction & "<PostUsername>" & pcPay_PaymentExpress_TestUsername & "</PostUsername>"
		Else
			sXmlAction = sXmlAction & "<PostUsername>" & pcPay_PaymentExpress_Username & "</PostUsername>"
		End If
		
		sXmlAction = sXmlAction & "<PostPassword>" & pcPay_PaymentExpress_Password & "</PostPassword>"
		sXmlAction = sXmlAction & "<ReceiptEmail>" & pcPay_PaymentExpress_ReceiptEmail & "</ReceiptEmail>"
		sXmlAction = sXmlAction & "<CardHolderName>" & pcBillingFirstName & " " & pcBillingLastName & "</CardHolderName>"
		sXmlAction = sXmlAction & "<CardNumber>" & pcv_CardNumber & "</CardNumber>"
		sXmlAction = sXmlAction & "<DateExpiry>" & pcv_ExpMonth & pcv_ExpYear & "</DateExpiry>"
		sXmlAction = sXmlAction & "<Amount>" & money(pcBillingTotal) & "</Amount>"
		
		If pcv_CVV = 1 Then
			sXmlAction = sXmlAction & "<Cvc2>" & pcv_CVN & "</Cvc2>"
		End If
	
		If pcPay_PaymentExpress_TransType = "SALE" Then
			pcPay_PaymentExpress_TransType = "Purchase"
		Else
			pcPay_PaymentExpress_TransType = "Auth"
		End If
		sXmlAction = sXmlAction & "<TxnType>" & pcPay_PaymentExpress_TransType & "</TxnType>"
	

		'######### Use the orderID + date for this, it will be suitable  ######### 
		sXmlAction = sXmlAction & "<TxnId>" & session("GWOrderId") & Now() & "</TxnId>"
		
		'######### Use this for the order id  ######### 
		sXmlAction = sXmlAction & "<MerchantReference>" & session("GWOrderId") & "</MerchantReference>"	

		sXmlAction = sXmlAction & "</Txn>"


		'// SEND THE XML REQUEST
		PostURL="https://www.paymentexpress.com/pxpost.aspx"

		Dim objXMLhttp 
		Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP"&scXML) 
		objXMLhttp.Open "POST", PostURL ,False 
		
		objXMLhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXMLhttp.send sXmlAction

		'// GRAB THE XML REQUEST
		strRetVal = objXMLhttp.responsetext
		Set objXMLhttp = nothing

		'// CHECK THE RESPONSE CODE
		'		>  The status of a transaction is indicated by the Authorized element 
		'		>  (0 = Declined, 1 = Accepted)
		' 
		Dim objOutputXMLDoc
		Set objOutputXMLDoc=Server.CreateObject("Microsoft.XMLDOM")
		objOutputXMLDoc.loadXML(strRetVal)

		if objOutputXMLDoc.parseError.errorcode <> 0 then
			' oops... error in xml
			responseCodeText = responseCodeText & "unspecified error. Please contact the store owner.  <br />"
		else
				
			'/ Parse the XML document "case by case"
			'
			' Possible Cases Include:
			' 1. Authorized 0 or 1
			' 2. TransactionId
			' 3. AuthCode
			' 4. MerchantResponseDescription
				
			Set Nodes = objOutputXMLDoc.selectNodes("//Transaction")
			For Each Node In Nodes
				Authorized = Node.selectSingleNode("Authorized").Text
				bankTransactionId = Node.selectSingleNode("TransactionId").Text
				bankApprovalCode = Node.selectSingleNode("AuthCode").Text
				MerchantResponseDescription = Node.selectSingleNode("MerchantResponseDescription").Text
			Next
		
		end if

		'// PROCESS THE TRANSACTION
		If Authorized = 0 Then
			DeclinedString="The transaction was declined by the payment processor for the following reason(s): " & MerchantResponseDescription
			response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& DeclinedString &"<br><br><a href="""&tempURL&"?psslurl=gwPaymentExpress.asp&idCustomer="&session("idCustomer")&"&idOrder="&	session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		Else
			'create sessions
			session("GWTransId")=bankTransactionId
			session("GWAuthCode")=bankApprovalCode
			session("GWTransType")=pcPay_PaymentExpress_TransType
			
			call closedb()

			'Redirect to complete order
			response.redirect "gwReturn.asp?s=true&gw=PaymentExpress"
		End If
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	End If ' end data integrity check

'*************************************************************************************
' END
'*************************************************************************************
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
					<% if pcPay_PaymentExpress_TestMode = 1 then %>
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
					<% If pcv_CVV = 1 Then %>
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
<% function IsCreditCard(ByRef anCardNumber)
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
%>