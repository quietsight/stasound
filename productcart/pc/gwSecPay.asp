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
<!--#include file="gwSecPay_xmlrpc.asp" -->
<%
'//Set redirect page to the current file name
session("redirectPage")="gwSecPay.asp"

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
query="SELECT pcPay_SecPay_TransType, pcPay_SecPay_Username, pcPay_SecPay_Password, pcPay_SecPay_TestMode, pcPay_SecPay_Cvc2 FROM pcPay_SecPay WHERE pcPay_SecPay_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_SecPay_TransType=rs("pcPay_SecPay_TransType") ' auth or sale
pcPay_SecPay_Username=rs("pcPay_SecPay_Username") ' Username
pcPay_SecPay_Password=rs("pcPay_SecPay_Password") ' Password
pcPay_SecPay_TestMode=rs("pcPay_SecPay_TestMode")  ' test mode or live mode
pcv_CVV=rs("pcPay_SecPay_Cvc2") ' not supported by SECPAy at this time. Leave for scalability. //rs("pcPay_SecPay_Cvc2") ' cvc "on" or "off"
set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	dim strPostURL
	strPostURL="https://www.secpay.com/secxmlrpc/make_call"

	pcv_CardNumber=request.Form("CardNumber")
	pcv_ExpYear=request.Form("expYear")
	pcv_ExpMonth=request.Form("expMonth")
	pcv_CVN=request.Form("CVV")
	pcv_CardType=request.Form("CardType")

	' validates expiration
	if DateDiff("d", Month(Now)&"/"&Year(now), pcv_ExpMonth&"/20"&pcv_ExpYear)<=-1 then
		response.redirect "msg.asp?message=67"
	end if

	if not IsCreditCard(pcv_CardNumber) then
		response.redirect "msg.asp?message=68"
	end if
	' (2a) Check the integrity of the data
	'		> Do we have everything that we need?
	'
	reqFieldsOK = true

	'/  card number
	If reqFieldsOK Then
		retVal = pcv_CardNumber
		if (retVal = "") then
			DeclinedString="Invalid credit card number"
			reqFieldsOK = false
		end if
	End If

	If reqFieldsOK Then  ' Data is accurate
		Dim paramList(14)

		If pcPay_SecPay_TestMode = 1 Then
			paramList(0)= "secpay"
		Else
			paramList(0)= pcPay_SecPay_Username
		End If

		'paramList(1)
		' >> Testmode vpn pass or live mode vpn pass
		If pcPay_SecPay_TestMode = 1 Then
			paramList(1)= "secpay"
		Else
			paramList(1)= pcPay_SecPay_Password
		End If

		'paramList(2)
		' >> Testmode transid or live mode transid
		If pcPay_SecPay_TestMode = 1 Then
			paramList(2)= "testxml"
		Else
			paramList(2)= session("GWOrderId")
		End If

		paramList(3)= pcCustIpAddress
		paramList(4)= pcBillingFirstName & " " & pcBillingLastName
		paramList(5)= pcv_CardNumber
		paramList(6)= replace(money(pcBillingTotal),",","")
		paramList(7)= pcv_ExpMonth & "/" & pcv_ExpYear
		paramList(8)= ""
		paramList(9)= ""
		paramList(10)= ""'"prod=funny_book,amount=18.50;prod=sad_book,amount=16.50x3"
		paramList(11)= ""'"name=CONTACT,company=COMPANY,addr_1=ADDRESSLINE1,addr_2=ADDRESSLINE2,city=CITY,state=COUNTY,country=COUNTY,post_code=POST_CODE,tel=TELEPHONE,email=EMAIL,url=URL"
		paramList(12)= ""'"name=CONTACT,company=COMPANY,addr_1=ADDRESSLINE1,addr_2=ADDRESSLINE2,city=CITY,state=COUNTY,country=COUNTY,post_code=POST_CODE,tel=TELEPHONE,email=EMAIL,url=URL"

		'// Is it sale or auth (realtime or deferred)
		If pcPay_SecPay_TransType = "SALE" Then
			str_tmpval=""
		Else
			'str_tmpval="deferred=true,"
			str_tmpval="deferred=reuse,"
		End If

		If pcPay_SecPay_TestMode = 1 Then
			paramList(13)= str_tmpval & "test_status=true,dups=false,card_type=Visa,cv2=123"
		Else
			paramList(13)= str_tmpval & "test_status=live,dups=false,card_type="&pcv_CardType&",cv2="&pcv_CVN
		End If

		' Set the Method
		If pcPay_SecPay_TransType = "SALE" Then
			pcPay_SecPay_TransType = "SECVPN.validateCardFull"
		Else
			pcPay_SecPay_TransType = "SECVPN.validateCardFull"
		End If

		'//SEND THE XML REQUEST
		transaction = xmlRPC (strPostURL, pcPay_SecPay_TransType, paramList)

		'// GRAB THE XML REQUEST
		strRetVal = transaction

		'// CHECK THE RESPONSE CODE
		'		>  The status of a transaction is indicated by the Authorized element
		'		>  (0 = Declined, 1 = Accepted)
		'
		' set to not authorized initially, until we do some checks
		Authorized = 0

		'/ is there a struct, no struct, or error?
		If instr(1,strRetVal,"valid=",1) = 0 OR strRetVal = "" Then
			' oops... no struct or error

			' Are we in test mode?
			if pcPay_SecPay_TestMode = 1 then
				' show testing error for developer
				' note: error message can be expanded upon further by uncommenting the error call in gwSecPay_xmlrpc.asp
				MerchantResponseDescription = " XML RPC ERROR (" & strRetVal & ")"
			else
				' show friendly error for customer
				MerchantResponseDescription = " unspecified error. Please contact the store. "
			end if

		Else

			'/ Check for success or failure
			' Possible Results Include:
			' 1. valid=true (Authorized)
			' 2. valid=false (Declined)
			str_valid = trimstring(str_parse(strRetVal, "valid"))

			if str_valid = "true" then

				' Possible Cases Include:
				' 1. ?valid=true&trans_id=testxml&code=A&auth_code=9999&message=TEST AUTH&amount=99.0&test_status=true
				' 2. ?valid=false&trans_id=testxml&code=N&message=[some error message]&resp_code=5

				' authorize the transaction
				Authorized = 1
				' parse out some data with reg exp
				bankApprovalCode =  trimstring(str_parse(strRetVal, "auth_code"))
				bankTransactionId =  trimstring(str_parse(strRetVal, "trans_id"))

			else
				' decline the transaction, parse out an error message

				' show friendly error for customer
				str_message =    trimstring(str_parse(strRetVal, "message"))
				str_trans_id =    trimstring(str_parse(strRetVal, "trans_id"))
				str_code =    trimstring(str_parse(strRetVal, "code"))
				MerchantResponseDescription = "<br><b>Reason: </b>" & str_message & "."
				MerchantResponseDescription = MerchantResponseDescription& "<br><b>Error Code: </b>" & str_code
				MerchantResponseDescription = MerchantResponseDescription& "<br><b>Transaction ID: </b>" & str_trans_id
			end if

		End If

		'// PROCESS THE TRANSACTION
		'Authorized = 1
		'
		If Authorized = 0 Then
			DeclinedString="The transaction was declined by the payment processor for the following reason(s): " & MerchantResponseDescription
			response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& DeclinedString &"<br><br><a href="""&tempURL&"?psslurl=gwSecPay.asp&idCustomer="&session("idCustomer")&"&idOrder="&	session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			response.end
		Else
			'create sessions
			session("GWTransId")=bankTransactionId
			session("GWAuthCode")=bankApprovalCode
			session("GWTransType")=pcPay_SecPay_TransType
			call closedb()

			'Redirect to complete order
			response.redirect "gwReturn.asp?s=true&gw=SecPay"

		End If
	End If
'*************************************************************************************
' END
'*************************************************************************************
end if

' strips out the unwanted text
function trimstring(strQ)
	nIndex = InStrRev(strQ,"=")
	If (nIndex>0) Then
		strQ = Right(strQ,Len(strQ)-nIndex)
	End If
	strQ = replace(strQ,"%20"," ")
	strQ = replace(strQ,"%3B",";")
	strQ = replace(strQ,"&","")
	strQ = replace(strQ,"+"," ")
	trimstring = strQ
end function

' can parse out a value from a string
Function str_parse(str, str_value)
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = str_value & "=(.*?)[\&$]"
  Set ExpMatch = ExpReg.Execute(str)
  If ExpMatch.count > 0 Then
	For each ExpMatched in ExpMatch
		str_parse = ExpMatched.Value
	Next
  Else
	str_parse = Null
  End If
  Set ExpReg = Nothing
End Function

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
						<td><p><%=pcBillingFirstName&" "&pcBillingLastName%></p></td>
					</tr>
					<tr>
						<td><p><%=pcBillingAddress%></p></td>
					</tr>
					<% if pcBillingAddress2<>"" then %>
					<tr>
						<td><p><%=pcBillingAddress2%></p></td>
					</tr>
					<% end if %>
					<tr>
						<td><p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% if pcPay_SecPay_TestMode = 1 then %>
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
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
						<td>
						<select name="CardType">
							<option value="Delta">Delta</option>
							<option value="Master Card Credit">Master Card Credit</option>
							<option value="Debit Master Card">Debit Master Card</option>
							<option value="Solo">Solo</option>
							<option value="Maestro">Maestro</option>
							<option value="Visa">Visa</option>
							<option value="Laser">Laser</option>
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