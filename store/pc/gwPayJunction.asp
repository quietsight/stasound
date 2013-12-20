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
'======================================================================================
'// Set redirect page
'======================================================================================
' The redirect page tells the form where to post the payment information. Most of the 
' time you will redirect the form back to this page.
'======================================================================================
session("redirectPage")="gwPayJunction.asp"  'ALTER

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE	
'======================================================================================

': Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
': End Declare and Retrieve Customer's IP Address	

': Declare URL path to gwSubmit.asp	
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
': End Declare URL path to gwSubmit.asp

': Get Order ID and Set to session
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if
': End Get Order ID
	
': Get customer and order data from the database for this order	
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<%
': End Get customer and order data


': Reset customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
': End Reset customer session

': Open Connection to the DB
dim connTemp, rs 'DELETE FOR HARD CODED VARS
call openDb() 'DELETE FOR HARD CODED VARS
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'======================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your 
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
	query="SELECT pcPay_PJ_MerchantID,pcPay_PJ_MerchantPassword,pcPay_PJ_TransType,pcPay_PJ_cardTypes,pcPay_PJ_CVC,pcPay_PJ_TestMode FROM pcPay_PayJunction Where pcPay_PJ_ID=1;"
'ALTER :: DELETE FOR HARD CODED VARS
'======================================================================================
'// End custom query
'======================================================================================

': Create recordset and execute query
set rs=server.CreateObject("ADODB.RecordSet") 'DELETE FOR HARD CODED VARS
set rs=connTemp.execute(query) 'DELETE FOR HARD CODED VARS

': Capture any errors
if err.number<>0 then 'DELETE FOR HARD CODED VARS
	call LogErrorToDatabase() 'DELETE FOR HARD CODED VARS
	set rs=nothing 'DELETE FOR HARD CODED VARS
	call closedb() 'DELETE FOR HARD CODED VARS
	response.redirect "techErr.asp?err="&pcStrCustRefID 'DELETE FOR HARD CODED VARS
end if 'DELETE FOR HARD CODED VARS

'======================================================================================
'// Set gateway specific variables - These can be your "hard coded variables" or 
'// Variables retrieved from the database.
'======================================================================================
	pcPay_PJ_MerchantID=rs("pcPay_PJ_MerchantID")
	pcPay_PJ_MerchantID=enDeCrypt(pcPay_PJ_MerchantID, scCrypPass)
	pcPay_PJ_MerchantPassword=rs("pcPay_PJ_MerchantPassword")
	pcPay_PJ_MerchantPassword=enDeCrypt(pcPay_PJ_MerchantPassword, scCrypPass) 
	pcPay_PJ_TransType = rs("pcPay_PJ_TransType")
	pcPay_PJ_TestMode = rs("pcPay_PJ_TestMode")
	pcPay_PJ_cardTypes = rs("pcPay_PJ_cardTypes")
	pcPay_PJ_CVC = rs("pcPay_PJ_CVC")
'======================================================================================
'// End gateway specific variables
'======================================================================================

': Clear recordset and close db connection
set rs=nothing 'DELETE FOR HARD CODED VARS
call closedb() 'DELETE FOR HARD CODED VARS

'======================================================================================
'// If you are posting back to this page from the gateway form, all actions will happen 
'// here. 
'======================================================================================
if request("PaymentSubmitted")="Go" then
  
			
			'*************************************************************************************
			'// This is where you would post and retrieve info to and from the gateway
			'// START below this line
			'*
			' Link these to your form elements. Make sure to error check if needed.
			if pcPay_PJ_TestMode = "0" Then
			dc_test = "no"
			Else 
			dc_test = "yes"
			End if 
			dc_logon = pcPay_PJ_MerchantID
			dc_password = pcPay_PJ_MerchantPassword
			dc_transaction_type = pcPay_PJ_TransType
			dc_first_name = pcBillingFirstName
			dc_last_name = pcBillingLasttName
			dc_name = dc_first_name &" "& dc_last_name
			dc_transaction_amount = money(pcBillingTotal)
			dc_number = getUserInput(Request.form("CardNumber"),16)
			dc_expiration_month = Request.form("expMonth")
			dc_expiration_year = Request.form("expYear")
			dc_address = pcBillingAddress & " " & pcBillingAddrss2
			dc_city = pcBillingCity
			if pcBillingStateCode  <> "" Then 
			dc_state = pcBillingStateCode
			else
			dc_state = pcBillingProvince
			end if 
			dc_zipcode = pcBillingPostalCode
			dc_country = pcBillingCountryCode
			dc_cvv = getUserInput(request.form("CVV"),4)
			dc_note = session("gwOrderId") & " " &  time() 
			
			
				
							
				pExpiration=dc_expiration_month & "/01/" & dc_expiration_year		
				' validates expiration
			    if DateDiff("d", Month(Now)&"/01/"&Year(now),pExpiration)<=-1 then
				 	response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_6")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
			    end if
		        'validate card
			    if not IsCreditCard(dc_number, request.form("CardType")) then
					response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_5")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
			    end if 
			
			 ' URL to the PayJunction Server
			 'URL = "https://payjunction.com:12002/live/vendor/quick_link/transact?"
			URL ="https://payjunction.com/quick_link?"
			'Build and encode the GET string
			URL = URL & "dc_test=" & server.URLencode(dc_test)
			URL = URL & "&dc_logon=" & server.URLEncode(dc_logon)
			URL = URL & "&dc_password=" & server.URLEncode(dc_password)
			URL = URL & "&dc_name=" & server.URLEncode(dc_name)
			URL = URL & "&dc_first_name=" & server.URLEncode(dc_first_name)
			URL = URL & "&dc_last_name=" & server.URLEncode(dc_last_name)
			URL = URL & "&dc_transaction_type=" & server.URLEncode(dc_transaction_type)
			URL = URL &"&dc_transaction_amount=" & server.URLEncode(dc_transaction_amount)
			URL = URL & "&dc_number=" & server.URLEncode(dc_number)
			URL = URL & "&dc_expiration_month=" & server.URLEncode(dc_expiration_month)
			URL = URL & "&dc_expiration_year=" & server.URLEncode(dc_expiration_year)
			URL = URL & "&dc_address=" & server.URLEncode(dc_address)
			URL = URL & "&dc_city=" & server.URLEncode(dc_city)
			URL = URL & "&dc_state=" & server.URLEncode(dc_state)
			URL = URL & "&dc_zipcode=" & server.URLEncode(dc_zipcode)
			URL = URL & "&dc_country=" & server.URLEncode(dc_country)
			URL = URL & "&dc_version=" & server.URLEncode("1.2")
			URL = URL & "&dc_note=" & server.URLEncode(dc_note)
			if pcPay_PJ_CVC = "1" then 
			 URL = URL & "&dc_verification_number=" & server.URLEncode(dc_cvv)
			End if 
		    			
			' New WinHTTP v5.0 - ships with MSXML 4.0 RTM
			' (5.1 should ship installed on XP Server)
			'
			' Download from (this should be on 1 line):
			
			'
			 Dim objWinHttp
			 on error resume next
			 err.number = 0
			 Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5")
			 if err.number <> 0 then			
			   Set objWinHttp = nothing 
			   Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
			 End if 
			
			' Full Docs (this should be on 1 line):
			' http://msdn.microsoft.com/library/default.asp?url=/workshop
			' /networking/winhttp/winhttp_node_entry.asp
			
			' If you have trouble or are getting connection errors,
			' try using the proxycfg.exe tool. Download from
			' (this should be on 1 line):
			' http://msdn.microsoft.com/downloads/default.asp?url=/downloads
			' /sample.asp?url=/msdn-files/027/001/766/msdncompositedoc.xml
			
			' Here we get the request ready to be sent.
			' First 2 parameters indicate method and URL.
			' The third is optional and indicates whether or not to
			' open the request in asyncronous mode (wait for a response
			' or not).  The default is False = syncronous = wait.
			' Syntax:
			'   .Open(bstrMethod, bstrUrl [, varAsync])
			objWinHttp.Open "GET", URL
			' Send it on it's merry way.
			objWinHttp.Send
			
			' Print out the request status:
	
			strStatus = objWinHttp.Status
			'store the response
						
			' Get the text of the response.
			if strStatus = 200 then							
				strRetVal = objWinHttp.ResponseText				
				Set objWinHttp = Nothing
				'store the response
				Response.write "<PRE>" & strRetVal &"</Pre>"
				
				If strRetVal <> ""  Then											
					' Split the response on the FS (0x1C)
					Dim responseArray
									
					reponseArray = SPLIT(strRetVal, Chr(28))				
					
					approvalCode = reponseArray(9)					
					responseCode = reponseArray(10) 
					responseMessage = reponseArray(11) 
					transactionID = reponseArray(12) 
					
					
					pcResultApproval = replace(approvalCode,"dc_approval_code=", "")					
					pcResultResponseMess = replace(responseMessage, "dc_response_message=","")
					pcResultResponseCode = replace(responseCode,"dc_response_code=","")
					pcResultTransRefNumber = replace(transactionID,"dc_transaction_id=","")
									
				Else
					'//ERROR
					pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"	
					response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& pcResultErrorMsg &"<br><br><a href="""&tempURL&"?psslurl=gwPayJunction.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
					response.end
				End If
				If pcResultResponseCode = "00" or pcResultResponseCode = "85" then
					session("GWAuthCode")=pcResultApproval
					session("GWTransId")=pcResultTransRefNumber
					response.redirect "gwReturn.asp?s=true&gw=PayJunction"
				Else				
					pcResultErrorMsg = pcResultResponseMess
					if pcResultErrorMsg="" then
					  pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"              	
					 end if
					Msg=pcResultErrorMsg
					response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwPayJunction.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
					Response.end	
				End if
			Else
		 
			if pcResultErrorMsg="" then
				pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"					
			end if
			Msg=pcResultErrorMsg
			response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwPayJunction.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			Response.end	
		 
			End if 
			Dim pcv_SuccessURL
			If scSSL="" OR scSSL="0" Then
				pcv_SuccessURL=replace((scStoreURL&"/"&scPcFolder&"/pc/pcPay_TD_Receipt.asp"),"//","/")
				pcv_SuccessURL=replace(pcv_SuccessURL,"https:/","https://")
				pcv_SuccessURL=replace(pcv_SuccessURL,"http:/","http://") 
			Else
				pcv_SuccessURL=replace((scSslURL&"/"&scPcFolder&"/pc/pcPay_TD_Receipt.asp"),"//","/")
				pcv_SuccessURL=replace(pcv_SuccessURL,"https:/","https://")
				pcv_SuccessURL=replace(pcv_SuccessURL,"http:/","http://")
			End If

	

	'*************************************************************************************
	' END
	'*************************************************************************************
	
end if 
'======================================================================================
'// End post back 
'======================================================================================


'======================================================================================
'// Show customer the payment form 
'======================================================================================
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
					<% 
					'======================================================================================
					'// If your gateway supports a Testing environment, create a variable for it and then
					'// if the cart is in testmode, alert the customer that this is not a live transaction.
					'// NOTE :: If no testing environment exists, delete the table row below
					'======================================================================================
					if pcPay_PJ_Testmode=1 then %>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
					<% end if
					'======================================================================================
					'// End Testing environment variable
					'// NOTE :: If no testing environment exists, delete the table row above
					'======================================================================================
					%>
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
								<% dim ArryCardTypes, strCardType, j
                                ArryCardTypes=split(pcPay_PJ_CardTypes,", ")
                                for j=0 to ubound(ArryCardTypes) 
									strCardType=ArryCardTypes(j) 
									select case strCardType
										case "VISA"
											response.write "<option value='VISA'>VISA</option>"
										case "MAST"
											response.write "<option value='MAST'>Master Card</option>"
										case "AMER"
											response.write "<option value='AMER'>American Express</option>"
										case "DISC"
											response.write "<option value='DISC'>Discover Card</option>"
										case "DINE"
											response.write "<option value='DINE'>Diners Club</option>"
										case "JCB"
											response.write "<option value='JCB'>JCB</option>"
									end select
								next %>
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
					<% 
					'======================================================================================
					'// If your gateway supports Credit Card Security Code (such as CVV and CVV2), create
					'// a variable for it and then show the row below.
					'// NOTE :: If no Security Code support exists, delete the table row below
					'======================================================================================
					If pcPay_PJ_CVC="1" Then %>
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
					<% end if
					'======================================================================================
					'// End Security Code support
					'// NOTE :: If no Security Code support exists, delete the table row above
					'======================================================================================
				 	%>
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
<% 
'======================================================================================
'// End Show customer the payment form 
'======================================================================================

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