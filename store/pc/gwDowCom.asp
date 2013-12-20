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
<% 'Gateway specific filesdim connTemp, rs
session("redirectPage")="gwDowCom.asp" 

Dim pcCustIpAddress
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


': Open Connection to the DB
dim connTemp, rs 'DELETE FOR HARD CODED VARS
call openDb() 'DELETE FOR HARD CODED VARS
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'//Functions parse response string
'======================================================================================
Function DecodeQueryValue(qValue)
	'Purpose: To URL decode a string
	'Pre: qValue is set to a url encoded value of a query string parameter. ex: "one+two"
	'Post: none
	'Returns: Returns the url decoded value of qValue. ex: "one two"
	Dim i
	Dim qChar
	dim newString
	if IsNull(qValue) = false then
		For i = 1 To Len(qValue)
		qChar = Mid(qValue, i, 1)
			If qChar = "%" Then
			on error resume next
			newString = newString & Chr("&H" & Mid(qValue, i + 1, 2))
			on error goto 0
			i = i + 2
			ElseIf qChar = "+" Then
			newString = newString & " "
			Else
			newString = newString & qChar
			End If
		Next
		
		DecodeQueryValue = Replace(newString, "&lt;", "<")
	else
		DecodeQueryValue = ""
	end if
End Function

Function GetQueryValue(queryString, paramName)
	'Purpose: To return the value of a parameter in an HTTP query string.
	'Pre: queryString is set to the full query string of url encoded name value pairs. ex:
	'"value1=one&value2=two&value3=3"
	' paramName is set to the name of one of the parameters in the queryString. ex: "value2"
	'Post: None
	'Returns: The function returns the query string value assigned to the paramName parameter. ex: "two"
	Dim pos1
	dim pos2
	Dim qString
	qString = "&" & queryString & "&"
	pos1 = InStr(1, qString, paramName & "=")
	If pos1 > 0 Then
		pos1 = pos1 + Len(paramName) + 1
		pos2 = InStr(pos1, qString, "&")
		If pos2 > 0 Then
		GetQueryValue = DecodeQueryValue(Mid(qString, pos1, pos2 - pos1))
		End If
	End If
End Function

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your 
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
query= "SELECT pcPay_Dow_TransType,pcPay_Dow_MerchantID,pcPay_Dow_MerchantPassword,pcPay_Dow_CardTypes,pcPay_Dow_CVC,pcPay_Dow_TestMode FROM pcPay_DowCom where pcPay_Dow_ID=1"
set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
	
pcPay_Dow_TransType=rs("pcPay_Dow_TransType")
pcPay_Dow_MerchantID=rs("pcPay_Dow_MerchantID")
'decrypt
pcPay_Dow_MerchantID=enDeCrypt(pcPay_Dow_MerchantID, scCrypPass)
pcPay_Dow_MerchantPassword=rs("pcPay_Dow_MerchantPassword")
'decrypt
pcPay_Dow_MerchantPassword=enDeCrypt(pcPay_Dow_MerchantPassword, scCrypPass)
pcPay_Dow_CardTypes=rs("pcPay_Dow_CardTypes")
pcPay_Dow_CVC=rs("pcPay_Dow_CVC")
pcPay_Dow_TestMode=rs("pcPay_Dow_TestMode")

M="0"
V="0"
A="0"
D="0"
set rs=nothing

if request("PaymentSubmitted")="Go" then 
	Dim objXMLHTTP, xml
	intCVV=request.form("CVV")
	strCardNumber=request.form("CardNumber")
	strExpMonth=request.form("expMonth")
	strExpYear=request.form("expYear")
	strCardType=request.form("x_Card_Type")
	
	if not IsCreditCard(strCardNumber, strCardType) then
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_5")&"<br><br><a href="""&tempURL&"?psslurl=gwDowCom.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
	end if 
	
	' validates expiration
	if DateDiff("d", Month(Now)&"/"&Year(now), strExpMonth&"/20"&strExpYear)<=-1 then
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_6")&"<br><br><a href="""&tempURL&"?psslurl=gwDowCom.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
	end if

	if pcPay_Dow_CVC = "1" and (not isNumeric(intCVV) or  len(intCVV) < 3 ) Then
		Msg = "Please Supply a Security Code."
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwDowCom.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		Response.end				
	End if 
				 
	DataToSend = "type=" & Server.URLEncode(pcPay_Dow_TransType) &_
	"&username=" & Server.URLEncode(pcPay_Dow_MerchantID) &_
	"&password=" & Server.URLEncode(pcPay_Dow_MerchantPassword) &_
	"&ccnumber=" & Server.URLEncode(strCardNumber) &_
	"&ccexp=" & Server.URLEncode(strExpMonth & strExpYear) &_
	"&amount=" & Server.URLEncode(pcBillingTotal) &_
	"&firstname=" & Server.URLEncode(pcBillingFirstName) &_
	"&lastname=" & Server.URLEncode(pcBillingLastName) &_
	"&phone=" & Server.URLEncode(pcBillingPhone)& _
	"&address1=" & Server.URLEncode(pcBillingAddress) &_
	"&city=" & Server.URLEncode(pcBillingCity) &_
	"&state=" & Server.URLEncode(pcBillingState) &_ 
	"&zip=" & Server.URLEncode(pcBillingPostalCode)& _
	"&country=" & Server.URLEncode(pcBillingCountryCode)& _ 
	"&ipaddress=" & Server.URLEncode(pcCustIpAddress)& _
	"&orderid=" & Server.URLEncode(session("GWOrderId"))& _ 
	 "&email=" & Server.URLEncode(pcCustomerEmail)
	 
	 if pcPay_Dow_CVC = "1" then 				 
		DataToSend = DataToSend & 	"&cvv=" & Server.URLEncode(intCVV)
	 End if 
	'Response.write DataToSend &"<BR><BR><BR>"
	'Send the transaction info as part of the querystring
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.open "POST", "https://secure.dowcommerce.net/api/transact.php", false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
	xml.send(DataToSend)			
	if err.number<>0 then
		pcResultErrorMsg = err.description
	end if
	strStatus = xml.Status			
	
	if strStatus = 200 then 	
		'store the response
		strRetVal = xml.responseText	
		'Response.write strRetVal
		'Response.end 			
		
		if strRetVal <> "" Then		
			pcResultResponseCode = GetQueryValue(strRetVal,"response")
			pcResultTransRefNumber = GetQueryValue(strRetVal,"transactionid") 
			pcResultErrorMsg = GetQueryValue(strRetVal,"responsetext")  
			pcResultAuthCode =  GetQueryValue(strRetVal,"authcode")
		Else
			'//ERROR
			pcResultErrorMsg = "Transaction error or declined.  Error Message: " & pcResultErrorMsg
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwDowCom.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			Response.end	
		End If
		If pcResultResponseCode = "1"  then
			session("GWAuthCode")=pcResultAuthCode
			session("GWTransId")=pcResultTransRefNumber
			response.redirect "gwReturn.asp?s=true&gw=DowCom"
		Else				
			if pcResultErrorMsg="" then
			  pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"              	
			 end if
			Msg=pcResultErrorMsg
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwDowCom.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			Response.end	
		End if
	Else		  
		if pcResultErrorMsg="" then
			pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"					
		end if
		Msg=pcResultErrorMsg
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwDowCom.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		Response.end			 
	End if 
End if 
			
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
        <form action="gwDowCom.asp" method="POST" name="form1" class="pcForms">
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
			<% if pcPay_Dow_TestMode="1" then %>
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
                    	<% 	pcPay_Dow_TransTypeArray=Split(pcPay_Dow_CardTypes,", ")
                    
                        i=ubound(pcPay_Dow_TransTypeArray)
                        cardCnt=0
                        do until cardCnt=i+1
                            cardVar=pcPay_Dow_TransTypeArray(cardCnt)
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
						%>
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
            <% If pcPay_Dow_CVC="1" Then %>
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
                <td colspan="2" class="pcSpacer"></td>
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