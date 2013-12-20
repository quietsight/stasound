<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<!--#include file="gwBluePayMD5.asp"-->
<% 
session("redirectPage")="gwBluePay.asp"
		
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
<%
session("idCustomer")=pcIdCustomer

dim connTemp, rs
call openDB()

'//Get the Admin Settings / BluePay data
query="SELECT BPMerchant,BPTestmode,BPTransType,BPSECRET_KEY FROM BluePay WHERE idBluePay=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set Admin Settings / BluePay data
pcBPMerchant=rs("BPMerchant")
pcBPTestmode=rs("BPTestmode")
if pcBPTestmode="0" then
	pcBPTestmode="LIVE"
else
	pcBPTestmode="TEST"
end if
pcBPTransType=rs("BPTransType")
pcBPSecretKey=rs("BPSECRET_KEY")
set rs=nothing
call closedb()

'*************************************************************************************
' Post_Back
' START
'*************************************************************************************
if request("PaymentSubmitted")="Go" then
	
	'// Handle the Requests
	pcStrCardNumber=request.Form("CardNumber")
	pcStrExpMonth=request.Form("expMonth")
	pcStrExpYear=request.Form("expYear")
	CC_Expires=pcStrExpMonth&"/"&pcStrExpYear
	CVCCVV2=request.Form("CVCCVV2")
	
	pcStrName=pcBillingFirstName&" "&pcBillingLastName
	
  BP_MISSING_URL = "http://myserver/bogusinfo/missinginfo.asp"
  BP_APPROVED_URL = "http://myserver/bogusinfo/approved.asp"
  BP_DECLINED_URL = "http://myserver/bogusinfo/declined.asp"
  BP_SUBMIT_URL = "https://secure.bluepay.com/bp10emu"

	'// Process The Transaction
	'
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' MD5 Tamper Proof Seal
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  md5input = pcBPSecretKey & pcBPMerchant & pcBPTransType & pcBillingTotal & BP_REBILLING & _
             BP_REB_FIRST & BP_REB_EXPR & BP_REB_CYCLES & BP_REB_AMOUNT & _
             BP_RRNO & BP_AVS_ALLOWED & BP_AUTOCAP & pcBPTestmode
			 
  sDigest = md5(md5input)
  
	BP_TAMPER_PROOF_SEAL = sDigest
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Dim bpHTTPObj
  Dim bpPostData
  Dim sDigest

  Set bpHTTPObj = server.CreateObject("WinHttp.WinHttpRequest.5.1")

  ' URI Escape is Server.URLEncode(var)
  '
  bpPostData = "MERCHANT="   & Server.URLEncode(pcBPMerchant) & _
		"&MISSING_URL="       & Server.URLEncode(BP_MISSING_URL) & _
		"&APPROVED_URL="      & Server.URLEncode(BP_APPROVED_URL) & _
		"&DECLINED_URL="      & Server.URLEncode(BP_DECLINED_URL) & _
		"&MODE="              & Server.URLEncode(pcBPTestmode) & _
		"&TAMPER_PROOF_SEAL=" & Server.URLEncode(BP_TAMPER_PROOF_SEAL) & _
		"&TRANSACTION_TYPE="  & Server.URLEncode(pcBPTransType) & _
		"&CC_NUM="            & Server.URLEncode(pcStrCardNumber) & _
		"&CVCCVV2="           & Server.URLEncode(CVCCVV2) & _
		"&CC_EXPIRES="        & Server.URLEncode(CC_Expires) & _
		"&AMOUNT="            & Server.URLEncode(pcBillingTotal) & _
		"&Order_ID="          & Server.URLEncode(session("GWOrderID")) & _
		"&NAME="              & Server.URLEncode(pcStrName) & _
		"&Addr1="             & Server.URLEncode(pcBillingAddress) & _
		"&Addr2="             & Server.URLEncode(pcBillingAddress2) & _
		"&CITY="              & Server.URLEncode(pcBillingCity) & _
		"&STATE="             & Server.URLEncode(pcBillingState) & _
		"&ZIPCODE="           & Server.URLEncode(pcBillingPostalCode) & _
		"&COMMENT="           & Server.URLEncode(BP_COMMENT) & _
		"&PHONE="             & Server.URLEncode(pcBillingPhone) & _
		"&EMAIL="             & Server.URLEncode(pcCustomerEmail) & _
		"&REBILLING="         & Server.URLEncode(BP_REBILLING) & _
		"&REB_FIRST_DATE="    & Server.URLEncode(BP_REB_FIRST) & _
		"&REB_EXPR="          & Server.URLEncode(BP_REB_EXPR) & _
		"&REB_CYCLES="        & Server.URLEncode(BP_REB_CYCLES) & _
		"&REB_AMOUNT="        & Server.URLEncode(BP_REB_AMOUNT) & _
		"&RRNO="              & Server.URLEncode(BP_RRNO) & _
		"&AUTOCAP="           & Server.URLEncode(BP_AUTOCAP) & _
		"&AVS_ALLOWED="       & Server.URLEncode(BP_AVS_ALLOWED)

	' here we perform a POST; the string we've just created goes in the BODY of the POST:
	BP_SUBMIT_URL="https://secure.bluepay.com/bp10emu"
  bpHTTPObj.Open "POST" ,BP_SUBMIT_URL, False
  bpHTTPObj.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  bpHTTPObj.Send(bpPostData)

	' We return Bluepay's Response; if your program doesn't want to parse that, then
	' you may use the convenience functions which follow:
  
  '1.) header
	BP_RESPONSE = bpHTTPObj.GetResponseHeader("location")
  ' response.write(bpHTTPObj.GetAllResponseHeaders()) & "<br>" 
  
  '2.) full response
  str_response = bpHTTPObj.GetAllResponseHeaders()

	Set bpHTTPObj = Nothing

	' response.write BP_RESPONSE & "<br>"

	'// Handle Errors "ERROR" or "APPROVED" or "DECLINED" or "MISSING"
	' GIVE MESSAGE AND SHOW FORM AGAIN
	
	' Get the result 
	Result = trimstring(bp_get_status)
	BPDeclinedString = ""

	If Result = "MISSING" Then
		BPDeclinedString = "All fields are required."

		strMissing = trimstring(bp_get_missing)
	
		strRMissing = ""
		if strMissing<>"" then
			select case ucase(strMissing)
				case "CardNumber"
					strRMissing="Credit Card Number is a required field."
				case "CC_EXPIRES"
					strRMissing="Credit Card expiration date is a required field."
				case "CVCCVV2"
					strRMissing="Security Code is a required field."
				case "TRANSACTION_TYPE"
					strRMissing="The type of transaction is not specified. Contact the store owner."			
			end select
			BPDeclinedString=BPDeclinedString&"&nbsp;"&strRMissing
		end if
	End If

	If Result="DECLINED" Then
		response.write "DECLINED"
		BPDeclinedString="Your tranaction was declined by the payment processor. Please check over your information to ensure that it is correct."
	End If

	If Result="ERROR" Then
		BPDeclinedString="Your tranaction was declined by the payment processor for the following reason: " & trimstring(bp_get_error)
	End If
	'// Handle the transaction approved
	' redirect to gwReturn.asp with proper values
	If Result="APPROVED" Then
		session("GWAuthCode")= trimstring(bp_get_RRNO) ' *** There is no approval code, using the trans id instead. ***
		'session("RRNO")= trimstring(bp_get_RRNO)
		'session("AVSResult")= trimstring(bp_get_AVS)
		'session("CVV2Result")= trimstring(bp_get_CVV2)
		
		BPDescription = trimstring(bp_get_approval)
		IF BPDescription = "DUPLICATE" THEN
			BPDeclinedString="You have submitted a duplicate transaction. Your order can not processed at this time."
		ELSE
			call closedb()
			session("GWTransType")=pcBPTransType
			response.redirect "gwReturn.asp?s=true&gw=BluePay"
		END IF          
	End If

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' end V2
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
end if

'*************************************************************************************
' Post_Back
' END
'*************************************************************************************
' Returns the status: "APPROVED", "DECLINED", "MISSING", "ERROR"
Function bp_get_status()
  Set ExpReg = new RegExp
  ExpReg.pattern = "Result=(\w+)"
 Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_status = ExpMatched.Value
	Next
  Else
  	bp_get_status = Null
  End If
  Set ExpReg = Nothing
End Function

'"MISSING"
Function bp_get_missing()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "Missing=(\w+)"
 Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_missing = ExpMatched.Value
	Next
  Else
  	bp_get_missing = Null
  End If
  Set ExpReg = Nothing
End Function

'### NS Message ###
' Returns the message - describes the transaction.
Function bp_get_message()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern="MESSAGE=(.*?)[\&$]"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_message = ExpMatched.Value
	Next
  Else
  	bp_get_message = Null
  End If
  Set ExpReg = Nothing
End Function

'### Error Message ###
Function bp_get_error()
  Set ExpReg = new RegExp
  ExpReg.pattern = "MESSAGE=(\S+)"
 Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_error = ExpMatched.Value
	Next
  Else
  	bp_get_error = Null
  End If
  Set ExpReg = Nothing
End Function


'### Approved Message ###
Function bp_get_approval()
  Set ExpReg = new RegExp
  ExpReg.pattern = "MESSAGE=(\S+)"
 Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_approval = ExpMatched.Value
	Next
  Else
  	bp_get_approval = Null
  End If
  Set ExpReg = Nothing
End Function

'### Returns the RRNO, if any. ###
Function bp_get_RRNO()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "RRNO=(\d+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_RRNO = ExpMatched.Value
	Next
  Else
  	bp_get_RRNO = Null
  End If
  Set ExpReg = Nothing
End Function


'### Returns the AVS Code. ###
Function bp_get_AVS()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "AVS=(\w+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_AVS = ExpMatched.Value
	Next
  Else
  	bp_get_AVS = Null
  End If
  Set ExpReg = Nothing
End Function

'### Returns the CVV2 Code. ###
Function bp_get_CVV2()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "AVS=(\w+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_CVV2 = ExpMatched.Value
	Next
  Else
  	bp_get_CVV2 = Null
  End If
  Set ExpReg = Nothing
End Function


'### Returns the ApprovalCode Code. ###
Function bp_get_ApprovalCode()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "ApprovalCode=(\w+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_ApprovalCode = ExpMatched.Value
	Next
  Else
  	bp_get_ApprovalCode = Null
  End If
  Set ExpReg = Nothing
End Function

' trim up string to get just the value
function trimstring(strQ)
	nIndex = InStrRev(strQ,"=")
	If (nIndex>0) Then
		strQ = Right(strQ,Len(strQ)-nIndex) 	
	End If
	strQ = replace(strQ,"%20"," ")
	strQ = replace(strQ,"%3B",";")
	trimstring = strQ
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
				<form method="POST" action="gwBluePay.asp" name="BPForm" class="pcForms">
					<input type="hidden" name="PaymentSubmitted" value="Go">
					<table class="pcShowContent">
			
					<% if BPDeclinedString<>"" then %>
						<tr valign="top"> 
							<td colspan="2">
								<div class="pcErrorMessage"><%=BPDeclinedString%></div>
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
					<% if pcBPTestmode="TEST" then %>
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
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
						<td> 
							<input name="CVCCVV2" type="text" id="CVCCVV2" value="" size="4" maxlength="4">
						</td>
					</tr>
					<tr> 
						<td>&nbsp;</td>
						<td><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
					</tr>
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