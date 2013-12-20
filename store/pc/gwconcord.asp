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
session("redirectPage")="gwConcord.asp"

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
query="SELECT StoreID, StoreKey, testmode, Curcode, CVV, MethodName FROM concord Where idConcord=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_StoreID=rs("StoreID")
pcv_StoreKey=rs("StoreKey")
'decrypt
pcv_StoreKey=enDeCrypt(pcv_StoreKey, scCrypPass)
pcv_CVV=rs("CVV")
pcv_Curcode=rs("Curcode")
pcv_TestMode=rs("testmode")
pcv_MethodName=rs("MethodName")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Quote = String(1,34) 
	If pcv_TestMode =1 Then
		strDLLUrl   = "https://stg.dw.us.fdcnet.biz/efsnet.dll"      
	Else ' Live Mode
		strDLLUrl = "https://prod.dw.us.fdcnet.biz/efsnet.dll"
	End If   

	'--------------------------------------------------------------------------
	' For simplicity in this sample, we'll only populate variables for a 
	' SystemCheck.  For other methods, required fields such as 
	' "strTransactionAmount" would need to be populated before calling 
	' ProcessCGIRequest. Notice how easy it is to use any EFSnet method 
	' by simply changing the method name passed to ProcessCGIRequest.
	'--------------------------------------------------------------------------
	' Populate Transaction Request Variables
	If pcv_TestMode =1 Then
		strApplicationID = "ProductCart Test"
	else
		strApplicationID = "ProductCart"
	end if
	
	Select Case pcv_MethodName
		' Build a properly formatted EFSnet Cgi Request Message 
		Case "SysCheck"  ProcessCGIRequest("SystemCheck")  
		Case "Authorize" ProcessCGIRequest("CreditCardAuthorize")
		Case "Settle"    ProcessCGIRequest("CreditCardSettle")  
		Case "Charge"    ProcessCGIRequest("CreditCardCharge") 'Auth/Settle
		Case "Refund"    ProcessCGIRequest("CreditCardRefund")  
	End Select

	'--------------------------------------------------------------------------
	' HELPER FUNCTIONS
	'--------------------------------------------------------------------------
	
	'--------------------------------------------------------------------------
	Public Sub ProcessCGIRequest(strMethod)
		 Dim objHTTP, strRequest 
	
		 strRequest = "Method="       & strMethod										& _
			"&StoreID="               & pcv_StoreID              						& _
			"&StoreKey="              & pcv_StoreKey              						& _   
			"&ApplicationID="         & strApplicationID       	  						& _   
			"&AccountNumber="         & request.Form("CardNumber")						& _   
			"&ExpirationMonth="       & request.Form("expMonth")  						& _
			"&ExpirationYear="        & request.Form("expYear")							& _
			"&CardVerificationValue=" & strCardVerificationValue						& _
			"&Track1="                & strTrack1                						& _
			"&Track2="                & strTrack2                						& _
			"&TerminalID="            & strTerminalID            						& _
			"&CashierNumber="         & strCashierNumber         						& _ 
			"&ReferenceNumber="       & session("GWOrderId")	 						& _  
			"&TransactionAmount="     & pcBillingTotal				    				& _
			"&SalesTaxAmount="        & strSalesTaxAmount       						& _
			"&Currency="              & pcv_Curcode               						& _
			"&BillingName="           & pcBillingFirstName& " "&pcBillingLastName		& _
			"&BillingAddress="        & pcBillingAddress        		 				& _
			"&BillingCity="           & pcBillingCity           						& _
			"&BillingState="          & pcBillingStateCode          		 			& _
			"&BillingPostalCode="     & pcBillingPostalCode     				 		& _
			"&BillingCountry="        & pcBillingCountryCode        		 			& _
			"&BillingPhone="          & pcBillingPhone         		 					& _
			"&BillingEmail="          & pcCustomerEmail          		 				& _
			"&ShippingName="          & pcShippingFirstName&" "&pcShippingLastName		& _
			"&ShippingAddress="       & pcShippingAddress      							& _
			"&ShippingCity="          & pcShippingCity         							& _
			"&ShippingState="         & pcShippingStateCode          					& _
			"&ShippingPostalCode"     & pcShippingPostalCode			 				& _
			"&ShippingCountry="       & pcShippingCountryCode       					& _
			"&ShippingPhone="         & pcShippingPhone         						& _
			"&ShippingEmail="         & strShippingEmail         						& _
			"&ClientIPAddress="       & pcCustIpAddress

		 If VIEW_CGI_REQUEST Then
				Response.Write strRequest & "<BR>"   ' Debug only
		 End If

		 ' Create the WinHTTPRequest Object
		 Set objHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		 objHTTP.Open "POST", strDLLUrl, false
		 objHTTP.Send(strRequest)    ' Send the HTTP request.

		 If objHTTP.Status = 200 Then  ' HTTP_STATUS_OK=200 
			
		 Dim objDictResponse, intDelimiterPos, ResponseArray
		 Dim strResponse, strNameValuePair, strName, strValue
		 strResponse = objHTTP.ResponseText

		If VIEW_CGI_REQUEST Then 
			Response.Write strResponse & "<BR>"   ' Debug only
		End If 

		' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
		ResponseArray = Split(strResponse, "&") 
		Set objDictResponse = server.createobject("Scripting.Dictionary")
		For each ResponseItem in ResponseArray
			NameValue = Split(ResponseItem, "=")
			objDictResponse.Add NameValue(0), NameValue(1)
		Next
       
		' Parse the response into local vars
		strResponseCode        = objDictResponse.Item("ResponseCode")
		strResultCode          = objDictResponse.Item("ResultCode")
		strResultMessage       = objDictResponse.Item("ResultMessage")
		strTransactionID       = objDictResponse.Item("TransactionID")
		strAVSResponseCode     = objDictResponse.Item("AVSResponseCode")
		strCVVResponseCode     = objDictResponse.Item("CVVResponseCode")
		strApprovalNumber      = objDictResponse.Item("ApprovalNumber")
		strAuthorizationNumber = objDictResponse.Item("AuthorizationNumber")
		strTransactionDate     = objDictResponse.Item("TransactionDate")
		strTransactionTime     = objDictResponse.Item("TransactionTime")

		If strResponseCode = 0 Then 
			session("GWAuthCode")=strAuthorizationNumber
			session("GWTransId")=strTransactionID
			session("GWTransType")=pcv_MethodName
			Response.redirect "gwReturn.asp?s=true&gw=Concord"
		Else
			response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error:&nbsp;&nbsp;"&strResponseCode&"&nbsp;&nbsp;"&lcase(strResultMessage)&"<br><br><a href="""&tempURL&"?psslurl=gwconcord.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			response.end
		End If  
   
	Else   
		Response.write "Connection Failed..."
		Response.Write "<BR> Https Response <BR>" 
		Response.Write "Status     = " & objHttp.status       & "<BR>"
		Response.Write "StatusText = " & objHttp.statusText   & "<BR>"
		Response.Write "Header     = " & objHttp.getAllResponseHeaders & _
		"<BR>"
		Response.Write "RespText   = " & objHttp.responseText & "<BR>"
	End If   
	Set objHttp   = Nothing
End Sub
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
					<% if pcv_TestMode =1 then %>
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