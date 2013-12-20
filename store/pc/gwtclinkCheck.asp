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
session("redirectPage")="gwtclinkCheck.asp"

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
query="SELECT TCLinkid, TCLinkPassword, TCTestmode, TCCurcode, TranType, TCLinkCheckPending FROM tclink Where idTCLink=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
TCLinkid=rs("TCLinkid")
TCLinkPassword=rs("TCLinkPassword")
'decrypt
TCLinkPassword=enDeCrypt(TCLinkPassword, scCrypPass)
TCLinkCheckPending=rs("TCLinkCheckPending")
TCCurcode=rs("TCCurcode")
TCTestmode=rs("TCTestmode")
action="sale"
DLLUrl="https://vault.trustcommerce.com/trans"
If TCTestmode =1 Then
		demo="y"
Else ' Live Mode
			demo=""   
End If   

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim objHTTP, strRequest 
	
	strRequest = "?custid="&TCLinkid&_
	"&password="&TCLinkPassword&_
	"&action="&action&_
	"&media="&"ach"&_
	"&demo="&demo&_
	"&amount="&pcBillingTotal*100&_   
	"&routing="&request.Form("bank_aba_code")&_
	"&account="&request.Form("bank_acct_num")&_
	"&name="&pcBillingFirstName& " "&pcBillingLastName&_
	"&address1="&pcBillingAddress&_
	"&city="&pcBillingCity&_
	"&state="&pcBillingState&_
	"&zip="&pcBillingPostalCode&_
	"&country="&pcBillingCountryCode&_
	"&phone="&pcBillingPhone&_
	"&email="&pcCustomerEmail&_
	"&shipto_name="&pcShippingFirstName&" "&pcShippingLastName&_
	"&shipto_address1="&pcShippingAddress&_
	"&shipto_city="&pcShippingCity&_
	"&shipto_state="&pcShippingState&_
	"&shipto_zip"&pcShippingPostalCode&_
	"&shipto_country="&pcShippingCountryCode&_
	"&currency="&TCCurcode&_
	"&ClientIPAddress="&pcCustIpAddress
			
	' Create the WinHTTPRequest Object
	Set objHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
	objHTTP.Open "POST", DLLUrl & strRequest, false
	objHTTP.Send()    ' Send the HTTP request.
			
			
	If objHTTP.Status = 200 Then  ' HTTP_STATUS_OK=200 
	
		Dim objDictResponse, intDelimiterPos, ResponseArray
		Dim strResponse, strNameValuePair, strName, strValue
		strResponse = replace(objHTTP.ResponseText,chr(10)," ")
		strResponse = rtrim(strResponse)

		 ' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS

		 ResponseArray = Split(strResponse, " ")
		 Set objDictResponse = server.createobject("Scripting.Dictionary")
		 For each ResponseItem in ResponseArray
			 NameValue = Split(ResponseItem, "=")
			 objDictResponse.Add NameValue(0), NameValue(1)
		 Next
			
			' Parse the response into local vars
			strstatus        = objDictResponse.Item("status")
			strerror          = objDictResponse.Item("error")
			stroffenders       = objDictResponse.Item("offenders")
			strTransactionID       = objDictResponse.Item("transid")
			strAVSResponseCode     = objDictResponse.Item("avs")
				
			If strstatus = "accepted" Then 
				'tordnum=(int(strTransactionID)-scpre)
				'session("AuthorizationNumber")=strAuthorizationNumber
				session("GWTransId")=strTransactionID
				session("TranType")=action
				session("GWTransType")=TCLinkCheckPending
				session("GWAuthCode")=""
				
				Response.redirect "gwReturn.asp?s=true&gw=TCLinkCheck"
			Else
				response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>&lsaquo;Error:&nbsp;&nbsp;"&strstatus&"&nbsp;in&nbsp;"&lcase(stroffenders)&" &rsaquo;&nbsp;&lsaquo;Error Type:&nbsp;"&lcase(strerror)&"&rsaquo;<br><br><a href="""&tempURL&"?psslurl=gwtclinkCheck.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("idOrder")&"""><img src="""&rslayout("back")&""" border=0></a>")
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
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp">Edit</a></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% if TCTestMode=1 then %>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
                        </tr>
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
							<td colspan="2" align="center"><img src="images/sampleck.gif" width="390" height="230"></td>
						</tr>
					<tr> 
						<td><p>Name on the Account:</p></td>
						<td> 
							<input name="bank_acct_name" type="text" size="35" maxlength="50">
						</td>
					</tr>
					<tr> 
						<td><p>Bank Routing Number:</p></td>
						<td> 
							<input name="bank_aba_code" type="text" size="35">
						</td>
					</tr>
					<tr> 
						<td><p>Bank Account Number:</p></td>
						<td> 
							<input name="bank_acct_num" type="text" size="35">
						</td>
					</tr>
					<tr> 
						<td><p>Check Number:</p></td>
						<td> 
							<input name="check_num" type="text" size="15">
						</td>
					</tr>
					
					<tr> 
						<td><p>Bank Account Type:</p></td>
						<td>
							<select name="bank_acct_type">
								<option value="CHECKING">Checking Account</option>
								<option value="SAVINGS">Savings Account</option>
							</select>
						</td>
					</tr>
					<tr> 
						<td><p>Bank Name:</p></td>
						<td> 
								<input name="bank_name" type="text" size="20" maxlength="20">
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					
					<tr> 
						<td colspan="2" align="center">
						<% 
						If scSSL="" OR scSSL="0" Then
							tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp"),"//","/")
							tempURL=replace(tempURL,"https:/","https://")
							tempURL=replace(tempURL,"http:/","http://") 
						Else
							tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp"),"//","/")
							tempURL=replace(tempURL,"https:/","https://")
							tempURL=replace(tempURL,"http:/","http://")
						End If
						%> 
						<a href="<%=tempURL%>"><img src="<%=rslayout("back")%>"></a>&nbsp;&nbsp; 
						<input type="image" name="Continue" src="<%=rslayout("pcLO_placeOrder")%>" id="submit">
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->