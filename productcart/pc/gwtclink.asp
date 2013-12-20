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
session("redirectPage")="gwTCLink.asp"

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
query="SELECT TCLinkid, TCLinkPassword, TCTestmode, TCCurcode, CVV, avs, TranType FROM tclink Where idTCLink=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_TCLinkid=rs("TCLinkid")
pcv_TCLinkPassword=rs("TCLinkPassword")
'decrypt
pcv_TCLinkPassword=enDeCrypt(pcv_TCLinkPassword, scCrypPass)
pcv_TestMode=rs("TCTestmode")
pcv_CurCode=rs("TCCurcode")
pcv_CVV=rs("CVV")
pcv_AVS=rs("avs")
pcv_TransType=rs("TranType")
				
If pcv_TestMode =1 Then
	pcv_Demo="y"
Else ' Live Mode
	pcv_Demo=""   
End If   
			
If pcv_AVS ="1" Then
	pcv_AVS="y"
Else 
	pcv_AVS="n"   
End If
			
If pcv_CVV ="1" Then
	pcv_CheckCVV ="y"
Else 
	pcv_CheckCVV ="n"   
End If

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

'*************************************************************************************
' This is where you would post info to the gateway
' START
'*************************************************************************************
		strPostUrl="https://vault.trustcommerce.com/trans"
		
		strRequest = "?custid="&pcv_TCLinkid&_
		"&password="&pcv_TCLinkPassword&_
		"&action="&pcv_TransType&_
		"&media="&"cc"&_
		"&demo="&pcv_Demo&_
		"&amount="&pcBillingTotal*100&_   
		"&cc="&request.Form("CardNumber")&_ 
		"&avs="&pcv_AVS&_  
		"&checkcvv="&pcv_CheckCVV&_  
		"&cvv="& request.Form("CVV")&_	
		"&exp="&request.Form("expMonth")&request.Form("expYear")&_
		"&currency="&pcv_CurCode&_
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
		"&shipto_country="&pcShippingCountryCode'&_
		
		' Create the WinHTTPRequest Object
		Set objHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objHTTP.Open "POST", strPostUrl & strRequest, false
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
		strAuthorizationNumber = objDictResponse.Item("authcode")
		strstatus        = objDictResponse.Item("status")
		strerror          = objDictResponse.Item("error")
		stroffenders       = objDictResponse.Item("offenders")
		strTransactionID       = objDictResponse.Item("transid")
		strAVSResponseCode     = objDictResponse.Item("avs")
						
		If lcase(strstatus) = "approved" Then
			session("GWAuthCode")=strAuthorizationNumber
			session("GWTransId")=strTransactionID
			session("GWTransType")=pcv_TransType
			Response.redirect "gwReturn.asp?s=true&gw=TCLink"
		Else
			response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error:&nbsp;&nbsp;"&strstatus&"&nbsp;in&nbsp;"&lcase(stroffenders)&" --Error Type:&nbsp;"&lcase(strerror)&"<br><br><a href="""&tempURL&"?psslurl=gwtclink.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
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
					<% 'If x_CVV="1" Then %>
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
					<% 'End If %>
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