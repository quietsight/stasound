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
session("redirectPage")="gwTripleDeal.asp"

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
query="SELECT pcPay_TD_MerchantName, pcPay_TD_MerchantPassword, pcPay_TD_Profile, pcPay_TD_ClientLang, pcPay_TD_PayPeriod, pcPay_TD_TestMode FROM pcPay_TripleDeal WHERE (((pcPay_TD_ID)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_TD_MerchantName=rs("pcPay_TD_MerchantName")
pcPay_TD_MerchantName=enDeCrypt(pcPay_TD_MerchantName, scCrypPass)
pcPay_TD_MerchantPassword=rs("pcPay_TD_MerchantPassword")
pcPay_TD_MerchantPassword=enDeCrypt(pcPay_TD_MerchantPassword, scCrypPass)
pcPay_TD_Profile=rs("pcPay_TD_Profile")
pcPay_TD_ClientLang=rs("pcPay_TD_ClientLang")
pcPay_TD_PayPeriod=rs("pcPay_TD_PayPeriod")
pcPay_TD_TestMode=rs("pcPay_TD_TestMode")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim objXMLHTTP, xml
	
	'Send the request to the Authorize.NET processor.
	stext="command=new_payment_cluster"
	stext=stext &"&merchant_name="&pcPay_TD_MerchantName
	stext=stext &"&merchant_password="&pcPay_TD_MerchantPassword
	stext=stext &"&merchant_transaction_id="&session("GWOrderID")
	stext=stext &"&profile="&pcPay_TD_Profile
	stext=stext &"&client_id="&pcIdCustomer
	stext=stext &"&price="&pcBillingTotal
	stext=stext &"&cur_price=EUR"
	stext=stext &"&client_email="&pcCustomerEmail
	stext=stext &"&client_firstname="&pcBillingFirstName
	stext=stext &"&client_lastname="&pcBillingLastName
	stext=stext &"&client_address="&pcBillingAddress
	stext=stext &"&client_zip="&pcBillingPostalCode
	stext=stext &"&client_city="&pcBillingCity
	stext=stext &"&client_country="&pcBillingCountryCode
	stext=stext &"&client_language="&pcPay_TD_ClientLang
	stext=stext &"&description=Online Order - "&session("GWOrderID")
	stext=stext &"&days_pay_period="&pcPay_TD_PayPeriod
	stext=stext &"&include_costs=yes"

	'Send the transaction info as part of the querystring
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	if pcPay_TD_TestMode=1 then
		xml.open "POST", "https://test.tripledeal.com/ps/com.tripledeal.paymentservice.servlets.PaymentService?"& stext & "", false
	else
		xml.open "POST", "https://www.tripledeal.com/ps/com.tripledeal.paymentservice.servlets.PaymentService?"& stext & "", false
	end if
	
	xml.send ""
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText

	Set TDXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
	TDXMLdoc.async = false 
	if TDXMLdoc.loadXML(strRetVal) then ' if loading from a string
	
		set objLst=TDXMLdoc.getElementsByTagName("new_payment_cluster")
		for i = 0 to (objLst.length - 1)
			varFlag=0
			for j=0 to ((objLst.item(i).childNodes.length)-1)
				If objLst.item(i).childNodes(j).nodeName="errorlist" then
					for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
						if objLst.item(i).childNodes(j).childNodes(k).nodeName="error" then
							response.write objLst.item(i).childNodes(j).childNodes(k).Attributes.GetNamedItem("msg").Text
						end if
					next
				Else
					If objLst.item(i).childNodes(j).nodeName="key" then
						strKey = objLst.item(i).childNodes(j).Attributes.GetNamedItem("value").Text
						session("keyValue")=strKey
					End If
				End if
			next
		next
	end if
	Set xml = Nothing
	set TDXMLdoc = Nothing

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

	'Check the ErrorCode to make sure that the component was able to talk to the authorization network
	If (strStatus <> 200) Then
		Msg = "An error occurred during processing. Please try again later."
	Else
		If session("keyValue") <> "" Then

			'send key to gateway
			stext="command=show_payment_cluster"
			stext=stext &"&merchant_name="&pcPay_TD_MerchantName
			stext=stext &"&client_language="&pcPay_TD_ClientLang
			stext=stext &"&default_pm=banktransfer-nl"
			stext=stext &"&payment_cluster_key="&session("keyValue")
			stext=stext &"&return_url_success="&pcv_SuccessURL
			stext=stext &"&return_url_pending="&pcv_SuccessURL
			stext=stext &"&return_url_error="&pcv_SuccessURL

			'Send the transaction info as part of the querystring
			response.redirect "https://test.tripledeal.com/ps/com.tripledeal.paymentservice.servlets.PaymentService?"& stext
			response.end
		End If
	End If
	
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
					<% if pcPay_TD_TestMode=1 then %>
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
					<% If x_CVV="1" Then %>
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