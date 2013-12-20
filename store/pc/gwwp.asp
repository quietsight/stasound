<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'Gateway File: Worldpay
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
'// See if this is a response back from WorldPay
if request("status")="Y" then
	session("GWAuthCode")=getUserInput(request("rawAuthCode"),0)
	session("GWTransId")=getUserInput(request("transId"),0)
	if session("GWOrderId")="" then
		session("GWOrderId")=getUserInput(request("idOrder"),0)
	end if
	session("GWSessionID")=Session.SessionID 
	'// Payment is received - redirect to gwReturn.asp
	response.Redirect("gwReturn.asp?s=true&gw=WorldPay")
end if

'//Set redirect page to the current file name
session("redirectPage")="gwwp.asp"
session("redirectPage2")="https://select.worldpay.com/wcc/purchase"
'//secure-test.wp3.rbsworldpay.com

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
	session("GWOrderId")=getUserInput(request("idOrder"),0)
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
query="SELECT WP_instID, WP_Currency, WP_testmode FROM WorldPay WHERE wp_id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
WP_instID=rs("WP_instID")
WP_Currency=rs("WP_Currency")
WP_testmode=rs("WP_testmode")

set rs=nothing
call closedb()

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
				<form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="instId" value="<%=WP_instID%>"> 
					<input type="hidden" name="cartId" value="<%=session("GWOrderId")%>"> 
					<input type="hidden" name="amount" value="<%=pcBillingTotal%>"> 
					<input type="hidden" name="currency" value="<%=WP_Currency%>"> 
					<input type="hidden" name="desc" value="Online Order, ProductCart Store">
					<% if WP_testmode="YES" then %>
						<input type="hidden" name="testMode" value="100"> 
						<input type=hidden name="name" value="AUTHORISED">
					<% else %>
                        <input type="hidden" name="name" value="<%=pcBillingFirstName&" "&pcBillingLastName%>"> 
					<% end if %>
					<input type="hidden" name="address" value="<%=pcBillingAddress%>"> 
					<input type="hidden" name="postcode" value="<%=pcBillingPostalCode%>"> 
					<input type="hidden" name="country" value="<%=pcBillingCountryCode%>"> 
					<input type="hidden" name="tel" value="<%=pcBillingPhone%>"> 
					<input type="hidden" name="email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="city" value="<%=pcBillingCity%>">
					<input type="hidden" name="state" value="<%=pcBillingState%>">    
					<input type="hidden" name="MC_OrderID" value="<%=session("GWOrderId")%>">
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
					<tr>
						<td colspan="2"><img src="images/top_wplogo.gif" alt="WorldPay Payment Gateway" width="249" height="76"></td>
					</tr>
					<tr>
						<td class="pcSpacer" colspan="2"></td>
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
					<% if WP_testmode="YES" then %>
					<tr>
						<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% end if %>

					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					<tr>
						<td class="pcSpacer" colspan="2"></td>
					</tr>
					<tr>
						<td colspan="2">
						<p>NOTE: When you click on the 'Place Order' button, you will temporarily leave our Web site and will be taken to a secure payment page on the WorldPay Web site. You will be redirected back to our store once the transaction has been processed. We have partnered with WorldPay, a leader in secure Internet payment processing, to ensure that your transactions are processed securely and reliably.</p>
						</td>
					</tr>
					<tr>
						<td class="pcSpacer" colspan="2"></td>
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