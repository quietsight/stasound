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
pcv_PNREF = request("PNREF")

if pcv_PNREF<>"" then
	gwTransID=request("order_number")

	session("GWAuthCode")=request("AUTHCODE")
	session("GWTransId")=pcv_PNREF
	if session("GWOrderId")="" then
		session("GWOrderId")=request("INVOICE")
	end if
	session("GWSessionID")=Session.SessionID 
	session("GWTransType")=request("TYPE")
	
	Response.redirect "gwReturn.asp?s=true&gw=PFLink"
end if

'//Set redirect page to the current file name
session("redirectPage")="gwpfl.asp"
'session("redirectPage2")="https://payments.verisign.com/payflowlink"
'//New URL effective May 31, 2008
session("redirectPage2")="https://payflowlink.paypal.com"

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="0" Then
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
query="SELECT v_User,v_Partner,pfl_testmode,pfl_transtype,pfl_CSC FROM verisign_pfp WHERE id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_User=rs("v_User")
pcv_Partner=rs("v_Partner")
pfl_testmode=rs("pfl_testmode")
pcv_transtype=rs("pfl_transtype")
pcv_CVV=rs("pfl_CSC")

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
					<input type="hidden" name="ADDRESS" value="<%=pcBillingAddress%>">
					<input type="hidden" name="CITY" value="<%=pcBillingCity%>">
					<input type="hidden" name="LOGIN" value="<%=pcv_User%>">
					<input type="hidden" name="PARTNER" value="<%=pcv_Partner%>">
					<input type="hidden" name="AMOUNT" value="<%=pcBillingTotal%>">
					<input type="hidden" name="TYPE" value="<%=pcv_transtype%>">
					<input type="hidden" name="METHOD" value="CC">
					<input type="hidden" name="ZIP" value="<%=pcBillingPostalCode%>">
					<input type="hidden" name="CUSTID" value="<%=session("GWOrderId")%>">
					<input type="hidden" name="EMAIL" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="INVOICE" value="<%=session("GWOrderId")%>">
					<input type="hidden" name="NAME" value="<%=pcBillingFirstName&" "&pcBillingLastName%>">
					<input type="hidden" name="PHONE" value="<%=pcBillingPhone%>">
					<input type="hidden" name="STATE" value="<%=pcBillingState%>">
                    <input type="hidden" name="COUNTRY" value="<%=pcBillingCountryCode%>">
                    <% if len(pcShippingFullName)> 0 then %>
						<input type="hidden" name="NAMETOSHIP" value="<%=pcShippingFullName%>">
                    <% end if %>
                    <% if len(pcShippingAddress)> 0 then %>
						<input type="hidden" name="ADDRESSTOSHIP" value="<%=pcShippingAddress%>">
                    <% end if %>
                    <% if len(pcShippingCity)> 0 then %>
						<input type="hidden" name="CITYTOSHIP" value="<%=pcShippingCity%>">
                    <% end if %>					
                    <% if len(pcShippingCountryCode)> 0 then %>
						<input type="hidden" name="COUNTRYCODE" value="<%=pcShippingCountryCode%>">
                    <% end if %>
                    <% if len(pcShippingState)> 0 then %>
						<input type="hidden" name="STATETOSHIP" value="<%=pcShippingState%>">
                    <% end if %>
                    <% if len(pcShippingPostalCode)> 0 then %>
						<input type="hidden" name="ZIPTOSHIP" value="<%=pcShippingPostalCode%>">
                    <% end if %>
                    <% if len(pcShippingEmail)> 0 then %>
                    	<input type="hidden" name="EMAILTOSHIP" value="<%=pcShippingEmail%>">
                    <% end if %>
                    <% if len(pcShippingFax)> 0 then %>
                    	<input type="hidden" name="FAXTOSHIP" value="<%=pcShippingFax%>">
                    <% end if %>
                    <% if len(pcShippingPhone)> 0 then %>
                    	<input type="hidden" name="PHONETOSHIP" value="<%=pcShippingPhone%>">
                    <% end if %>

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
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td> 
							<input type="text" name="CARDNUM" value="">
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td><input name="EXPDATE" type="text" id="EXPDATE" value="" size="7" maxlength="7">&nbsp;(mmyy)		</td>
					</tr>
					<% If pcv_CVV="YES" Then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
								<input name="CSC" type="text" id="CSC" value="" size="4" maxlength="4">
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