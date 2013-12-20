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
If request("xid")<>"" then
		session("GWAuthCode")=request("authcode")
		session("GWTransId")=request("xid")
		session("GWTransType")=request("xxauth")
		if session("GWOrderId")="" then
			session("GWOrderId")=request("idOrder")
		end if
		session("GWSessionID")=Session.SessionID 
		response.redirect "gwReturn.asp?s=true&gw=iTransact"
end if

'//Set redirect page to the current file name
session("redirectPage")="gwit.asp"
session("redirectPage2")="https://secure.paymentclearing.com/cgi-bin/rc/ord.cgi"

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
query="SELECT URL, Gateway_ID, it_amex, it_diner, it_disc, it_mc, it_visa, ReqCVV FROM ITransact WHERE id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_GatewayId=rs("Gateway_ID")
it_amex=rs("it_amex")
it_diner=rs("it_diner")
it_disc=rs("it_disc")
it_mc=rs("it_mc")
it_visa=rs("it_visa")
pcv_CVV=rs("ReqCVV")

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
				<input type="hidden" name="passback" value="idOrder">
				<input type="hidden" name= "idOrder" value="<%=session("GWOrderId")%>">
				
				<input type="hidden" name="PaymentSubmitted" value="Go">
				<input type="hidden" name="vendor_id" value="<%=pcv_GatewayId%>">
				<% pcv_HomePageURL=replace((scStoreURL&"/"&scPcFolder),"//","/")
				pcv_HomePageURL=replace(pcv_HomePageURL,"http:/","http://")
				pcv_HomePageURL=replace(pcv_HomePageURL,"https:/","https://") %>
				<input type="hidden" name="home_page" value="<%=pcv_HomePageURL&"/pc/default.asp"%>">
				<input type="hidden" name="ret_addr" value="<%=pcv_HomePageURL&"/pc/gwIT.asp"%>">
				
				<input type="hidden" name="1_desc" value="Online Order">
				<input type="hidden" name="1_cost" value="<%=replace(money(pcBillingTotal),",","")%>">
				<input type="hidden" name="1_qty" value="1">

				<input type="hidden" name="showaddr" value="1"> 
				<% if pcv_CVV="1" then %>
				<input type="hidden" name="showcvv" value="1">
				<% end if %>
				<input type="hidden" name="mername" value="<%=scCompanyName%>"> 
				<input type="hidden" name="acceptcards" value="1"> 
				<input type="hidden" name="acceptchecks" value="0"> 
				<input type="hidden" name="accepteft" value="0"> 
				<input type="hidden" name="altaddr" value="1"> 
				<input type="hidden" name="nonum" value="1"> 
				
				<input type="hidden" name="lookup" value="xid">
				<input type="hidden" name="lookup" value="authcode">
				<input type="hidden" name="lookup" value="avs_response">
				<input type="hidden" name="lookup" value="cvv2_response">
				<input type="hidden" name="formtype" value="2">
				<input type="hidden" name="cust_id" value="<%=session("idCustomer")%>">
					 
				<!-- SSS ON CUSTOMER'S SITE - REQUIRED -->
				<input type="hidden" name="first_name" value="<%=pcBillingFirstName%>"> 
				<input type="hidden" name="last_name" value="<%=pcBillingLastName%>"> 
				<input type="hidden" name="address" value="<%=pcBillingAddress%>"> 
				<input type="hidden" name="city" value="<%=pcBillingCity%>"> 
				<input type="hidden" name="state" value="<%=pcBillingState%>"> 
				<input type="hidden" name="zip" value="<%=pcBillingPostalCode%>"> 
				<input type="hidden" name="country" value="<%=pcBillingCountryCode%>"> 
				<input type="hidden" name="phone" value="<%=pcBillingPhone%>"> 
				<input type="hidden" name="email" value="<%=pcCustomerEmail%>"> 
				<input type="hidden" name="sfname" value="<%=pcShippingFirstName%>"> 
				<input type="hidden" name="slname" value="<%=pcShippingLastName%>"> 
				<input type="hidden" name="saddr" value="<%=pcShippingAddress%>"> 
				<input type="hidden" name="scity" value="<%=pcShippingCity%>"> 
				<input type="hidden" name="sstate" value="<%=pcShippingState%>"> 
				<input type="hidden" name="szip" value="<%=pcShippingPostalCode%>"> 
				<input type="hidden" name="sctry" value="<%=pcShippingCountryCode%>"> 


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
							<input type="text" name="ccnum" value="">
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="ccmo">
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
							<select name="ccyr">
								<option value="<%=right(dtCurYear,4)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,4)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,4)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,4)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,4)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,4)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,4)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,4)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,4)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,4)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,4)%>"><%=dtCurYear+10%></option>
							</select>
						</td>
					</tr>
					<% If pcv_CVV="1" Then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
							<INPUT NAME="cvv2_number" SIZE=5>
							Code is present on my card but illegible: 
							<INPUT type=checkbox NAME="cvv2_illegible" value="1">
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