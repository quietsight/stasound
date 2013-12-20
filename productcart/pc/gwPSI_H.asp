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
<% If request("Approved")= "APPROVED" AND request("OrderID")<>"" then
	session("GWAuthCode")=request("TransRefNumber")
	session("GWTransId")=request("TransRefNumber")
	session("GWTransType")=request("TransactionType")
	session("GWSessionID")=Session.SessionID 
	if session("GWOrderId")="" then
		session("GWOrderId")=request("OrderID")
	end if
		
	response.redirect "gwReturn.asp?s=true&gw=PSI"
End If

'//Set redirect page to the current file name
session("redirectPage")="gwPSI_H.asp"


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
query="SELECT Userid,[Mode],psi_post,psi_testmode FROM PSIGate WHERE id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_Userid=rs("Userid")
pcv_TransType=rs("Mode")
pcv_PSI_Post=rs("psi_post")
pcv_PSI_TestMode=rs("psi_testmode")

if pcv_PSI_TestMode="YES" then 
  session("redirectPage2")= "https://devcheckout.psigate.com/HTMLPost/HTMLMessenger"
Else
 session("redirectPage2")= "https://checkout.psigate.com/HTMLPost/HTMLMessenger" '"https://order.psigate.com/psigate.asp"
end if 

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
					<input type="hidden" name="StoreKey" value="<%=pcv_Userid%>">
					<INPUT TYPE="hidden" NAME="PaymentType" VALUE="CC">
					<INPUT TYPE="hidden" NAME="CustomerRefNo" VALUE="<%=session("idCustomer")%>">
					<input type="hidden" name="Bcompany" value="<%=pcBillingCompany%>">
					<input type="hidden" name="Bname" value="<%=pcBillingFirstName&" "&pcBillingLastName%>" size="45">
					<input type="hidden" name="Baddress1" value="<%=pcBillingAddress%>">
					<INPUT TYPE="hidden" NAME="Baddress2" VALUE="<%=pcBillingAddress2%>">
					<input type="hidden" name="Bcity" value="<%=pcBillingCity%>">
					<% if pcBillingStateCode = "" then pcBillingStateCode= pcBillingProvince End if %>
					<input type="hidden" name="Bprovince" value="<%=pcBillingStateCode%>">
					<input type="hidden" name="Bcountry" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="Email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="Bpostalcode" value="<%=pcBillingPostalCode%>" size="15">
					<input type="hidden" name="Phone" value="<%=pcBillingPhone%>" size="20">
					<input type="hidden" name="OrderID" value="<%=session("GWOrderID")%>"> 
					<input type="hidden" name="Userid" value="HTML Posting">
					<input type="hidden" name="CardAction" value="<%=pcv_TransType%>">
					<% if pcv_PSI_TestMode="YES" then %>
						<input type="hidden" name="TestResult" value="A"> <% 'test only %>
					<% end if %>
						<input type="hidden" name="items" value="1">
						<input type="hidden" name="ItemID1" value="Online Order">
					<% If scCompanyName="" then %>
						<input type="hidden" name="Description1" value="Shopping Cart"> 
					<%else %>
						<input type="hidden" name="Description1" value="<%=scCompanyName%>"> 
					<% end if %>
					<input type="hidden" name="Price1" value="<%=pcBillingTotal%>">
					<input type="hidden" name="Quantity1" value="1">
					<%
					pcv_ThanksURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwPSI_H.asp"),"//","/")
					pcv_ThanksURL=replace(pcv_ThanksURL,"https:/","https://")
					pcv_ThanksURL=replace(pcv_ThanksURL,"http:/","http://")
					%>
					<input type="hidden" name="Subtotal" value="<%=money(pcBillingTotal)%>">
					<input type="hidden" name="ThanksURL" value="<%=pcv_ThanksURL%>">
					<%
					pcv_SorryURL=replace((scStoreURL&"/"&scPcFolder&"/pc/sorry_psi.asp"),"//","/")
					pcv_SorryURL=replace(pcv_SorryURL,"https:/","https://")
					pcv_SorryURL=replace(pcv_SorryURL,"http:/","http://")
					%>
					<input type="hidden" name="NoThanksURL" value="<%=pcv_SorryURL%>">
					<input type="hidden" name="Sname" value="<%=pcShippingFirstName&" "&pcShippingLastName%>">
					<input type="hidden" name="Saddress1" value="<%=pcShippingAddress%>">
					<input type="hidden" name="Saddress2" value="<%=pcShippingAddress2%>">
					<input type="hidden" name="Scity" value="<%=pcShippingCity%>">
					<% if pcshippingStateCode = "" then pcshippingStateCode= pcShippingProvince End if %>
					<input type="hidden" name="Sprovince" value="<%=pcshippingStateCode%>">
					<input type="hidden" name="Spostalcode" value="<%=pcShippingPostalCode%>">
					<input type="hidden" name="Scountry" value="<%=pcShippingCountryCode%>">
					<input type="hidden" name="Comments" value="none"> 
					<INPUT TYPE="hidden" NAME="ResponseFormat" VALUE="HTML">
					<%'Response.end%>
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
                    
					<% if pcv_PSI_TestMode="YES" then %>
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
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td> 
							<input type="text" name="CardNumber" value="">
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="CardExpMonth">
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
							<select name="CardExpYear">
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
								<input name="CardIDNumber" type="text" id="CardIDNumber" value="" size="4" maxlength="4">
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