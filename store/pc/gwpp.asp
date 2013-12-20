<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/shipFromsettings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="header.asp"-->
<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalStandard"

'******************************************************************
'// PayPal Itemized Order
'// To change this value from the default "non-Itemized Order"
'// you will need to change the variable below to the value of 1.
'//
'// For Example: 
'// pcv_strItemizeOrder = 1

'******************************************************************
'// Set to "non-Itemized Order" by Default
pcv_strItemizeOrder = 0	
'******************************************************************


'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass


'//Set redirect page to the current file name
session("redirectPage")="gwpp.asp"
session("redirectPage2")="https://www.paypal.com/cgi-bin/webscr"

session("PP_SendMail_A")="0"
session("PP_SendMail_C")="0"
session("PP_SendMail_AF")="0"  


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
<% 
'// Fix the phone number
pcBillingPhoneA=""
pcBillingPhoneB=""
pcBillingPhoneC=""

if pcBillingPhone<>"" AND isNULL(pcBillingPhone)=False AND pcBillingCountryCode="US" then
	pcBillingPhone=replace(pcBillingPhone,"-","")
	pcBillingPhone=replace(pcBillingPhone,".","")
	pcBillingPhone=replace(pcBillingPhone," ","")
	pcBillingPhone=replace(pcBillingPhone,"(","")
	pcBillingPhone=replace(pcBillingPhone,")","")
	pcBillingPhoneLength=len(pcBillingPhone)	
	if pcBillingPhoneLength=10 then
		pcBillingPhoneA=left(pcBillingPhone,3)
		pcBillingPhoneB=left(right(pcBillingPhone,7),3)
		pcBillingPhoneC=right(pcBillingPhone,4)
	end if
end if

'//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT Pay_To, PP_Currency, PP_Sandbox FROM paypal WHERE ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_PayTo=rs("Pay_To")
pcv_PPCurrency=rs("PP_Currency")
pcv_PP_Sandbox=rs("PP_Sandbox")

'// Check For Test Mode
if pcv_PP_Sandbox=1 then
	session("redirectPage2")="https://www.sandbox.paypal.com/cgi-bin/webscr"
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
					<%
					call opendb()
					
					'// Check for Discounts that are not compatible with "Itemization"
					query="SELECT orders.discountDetails, orders.pcOrd_CatDiscounts FROM orders WHERE orders.idOrder="&(int(session("GWOrderId"))-scpre)&";"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if not rs.eof then
						pcv_strDiscountDetails=rs("discountDetails")
						pcv_CatDiscounts=rs("pcOrd_CatDiscounts")						
					end if
					
					set rs=nothing
					call closedb()
					
					if pcv_CatDiscounts>0 or trim(pcv_strDiscountDetails)<>"No discounts applied." then
						pcv_strItemizeOrder = 0
					end if
					
					IF pcv_strItemizeOrder = 1 THEN
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Start: Itemized Order
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					%>
					<input type="hidden" name="cmd" value="_cart">
					<input type="hidden" name="upload" value="1">									
					<!--#include file="pcPay_PayPal_Itemize.asp"-->
					<%
					'// PayPal requires two decimal places with a "." decimal separator.
					pcv_strFinalTotal= pcf_CurrencyField(money(pcv_strFinalTotal))
					pcv_strFinalShipCharge= pcf_CurrencyField(money(pcv_strFinalShipCharge))
					pcv_strFinalServiceCharge= pcf_CurrencyField(money(pcv_strFinalServiceCharge))
					pcv_strFinalTax= pcf_CurrencyField(money(pcv_strFinalTax))
					ItemTotal= pcf_CurrencyField(money(ItemTotal))
					%>
					<input type="hidden" name="shipping_1" value="<%=pcv_strFinalShipCharge%>">
					<input type="hidden" name="shipping2_1" value="0">
					<input type="hidden" name="handling_1" value="<%=pcv_strFinalServiceCharge%>">
					<%	
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' End: Itemized Order
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
					ELSE		
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Start: Totaled Order
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
					%>	
					<input type="hidden" name="cmd" value="_ext-enter">				
					<input type="hidden" name="quantity" value="1">
					<input type="hidden" name="item_name" value="<%=scCompanyName & " Items"%>">
					<input type="hidden" name="item_number" value="<%=session("GWOrderId")%>">				
					<% if scDecSign = "," then %>
					<input type="hidden" name="amount" value="<%=pcBillingTotal%>">
					<% else %>
					<input type="hidden" name="amount" value="<%=money(pcBillingTotal)%>">
					<% end if %>
					<input type="hidden" name="shipping" value="0">
					<input type="hidden" name="shipping2" value="0">
					<input type="hidden" name="handling" value="">
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' End: Totaled Order
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
					END IF
					%>	
					<% 
					If (scSSL<>"" AND scSSL<>"0" AND scCompanyLogo<>"") Then
						tempURL=scSslURL &"/"& scPcFolder & "/pc/" & "catalog/" & scCompanyLogo
						tempURL=replace(tempURL,"///","/")
						tempURL=replace(tempURL,"//","/")
						tempURL=replace(tempURL,"https:/","https://")
						tempURL=replace(tempURL,"http:/","http://")
						logoURL		= tempURL 
					End If
					if logoURL<>"" then %>
					<input type="hidden" name="cpp_header_image" value="<%=logoURL%>">
					<% end if %>		
					<input type="hidden" name="business" value="<%=pcv_PayTo%>">			
					<input type="hidden" name="redirect_cmd" value="_xclick">					
					<input type="hidden" name="currency_code" value="<%=pcv_PPCurrency%>">					
					<input type="hidden" name="invoice" value="<%=session("GWOrderId")%>">
					<input type="hidden" name="custom" value="">					
					<%
					NotifyUrl=replace((scStoreURL&"/"&scPcFolder&"/pc/paypalOrdConfirm.asp"),"//","/")
					NotifyUrl=replace(NotifyUrl,"http:/","http://")
					NotifyUrl=replace(NotifyUrl,"https:/","https://")
					%>
					<input type="hidden" name="notify_url" value="<%=NotifyUrl%>">
					<input type="hidden" name="return" value="<%=NotifyUrl%>?pcOID=<%=session("GWOrderId")%>">					
					<input type="hidden" name="rm" value="2">
					<input type="hidden" name="cancel_return" value="<%=NotifyUrl%>">
					<input type="hidden" name="BN" value="ProductCart_Cart_STD_US">
					<input type="hidden" name="first_name" value="<%=pcBillingFirstName%>">
					<input type="hidden" name="last_name" value="<%=pcBillingLastName%>">
					<input type="hidden" name="address1" value="<%=pcBillingAddress%>">
					<input type="hidden" name="address2" value="<%=pcBillingAddress2%>">
					<input type="hidden" name="city" value="<%=pcBillingCity%>">
					<input type="hidden" name="state" value="<%=pcBillingState%>">
					<input type="hidden" name="zip" value="<%=pcBillingPostalCode%>">
					<input type="hidden" name="country" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="night_phone_a" value="<%=pcBillingPhoneA%>">
					<input type="hidden" name="night_phone_b" value="<%=pcBillingPhoneB%>">
					<input type="hidden" name="night_phone_c" value="<%=pcBillingPhoneC%>">
					<input type="hidden" name="email" value="<%=pcCustomerEmail%>">					
					<INPUT type="hidden" name="cbt" value="<%=dictLanguage.Item(Session("language")&"_GateWay_23")%>">
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
						<td colspan="2"><p><img src="images/PayPal_mark_50x34.gif" width="50" height="34" alt="PayPal"></p><br>
</td>
					</tr>
					<tr> 
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_22")%></p>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td width="5%" nowrap><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td width="95%"><%=scCurSign & money(pcBillingTotal)%></td>
					</tr>
					<tr>
						<td colspan="2" align="center">&nbsp;</td>
					</tr>
					<tr> 
						<td colspan="2" align="left">
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