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
'//Check if this is a post-back
pcv_Response_StatusCode=request("Ecom_Ezic_Response_StatusCode")
if pcv_Response_StatusCode<>"" then
	pcv_Response_StatusCode=request("Ecom_Ezic_Response_StatusCode")
	pcv_Response_AuthCode=request("Ecom_Ezic_Response_AuthCode")
	pcv_Response_AuthMessage=request("Ecom_Ezic_Response_AuthMessage")
	pcv_Response_TransactionID=request("Ecom_Ezic_Response_TransactionID")
	pcv_Response_Card_AVSCode=request("Ecom_Ezic_Response_Card_AVSCode")
	pcv_Response_Card_VerificationCode=request("Ecom_Ezic_Response_Card_VerificationCode")
	pcv_Response_IssueDate=request("Ecom_Ezic_Response_IssueDate")
	'rt_gateway="FastTransact"
	if pcv_Response_StatusCode="F" OR pcv_Response_StatusCode="0" OR pcv_Response_StatusCode="D" then
		Msg=pcv_Response_AuthMessage
		'response.redirect "fasttransact_giveup.asp?msg="&pcv_Response_AuthMessage
	end if
	session("GWAuthCode")=pcv_Response_AuthCode
	session("GWTransId")=pcv_Response_TransactionID
	session("GWSessionID")=Session.SessionID 

	Response.redirect "gwReturn.asp?s=true&gw=FastTransact"
end if

'//Set redirect page to the current file name
session("redirectPage")="gwfast.asp"
session("redirectPage2")="https://secure.fasttransact.com/gw/native/interactive2.2"

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
query="SELECT AccountID, SiteTag, tran_type, card_types, CVV2 FROM fasttransact Where id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
dim pcv_AccountID, pcv_SiteTag, pcv_TranType, pcv_CardTypes, pcv_CVV

pcv_AccountID=rs("AccountID")
pcv_SiteTag=rs("SiteTag")
pcv_TranType=rs("tran_type")
pcv_CardTypes=rs("card_types")
pcv_CVV=rs("CVV2")

set rs=nothing
call closedb()

Dim strReturnURL
If scSSL="" OR scSSL="0" Then
	strReturnURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwFast.asp"),"//","/")
	strReturnURL=replace(strReturnURL,"https:/","https://")
	strReturnURL=replace(strReturnURL,"http:/","http://") 
Else
	strReturnURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwFast.asp"),"//","/")
	strReturnURL=replace(strReturnURL,"https:/","https://")
	strReturnURL=replace(strReturnURL,"http:/","http://")
End If

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
					<input type="hidden" name="Ecom_Ezic_Fulfillment_ReturnURL" value="<%=strReturnURL%>">
					<input type="hidden" name="Ecom_Ezic_Fulfillment_GiveUpURL" value="<%=strReturnURL%>">
					<input type="hidden" name="Ecom_Receipt_Description" value="Online Store Order ID: <%=session("GWOrderId")%>">
					<input type="hidden" name="Ecom_Ezic_AccountAndSitetag" value="<%=pcv_AccountID&":"&pcv_SiteTag%>">
					<input type="hidden" name="Ecom_Cost_Total" value="<%=pcBillingTotal%>">
					<input type="hidden" name="Ezic_HideForm" value="TRUE">
					<input type="hidden" name="Ecom_Ezic_Payment_AuthorizationType" value="<%=pcv_TranType%>">
					<input type="hidden" name="Ecom_ConsumerOrderID" value="<%=session("GWOrderId")%>">
					<input type="hidden" name="Ecom_BillTo_Online_Email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_Name_First" value="<%=pcShippingFirstName%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_Name_Last" value="<%=pcShippingLastName%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_Street_Line1" value="<%=pcShippingAddress%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_Street_Line2" value="<%=pcShippingAddress2%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_City" value="<%=pcShippingCity%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_StateProv" value="<%=pcShippingState%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_PostalCode" value="<%=pcShippingPostalCode%>">
					<input type="hidden" name="Ecom_ShipTo_Postal_CountryCode" value="<%=pcShippingCountryCode%>">
					<input type="hidden" name="Ecom_ShipTo_Telecom_Phone_Number" value="<%=pcShippingPhone%>">
					<input type="hidden" name="Ecom_ShipTo_Online_Eamil" value="">
					<input type="hidden" name="Ecom_BillTo_Postal_Name_First" value="<%=pcBillingFirstName%>">
					<input type="hidden" name="Ecom_BillTo_Postal_Name_Last" value="<%=pcBillingLastName%>">
					<input type="hidden" name="Ecom_BillTo_Postal_Street_Line1" value="<%=pcBillingAddress%>">
					<input type="hidden" name="Ecom_BillTo_Postal_Street_Line2" value="<%=pcBillingAddress2%>"> 
					<input type="hidden" name="Ecom_BillTo_Postal_City" value="<%=pcBillingCity%>"> 
					<input type="hidden" name="Ecom_BillTo_Postal_StateProv" value="<%=pcBillingState%>"> 
					<input type="hidden" name="Ecom_BillTo_Postal_PostalCode" value="<%=pcBillingPostalCode%>">
					<input type="hidden" name="Ecom_BillTo_Postal_CountryCode" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="Ecom_BillTo_Telecom_Phone_Number" value="<%=pcBillingPhone%>"> 
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
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
							<td>  
								<select name="Ecom_Payment_Card_Type">
									<% cardTypeArray=split(pcv_CardTypes,", ")
									dim i
									for i = lbound(cardTypeArray) to ubound(cardTypeArray)
										select case cardTypeArray(i)
											case "V"
												response.write "<option value=""V"" selected>Visa</option>"
											case "M"
												response.write "<option value=""M"">MasterCard</option>"
											case "A"
												response.write "<option value=""A"">American Express</option>"
											case "D"
												response.write "<option value=""D"">Discover</option>"
										end select
									next
									%>
								</select>
							</td>
						</tr>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
							<td> 
								<input type="text" name="Ecom_Payment_Card_Number" value="">
							</td>
						</tr>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
							<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
								<select name="Ecom_Payment_Card_ExpDate_Month">
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
								<select name="Ecom_Payment_Card_ExpDate_Year">
									<option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
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
									<input name="Ecom_Payment_Card_Verification" type="text" id="Ecom_Payment_Card_Verification" value="" size="4" maxlength="4">
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