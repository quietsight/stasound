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
'====================================
'// Turn CVV on - "1"=on, "0"=off
pcv_CVV="1"
'====================================


'//Set redirect page to the current file name
session("redirectPage")="gwPSI_XML.asp"

if session("GWOrderDone")="YES" then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
	session("GWOrderDone")=""
	response.redirect tempURL
end if

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
query="SELECT Config_File_Name, Userid, [Mode], psi_testmode FROM PSIGate WHERE (((id)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
psi_XMLPassPhrase=rs("Config_File_Name")
psi_XMLStoreID=rs("Userid")
psi_XMLTransType=rs("Mode")
psi_XMLTestmode=rs("psi_testmode")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	if psi_XMLTestmode="YES" then
		PSiGateURL = "https://dev.psigate.com:7989/Messenger/XMLMessenger"
		psi_XMLPassPhrase="psigate1234"
		psi_XMLStoreID="teststore"
	else
		PSiGateURL = "https://secure.psigate.com:7934/Messenger/XMLMessenger"
	end if
	
	Dim SrvPSiGateXmlHttp, pcPSiGateXMLPostData
	pcPSiGateXMLPostData=""
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<?xml version=""1.0"" encoding=""UTF-8""?>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Order>"
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
		
	pcv_strItemizeOrder = 1

	if pcv_CatDiscounts>0 or trim(pcv_strDiscountDetails)<>"No discounts applied." then
		pcv_strItemizeOrder = 0
	end if
	
	IF pcv_strItemizeOrder = 1 THEN
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Itemized Order
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>
		<!--#include file="pcPay_PSiGate_Itemize.asp"-->
	<%	end if
	
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<StoreID>"&Server.HTMLEncode(psi_XMLStoreID)&"</StoreID>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Passphrase>"&Server.HTMLEncode(psi_XMLPassPhrase)&"</Passphrase>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Tax1>"&Server.HTMLEncode(pcv_strFinalTax)&"</Tax1>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<ShippingTotal>"&Server.HTMLEncode(pcv_strFinalShipCharge)&"</ShippingTotal>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Subtotal>"&Server.HTMLEncode(pcBillingTotal)&"</Subtotal>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<PaymentType>"&Server.HTMLEncode("CC")&"</PaymentType>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardAction>"&Server.HTMLEncode(psi_XMLTransType)&"</CardAction>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardNumber>"&Server.HTMLEncode(Request.Form("Cardnumber"))&"</CardNumber>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardExpMonth>"&Server.HTMLEncode(Request.Form("expMonth"))&"</CardExpMonth>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardExpYear>"&Server.HTMLEncode(Request.Form("expYear"))&"</CardExpYear>"
	If pcv_CVV="1" Then
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardIDCode>1</CardIDCode>"
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CardIDNumber>"&Server.HTMLEncode(Request.Form("CVV"))&"</CardIDNumber>"
	end if
	if psi_XMLTestmode="YES" then
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<TestResult>A</TestResult>"
	end if
	if psi_XMLTestmode="YES" then
		pcTestOrderID = Hour(Now) & Minute(Now) & Second(Now)
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<OrderID>"&pcTestOrderID&"PCTEST"&Server.HTMLEncode(session("GWOrderId"))&"</OrderID>"
	else
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<OrderID>"&Server.HTMLEncode(session("GWOrderId"))&"</OrderID>"
	end if
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<UserID>"&Server.HTMLEncode(session("idCustomer"))&"</UserID>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bname>"&Server.HTMLEncode(pcBillingFirstName&" "&pcBillingLastName)&"</Bname>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bcompany>"&Server.HTMLEncode(pcBillingCompany)&"</Bcompany>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Baddress1>"&Server.HTMLEncode(pcBillingAddress)&"</Baddress1>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Baddress2>"&Server.HTMLEncode(pcBillingAddress2)&"</Baddress2>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bcity>"&Server.HTMLEncode(pcBillingCity)&"</Bcity>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bprovince>"&Server.HTMLEncode(pcBillingState)&"</Bprovince>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bpostalcode>"&Server.HTMLEncode(pcBillingPostalCode)&"</Bpostalcode>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Bcountry>"&Server.HTMLEncode(pcBillingCountry)&"</Bcountry>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Sname>"&Server.HTMLEncode(pcShippingFirstName&" "&pcShippingLastName)&"</Sname>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Scompany></Scompany>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Saddress1>"&Server.HTMLEncode(pcShippingAddress)&"</Saddress1>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Saddress2>"&Server.HTMLEncode(pcShippingAddress2)&"</Saddress2>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Scity>"&Server.HTMLEncode(pcShippingCity)&"</Scity>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Sprovince>"&Server.HTMLEncode(pcShippingState)&"</Sprovince>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Spostalcode>"&Server.HTMLEncode(pcShippingPostalCode)&"</Spostalcode>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Scountry>"&Server.HTMLEncode(pcShippingCountryCode)&"</Scountry>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Phone>"&Server.HTMLEncode(pcBillingPhone)&"</Phone>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Email>"&Server.HTMLEncode(pcCustomerEmail)&"</Email>"
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<Comments></Comments>"
	if psi_XMLTestmode="YES" then
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CustomerIP>66.249.66.203</CustomerIP>"
	else
		pcPSiGateXMLPostData=pcPSiGateXMLPostData&"<CustomerIP>"&Server.HTMLEncode(pcCustIpAddress)&"</CustomerIP>"
	end if
	
	pcPSiGateXMLPostData=pcPSiGateXMLPostData&"</Order>"

	Set SrvPSiGateXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	err.clear
	SrvPSiGateXmlHttp.open "POST", PSiGateURL, false
	SrvPSiGateXmlHttp.send(pcPSiGateXMLPostData)
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	PSiGateResult = SrvPSiGateXmlHttp.responseText

	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
	xmlDoc.async = False
	If xmlDoc.loadXML(SrvPSiGateXmlHttp.responseText) Then
		' Get the results
		pcResultApproved = xmlDoc.documentElement.selectSingleNode("/Result/Approved").Text
		pcResultErrorMsg = xmlDoc.documentElement.selectSingleNode("/Result/ErrMsg").Text
		pcResultTransRefNumber = xmlDoc.documentElement.selectSingleNode("/Result/TransRefNumber").Text
		pcResultCardAuthNumber = xmlDoc.documentElement.selectSingleNode("/Result/CardAuthNumber").Text
		pcResultCardRefNumber = xmlDoc.documentElement.selectSingleNode("/Result/CardRefNumber").Text
	Else
		'//ERROR
		Response.Write "Transaction error or declined.  Error Message: " & pcResultErrorMsg
		response.end
	End If
	If pcResultApproved="APPROVED" then
		session("GWAuthCode")=pcResultCardAuthNumber
		session("GWTransId")=pcResultTransRefNumber
		response.redirect "gwReturn.asp?s=true&gw=PSIGate"
	Else
		if pcResultErrorMsg="" then
			pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"&PSiGateResult
		end if
		Msg=pcResultErrorMsg
	End if

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
					<input type="hidden" name="ItemID1" value="Online Order">
					<% If scCompanyName="" then %>
						<input type="hidden" name="Description1" value="Shopping Cart"> 
					<%else %>
						<input type="hidden" name="Description1" value="<%=scCompanyName%>"> 
					<% end if %>
					<input type="hidden" name="Price1" value="<%=pcBillingTotal%>"> 
					<input type="hidden" name="Quantity1" value="1">
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
					<% if psi_XMLTestmode="YES" then %>
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
							<input type="text" name="CardNumber" value="" size="20">
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