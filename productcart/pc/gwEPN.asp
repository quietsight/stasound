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
session("redirectPage")="gwEPN.asp"

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
query="SELECT pcPay_EPN_Account, pcPay_EPN_RestrictKey, pcPay_EPN_CVV, pcPay_EPN_testmode FROM pcPay_EPN Where pcPay_EPN_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_EPN_Account=rs("pcPay_EPN_Account")
pcPay_EPN_RestrictKey=rs("pcPay_EPN_RestrictKey")
pcv_CVV=rs("pcPay_EPN_CVV")
pcPay_EPN_testmode=rs("pcPay_EPN_testmode")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim objXMLHTTP, xml

	sRemoteURL = "https://www.eProcessingNetwork.Com/cgi-bin/tdbe/transact.pl"

	if pcPay_EPN_testmode=1 then
		pcPay_EPN_Account="080880"
		pcPay_EPN_RestrictKey="yFqqXJh9Pqnugfr"
		pcCardNumber="4111111111111111"
	else
		pcCardNumber = request.Form("CardNumber")
		pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)
		pcPay_EPN_RestrictKey=enDeCrypt(pcPay_EPN_RestrictKey, scCrypPass)
	end if

	'Get form variables
	stext="ePNAccount="&pcPay_EPN_Account
	stext=stext & "&RestrictKey="&pcPay_EPN_RestrictKey
	stext=stext & "&CardNo="&pcCardNumber
	stext=stext & "&ExpMonth="&request.Form("expMonth")
	stext=stext & "&ExpYear="&request.Form("expYear")
	stext=stext & "&Total="&pcBillingTotal
	stext=stext & "&Address="&pcBillingAddress
	stext=stext & "&Zip="&pcBillingPostalCode
	stext=stext & "&EMail="&pcCustomerEmail
	stext=stext & "&FirstName="&pcBillingFirstName
	stext=stext & "&LastName="&pcBillingLastName
	if pcv_CVV=1 then
		stext=stext & "&CVV2Type=1"
		stext=stext & "&CVV2="&request.Form("CVV")
	else
		stext=stext & "&CVV2Type=0"
	end if
	stext=stext & "&HTML=No"

	'Create & initialize the XMLHTTP object
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	'Open the connection to the remote server
	xml.Open "POST", sRemoteURL, False
	'Send the request to the eProcessingNetwork Transparent Database Engine
	xml.Send stext

	'store the response
	sResponse = xml.responseText

	'parse the response string and handle appropriately
	sApproval = mid(sResponse, 2, 1)

	if sApproval = "Y" then
	elseif sApproval = "N" then
		sDeclineReason = "Your transaction has been declined with the " & _
				 "following response: <b>" & mid(sResponse, 3, 16) & "</b><br>"
	else
		sDeclineReason = "The processor was unable to handle your " & _
				 "transaction, having returned the following response: <b>" & _
				 sResponse & "</b><br>"
	end if
	strStatus = xml.Status
	Set xml = Nothing

	'save and update order
	if sApproval = "Y" then
		Response.redirect "gwReturn.asp?s=true&gw=EPN"
	else
		response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& sDeclineReason &"<br><br><a href="""&tempURL&"?psslurl=gwEPN.asp&idCustomer="&session("idCustomer")&"&idOrder="&	session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		response.end
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
					<% if pcPay_EPN_TestMode=1 then %>
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
								<option value="1">1</option>
								<option value="2">2</option>
								<option value="3">3</option>
								<option value="4">4</option>
								<option value="5">5</option>
								<option value="6">6</option>
								<option value="7">7</option>
								<option value="8">8</option>
								<option value="9">9</option>
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