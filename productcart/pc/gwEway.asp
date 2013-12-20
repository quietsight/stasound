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
session("redirectPage")="gwEway.asp"

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

pcv_BeagleNotAvailable=0

'//Check if Beagle Field Exists
on error resume next
err.clear
query="SELECT * FROM eWay;"
set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query)
eWay_BeagleActive=rstemp("eWayBeagleActive")
if err.number<>0 then
	pcv_BeagleNotAvailable=0
else
	pcv_BeagleNotAvailable=1
end if
set rstemp=nothing

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT eWayCustomerid, eWayPostMethod, eWayTestmode, eWayCVV"
if pcv_BeagleNotAvailable=1 then
	query=query&", eWayBeagleActive"
end if
query=query&" FROM eWay WHERE eWayID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcEwayCustomerid=rs("eWayCustomerid")
pcEwayPostMethod=rs("eWayPostMethod")
pcEwayPostMethod="XML"
pcEwayTestmode=rs("eWayTestmode")
pcEwayBillingTotal=pcBillingTotal
if pcEwayTestmode=1 then
	pcEwayCustomerid="87654321"
	pcEwayBillingTotal="10.00"
end if
pcEwayCVV = rs("eWayCVV")
if pcv_BeagleNotAvailable=1 then
	pcEwayBeagleActive = rs("eWayBeagleActive")
else
	pcEwayBeagleActive="0"
end if
pcEwayBillingTotal = replacecomma(pcEwayBillingTotal)
pcEwayBillingTotal = (pcEwayBillingTotal*100)

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	if pcEwayTestmode=1 then
	   if pcEwayCVV ="1" Then
	   		pcEwayURL = "https://www.eway.com.au/gateway_cvn/xmltest/testpage.asp"
	   Else
			pcEwayURL = "https://www.eway.com.au/gateway/xmltest/TestPage.asp"
	   End if 	
	else
	   if pcEwayCVV ="1" Then
	   		pcEwayURL = "https://www.eway.com.au/gateway_cvn/xmlpayment.asp"
			if pcEwayBeagleActive = "1" then
				pcEwayURL = "http://www.eway.com.au/gateway_cvn/xmlbeagle.asp"
			end if
	   Else
			pcEwayURL = "https://www.eway.com.au/gateway/xmlpayment.asp"
			if pcEwayBeagleActive = "1" then
				pcEwayURL = "http://www.eway.com.au/gateway_cvn/xmlbeagle.asp"
			end if
	   end if		
	end if

	Dim SrvEWayXmlHttp, pcEwayXMLPostData
	pcEwayXMLPostData=""
	pcEwayXMLPostData=pcEwayXMLPostData&"<?xml version=""1.0"" encoding=""UTF-8""?>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewaygateway>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerID>"&pcEwayCustomerid&"</ewayCustomerID>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayTotalAmount>"&pcEwayBillingTotal&"</ewayTotalAmount>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerFirstName>"&pcBillingFirstName&"</ewayCustomerFirstName>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerLastName>"&pcBillingLastName&"</ewayCustomerLastName>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerEmail>"&pcCustomerEmail&"</ewayCustomerEmail>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerAddress>"&pcBillingAddress&"</ewayCustomerAddress>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerPostcode>"&pcBillingPostalCode&"</ewayCustomerPostcode>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerInvoiceDescription>Online Order</ewayCustomerInvoiceDescription>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerInvoiceRef>"&session("GWOrderId")&"</ewayCustomerInvoiceRef>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardHoldersName>"&pcBillingFirstName&" "&pcBillingLastName&"</ewayCardHoldersName>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardNumber>"&request("CardNumber")&"</ewayCardNumber>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardExpiryMonth>"&request("expMonth")&"</ewayCardExpiryMonth>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCardExpiryYear>"&request("expYear")&"</ewayCardExpiryYear>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayTrxnNumber>"&session("GWOrderId")&"</ewayTrxnNumber>"
	
	if pcEwayCVV ="1" Then
		pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCVN>"&request.form("CVV")&"</ewayCVN>"
	End if 
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayOption1></ewayOption1>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayOption2></ewayOption2>"
	pcEwayXMLPostData=pcEwayXMLPostData&"<ewayOption3></ewayOption3>"
	'//eWay's Beagle Fraud Prevention
	if pcEwayBeagleActive = "1" then
		pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerIPAddress>"&xxx&"</ewayCustomerIPAddress>" 
		pcEwayXMLPostData=pcEwayXMLPostData&"<ewayCustomerBillingCountry>"&xxx&"</ewayCustomerBillingCountry>"
	end if 

	pcEwayXMLPostData=pcEwayXMLPostData&"</ewaygateway>"
	'response.write pcEwayXMLPostData&"<HR>"
	
	Set SrvEWayXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	err.clear
	SrvEWayXmlHttp.open "POST", pcEwayURL, false
	SrvEWayXmlHttp.send(pcEwayXMLPostData)
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	EWayResult = SrvEWayXmlHttp.responseText
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
	xmlDoc.async = False
	If xmlDoc.loadXML(SrvEWayXmlHttp.responseText) Then
		' Get the results
		pcResultEwayTrxnStatus = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnStatus").Text
		pcResultEwayTrxnNumber = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnNumber").Text
		pcResultEwayTrxnOption1 = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnOption1").Text
		pcResultEwayTrxnOption2 = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnOption2").Text
		pcResultEwayTrxnOption3 = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnOption3").Text
		pcResultEwayTrxnReference = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnReference").Text
		pcResultEwayAuthCode = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayAuthCode ").Text
		pcResultEwayReturnAmount = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayReturnAmount ").Text
		pcResultEwayTrxnError = xmlDoc.documentElement.selectSingleNode("/ewayResponse/ewayTrxnError").Text
	Else
		'//ERROR
		response.write "Failed to process response"
		response.end
	End If
	'response.write pcResultEwayTrxnReference&"<BR>"
	'response.write pcResultEwayTrxnNumber
	'response.end
	if ucase(pcResultEwayTrxnStatus)="TRUE" then
		session("GWAuthCode")=pcResultEwayTrxnReference
		session("GWTransId")=pcResultEwayTrxnNumber
		Response.redirect "gwReturn.asp?s=true&gw=eWay"
		Set eWay = Nothing
		response.end
	else
		Set eWay = Nothing
		response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& pcResultEwayTrxnError &"<br><br><a href="""&tempURL&"?psslurl=gwEway.asp&idCustomer="&session("idCustomer")&"&idOrder="&	session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		%>
		<%
		response.end
	end if

	'//redirect to gwReturn.asp
	
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
					<% if pcEwayTestmode=1 then %>
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
					<% If pcEwayCVV="1" Then %>
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