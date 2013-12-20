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
session("redirectPage")="gwNETOne.asp"

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
query="SELECT pcPay_NETOne_Mid, pcPay_NETOne_Mkey, pcPay_NETOne_Tcode, pcPay_NETOne_CVV, pcPay_NETOne_CardTypes FROM pcPay_NETOne WHERE pcPay_NETOne_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_NETOne_Mid=rs("pcPay_NETOne_Mid")
'decrypt
pcPay_NETOne_Mid=enDeCrypt(pcPay_NETOne_Mid, scCrypPass)
pcPay_NETOne_Mkey=rs("pcPay_NETOne_Mkey")
pcPay_NETOne_Mkey=enDeCrypt(pcPay_NETOne_Mkey, scCrypPass)
pcPay_NETOne_Tcode=rs("pcPay_NETOne_Tcode")
pcv_CVV=rs("pcPay_NETOne_CVV")
pcPay_NETOne_CardTypes=rs("pcPay_NETOne_CardTypes")
set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

'*************************************************************************************
' This is where you would post info to the gateway
' START
'*************************************************************************************

	Dim objXMLHTTP, xml
					
	'Send the request to the NET1 processor.
	'=======================================================
	strData = "M_id="&pcPay_NETOne_Mid
	strData = strData & "&" & "M_key="&pcPay_NETOne_Mkey
					strData = strData & "&" & "C_name=" & pcBillingFirstName&" "&pcBillingLastName
					strData = strData & "&" & "C_address=" & pcBillingAddress
					strData = strData & "&" & "C_city=" & pcBillingCity
					strData = strData & "&" & "C_state=" & pcBillingState
					strData = strData & "&" & "C_zip=" & pcBillingPostalCode
					strData = strData & "&" & "C_country=" & pcBillingCountryCode
					strData = strData & "&" & "C_email=" & pcCustomerEmail
					strData = strData & "&" & "C_cardnumber=" & request.Form("CardNumber")
					strData = strData & "&" & "C_exp=" & request.Form("expMonth") & request.Form("expYear")
					strData = strData & "&" & "T_code=" & pcPay_NETOne_Tcode
					strData = strData & "&" & "C_cvv=" & request.Form("CVV")
					strData = strData & "&" & "T_amt=" & pcBillingTotal
					strData = strData & "&" & "T_Ordernum=" &session("GWOrderId")
					strData = strData & "&" & "C_ship_name=" &pcShippingFirstName&" "&pcShippingLastName
					strData = strData & "&" & "C_ship_address" &pcShippingAddress&", "&pcShippingAddress2
					strData = strData & "&" & "C_ship_city=" &pcShippingCity
					strData = strData & "&" & "C_ship_state=" &pcShippingStateCode
					strData = strData & "&" & "C_ship_zip=" &pcShippingPostalCode
					strData = strData & "&" & "C_ship_country=" &pcShippingCountryCode
					strData = strData & "&" & "C_telephone=" &pcBillingPhone
				
					Set obj = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
					obj.Open "POST", "https://va.eftsecure.net/cgi-bin/eftBankcard.dll?transaction", False
					obj.Send strData
					strStatus = obj.Status
				
					response_string = obj.responseText 
					'Parse Response and Print Out Results Per Spec 
					
					'Response.Write "Full Packet Response: " & response_string & "<BR>"
					strApprovalIndicator=mid(response_string,2,1)
				
					'Response.Write "Approval Indicator: " & strApprovalIndicator & "<BR>"
					strErrorCode=mid(response_string,3,6)
					'Response.Write "Approval/Error Code: " & mid(response_string,3,6) & "<BR>"
					strErrorMessage= mid(response_string,9,32)
					'Response.Write "Approval/Error Message: " & mid(response_string,9,32) & "<BR>"
					'Response.Write "Front-End Indicator: " & mid(response_string,41,2) & "<BR>"
					'Response.Write "CVV Indicator: " & mid(response_string,43,1) & "<BR>"
					'Response.Write "AVS Indicator: " & mid(response_string,44,1) & "<BR>"
					'Response.Write "Risk Indicator: " & mid(response_string,45,2) & "<BR>"
					strReference=mid(response_string,47,10)
					'Response.Write "Reference: " & mid(response_string,47,10) & "<BR>"
					intOrderNumber=mid(response_string, _
					 (InStr(1, response_string, Chr(28)) + 1), InStr((InStr(1, response_string, Chr(28)) + 1), _
						response_string, Chr(28)) - (InStr(1, response_string, Chr(28)) + 1))
				
					Set obj = Nothing
					
					'Check the ErrorCode to make sure that the component was able to talk to the authorization network
					If (strStatus <> 200) Then
						Response.Write "An error occurred during processing. Please try again later."
					else
						'save and update order 
						If strApprovalIndicator = "A" Then
							session("GWAuthCode")=strReference
							session("GWTransId")=strReference
							session("GWTransType")=pcPay_NETOne_Tcode
							
							Response.redirect "gwReturn.asp?s=true&gw=NETOne"
						elseif strApprovalIndicator = "X" or strApprovalIndicator = "E" then
							response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;"&strErrorCode&"</b>: "& strErrorMessage&"<br><br><a href="""&tempURL&"?psslurl=gwNETOne.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
				
							response.end
						End If
					end if
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