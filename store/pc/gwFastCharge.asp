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
session("redirectPage")="gwFastCharge.asp"

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
query="SELECT pcPay_FAC_ATSID, pcPay_FAC_TransType, pcPay_FAC_CVV FROM pcPay_FastCharge WHERE pcPay_FAC_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_FAC_ATSID=rs("pcPay_FAC_ATSID")
pcPay_FAC_ATSID=enDeCrypt(pcPay_FAC_ATSID, scCrypPass)
pcPay_FAC_TransType=rs("pcPay_FAC_TransType")
pcv_CVV=rs("pcPay_FAC_CVV")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************

	Set gwObject = Server.CreateObject("ATS.SecurePost")
	if err.number<>0 then
		FAC_success=0
		strErrorInfo="Unable to send payment information, the required COM Object is not installed on this server."
	else
		gwObject.ATSID = pcPay_FAC_ATSID
		gwObject.Amount = Round(pcBillingTotal,2)
		gwObject.CCName = pcBillingFirstName & " " & pcBillingLastName
		gwObject.CCNumber = Request.Form( "CardNumber" )
		gwObject.ExpMonth = Request.Form( "expMonth" )
		gwObject.ExpYear = Request.Form( "expYear" )
		if pcv_CVV="1" then
			gwObject.CVV2 = Request.Form( "CVV" )
		end if
		gwObject.CI_IPAddress = pcCustIpAddress
		gwObject.MerchantOrderNumber = "ORD-" & session("GWOrderID")
		gwObject.CI_CompanyName = pcBillingCompany
		gwObject.CI_BillAddr1 = pcBillingAddress
		gwObject.CI_BillAddr2 = pcBillingAddress2
		gwObject.CI_BillCity = pcBillingCity
		gwObject.CI_BillState = pcBillingState
		gwObject.CI_BillZip = pcBillingPostalCode
		gwObject.CI_BillCountry = pcBillingCountryCode
		gwObject.CI_Phone = pcBillingPhone
		gwObject.CI_Email = pcCustomerEmail
		gwObject.CI_ShipAddr1 = pcShippingAddress
		gwObject.CI_ShipAddr2 = pcShippingAddress2
		gwObject.CI_ShipCity = pcShippingCity
		gwObject.CI_ShipState = pcShippingState
		gwObject.CI_ShipZip = pcShippingPostalCode
		gwObject.CI_ShipCountry = pcShippingCountryCode
		if pcPay_FAC_TransType="1" then
			gwObject.ProcessSale
		else
			gwObject.ProcessAuth
		end if

		response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
	
		FAC_success=0
		
		If gwObject.ResultAccepted Then
			 FAC_success=1
		End if
	
		Dim FAC_rd_successurl, FAC_rd_resultfailurl
	
		If FAC_success=1 Then
			session("GWAuthCode")=gwObject.ResultAuthCode
			session("GWTransId")=gwObject.ResultRefCode
	
			FAC_rd_successurl="gwReturn.asp?s=true&gw=FAC"
			if pcPay_FAC_TransType<>"1" then
				session("GWTransType")="yes"
			end if
		end if

		If (FAC_success <> 1) then
			strErrorInfo=""
			If gwObject.ResultErrorFlag Then
				strErrorInfo="Error: " & gwObject.LastError
			Else
				strErrorInfo="Declined: " & gwObject.ResultAuthCode
			End If
	
			If (strErrorInfo="") Then
				strErrorInfo="There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."
			End if
		End if
	
	End if

	FAC_rd_resultfailurl="msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error</b>: "&strErrorInfo&"<br><br><a href="""&tempURL&"?psslurl=gwFastCharge.asp&idCustomer="&session("idcustomer")&"&idOrder="&Session("GWOrderID")&"""><img src="""&rslayout("back")&""" border=0></a>")
	
	If FAC_success <> 1 Then
		Response.Redirect FAC_rd_resultfailurl
	ElseIf FAC_success=1 Then
		call closeDb()
		Response.Redirect FAC_rd_successurl
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