<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwCyberSource.asp"

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
query="SELECT pcPay_Cys_MerchantId, pcPay_Cys_TransType, pcPay_Cys_CardType, pcPay_Cys_CVV, pcPay_Cys_TestMode FROM pcPay_CyberSource WHERE pcPay_Cys_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_Cys_MerchantId=rs("pcPay_Cys_MerchantId")
pcPay_Cys_MerchantId=enDeCrypt(pcPay_Cys_MerchantId, scCrypPass)
pcPay_Cys_TransType=rs("pcPay_Cys_TransType")
pcPay_Cys_CardType=rs("pcPay_Cys_CardType")
x_CVV=rs("pcPay_Cys_CVV")
pcPay_Cys_TestMode=rs("pcPay_Cys_TestMode")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then
	dim varReply, nStatus, strErrorInfo	
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	dim oMerchantConfig
	set oMerchantConfig = Server.CreateObject( "CyberSourceWS.MerchantConfig" )
	if err.number<>0 then
		response.write err.description
		response.End()
	end if
	oMerchantConfig.MerchantID = pcPay_Cys_MerchantId
	if PPD="1" then
		filename="/"&scPcFolder&"/" & scAdminFolderName
	else
		filename="../"&scAdminFolderName
	end if

	oMerchantConfig.KeysDirectory = Server.MapPath (filename) 
	if pcPay_Cys_TestMode="1" then
		oMerchantConfig.SendToProduction = "1"
	else
		oMerchantConfig.SendToProduction = "0"
	end if
	oMerchantConfig.TargetAPIVersion = "1.7"
	oMerchantConfig.EnableLog = "1"
	
	if PPD="1" then
		filename="/"&scPcFolder&"/includes"
	else
		filename="../includes"
	end if
	oMerchantConfig.LogDirectory = Server.MapPath(filename) 

	' set up the request by creating a Hashtable and adding fields to it
	dim oRequest	
	set oRequest = Server.CreateObject( "CyberSourceWS.Hashtable" )

	oRequest( "ccAuthService_run" ) = "true"

	if pcPay_Cys_TransType="2" then
		oRequest( "ccCaptureService_run" ) = "true"
	end if
	' we will let the Client get the merchantID from the MerchantConfig object
	' and insert it into the Hashtable.

	oRequest( "merchantReferenceCode" ) = "ORD-" & session("GWOrderId")
	
	oRequest( "clientApplication" ) = "ProductCart"
	oRequest( "clientApplicationVersion" ) = "v3"
	
	oRequest( "billTo_firstName" ) = pcBillingFirstName
	oRequest( "billTo_lastName" ) = pcBillingLastName
	oRequest( "billTo_street1" ) = pcBillingAddress
	oRequest( "billTo_city" ) = pcBillingCity
	oRequest( "billTo_state" ) = pcBillingState
	oRequest( "billTo_postalCode" ) = pcBillingPostalCode
	oRequest( "billTo_country" ) = pcBillingCountryCode
	oRequest( "billTo_email" ) = pcCustomerEmail
	oRequest( "billTo_phoneNumber" ) = pcBillingPhone
	oRequest( "card_cardType" ) = Request.Form( "CardType" )
	oRequest( "card_accountNumber" ) = Request.Form( "CardNumber" )
	oRequest( "card_expirationMonth" ) = Request.Form( "expMonth" )
	oRequest( "card_expirationYear" ) = Request.Form( "expYear" )
	if clng(x_CVV)=1 then
		oRequest( "card_cvNumber" ) = Request.Form( "CVV" )
	end if
	oRequest( "purchaseTotals_currency" ) = "USD"
	oRequest( "purchaseTotals_grandTotalAmount" ) = pcBillingTotal
	oRequest("shipTo_firstName") = pcShippingFirstName
	oRequest("shipTo_lastName") = pcShippingLastName
	oRequest("shipTo_street1")  = pcShippingAddress
	oRequest("shipTo_city") =  pcShippingCity
	oRequest("shipTo_state") =  pcShippingState
	oRequest("shipTo_postalCode") = pcShippingPostalCode
	oRequest("shipTo_country") = pcShippingCountryCode
	if pcCustIpAddress <> "" then
		oRequest( "billTo_ipAddress" ) = pcCustIpAddress
	end if
	
	' create Client object
	dim oClient
	set oClient = Server.CreateObject( "CyberSourceWS.Client" )
	
	' send request now
	nStatus = oClient.RunTransaction( _
	oMerchantConfig, Nothing, Nothing, _
	oRequest, varReply, strErrorInfo )

	response.buffer=true
	response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
	
	cys_success=0

	select case nStatus
	
		case 0:
			dim decision
			decision = UCase( varReply( "decision" ) )
			
			if decision = "ACCEPT" then
				cys_success=1
			end if
	
	end select

	Dim cys_rd_successurl, cys_rd_resultfailurl

	If cys_success=1 Then
		Cys_AuthCode=varReply( "ccAuthReply_authorizationCode" )
		Cys_TransId=varReply( "requestID" )
		session("GWAuthCode")=Cys_AuthCode
		session("GWTransId")=Cys_TransId
		session("GWTransType")=pcPay_Cys_TransType

		cys_rd_successurl="gwReturn.asp?s=true&gw=CYS"
	end if
	
	If (cys_success <> 1) and (strErrorInfo="") Then
		strErrorInfo="There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."
	End if
	
	cys_rd_resultfailurl="msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error</b>: "&strErrorInfo&"<br><br><a href="""&tempURL&"?psslurl=gwCyberSource.asp&idCustomer="&session("idcustomer")&"&idOrder="&session("GWOrderID")&"""><img src="""&rslayout("back")&""" border=0></a>")

	If cys_success <> 1 Then
		Response.Redirect cys_rd_resultfailurl
	ElseIf cys_success=1 Then
		call closeDb()
		Response.Redirect cys_rd_successurl
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
					<% if pcPay_Cys_TestMode="0" then %>
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
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
						<td> 
						<select name="CardType">
							<%if pcPay_Cys_CardType="V" OR (instr(pcPay_Cys_CardType,"V,")>0) or (instr(pcPay_Cys_CardType,", V")>0) then%>
								<option value="001">Visa</option>
								<%end if%>
								<%if pcPay_Cys_CardType="M" OR (instr(pcPay_Cys_CardType,"M,")>0) or (instr(pcPay_Cys_CardType,", M")>0) then%>
								<option value="002">MasterCard</option>
								<%end if%>
								<%if pcPay_Cys_CardType="A" OR (instr(pcPay_Cys_CardType,"A,")>0) or (instr(pcPay_Cys_CardType,", A")>0) then%>
								<option value="003">American Express</option>
								<%end if%>
								<%if pcPay_Cys_CardType="D" OR (instr(pcPay_Cys_CardType,"D,")>0) or (instr(pcPay_Cys_CardType,", D")>0) then%>
								<option value="004">Discover</option>
								<%end if%>
								<%if pcPay_Cys_CardType="E" OR (instr(pcPay_Cys_CardType,"E,")>0) or (instr(pcPay_Cys_CardType,", E")>0) then%>
								<option value="005">Diners</option>
							<%end if%>
						</select>
						</td>
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
					<% If x_CVV="1" Then %>
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