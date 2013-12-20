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
<!--#include file="gwUSAePay_xcenums.inc"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwUSAePayCheck.asp"

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
query="SELECT pcPay_Uep_SourceKey,pcPay_Uep_CheckPending,pcPay_Uep_TestMode FROM pcPay_USAePay WHERE pcPay_Uep_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_Uep_SourceKey=rs("pcPay_Uep_SourceKey")
pcPay_Uep_SourceKey=enDeCrypt(pcPay_Uep_SourceKey, scCrypPass)
pcPay_Uep_CheckPending=rs("pcPay_Uep_CheckPending")
pcPay_Uep_TestMode=rs("pcPay_Uep_TestMode")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Set XCharge1 = Server.CreateObject("USAePayXChargeCom2.XChargeCom2")
	
	if request.Form("CHECKTYPE")="1" then
		XCharge1.Command = checkcredit
	else
		XCharge1.Command = check
	end if

	XCharge1.Sourcekey = pcPay_Uep_SourceKey
	XCharge1.IP = pcCustIpAddress

	if pcPay_Uep_TestMode="1" then
		XCharge1.Testmode = True
	else
		XCharge1.Testmode = False
	end if

	XCharge1.Routing = Request.Form( "BANKROUTING" )
	XCharge1.Account = Request.Form( "CHECKACCT" )
	XCharge1.SSN = Request.Form( "SSN" )
	XCharge1.DLNum=Request.Form( "DLNUM" )
	XCharge1.DLState=Request.Form( "DLSTATE" )

	XCharge1.Amount = pcBillingTotal
	
	XCharge1.Invoice = "ORD-" & session("GWOrderId")
	XCharge1.Description = "ORDER ID: #" & session("GWOrderId")
	
	XCharge1.TransHolderName = pcBillingFirstName & " " & pcBillingLastName
	XCharge1.Street = pcBillingAddress
	XCharge1.Zip = pcBillingPostalCode
	
	XCharge1.BillFName = pcBillingFirstName
	XCharge1.BillLName = pcBillingLastName
	XCharge1.BillCompany = pcBillingCompany
	XCharge1.BillStreet = pcBillingAddress
	XCharge1.BillStreet2 = pcBillingAddress2
	XCharge1.BillCity = pcBillingCity
	XCharge1.BillState = pcBillingState
	XCharge1.BillZip = pcBillingPostalCode
	XCharge1.BillCountry = pcBillingCountryCode
	XCharge1.BillPhone = pcBillingPhone
	XCharge1.Email = pcCustomerEmail
				
	XCharge1.ShipFName = pcShippingFirstName
	XCharge1.ShipLName = pcShippingLastName
	XCharge1.ShipCompany = pcShippingCompany
	XCharge1.ShipStreet = pcShippingAddress
	XCharge1.ShipStreet2 = pcShippingAddress2
	XCharge1.ShipCity = pcShippingCity
	XCharge1.ShipState = pcShippingState
	XCharge1.ShipZip = pcShippingPostalCode
	XCharge1.ShipCountry = pcShippingCountryCode
	XCharge1.ShipPhone = pcShippingPhone
	XCharge1.Process
	
	response.buffer=true
	response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
				
	uep_success=0
				
	Select Case XCharge1.ResponseStatus
		Case Approved
		uep_success=1
	End Select

	Dim uep_rd_successurl, uep_rd_resultfailurl

	If uep_success=1 Then
		session("GWAuthCode")=XCharge1.ResponseAuthCode
		session("GWTransId")=XCharge1.ResponseReferenceNum

		uep_rd_successurl="gwReturn.asp?s=true&gw=UEP&c=1"
		if pcPay_Uep_CheckPending="1" then
			uep_rd_successurl=uep_rd_successurl
		end if
	end if
				
	If (uep_success <> 1) then
		
		strErrorInfo=""
		
		If XCharge1.ErrorExists = True Then            
			Dim XError
			
			For Each XError In XCharge1.Errors
				strErrorInfo=strErrorInfo&"<br>"
				strErrorInfo=strErrorInfo & "Error code: " & XError.ErrorCode  & " - Error Message: " & XError.ErrorText
			Next
		End If
					
		If (strErrorInfo="") Then
			strErrorInfo="There was a problem completing your order. We apologize for the inconvenience. Please contact customer support to review your order."
		End if
	End if
				
	uep_rd_resultfailurl="msgb.asp?message="&server.URLEncode("<b>Error</b>: "&strErrorInfo&"<br><br><a href="""&tempURL&"?psslurl=gwUSAePayCheck.asp&idCustomer="&session("idcustomer")&"&idOrder="&Session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
			
	If uep_success <> 1 Then
		Response.Redirect uep_rd_resultfailurl
	ElseIf uep_success=1 Then
		call closeDb()
		Response.Redirect uep_rd_successurl
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
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp">Edit</a></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
                   
					<%  if pcPay_Uep_TestMode="1" then %>
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
						<td colspan="2" align="center"><img src="images/sampleck.gif" width="390" height="230"></td>
					</tr>
					<tr> 
						<td><p>Bank Routing Number:</p></td>
						<td> 
							<input name="BANKROUTING" type="text" size="35" maxlength="50">
						</td>
					</tr>
					
					<tr> 
						<td><p>Checking Account Number:</p></td>
						<td><input name="CHECKACCT" type="text" size="35"></td>
					</tr>
					<tr> 
						<td><p>Account Type:</p></td>
						<td>
							<select name="CHECKTYPE">
								<option value="0">Check</option>
								<option value="1">Checkcredit</option>
							</select>
						</td>
					</tr>
					<tr> 
						<td><p>Social Security number:</p></td>
						<td><input name="SSN" type="text" size="20" maxlength="35"></td>
					</tr>
					<tr> 
						<td><p>Drivers License Number:</p></td>
						<td><input name="DLNUM" type="text" size="20" maxlength="35"></td>
					</tr>
					<tr> 
						<td><p>Drivers License Issuing State:</p></td>
						<td><input name="DLSTATE" type="text" size="20" maxlength="35"></td>
					</tr>
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