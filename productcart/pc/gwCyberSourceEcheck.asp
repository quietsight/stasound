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
session("redirectPage")="gwCyberSourceEcheck.asp"

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
	oMerchantConfig.TargetAPIVersion = "1.32"
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

	'oRequest( "ccAuthService_run" ) = "true"
	
	 oRequest( "ecDebitService_commerceIndicator" ) = "internet"
	 oRequest( "ecDebitService_run" ) = "true"
	
	 
	'if pcPay_Cys_TransType="2" then
		'oRequest( "ccCaptureService_run" ) = "true"
	'end if
	' we will let the Client get the merchantID from the MerchantConfig object
	' and insert it into the Hashtable.


	oRequest( "merchantReferenceCode" ) = "ORD-" & session("GWOrderId")
	oRequest( "merchantID" ) =pcPay_Cys_MerchantId
	oRequest( "clientApplication" ) = "ProductCart"
	oRequest( "clientApplicationVersion" ) = "v3.11"
	nameoncheck = split(Request.Form( "x_bank_acct_name" ), " ")
	if ubound(nameoncheck) > 0 Then
		oRequest( "billTo_firstName" ) = nameoncheck(0)
		oRequest( "billTo_lastName" ) = nameoncheck(1)		
	end if 
	
	oRequest( "billTo_company" ) = Request.Form( "x_bank_acct_name" )
	oRequest( "billTo_street1" ) = pcBillingAddress
	oRequest( "billTo_city" ) = pcBillingCity
	oRequest( "billTo_state" ) = pcBillingState
	oRequest( "billTo_postalCode" ) = pcBillingPostalCode
	oRequest( "billTo_country" ) = pcBillingCountryCode
	oRequest( "billTo_email" ) = pcCustomerEmail
	oRequest( "billTo_phoneNumber" ) = pcBillingPhone
	
	oRequest("ecDebitService_paymentMode") = "0"
	oRequest( "billTo_companyTaxID" ) = Request.Form( "x_customer_tax_id" )
	oRequest( "businessRules_declineAVSFlags" ) ="n"
	
	oRequest( "billTo_driversLicenseNumber" ) = Request.Form( "x_drivers_license_num" )
	oRequest( "billTo_driversLicenseState" ) = Request.Form( "x_drivers_license_state" )
	if isdate(Request.form("x_drivers_license_dob")) Then
		if Month(Request.form("x_drivers_license_dob")) < 10 Then 
			dtMonth = "0" & Month(Request.form("x_drivers_license_dob"))
		else
			dtMonth =  Month(Request.form("x_drivers_license_dob"))
		end if 
		if Day(Request.form("x_drivers_license_dob")) < 10 Then 
			dtDay = "0" & Day(Request.form("x_drivers_license_dob"))
		else
			dtDay = Day(Request.form("x_drivers_license_dob"))
		end if 
		oRequest( "billTo_dateOfBirth" ) =  year(Request.form("x_drivers_license_dob")) &"-" & dtMonth&"-" & dtDay
	else
		oRequest( "billTo_dateOfBirth" ) = "1970-01-01"	
	End if 
	
	


	
	' Check info 
	oRequest( "check_accountNumber" ) = Request.Form( "x_bank_acct_num" )
	oRequest( "check_accountType" ) = Request.Form( "x_bank_acct_type" )
	oRequest( "check_bankTransitNumber" ) = Request.Form( "x_bank_aba_code" )
	'oRequest( "check_checkNumber" ) = "1040"
	

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
	
	'names = oRequest.Names

	'For Each name in names

  	'Response.Write name & "=" & oRequest.Value(name) &"<BR>"

	'Next
 

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
	Response.write nStatus &" stat<BR>"

	select case nStatus
	
		case 0:
			dim decision
			decision = UCase( varReply( "decision" ) )
			'Response.write UCase( varReply( "decision" ) ) &"<BR>"
			'Response.write UCase( varReply( "ecDebitReply_reasonCode" ) ) &"<BR>"
			'Response.write UCase( varReply( "requestID" ) ) &"<BR>"
			'Response.write UCase( varReply( "invalidField_0" ) ) &"<BR>"
	        ' Response.end
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
		session("GWTransType")="Echeck"

		cys_rd_successurl="gwReturn.asp?s=true&gw=CYS"
	end if
	
	If (cys_success <> 1) and (strErrorInfo="") Then
		strErrorInfo="There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."
	End if
	
	cys_rd_resultfailurl="msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error</b>: "&strErrorInfo&"<br><br><a href="""&tempURL&"?psslurl=gwCyberSourceEcheck.asp&idCustomer="&session("idcustomer")&"&idOrder="&session("GWOrderID")&"""><img src="""&rslayout("back")&""" border=0></a>")

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
		<td><img src="images/checkout_bar_step5.gif" alt=""></td>
	</tr>
	<tr>
		<td class="pcSpacer"></td>
	</tr>
	<tr>
		<td>
			
					<form action="<%=session("redirectPage")%>" method="POST" name="form1" class="pcForms">
				

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
					<% if x_testmode="1" then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						</tr>
					<% end if %>
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
							<td colspan="2" align="center"><img src="images/sampleck.gif" width="390" height="230"></td>
						</tr>
					    <tr>
					      <td>&nbsp;</td>
					      <td><a href="http://www.achex.com/html/NSF_pop.jsp" target="_blank">Returned Check Fees</a></td>
				      </tr>
				      <tr> 
						<td> 
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_12")%></p>
						</td>
						<td>
							<input name="x_bank_acct_name" type="text" size="35" maxlength="50">
						</td>
					</tr>
					<tr> 
						<td> 
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_13")%></p>
						</td>
						<td>  
							<input name="x_bank_aba_code" type="text" size="35">
						</td>
					</tr>
					<tr> 
						<td> 
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_14")%></p>
						</td>
						<td>  
							<input name="x_bank_acct_num" type="text" size="35">
						</td>
					</tr>
					<tr> 
						<td> 
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_15")%></p>
						</td>
						<td>
							<select name="x_bank_acct_type">
								<option value="C">Checking Account</option>
								<option value="S">Savings Account</option>
								<option value="X">Corporate Checking Account</option>
							</select>
						</td>
					</tr>
					<!--tr> 
						<td>
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_16")%></p>
						</td>
						<td>  
							<input name="x_bank_name" type="text" size="20" maxlength="20">
						</td>
					</tr-->
<% if x_secureSource="1" then %>
						<!--tr> 
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_17")%></p>
							</td>
							<td> 
								<input type="radio" name="x_customer_organization_type" value="I" class="clearBorder">Individual 
								<input type="radio" name="x_customer_organization_type" value="B" class="clearBorder">Business
							</td>
						</tr-->
					<% end if %>
					<tr>
					  <td colspan="2">
							<p><b><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_18")%></b></p>
						</td>
			  	</tr>
					<tr> 
						<td>
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_19")%></p>
						</td>
						<td>
							<input name="x_customer_tax_id" type="text" size="9" maxlength="9">
						</td>
					</tr>
					<tr>
					  <td colspan="2">
							<p><b><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_20")%></b></p>
						</td>
			  	</tr>
					<tr> 
						<td>
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_21")%></p>
						</td>
						<td>
							<input name="x_drivers_license_num" type="text" size="35" maxlength="50">
						</td>
					</tr>
					<tr> 
						<td>
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_22")%></p>
						</td>
						<td>
							<input name="x_drivers_license_state" type="text" size="2" maxlength="2"> 
							<span class="pcSmallText"><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_23")%></span>
						</td>
					</tr>
					<tr> 
						<td>
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_24")%></p>
						</td>
						<td>
							<input name="x_drivers_license_dob" type="text" size="10" maxlength="10"> 
							<span class="pcSmallText"><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_25")%></span>
						</td>
					</tr>
					<tr>
					  <td colspan="2" class="pcSpacer"></td>
			  	</tr>
					
					
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					<tr>
					  <td colspan="2" class="pcSpacer"></td>
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