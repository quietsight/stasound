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
session("redirectPage")="gwAuthorizeCheckConfirm.asp"

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
query="SELECT x_Type, x_Login, x_Password, x_Curcode, x_Method, x_AIMType, x_testmode, x_eCheck, x_secureSource, x_eCheckPending FROM authorizeNet Where id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_Type=rs("x_Type")
pcv_Login=rs("x_Login")
'decrypt
pcv_Login=enDeCrypt(pcv_Login, scCrypPass)
pcv_Password=rs("x_Password")
'decrypt
pcv_Password=enDeCrypt(pcv_Password, scCrypPass)
pcv_Curcode=rs("x_Curcode")
pcv_Method=rs("x_Method")
pcv_AIMType=rs("x_AIMType")
pcv_testmode=rs("x_testmode")
pcv_eCheck=rs("x_eCheck")
pcv_secureSource=rs("x_secureSource")
pcv_eCheckPending=rs("x_eCheckPending")
session("x_eCheckPending")=pcv_eCheckPending
pcv_TypeArray=Split(pcv_Type,"||")
pcv_Type1=pcv_TypeArray(0)

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	'expdate=expmonth & right(expyear, 2)
	Dim objXMLHTTP, xml
	
	'Send the request to the Authorize.NET processor.
	stext="x_version=3.1"
	stext=stext & "&x_delim_data=True"
	stext=stext & "&x_delim_char=,"
	stext=stext & "&x_method=ECHECK"
	if pcv_testmode="1" then
		stext=stext & "&x_Test_Request=True"
	else
		stext=stext & "&x_Test_Request=False"
	end if
	stext=stext & "&x_relay_response=FALSE"
	stext=stext & "&x_login=" & pcv_Login
	stext=stext & "&x_tran_key=" & pcv_Password
	stext=stext & "&x_amount=" & pcBillingTotal
	'check data
	stext=stext & "&x_bank_acct_name="& request.Form("x_bank_acct_name")
	stext=stext & "&x_bank_aba_code="& request.Form("x_bank_aba_code")
	stext=stext & "&x_bank_acct_num="& request.Form("x_bank_acct_num")
	stext=stext & "&x_bank_acct_type="& request.Form("x_bank_acct_type")
	stext=stext & "&x_bank_name="& request.Form("x_bank_name")
	stext=stext & "&x_customer_tax_id="& request.Form("x_customer_tax_id")
	if request.Form("x_customer_tax_id")="" then
		stext=stext & "&x_drivers_license_num="& request.Form("x_drivers_license_num")
		stext=stext & "&x_drivers_license_state="& request.Form("x_drivers_license_state")
		stext=stext & "&x_drivers_license_dob="& request.Form("x_drivers_license_dob")
	end if
	stext=stext & "&x_customer_ip=" & pcCustIpAddress
	if pcv_secureSource="1" then
		stext=stext & "&x_customer_organization_type=" & request.Form("customer_organization_type")
	end if
	stext=stext & "&x_type=AUTH_CAPTURE"
	stext=stext & "&x_echeck_type=WEB"
	stext=stext & "&x_recurring_billing=NO"
	stext=stext & "&x_Currency_Code=" & pcv_Curcode
	stext=stext & "&x_Description=" & replace(scCompanyName,",","-") & " Order: " & session("GWOrderID")
	stext=stext & "&x_Invoice_Num=" & session("GWOrderID")
	stext=stext & "&x_Cust_ID=" & session("idCustomer")
	stext=stext & "&x_first_name=" & pcBillingFirstName
	stext=stext & "&x_last_name=" & pcBillingLastName
	stext=stext & "&x_company=" & replace(pcBillingCompany,",","||")
	stext=stext & "&x_address=" & replace(pcBillingAddress,",","||")
	stext=stext & "&x_city=" & pcBillingCity
	stext=stext & "&x_state=" & pcBillingState
	stext=stext & "&x_zip=" & pcBillingPostalCode
	stext=stext & "&x_country=" & pcBillingCountryCode
	stext=stext & "&x_phone=" & pcBillingPhone
	stext=stext & "&x_email=" & pcCustomerEmail
	stext=stext & "&x_Ship_To_First_Name=" & pcShippingFirstName
	stext=stext & "&x_Ship_To_Last_Name=" & pcShippingLastName
	stext=stext & "&x_Ship_To_Address=" & replace(pcShippingAddress,",","||")
	stext=stext & "&x_Ship_To_City=" & pcShippingCity
	stext=stext & "&x_Ship_To_State=" & pcShippingState
	stext=stext & "&x_Ship_To_Zip=" & pcShippingPostalCode
	stext=stext & "&x_Ship_To_Country=" & pcShippingCountryCode
	
	'Send the transaction info as part of the querystring
	set xml =  Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.open "POST", "https://secure.authorize.net/gateway/transact.dll?" & stext & "", false
	xml.send ""
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText
	Set xml = Nothing
	
	strArrayVal = split(strRetVal, ",", -1)
	session("x_response_code")=strArrayVal(0)
	session("x_response_subcode")=strArrayVal(1)
	session("x_response_reason_code")=strArrayVal(2)
	session("x_response_reason_text")=strArrayVal(3)
	session("GWAuthCode")=strArrayVal(4)    '6 digit approval code
	session("x_avs_code")=strArrayVal(5)
	session("GWTransId")=strArrayVal(6)    'transaction id
	session("x_invoice_num")=strArrayVal(7)
	session("x_description")=strArrayVal(8)
	session("x_amount")=strArrayVal(9)
	session("x_method")=strArrayVal(10)
	session("x_type")=strArrayVal(11)
	session("x_cust_id")=strArrayVal(12)
	session("x_first_name")=strArrayVal(13)
	session("x_last_name")=strArrayVal(14)
	pcv_company = strArrayVal(15)
	session("x_address")=strArrayVal(16)
	pcv_city = strArrayVal(17)
	pcv_state= strArrayVal(18)
	session("x_zip")=strArrayVal(19)
	
	pcv_country                 = strArrayVal(20)
	pcv_phone                   = strArrayVal(21)
	pcv_fax                     = strArrayVal(22)
	pcv_email                   = strArrayVal(23)
	pcv_ship_to_first_name      = strArrayVal(24)
	pcv_ship_to_last_name       = strArrayVal(25)
	pcv_ship_to_company         = strArrayVal(26)
	pcv_ship_to_address         = strArrayVal(27)
	pcv_ship_to_city            = strArrayVal(28)
	pcv_ship_to_state           = strArrayVal(29)
	pcv_ship_to_zip             = strArrayVal(30)
	pcv_ship_to_country         = strArrayVal(31)

	'Check the ErrorCode to make sure that the component was able to talk to the authorization network
	If (strStatus <> 200) Then
		Response.Write "An error occurred during processing. Please try again later."
	else
		'save and update order 
		If session("x_response_code") = 1 Then
			'save info in authOrders
			tordnum=(int(session("x_invoice_num"))-scpre)
			call opendb()
			query="INSERT INTO authorders (idOrder, amount, paymentmethod, transtype, authcode, ccnum, ccexp, idCustomer, fname, lname, address, zip, captured) VALUES ("&tordnum&", "&session("x_amount")&", 'ECHECK', '"&x_Type1&"', '"&session("GWAuthCode")&"','1111111111111111','0000',"&session("x_cust_id")&",'"&replace(session("x_first_name"),"'","''")&"', '"&replace(session("x_last_name"),"'","''")&"', '"&replace(session("x_address"),"'","''")&"', '"&session("x_zip")&"',0);"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closedb()
			Response.redirect "gwReturn.asp?s=true&gw=AIM&c=true"
			
		elseif session("x_response_code")<>1 then
			response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;"&session("x_response_code")&"</b>: "& session("x_response_reason_text")&"<br><br><a href="""&tempURL&"?psslurl=gwAuthorizeCheck.asp&idCustomer="&session("x_cust_id")&"&idOrder="&session("x_invoice_num")&"""><img src="""&rslayout("back")&""" border=0></a>")
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
					<% if pcv_testmode="1" then %>
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
								<option value="CHECKING">Checking Account</option>
								<option value="SAVINGS">Savings Account</option>
							</select>
						</td>
					</tr>
					<tr> 
						<td>
							<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_16")%></p>
						</td>
						<td>  
							<input name="x_bank_name" type="text" size="20" maxlength="20">
						</td>
					</tr>
<% if x_secureSource="1" then %>
						<tr> 
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_17")%></p>
							</td>
							<td> 
								<input type="radio" name="x_customer_organization_type" value="I" class="clearBorder">Individual 
								<input type="radio" name="x_customer_organization_type" value="B" class="clearBorder">Business
							</td>
						</tr>
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
                    <tr><td colspan="2" class="pcSpacer"></td></tr>
					<tr>
                    <td colspan="2"><strong>By clicking the button below, I authorize <%=scCompanyName%> to charge my account specified above on <%=Now() %> for the amount of <%=money(pcBillingTotal)%> for the contents of my cart in <%=scCompanyName%> store.</strong></td></tr>
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