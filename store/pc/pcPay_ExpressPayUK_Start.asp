<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/pcPayPalUKClass.asp"-->
<!--#include file="DBsv.asp"-->
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<% response.Buffer = true %>
<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalExp"

'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass




'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************


'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=Session("pcCartIndex")

If session("customerType")=1 Then
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
	'Wholesale minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=205"
	end if
Else
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then
	'Retail minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=206"
	end if
End If

'///////////////////////////////////////////////////////////////////////////////
'// START: GET DATA FROM DB
 '///////////////////////////////////////////////////////////////////////////////

'// Open Db
call opendb()

'// Declare Local Variables at once
'>>> pcPay_PayPal_TransType
'>>> PaymentAction
'>>> pcPay_PayPal_Username
'>>> pcPay_PayPal_Password
'>>> pcPay_PayPal_Sandbox
'>>> pcPay_PayPal_Method
'>>> pcPay_PayPal_Signature
objPayPalClass.pcs_SetAllVariables()

'// Close Db
call closedb()

'///////////////////////////////////////////////////////////////////////////////
'// END: GET DATA FROM DB
'///////////////////////////////////////////////////////////////////////////////



'///////////////////////////////////////////////////////////////////////////////
'// START: GET ORDER DETAILS
'///////////////////////////////////////////////////////////////////////////////
'// Order Total
if session("pcPay_PayPalExp_OrderTotal")="" OR session("pcPay_PayPalExp_OrderTotal")=0 then
	session("pcPay_PayPalExp_OrderTotal")=calculateCartTotal(pcCartArray, ppcCartIndex)
end if
OrderTotal=session("pcPay_PayPalExp_OrderTotal")
if OrderTotal="" then
	OrderTotal=0
end if
OrderTotal=money(OrderTotal)
OrderTotal=pcf_CurrencyField(OrderTotal)

'// Currency Code Type
currencyCodeType = pcPay_PayPal_Currency

'// Express URLs
url = objPayPalClass.GetURL()
returnURL	= url & "pcPay_ExpressPayUK_Start.asp?currencyCodeType=" &  currencyCodeType & "&paymentAmount=" & OrderTotal & "&paymentType=" &PaymentAction 
cancelURL	= url & "viewcart.asp?cmd=_express-checkout"

If (scSSL<>"" AND scSSL<>"0" AND scCompanyLogo<>"") Then
	tempURL=scSslURL &"/"& scPcFolder & "/pc/" & "catalog/" & scCompanyLogo
	tempURL=replace(tempURL,"///","/")
	tempURL=replace(tempURL,"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
	logoURL		= tempURL 
End If

'// Sandbox or Live URL
pcv_PayPal_URL	= objPayPalClass.GetECURL(pcPay_PayPal_Method)
pcv_PayPal_URL = pcv_PayPal_URL & "?cmd=_express-checkout&token="	


'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
'///////////////////////////////////////////////////////////////////////////////
'// END: GET ORDER DETAILS
'///////////////////////////////////////////////////////////////////////////////




'///////////////////////////////////////////////////////////////////////////////
'// START: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////

'// Set our token
Dim Token
Token=Request.Querystring("TOKEN")
session("PayPalExpressToken")=Token

'// Begin Post If No Token
If  Request.QueryString("token")="" Then
	'***********************************************************************
	'// Start: Posting Details to PayPal
	'***********************************************************************

		
		'---------------------------------------------------------------------------
		' Construct the parameter string that describes the PayPal payment the varialbes 
		' were set in the web form, and the resulting string is stored in nvpstr
		'
		' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
		'---------------------------------------------------------------------------
		nvpstr="" '// clear 
		objPayPalClass.AddNVP "CLIENTIP", pcCustIpAddress
		objPayPalClass.AddNVP "AMT", OrderTotal		
		objPayPalClass.AddNVP "TENDER", "P" '// C = credit card, P = PayPal
		objPayPalClass.AddNVP "ACTION", "S" '// S = Set, G = Get, D = Do
		objPayPalClass.AddNVP "TRXTYPE", PaymentAction '// S = Sale transaction, A = Authorisation, C = Credit, D = Delayed Capture, V = Void 
		objPayPalClass.AddNVP "CURRENCY", pcPay_PayPal_Currency		
		objPayPalClass.AddNVP "RETURNURL", returnURL
		objPayPalClass.AddNVP "CANCELURL", cancelURL
		
		if logoURL<>"" then
		'	objPayPalClass.AddNVP "HDRIMG", logoURL
		end if
		'response.Write(nvpstr)
		'response.End()
		
		'--------------------------------------------------------------------------- 
		' Make the call to PayPal to set the Express Checkout token
		' If the API call succeded, then redirect the buyer to PayPal
		' to begin to authorize payment.  If an error occurred, show the
		' resulting errors
		'---------------------------------------------------------------------------
		Set resArray = objPayPalClass.hash_call("SetExpressCheckout",nvpstr)
		Set Session("nvpResArray")=resArray
		ack = UCase(resArray("RESPMSG"))

		if err.number <> 0 then			
			'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
			Dim pcv_PayPalErrMessage
			%><!--#include file="../includes/pcPayPalErrors.asp"--><%	
			session("ExpressCheckoutPayment")=""							
		end if

		If instr(ack,"APPROVED")>0 Then
		
				'// Redirect to paypal.com here
				token = resArray("TOKEN")
				payPalURL = pcv_PayPal_URL & token
				objPayPalClass.ReDirectURL(payPalURL)
				
		Else 

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start: Error Reporting
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			'// Combine User Friendly Errors from "pcPay_PayPal_Errors.asp"
			'// with Code errors from string "Declined String".
			'// Return a formatted error report as the string "pcv_PayPalErrMessage".
			objPayPalClass.GenerateErrorReport()
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' End: Error Reporting
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			session("ExpressCheckoutPayment")=""
				
		End If
	
	'***********************************************************************
	'// End: Posting Details to PayPal
	'***********************************************************************
Else
	'***********************************************************************
	'// Start: Get Details from PayPal
	'***********************************************************************

	'// Create a Session Flag
	session("ExpressCheckoutPayment")="YES"	
	
	'---------------------------------------------------------------------------
	' At this point, the buyer has completed in authorizing payment
	' at PayPal.  The script will now call PayPal with the details
	' of the authorization, incuding any shipping information of the
	' buyer.  Remember, the authorization is not a completed transaction
	' at this state - the buyer still needs an additional step to finalize
	' the transaction
	'---------------------------------------------------------------------------	
	Session("currencyCodeType") = getUserInput(Request.Querystring("currencyCodeType"),0)
	Session("paymentAmount") = getUserInput(Request.Querystring("paymentAmount"),0)
	Session("PaymentType")= getUserInput(Request.Querystring("PaymentType"),0)
	Session("PayerID")= getUserInput(Request.Querystring("PayerID"),0)


	'---------------------------------------------------------------------------
	' Build a second API request to PayPal, using the token as the
	' ID to get the details on the payment authorization
	'
	' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
	'---------------------------------------------------------------------------
	nvpstr="" '// clear 
	objPayPalClass.AddNVP "IPADDRESS", pcCustIpAddress
	objPayPalClass.AddNVP "TENDER", "P" '// C = credit card, P = PayPal
	objPayPalClass.AddNVP "ACTION", "G" '// S = Set, G = Get, D = Do
	objPayPalClass.AddNVP "TRXTYPE", "S" '// PaymentAction '// S = Sale transaction, A = Authorisation, C = Credit, D = Delayed Capture, V = Void 
	objPayPalClass.AddNVP "TOKEN", session("PayPalExpressToken")
	
	'---------------------------------------------------------------------------
	' Make the API call and store the results in an array.  If the
	' call was a success, show the authorization details, and provide
	' an action to complete the payment.  If failed, show the error
	'---------------------------------------------------------------------------
	Set resArray = objPayPalClass.hash_call("GetExpressCheckoutDetails",nvpstr)
	ack = UCase(resArray("RESPMSG"))
	Set Session("nvpResArray")=resArray
	
	'// Successful Get Express Details
	If ack="APPROVED" Then


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Set Express Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
		pcStrShippingPhone= getUserInput(resArray("PHONENUM"),0)
		pcv_Payer= getUserInput(resArray("EMAIL"),0)
		session("Payer")=pcv_Payer
		pcv_PayerID= getUserInput(resArray("PAYERID"),0)
		session("PayerId")=pcv_PayerID
		pcv_PayerStatus= getUserInput(resArray("PAYERSTATUS"),0)
		pcv_PayerBusiness= getUserInput(resArray("BUSINESS"),0)	
		pcv_FirstName= getUserInput(resArray("FIRSTNAME"),0)
		pcv_LastName= getUserInput(resArray("LASTNAME")	,0)	
		pcv_FullName= pcv_FirstName & " " & pcv_LastName	
		pcv_ShipToName =  getUserInput(resArray("SHIPTONAME"),0)	
		pcv_Street1= getUserInput(resArray("SHIPTOSTREET"),0)
		pcv_Street2= getUserInput(resArray("SHIPTOSTREET2"),0)
		pcv_CityName= getUserInput(resArray("SHIPTOCITY"),0)
		pcv_StateOrProvince= getUserInput(resArray("SHIPTOSTATE"),0)
		pcv_StateCode= getUserInput(resArray("SHIPTOSTATE"),0)
		pcv_Country=getUserInput(resArray("SHIPTOCOUNTRY"),0)
		if pcv_Country = "AU" or pcv_Country = "CA" then 
		  call opendb()
			query="SELECT stateCode,stateName FROM states WHERE pcCountryCode = '"&pcv_Country&"' and stateName='"&pcv_StateCode&"'"		
			set rsStates=server.CreateObject("ADODB.RecordSet")
			set rsStates=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsStates=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if not rsstates.eof then 		
			  pcv_StateCode = rsStates("stateCode")
			  pcv_StateOrProvince = rsStates("stateCode")
			End if
			set rsStates = nothing
			call closedb() 
		End if 
		pcv_CountryName= getUserInput(resArray("SHIPTOCOUNTRYNAME"),0)
		pcv_PostalCode= getUserInput(resArray("SHIPTOZIP"),0)

		strEmail=session("Payer")
		strPassword=randomNumber(9999999)
		strPassword=enDeCrypt(strPassword, scCrypPass)
		pCustomerType = 0
		pIdRefer = 0
		pRecvNews = 0
		pcv_strPhoneQuery = ""
		
		if len(pcv_StateCode)>4 then
			pcv_StateCode="" '// Show Province Field, This is not a valid ISO Code
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Set Express Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Open Db
		call opendb()
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
		
		'// Customer Logged into ProductCart
		if session("idCustomer")<>"" and session("idCustomer")<>0 then
			
			query="UPDATE pcCustomerSessions SET idCustomer="&session("idCustomer")&", pcCustSession_ShippingNickName='"& pcf_SanitizeApostrophe(pcv_ShipToName)&"', pcCustSession_ShippingFirstName='"&pcf_SanitizeApostrophe(pcv_FirstName)&"', pcCustSession_ShippingLastName='"&pcf_SanitizeApostrophe(pcv_LastName)&"', pcCustSession_ShippingCompany='"&pcf_SanitizeApostrophe(pcv_PayerBusiness)&"', pcCustSession_ShippingAddress='"&pcf_SanitizeApostrophe(pcv_Street1)&"', pcCustSession_ShippingPostalCode='"&pcf_SanitizeApostrophe(pcv_PostalCode)&"', pcCustSession_ShippingStateCode='"&pcf_SanitizeApostrophe(pcv_StateCode)&"', pcCustSession_ShippingProvince='"&pcf_SanitizeApostrophe(pcv_StateOrProvince)&"', pcCustSession_ShippingPhone='"&pcf_SanitizeApostrophe(pcStrShippingPhone)&"',  pcCustSession_ShippingCity='"&pcf_SanitizeApostrophe(pcv_CityName)&"', pcCustSession_ShippingCountryCode='"&pcf_SanitizeApostrophe(pcv_Country)&"', pcCustSession_ShippingAddress2='"&pcf_SanitizeApostrophe(pcv_Street2)&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
			set rs=server.CreateObject("ADODB.RecordSet")	
		
			set rs=conntemp.execute(query)			
			set rs=nothing
			call closedb()								
			response.redirect "OnePageCheckout.asp"
		
		'// Customer NOT Logged into ProductCart
		else

			'// Check if Customer Exists
			query="SELECT idCustomer, pcCust_Guest FROM customers WHERE email='"&strEmail&"' AND pcCust_Guest=0;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)				
			
			'// Email Does Not Exist - Create New Customer
			if rs.eof then		
			
				pcv_dateCustomerRegistration=Date()
				if SQL_Format="1" then
					pcv_dateCustomerRegistration=Day(pcv_dateCustomerRegistration)&"/"&Month(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
				else
					pcv_dateCustomerRegistration=Month(pcv_dateCustomerRegistration)&"/"&Day(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
				end if
			
				query="INSERT INTO customers (name, lastName, email, [password], city, zip, CountryCode, state, stateCode,shippingcity,shippingZip,shippingCountryCode, shippingState, shippingStateCode, phone, address, shippingAddress, customercompany, customerType, IDRefer, CI1, CI2, address2, shippingCompany, shippingAddress2,RecvNews,pcCust_DateCreated,pcCust_Guest) VALUES ('" &pcf_SanitizeApostrophe(pcv_FirstName)& "', '" &pcf_SanitizeApostrophe(pcv_LastName)& "', '" &pcf_SanitizeApostrophe(strEmail)& "', '" &pcf_SanitizeApostrophe(strPassword)&"','" &pcf_SanitizeApostrophe(pcv_CityName)& "','" &pcf_SanitizeApostrophe(pcv_PostalCode)& "','" &pcf_SanitizeApostrophe(pcv_Country)& "', '"&pcf_SanitizeApostrophe(pcv_StateOrProvince)&"', '" &pcf_SanitizeApostrophe(pcv_StateCode)& "','" &pcf_SanitizeApostrophe(pcv_CityName)& "','" &pcf_SanitizeApostrophe(pcv_PostalCode)& "','" &pcf_SanitizeApostrophe(pcv_Country)& "', '"&pcf_SanitizeApostrophe(pcv_StateOrProvince)&"', '" &pcf_SanitizeApostrophe(pcv_StateCode)& "', '" &pcf_SanitizeApostrophe(pcStrShippingPhone)& "', '" &pcf_SanitizeApostrophe(pcv_Street1)& "', '" &pcf_SanitizeApostrophe(pcv_Street1)& "', '"&pcf_SanitizeApostrophe(pcv_PayerBusiness)&"', " &pcf_SanitizeApostrophe(pCustomerType)& ","&pcf_SanitizeApostrophe(pIdRefer)&",'" &pcf_SanitizeApostrophe(pCI1)& "','" &pcf_SanitizeApostrophe(pCI2)& "', '" &pcf_SanitizeApostrophe(pcv_Street2)& "','','" &pcf_SanitizeApostrophe(pcv_Street2)& "',"&pcf_SanitizeApostrophe(pRecvNews)&",'" & pcf_SanitizeApostrophe(pcv_dateCustomerRegistration) & "',0)"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				
				set rstemp=conntemp.execute(query)	
				set rstemp=nothing

				query="SELECT idCustomer, pcCust_Guest FROM customers WHERE email='"&strEmail&"' AND pcCust_Guest=0 ORDER BY idCustomer DESC;"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)					
				session("idCustomer")=rstemp("idCustomer")	
				session("CustomerGuest")=rstemp("pcCust_Guest")	
				session("isCustomerNew")="YES"				
				set rstemp=nothing				
			
			'// Email Does Exist - Login Customer
			else 					
				intIdCustomer=rs("idCustomer")
				intCustomerGuest=rs("pcCust_Guest")	
				session("idCustomer")=intIdCustomer	
				session("CustomerGuest")=intCustomerGuest			
				set rs=nothing
			end if

		end if	
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Update Customer Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	
		
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Sessions
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		query="UPDATE pcCustomerSessions SET idCustomer="&session("idCustomer")&", pcCustSession_ShippingNickName='"&pcv_ShipToName&"', pcCustSession_ShippingFirstName='"&pcv_FirstName&"', pcCustSession_ShippingLastName='"&pcv_LastName&"', pcCustSession_ShippingCompany='"&pcv_PayerBusiness&"', pcCustSession_ShippingPhone='"&pcStrShippingPhone&"',  pcCustSession_ShippingAddress='"&pcv_Street1&"', pcCustSession_ShippingPostalCode='"&pcv_PostalCode&"', pcCustSession_ShippingStateCode='"&pcv_StateCode&"', pcCustSession_ShippingProvince='"&pcv_StateOrProvince&"', pcCustSession_ShippingCity='"&pcv_CityName&"', pcCustSession_ShippingCountryCode='"&pcv_Country&"', pcCustSession_ShippingAddress2='"&pcv_Street2&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Sessions
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

		set rs=nothing
		call closedb()	


		If session("customerType")=1 Then
			if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
			'Wholesale minimum not met, so customer cannot checkout -> show message
				response.redirect "msg.asp?message=205"
			end if
		Else
			if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then
			'Retail minimum not met, so customer cannot checkout -> show message
				response.redirect "msg.asp?message=206"
			end if
		End If
		
		
		response.redirect "OnePageCheckout.asp"
		
	'// Failed Get Express Details
	Else		
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Error Reporting
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		'// Combine User Friendly Errors from "pcPay_PayPal_Errors.asp"
		'// with Code errors from string "Declined String".
		'// Return a formatted error report as the string "pcv_PayPalErrMessage".
		objPayPalClass.GenerateErrorReport()
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Error Reporting
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~       
		session("ExpressCheckoutPayment")=""
		
	End If	
	'***********************************************************************
	'// End: Get Details from PayPal
	'***********************************************************************
End If
'///////////////////////////////////////////////////////////////////////////////
'// END: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////

function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function

%>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td> 
			<p>&nbsp;</p>
			<div class="pcErrorMessage"><%=pcv_PayPalErrMessage%></div>
			 <p>&nbsp;</p>
			</td>
		</tr>
	</table>
</div>
<!--#include file="footer.asp"-->
