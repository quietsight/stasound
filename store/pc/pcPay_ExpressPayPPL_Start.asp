<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/PPConstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/pcPayPalPPLClass.asp"-->
<!--#include file="DBsv.asp"-->
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<% response.Buffer = true %>
<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalExp"

session("ExpressPayPPL") = "YES"

'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	if PPD="1" then
		pcStrLogName=Server.Mappath ("/"&scPcFolder&"/includes/PFLECLog.txt")
	else
		pcStrLogName=Server.Mappath ("../includes/PFLECLog.txt")
	end if
	
	'// Create Log of request string and save in PPALog.txt
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set OutputFile = fs.OpenTextFile (pcStrLogName, 8, True)
End If
'//PAYPAL LOGGING END

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
session("pcPay_PayPalExp_OrderTotal")=calculateCartTotal(pcCartArray, ppcCartIndex)

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
returnURL	= url & "pcPay_ExpressPayPPL_Start.asp?currencyCodeType=" &  currencyCodeType & "&paymentAmount=" & OrderTotal & "&paymentType=" &PaymentAction 
lenReturnURL = len(returnURL)
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
If  Request.QueryString("token")&""="" Then
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

		if logoURL<>"" then
		'	objPayPalClass.AddNVP "HDRIMG", logoURL
		end if

		
		iTmpEcCount = 0					
		iTmpEcItemTotal = 0
		for ec=1 to ppcCartIndex
			tmpEcName = pcCartArray(ec,1)
			tmpEcQty = pcCartArray(ec,2)
			tmpEcUnitCost = Ccur(pcCartArray(ec,3))
			tmpEcOptions = pcCartArray(ec,5)
			tmpEcSku = pcCartArray(ec,7)
			If not isNumeric(tmpEcOptions) Then
				tmpEcOptions = Ccur(0)
			Else
				tmpEcOptions = Ccur(tmpEcOptions)
			End If
			tmpUnitTotal =  tmpEcUnitCost + tmpEcOptions
			iTmpEcItemTotal = iTmpEcItemTotal + (tmpUnitTotal * tmpEcQty)
			tmpEcName = replace(tmpEcName,"&quot;","")
			if instr(tmpEcName, "&") then
				LDescStr = "L_DESC["&len(tmpEcName)&"]"
			else
				LDescStr = "L_DESC"
			end if
			objPayPalClass.AddNVP LDescStr&iTmpEcCount, tmpEcName
			if instr(tmpEcName, "&") then
				LNameStr = "L_NAME["&len(tmpEcName)&"]"
			else
				LNameStr = "L_NAME"
			end if
			objPayPalClass.AddNVP LNameStr&iTmpEcCount, tmpEcName
			objPayPalClass.AddNVP "L_QTY"&iTmpEcCount, tmpEcQty				
			objPayPalClass.AddNVP "L_COST"&iTmpEcCount, tmpUnitTotal
			iTmpEcCount = iTmpEcCount + 1					
		Next
		
		objPayPalClass.AddNVP "ITEMAMT", iTmpEcItemTotal
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Itemized Order
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
		if lcase(request.QueryString("refer")) = "viewcart.asp" OR lcase(request.QueryString("refer")) = "onepagecheckout.asp" then
		else
			call opendb()
			
			query="SELECT pcCustSession_ShippingFirstName, pcCustSession_ShippingLastName, pcCustSession_ShippingCompany, pcCustSession_ShippingAddress, pcCustSession_ShippingPostalCode, pcCustSession_ShippingStateCode, pcCustSession_ShippingProvince, pcCustSession_ShippingPhone,  pcCustSession_ShippingCity, pcCustSession_ShippingCountryCode, pcCustSession_ShippingAddress2 FROM pcCustomerSessions WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
	
			pcCustSession_ShippingFirstName = rs("pcCustSession_ShippingFirstName")
			pcCustSession_ShippingLastName = rs("pcCustSession_ShippingLastName")
			pcCustSession_ShippingCompany = rs("pcCustSession_ShippingCompany")
			pcv_strShippingAddress = rs("pcCustSession_ShippingAddress")
			pcv_strShippingPostalCode = rs("pcCustSession_ShippingPostalCode")
			pcv_strShippingStateCode = rs("pcCustSession_ShippingStateCode")
			pcv_strShippingProvince = rs("pcCustSession_ShippingProvince")
			pcv_strShippingPhone = rs("pcCustSession_ShippingPhone")
			pcv_strShippingCity = rs("pcCustSession_ShippingCity")
			pcv_strShippingCountryCode = rs("pcCustSession_ShippingCountryCode")
			pcv_strShippingAddress2  = rs("pcCustSession_ShippingAddress2")
			pcv_strShippingFullName = pcCustSession_ShippingFirstName & " "&pcCustSession_ShippingLastName
			set rs=nothing
			call closedb()								
	
			if pcv_strShippingStateCode="" OR isNULL(pcv_strShippingStateCode)=True then
				pcv_strShippingStateCode=pcv_strShippingProvince
			end if
			if pcv_strShippingStateCode<>"" AND isNULL(pcv_strShippingStateCode)=False then
				objPayPalClass.AddNVP "SHIPTONAME", pcv_strShippingFullName
				objPayPalClass.AddNVP "SHIPTOSTREET", pcv_strShippingAddress
				objPayPalClass.AddNVP "SHIPTOCITY", pcv_strShippingCity
				objPayPalClass.AddNVP "SHIPTOSTATE", pcv_strShippingStateCode
				objPayPalClass.AddNVP "SHIPTOZIP", pcv_strShippingPostalCode
				objPayPalClass.AddNVP "SHIPTOCOUNTRYCODE", pcv_strShippingCountryCode
				objPayPalClass.AddNVP "SHIPTOSTREET2", pcv_strShippingAddress2
				objPayPalClass.AddNVP "SHIPTOPHONENUM", pcv_strShippingPhone
			end if
		end if
		'//PAYPAL LOGGING START
		If scPPLogging = "1" Then
			OutputFile.WriteLine now()
			OutputFile.WriteLine "_______________________________________________________________________"
			OutputFile.WriteBlankLines(1)
		End If
		'//PAYPAL LOGGING END
		'--------------------------------------------------------------------------- 
		' Make the call to PayPal to set the Express Checkout token
		' If the API call succeded, then redirect the buyer to PayPal
		' to begin to authorize payment.  If an error occurred, show the
		' resulting errors
		'---------------------------------------------------------------------------
		Set resArray = objPayPalClass.hash_call("SetExpressCheckout",nvpstr)
		Set Session("nvpResArray")=resArray

		'//PAYPAL LOGGING START
		If scPPLogging = "1" Then
			OutputFile.WriteLine "Response from PayPal to ProductCart for Payflow Link Express Purchase"
			OutputFile.WriteLine objPayPalHttp.responseText
			OutputFile.WriteBlankLines(2)
			OutputFile.Close
		End If
		'//PAYPAL LOGGING END

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
				pcv_PayPalPPA_URL = "https://www.sandbox.paypal.com/cgibin/webscr?cmd=_express-checkout&token="& token
				response.write pcv_PayPalPPA_URL
				response.write "<BR><BR>"
				'response.end
				objPayPalClass.ReDirectURL(pcv_PayPalPPA_URL)
				
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
	objPayPalClass.AddNVP "ACTION", "G" '// S = Set, G = Get, D = Do
	'---------------------------------------------------------------------------
	' Make the API call and store the results in an array.  If the
	' call was a success, show the authorization details, and provide
	' an action to complete the payment.  If failed, show the error
	'---------------------------------------------------------------------------
	ack=""
	
	Set resArray = objPayPalClass.hash_call("GetExpressCheckoutDetails",nvpstr)
	'//PAYPAL LOGGING START
	If scPPLogging = "1" Then
		OutputFile.WriteBlankLines(1)
		OutputFile.WriteLine "Response from PayPal to ProductCart for Payflow Link Express Purchase"
		OutputFile.WriteLine objPayPalHttp.responseText
		OutputFile.WriteBlankLines(2)
		OutputFile.Close
	End If
	'//PAYPAL LOGGING END

	ack = UCase(resArray("RESPMSG"))
	Set Session("nvpResArray")=resArray
	
	'// Successful Get Express Details
	If ack="APPROVED" Then
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Set Express Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
		pcStrShippingPhone=""
		pcv_Payer= getUserInput(resArray("EMAIL"),0)
		session("Payer")=pcv_Payer
		pcv_PayerID= getUserInput(resArray("PAYERID"),0)
		session("PayerId")=pcv_PayerID
		pcv_PayerStatus= getUserInput(resArray("PAYERSTATUS"),0)
		pcv_PayerBusiness= getUserInput(resArray("BUSINESS"),0)	'?
		pcv_FirstName= getUserInput(resArray("FIRSTNAME"),0)
		pcv_LastName= getUserInput(resArray("LASTNAME")	,0)	
		pcv_FullName= pcv_FirstName & " " & pcv_LastName	
		pcv_ShipToName =  getUserInput(resArray("SHIPTONAME"),0)
		pcv_ShipToBusiness =  getUserInput(resArray("SHIPTOBUSINESS"),0)
		pcv_Street1= getUserInput(resArray("SHIPTOSTREET"),0)
		pcv_Street2= getUserInput(resArray("SHIPTOSTREET2"),0) '?
		pcv_CityName= getUserInput(resArray("SHIPTOCITY"),0)
		pcv_StateOrProvince= getUserInput(resArray("SHIPTOSTATE"),0)
		pcv_StateCode= getUserInput(resArray("SHIPTOSTATE"),0)
		pcv_Country=getUserInput(resArray("SHIPTOCOUNTRY"),0)
		pcv_CountryName= getUserInput(resArray("SHIPTOCOUNTRYNAME"),0) '?
		pcv_PostalCode= getUserInput(resArray("SHIPTOZIP"),0)

		session("ppec_shipto_Name") = pcv_ShipToName
		session("ppec_shipto_Business") = pcv_ShipToBusiness
		session("ppec_shipto_Street1") = pcv_Street1
		session("ppec_shipto_Street2") = pcv_Street2
		session("ppec_shipto_City") = pcv_CityName
		session("ppec_shipto_StateCode") = pcv_StateCode
		session("ppec_shipto_Province") = pcv_StateOrProvince
		session("ppec_shipto_Country") = pcv_Country
		session("ppec_shipto_PostalCode") = pcv_PostalCode
		session("ppec_shipto_Phone") = pcStrShippingPhone
		session("ppec_shipto_Email") = pcv_Payer
		
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
		
		session("PPSA")="0"
		session("PPSAID") = ""
		If session("ppec_shipto_Name")&""<>"" then
			shipToNameArry = split(session("ppec_shipto_Name"), " ")
			shipToFirstNameTmp = shipToNameArry(0)
			shipToLastNameTmp = shipToNameArry(1)
			query="SELECT idRecipient, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Email, recipient_Phone, recipient_Fax, recipient_Company, recipient_Address, recipient_Address2, recipient_City, recipient_State, recipient_StateCode, recipient_Zip, recipient_CountryCode, Recipient_Residential FROM recipients WHERE recipient_NickName='PayPal Shipping Address' AND idCustomer="&session("idCustomer")&";"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)				
			If rs.eof then
				query = "INSERT INTO recipients (idCustomer, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Email, recipient_Phone, recipient_Fax, recipient_Company, recipient_Address, recipient_Address2, recipient_City, recipient_State, recipient_StateCode, recipient_Zip, recipient_CountryCode) VALUES ("&session("idCustomer")&",'PayPal Shipping Address', '"&shipToFirstNameTmp&"', '"&shipToLastNameTmp&"', '"&session("ppec_shipto_Email")&"', '"&session("ppec_shipto_Phone")&"', '', '"&session("ppec_shipto_Business")&"', '"&session("ppec_shipto_Street1")&"', '"&session("ppec_shipto_Street2")&"', '"&session("ppec_shipto_City")&"', '"&session("ppec_shipto_Province")&"', '"&session("ppec_shipto_StateCode")&"', '"&session("ppec_shipto_PostalCode")&"', '"&session("ppec_shipto_Country")&"');"
			Else
				query="UPDATE recipients SET recipient_FirstName='"&shipToFirstNameTmp&"', recipient_LastName='"&shipToLastNameTmp&"', recipient_Email='"&session("ppec_shipto_Email")&"', recipient_Phone='"&session("ppec_shipto_Phone")&"', recipient_Fax='', recipient_Company='"&session("ppec_shipto_Business")&"', recipient_Address='"&session("ppec_shipto_Street1")&"', recipient_Address2='"&session("ppec_shipto_Street2")&"', recipient_City='"&session("ppec_shipto_City")&"', recipient_State='"&session("ppec_shipto_Province")&"', recipient_StateCode='"&session("ppec_shipto_StateCode")&"', recipient_Zip='"&session("ppec_shipto_PostalCode")&"', recipient_CountryCode='"&session("ppec_shipto_Country")&"' WHERE recipient_NickName='PayPal Shipping Address' AND idCustomer="&session("idCustomer")&";"
			End If
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			query="SELECT idRecipient FROM recipients WHERE recipient_NickName='PayPal Shipping Address' AND idCustomer="&session("idCustomer")&";"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)				
			if NOT rs.eof then
				session("PPSA")="1"
				session("PPSAID") = rs("idRecipient")
			end if
		End If
		
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
