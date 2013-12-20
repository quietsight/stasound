<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/shipFromsettings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/pcPayPalUKClass.asp"-->
<!--#include file="header.asp"-->
<% response.Buffer = true %>
<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalExp"


'// Set the PayPal Class Obj
set objPayPalClass = New pcPayPalClass

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

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

dim query, rs, conntemp

'///////////////////////////////////////////////////////////////////////////////
'// START: GET DATA FROM DB
 '///////////////////////////////////////////////////////////////////////////////

'// Open Db
call opendb()

'// Declare Local Variables at once
objPayPalClass.pcs_SetAllVariables()
objPayPalClass.pcs_SetShipAddress((int(session("GWOrderId"))-scpre))

'// Close Db
call closedb()

'///////////////////////////////////////////////////////////////////////////////
'// END: GET DATA FROM DB
'///////////////////////////////////////////////////////////////////////////////




'///////////////////////////////////////////////////////////////////////////////
'// START: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////

'// Set our Token
Dim Token
PayerID			= Session("PayerId")
Token			= Session("PayPalExpressToken")
currCodeType	= Session("currencyCodeType")
paymentAmount	= pcBillingTotal
paymentType		= Session("PaymentType")

Session("GWTransType")=pcPay_PayPal_TransType

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
objPayPalClass.AddNVP "PAYERID", PayerID
objPayPalClass.AddNVP "AMT", pcf_CurrencyField(money(paymentAmount))		
objPayPalClass.AddNVP "TENDER", "P" '// C = credit card, P = PayPal
objPayPalClass.AddNVP "ACTION", "D" '// S = Set, G = Get, D = Do
objPayPalClass.AddNVP "TRXTYPE", PaymentAction '// S = Sale transaction, A = Authorisation, C = Credit, D = Delayed Capture, V = Void 
objPayPalClass.AddNVP "CURRENCY", pcPay_PayPal_Currency
objPayPalClass.AddNVP "TOKEN", Token
objPayPalClass.AddNVP "BUTTONSOURCE", "ProductCart_Cart_EC_US"
objPayPalClass.AddNVP "INVNUM", session("GWOrderId")
	
		
'***********************************************************************
'// Start: Address Override
'***********************************************************************	
		if pcv_strShippingStateCode="" OR isNULL(pcv_strShippingStateCode)=True then
			pcv_strShippingStateCode=pcv_strShippingProvince
		end if
		if pcv_strShippingStateCode<>"" AND isNULL(pcv_strShippingStateCode)=False then
			objPayPalClass.AddNVP "SHIPTONAME", pcv_strShippingFullName
			objPayPalClass.AddNVP "SHIPTOSTREET", pcv_strShippingAddress
			objPayPalClass.AddNVP "SHIPTOCITY", pcv_strShippingCity
			objPayPalClass.AddNVP "SHIPTOSTATE", pcv_strShippingStateCode
			objPayPalClass.AddNVP "SHIPTOZIP", pcv_strShippingPostalCode
			objPayPalClass.AddNVP "SHIPTOCOUNTRY", pcv_strShippingCountryCode
		end if
'***********************************************************************
'// End: Address Override
'***********************************************************************	



'--------------------------------------------------------------------------- 
' Make the call to PayPal to set the Express Checkout token
' If the API call succeded, then redirect the buyer to PayPal
' to begin to authorize payment.  If an error occurred, show the
' resulting errors
'---------------------------------------------------------------------------
Set resArray = objPayPalClass.hash_call("DoExpressCheckoutPayment",nvpstr)
Set Session("nvpResArray")=resArray
ack = UCase(resArray("RESULT"))

if err.number <> 0 then	
	'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
	Dim pcv_PayPalErrMessage
	%><!--#include file="../includes/pcPayPalErrors.asp"--><%
	session("ExpressCheckoutPayment")=""							
end if


If ack=0 Then

	TransactionID=resArray("CORRELATIONID")	
	session("GWTransId")=TransactionID
	session("AVSCode")=resArray("AVSADDR")
	session("CVV2Code")=resArray("CVV2MATCH")
			
	if session("GWTransId") <> "" then

		'// Save info in pcPay_PayPal_Authorize if "A"			
		If PaymentAction="A" Then
			call opendb()
			Dim pTodaysDate
			pTodaysDate=Date()
			if SQL_Format="1" then
				pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
			else
				pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
			end if
			if scDB="Access" then
				tmpStr="#"& pTodaysDate &"#"
			else
				tmpStr="'"& pTodaysDate &"'"
			end if
			query="INSERT INTO pcPay_PayPal_Authorize (idOrder, amount, paymentmethod, transtype, authcode, idCustomer, captured, AuthorizedDate, CurrencyCode) VALUES ("&pcGatewayDataIdOrder&", "&pcBillingTotal&", 'PayPalExp', '"&paymentAction&"', '"&session("GWTransId")&"', "&pcIdCustomer&", 0," & tmpStr & ", '"&pcPay_PayPal_Currency&"');"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)				
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			call closedb()
		End If		
					
		response.redirect "gwReturn.asp?s=true&gw=PayPalExp"			
				
	else			
		
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
		
	end if


'// Unsuccessful Express Checkout / Transaction Not Complete
Else	

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Error Reporting
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	'// Combine User Friendly Errors from "pcPay_PayPal_Errors.asp"
	'// with Code errors from string "DeclinedString".
	'// Return a formatted error report as the string "pcv_PayPalErrMessage".
	objPayPalClass.GenerateErrorReport()
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Error Reporting
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
	session("ExpressCheckoutPayment")=""

End If
'///////////////////////////////////////////////////////////////////////////////
'// END: Express Checkout Method
'///////////////////////////////////////////////////////////////////////////////
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
