<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/shipFromsettings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/PPConstants.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="header.asp"-->
<% response.Buffer = true %>
<%
Dim pcv_strPayPalmethod
pcv_strPayPalmethod = "PayPalExp"


'******************************************************************
'// PayPal Itemized Order
'// To change this value from the default "non-Itemized Order"
'// you will need to change the variable below to the value of 1.
'//
'// For Example: 
'// pcv_strItemizeOrder = 1

'******************************************************************
'// Set to "non-Itemized Order" by Default
pcv_strItemizeOrder = 1
'******************************************************************


'******************************************************************
'// PayPal Address Override
'// To change this value from the default "no Address Override"
'// you will need to change the variable below to the value of 1.
'//
'// For Example: 
'// pcv_strAddressOverride = 1

'******************************************************************
'// Set to "no Address Override" by Default
pcv_strAddressOverride = 1	
'******************************************************************
'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	if PPD="1" then
		pcStrLogName=Server.Mappath ("/"&scPcFolder&"/includes/ECLog.txt")
	else
		pcStrLogName=Server.Mappath ("../includes/ECLog.txt")
	end if

	'// Create Log of request string and save in PPALog.txt
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set OutputFile = fs.OpenTextFile (pcStrLogName, 8, True)
End If
'//PAYPAL LOGGING END

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
paymentAmount	= pcBillingTotal '// Session("paymentAmount")
paymentType		= Session("PaymentType")

Session("GWTransType")=pcPay_PayPal_TransType

'***********************************************************************
'// Start: Posting Details to PayPal
'***********************************************************************


'SB S
'// By pass PayPal if the immediate order value is 0 
If pcBillingTotal = 0 AND pcIsSubscription Then 	
		
	session("GWAuthCode")	= "AUTH-ARB" 
	session("GWTransId")	= "0" 
	Response.Redirect("gwReturn.asp?s=true&gw=PayPalExp&GWError=1")
	Response.End 
	
Else

	'// Normal Payment, Let Pass

End if 
'SB E

'---------------------------------------------------------------------------
' Construct the parameter string that describes the PayPal payment the varialbes 
' were set in the web form, and the resulting string is stored in nvpstr
'
' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
'---------------------------------------------------------------------------
nvpstr="" '// clear 
objPayPalClass.AddNVP "TOKEN", Token
objPayPalClass.AddNVP "PAYERID", PayerID
objPayPalClass.AddNVP "PAYMENTACTION", PaymentAction
objPayPalClass.AddNVP "AMT", pcf_CurrencyField(money(paymentAmount))
objPayPalClass.AddNVP "CURRENCYCODE", currCodeType
objPayPalClass.AddNVP "BUTTONSOURCE", "ProductCart_Cart_EC_US"
objPayPalClass.AddNVP "INVNUM", session("GWOrderId")

call opendb()

'// Check for Discounts that are not compatible with "Itemization"
query="SELECT orders.discountDetails, orders.pcOrd_CatDiscounts FROM orders WHERE orders.idOrder="&(int(session("GWOrderId"))-scpre)&";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if not rs.eof then
	pcv_strDiscountDetails=rs("discountDetails")
	pcv_CatDiscounts=rs("pcOrd_CatDiscounts")						
end if

set rs=nothing
call closedb()

if pcv_CatDiscounts>0 or trim(pcv_strDiscountDetails)<>"No discounts applied." then
	pcv_strItemizeOrder = 0
end if

IF pcv_strItemizeOrder = 1 THEN

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Itemized Order
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	%>
	<!--#include file="pcPay_PayPal_Itemize.asp"-->
	<%	
	'// PayPal requires two decimal places with a "." decimal separator.
	pcv_strFinalTotal= pcf_CurrencyField(money(pcv_strFinalTotal))
	pcv_strFinalShipCharge= pcf_CurrencyField(money(pcv_strFinalShipCharge))
	pcv_strFinalServiceCharge= pcf_CurrencyField(money(pcv_strFinalServiceCharge))
	pcv_strFinalTax= pcf_CurrencyField(money(pcv_strFinalTax))
	ItemTotal= pcf_CurrencyField(money(ItemTotal))
		
	objPayPalClass.AddNVP "ITEMAMT", ItemTotal
	objPayPalClass.AddNVP "SHIPPINGAMT", pcv_strFinalShipCharge
	objPayPalClass.AddNVP "HANDLINGAMT", pcv_strFinalServiceCharge
	objPayPalClass.AddNVP "TAXAMT", pcv_strFinalTax
	

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Itemized Order
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
	
End If	

'--------------------------------------------------------------------------- 
' Make the call to PayPal to set the Express Checkout token
' If the API call succeded, then redirect the buyer to PayPal
' to begin to authorize payment.  If an error occurred, show the
' resulting errors
'---------------------------------------------------------------------------
'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	OutputFile.WriteLine "Request from ProductCart to PayPal for Express Purchase"
	OutputFile.WriteBlankLines(1)
	OutputFile.WriteLine nvpstr
	OutputFile.WriteBlankLines(1)
End If
'//PAYPAL LOGGING END

Set resArray = objPayPalClass.hash_call("DoExpressCheckoutPayment",nvpstr)
Set Session("nvpResArray")=resArray

'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	OutputFile.WriteLine "Response from PayPal to ProductCart for Express Purchase"
	OutputFile.WriteLine objPayPalHttp.responseText
	OutputFile.WriteBlankLines(2)
	OutputFile.Close
End If
'//PAYPAL LOGGING END

ack = UCase(resArray("ACK"))

if err.number <> 0 then	
	'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
	Dim pcv_PayPalErrMessage
	%><!--#include file="../includes/pcPayPalErrors.asp"--><%
	session("ExpressCheckoutPayment")=""							
end if


If ack="SUCCESS" Then

	TransactionID=resArray("TRANSACTIONID")	
	session("GWTransId")=TransactionID
	session("AVSCode")=resArray("AVSCODE")
	session("CVV2Code")=resArray("CVV2MATCH")  
	
	if session("GWTransId") <> "" then

		'// Save info in pcPay_PayPal_Authorize if "Authorization"			
		If PaymentAction="Authorization" Then
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
		'session("ExpressCheckoutPayment")=""
		
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
	'session("ExpressCheckoutPayment")=""

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
				<div class="pcErrorMessage">
					<%=pcv_PayPalErrMessage%>
            		<br /><br />
                    <a href="OnePageCheckout.asp"><img src="<%=rslayout("back") %>"></a>
            	</div>
			 	<p>&nbsp;</p>
			</td>
		</tr>
	</table>
</div>
<!--#include file="footer.asp"-->
