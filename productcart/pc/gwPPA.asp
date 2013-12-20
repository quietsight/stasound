<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.CacheControl = "No-Store"
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/PPConstants.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<%
dim connTemp, rs
call openDb()

'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	if PPD="1" then
		pcStrLogName=Server.Mappath ("/"&scPcFolder&"/includes/PPALog.txt")
	else
		pcStrLogName=Server.Mappath ("../includes/PPALog.txt")
	end if
	
	'// Create Log of request string and save in PPALog.txt
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set OutputFile = fs.OpenTextFile (pcStrLogName, 8, True)
End If
'//PAYPAL LOGGING END

Dim PPAURL
If scSSL="" OR scSSL="0" Then
	PPAURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	PPAURL=replace(PPAURL,"https:/","https://")
	PPAURL=replace(PPAURL,"http:/","http://")
Else
	PPAURL=replace((scSslURL&"/"&scPcFolder&"/pc/"),"//","/")
	PPAURL=replace(PPAURL,"https:/","https://")
	PPAURL=replace(PPAURL,"http:/","http://")
End If

'//Set redirect page to the current file name
session("redirectPage")="gwPPA.asp"

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
	session("GWOrderId")=getUserInput(request("idOrder"),0)
end if

'//Retrieve customer data from the database using the current session id
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

call opendb()
'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT pcPay_PayPalAd_Partner, pcPay_PayPalAd_MerchantLogin, pcPay_PayPalAd_Vendor, pcPay_PayPalAd_User, pcPay_PayPalAd_Password, pcPay_PayPalAd_TransType, pcPay_PayPalAd_CSC, pcPay_PayPalAd_Sandbox FROM pcPay_PayPalAdvanced WHERE pcPay_PayPalAd_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_PayPalAd_Partner=rs("pcPay_PayPalAd_Partner")				
pcPay_PayPalAd_MerchantLogin=rs("pcPay_PayPalAd_MerchantLogin")	
pcPay_PayPalAd_MerchantLogin=enDeCrypt(pcPay_PayPalAd_MerchantLogin, scCrypPass)					
pcPay_PayPalAd_Vendor=rs("pcPay_PayPalAd_Vendor")	
pcPay_PayPalAd_Vendor=enDeCrypt(pcPay_PayPalAd_Vendor, scCrypPass)					
pcPay_PayPalAd_User=rs("pcPay_PayPalAd_User")
pcPay_PayPalAd_User=enDeCrypt(pcPay_PayPalAd_User, scCrypPass)					
pcPay_PayPalAd_Password=rs("pcPay_PayPalAd_Password")
pcPay_PayPalAd_Password=enDeCrypt(pcPay_PayPalAd_Password, scCrypPass)					
pcPay_PayPalAd_TransType=rs("pcPay_PayPalAd_TransType")			
pcPay_PayPalAd_CSC=rs("pcPay_PayPalAd_CSC")
pcPay_PayPalAd_Sandbox=rs("pcPay_PayPalAd_Sandbox")				

if ucase(pcPay_PayPalAd_Sandbox) = "YES" then
   SEcureTokenGatewayPayPalAdvancedURL="https://pilot-payflowpro.paypal.com" 'for testing
else
   SEcureTokenGatewayPayPalAdvancedURL="https://payflowpro.paypal.com" ' production
end if

newSecureTokenID = genrandomvalue(36)

'// ALTER FUNCTION AFTER TESTING  
function genrandomvalue(passwordLength)
   Dim sDefaultChars
   Dim iCounter
   Dim sMyPassword
   Dim iPickedChar
   Dim iDefaultCharactersLength
   Dim iPasswordLength

   sDefaultChars="abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
   iPasswordLength=passwordLength
   iDefaultCharactersLength = Len(sDefaultChars) 

   Randomize
   For iCounter = 1 To iPasswordLength
      iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1) 
      sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
   Next 
   genrandomvalue = sMyPassword
end function

'//SAVE TOKEN TO ORDER
query = "UPDATE orders SET pcPay_PayPal_Signature = '"&newSecureTokenID&"' WHERE idOrder="& pcGatewayDataIdOrder
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

set rs=nothing
call closedb()

if pcShippingFullName&""="" then
	pcShippingFullName = pcBillingFirstName &" "&pcBillingLastName
end if
if pcShippingAddress&""="" then
	pcShippingAddress = pcBillingAddress
	pcShippingCity = pcBillingCity
	pcShippingStateCode = pcBillingStateCode
	pcShippingPostalCode = pcBillingPostalCode
	pcShippingCountryCode = pcBillingCountryCode
	pcShippingAddress2 = pcBillingAddress2
	pcShippingPhone = pcBillingPhone
end if

stext = "USER="&trim(pcPay_PayPalAd_User)
stext = stext & "&VENDOR="&trim(pcPay_PayPalAd_MerchantLogin)
stext = stext & "&PARTNER="&pcPay_PayPalAd_Partner
stext = stext & "&PWD="&pcPay_PayPalAd_Password
stext = stext & "&TRXTYPE="&pcPay_PayPalAd_TransType
stext = stext & "&CREATESECURETOKEN=Y"
stext = stext & "&RETURNURL="&PPAURL&"gwPPAResults.asp"
stext = stext & "&CANCELURL="&PPAURL&"gwPPAResults.asp"
stext = stext & "&ERRORURL="&PPAURL&"gwPPAResults.asp"
stext = stext & "&URLMETHOD=POST"
stext = stext & "&SILENTPOST=False"
stext = stext & "&TEMPLATE=MINLAYOUT"
stext = stext & "&SECURETOKENID="&newSecureTokenID
stext = stext & "&INVNUM="&pcGatewayDataIdOrder
stext = stext & "&AMT="&pcBillingTotal
stext = stext & "&BILLTOFIRSTNAME="&pcBillingFirstName
stext = stext & "&BILLTOLASTNAME="&pcBillingLastName
stext = stext & "&BILLTOSTREET="&pcBillingAddress
stext = stext & "&BILLTOSTREET2="&pcBillingAddress2
stext = stext & "&BILLTOCITY="&pcBillingCity
stext = stext & "&BILLTOSTATE="&pcBillingStateCode
stext = stext & "&BILLTOZIP="&pcBillingPostalCode
stext = stext & "&BILLTOPHONENUM="&pcBillingPhone
stext = stext & "&EMAIL="&pcCustomerEmail
stext = stext & "&DISABLERECEIPT=TRUE"
stext = stext & "&ADDROVERRIDE=1"
stext = stext & "&SHIPTONAME="&pcShippingFullName
stext = stext & "&SHIPTOSTREET="&pcShippingAddress
stext = stext & "&SHIPTOCITY="&pcShippingCity
stext = stext & "&SHIPTOSTATE="&pcShippingStateCode
stext = stext & "&SHIPTOZIP="&pcShippingPostalCode
stext = stext & "&SHIPTOCOUNTRYCODE="&pcShippingCountryCode
stext = stext & "&SHIPTOSTREET2="&pcShippingAddress2
stext = stext & "&SHIPTOPHONENUM="&pcShippingPhone
stext = stext & "&BUTTONSOURCE=ProductCart_Cart_PPA"
%>
<!--#include file="pcPay_PPA_Itemize.asp"-->
<%	
'// PayPal requires two decimal places with a "." decimal separator.
pcv_strFinalTotal= pcf_CurrencyField(money(pcv_strFinalTotal))
pcv_strFinalShipCharge= pcf_CurrencyField(money(pcv_strFinalShipCharge))
pcv_strFinalServiceCharge= pcf_CurrencyField(money(pcv_strFinalServiceCharge))
pcv_strFinalTax= pcf_CurrencyField(money(pcv_strFinalTax))
ItemTotal= pcf_CurrencyField(money(ItemTotal))
'Temp unitl PayPal fixes this issue
If iQDString&""="" Then
	stext = stext & iString
	stext = stext & iQDString
	stext = stext & iPCString
	stext = stext & iCDString
	stext = stext & iGWString
End If
stext = stext & "&ITEMAMT="&ItemTotal
stext = stext & "&FREIGHTAMT="&pcv_strFinalShipCharge
stext = stext & "&HANDLINGAMT="&pcv_strFinalServiceCharge
stext = stext & "&TAXAMT="&pcv_strFinalTax

'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	OutputFile.WriteLine now()
	OutputFile.WriteLine "__________________________________________________________________"
	OutputFile.WriteLine "Request from ProductCart to PayPal for PayPal Advanced Payments"
	OutputFile.WriteBlankLines(1)
	OutputFile.WriteLine stext
	OutputFile.WriteBlankLines(1)
End If
'//PAYPAL LOGGING END

'Send the transaction info as part of the querystring
set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
'SB S
if ucase(pcPay_PayPalAd_Sandbox) = "YES" then
	xml.open "POST", "https://pilot-payflowpro.paypal.com", false
else
	xml.open "POST", "https://payflowpro.paypal.com", false
end if
'SB E

xml.Send stext
strStatus = xml.Status

'store the response
strRetVal = xml.responseText
Set xml = Nothing

'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	OutputFile.WriteLine "Response from PayPal to ProductCart for PayPal Advanced Payments"
	OutputFile.WriteLine strRetVal
	OutputFile.WriteBlankLines(1)
End If
'//PAYPAL LOGGING END

split_resultXML = split(strRetVal,"&")
j=0
for each item in split_resultXML
  split_param = split(split_resultXML(j),"=")
  formname = split_param(0)
  formvalue = split_param(1)
  if ucase(formname)  = "RESULT" then resultcode_pymt = formvalue
  if ucase(formname)  = "RESPMSG" then resultvalue_pymt = formvalue
  if ucase(formname)  = "SECURETOKEN" or ucase(formname) = "SECURETOKENID" then
	 if trim(iframeString) = "" then
		iframeString = formname &"="& formvalue
	 else
		iframeString = iframeString & "&"& formname &"="& formvalue
	 end if
  end if
  j = j + 1	  
next
dim ppmode
ppmode = ""
if ucase(pcPay_PayPalAd_Sandbox) = "YES" then ppmode = "MODE=TEST&"

'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	OutputFile.WriteLine "iFrame generated for PayPal Advanced Payments"
	OutputFile.WriteLine "https://payflowlink.paypal.com/?"&ppmode&iframeString
	OutputFile.WriteBlankLines(2)
	OutputFile.Close
	Set fs = nothing
End If
'//PAYPAL LOGGING END
	
err.number=0
err.clear

%>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
			<img src="images/checkout_bar_step5.gif" alt="">
			</td>
		</tr>
		<% if request("Message")&""<>"" then
			myMsg = getUserInput(request("Message"),0)
			
			if lcase(myMsg)="session" then
				myMsg = session("ppa_message")
			end if %>
            <tr valign="top"> 
                <td colspan="2">
                    <div class="pcErrorMessage"><%=myMsg%></div>
                </td>
            </tr>
		<% end if %>
		<tr>
			<td>
			<center><iframe src="https://payflowlink.paypal.com/?<%=ppmode%><%=iframeString%>" scrolling="no" frameborder="0" width="490" height="565" name="PPAFrame" ></iframe></center>
			</td>
		</tr>
	</table>
</div>
<!--#include file="footer.asp"-->