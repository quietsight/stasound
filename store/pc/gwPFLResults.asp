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
<!--#include file="../includes/PPConstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<%
dim rs, conntemp, query

Dim PFLURL
If scSSL="" OR scSSL="0" Then
	PFLURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	PFLURL=replace(PFLURL,"https:/","https://")
	PFLURL=replace(PFLURL,"http:/","http://")
Else
	PFLURL=replace((scSslURL&"/"&scPcFolder&"/pc/"),"//","/")
	PFLURL=replace(PFLURL,"https:/","https://")
	PFLURL=replace(PFLURL,"http:/","http://")
End If

'//PAYPAL LOGGING START
If scPPLogging = "1" Then
	if PPD="1" then
		pcStrLogName=Server.Mappath ("/"&scPcFolder&"/includes/PFLLog.txt")
	else
		pcStrLogName=Server.Mappath ("../includes/PFLLog.txt")
	end if
	
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set OutputFile = fs.OpenTextFile (pcStrLogName, 8, True)
	
	OutputFile.WriteLine now()
	OutputFile.WriteLine "Referrer:       " + Request.ServerVariables("HTTP_REFERER")
	OutputFile.WriteLine "Remote Address: " + Request.ServerVariables("REMOTE_ADDR")
	OutputFile.WriteLine "Content-Type:   " + Request.ServerVariables("CONTENT_TYPE")
	OutputFile.WriteLine "User-Agent:     " + Request.ServerVariables("HTTP_USER_AGENT")
	OutputFile.WriteBlankLines(2)
	OutputFile.WriteLine "All Server Variables:"
	OutputFile.WriteLine Request.ServerVariables("ALL_RAW")   
	OutputFile.WriteLine "Raw Posted Data: " & Request.Form
	OutputFile.WriteLine "Raw QueryString Data: " & Request.QueryString
End If
'//PAYPAL LOGGING END 

'//Check for a response
pfl_avszip = request("AVSZIP")
pfl_ppref = request("PPREF")
pfl_transactiontime = request("TRANSTIME")
pfl_ziptoship = request("ZIPTOSHIP")
pfl_lastname = request("LASTNAME")
pfl_pnref = request("PNREF")
pfl_avsdata = request("AVSDATA")
pfl_type = request("TYPE")
pfl_citytoship = request("CITYTOSHIP")
pfl_payerid = request("PAYERID")
pfl_tender = request("TENDER")
pfl_pendingreason = request("PENDINGREASON")
pfl_token = request("TOKEN")
pfl_method = request("METHOD")
pfl_avsaddr = request("AVSADDR")
pfl_addresstoship = request("ADDRESSTOSHIP")
pfl_securetoken = request("SECURETOKEN")
pfl_securetokenid = request("SECURETOKENID")
pfl_responsemessage = request("RESPMSG")
pfl_firstname = request("FIRSTNAME")
pfl_correlationid = request("CORRELATIONID")
pfl_countrytoship = request("COUNTRYTOSHIP")
pfl_statetoship = request("STATETOSHIP")
if request("RESULT")&""<>"" then
	pfl_result = request("RESULT")
else
	pfl_result="NONE"
end if
pfl_cancelflag = request("cancel_ec_trans")
pfl_prefpsmsg = request("PREFPSMSG")
pfl_hostcode = request("HOSTCODE")
pfl_invoice = request("INVOICE")
PFLpcOrderId=cLng(session("GWOrderId"))-cLng(scPre)
if clng(PFLpcOrderId)<0 then
	PFLpcOrderId = pfl_invoice
	session("GWOrderId") = PFLpcOrderId
end if
pfl_postfpsmsg = request("POSTFPSMSG") 'Review
pfl_acct = request("ACCT") '7930
pfl_proccvv2 = request("PROCCVV2") 'M
pfl_cvv2match = request("CVV2MATCH") 'Y
pfl_email = request("EMAIL") 'spark@earlyimpact.com
pfl_phone = request("PHONE") '1231231231
pfl_amt = request("AMT") '70.04
pfl_zip = request("ZIP") '92506
pfl_authcode = request("AUTHCODE") '111111
pfl_expdate = request("EXPDATE") '1017
pfl_iavs = request("IAVS") 'N
pfl_tax = request("TAX") '0.00
pfl_cardtype = request("CARDTYPE") '0
pfl_procavs = request("PROCAVS") 'X
pfl_prefpsmsg = request("PREFPSMSG") 'Review%3A+More+than+one+rule+was+triggered+for+Review
pfl_invnum = request("INVNUM") '29

if pfl_securetokenid&""<>"" then 
	call opendb()
	
	query = "SELECT pcPay_PayPal_Signature FROM orders WHERE idOrder="& PFLpcOrderId
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	pcPay_PayPal_SecureTokenId = rs("pcPay_PayPal_Signature")
	
	set rs=nothing
	call closedb()
	%>
    <html><body><p><center>Your payment is currently being processed.<br />It can take up to 2 minutes to complete.</center></p>
	<%
	'//PAYPAL LOGGING START
	If scPPLogging = "1" Then
		OutputFile.WriteBlankLines(1)
		OutputFile.WriteLine "pcPay_PayPal_SecureTokenId: "&pcPay_PayPal_SecureTokenId
		OutputFile.WriteLine "pfl_securetokenid: "&pfl_securetokenid
	End If
	'//PAYPAL LOGGING END
	
	if pcPay_PayPal_SecureTokenId <> pfl_securetokenid then
		pfl_message = "Invalid Secure Token!"
		%>
		<script language="javascript" type="text/javascript">window.parent.location.href='gwReturn.asp?s=true&gw=PayPalAdvanced&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwpflEB.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if
	
	if lcase(ppa_cancelflag) = "true" then
		'Customer cancelled payment."
		pfl_message = "Result: Customer Canceled Payment"
		
		'//PAYPAL LOGGING START
		If scPPLogging = "1" Then
			OutputFile.WriteBlankLines(1)
			OutputFile.WriteLine "Result: " & pfl_message
			OutputFile.WriteBlankLines(2)
			OutputFile.Close
			Set fs = nothing
		End If
		'//PAYPAL LOGGING END
		
		err.number=0
		err.clear
		%>
		<script language="javascript" type="text/javascript">window.parent.location.href='<%=PFLURL&"gwpflEB.asp?Message="&pfl_message%>';</script>
        <noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwpflEB.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	end if

	if cstr(pfl_result) = "0" then
		
		call opendb()
		'//Update the customer's shipping address for this order
		query = "Update orders set shippingAddress='"&pfl_addresstoship&"', shippingStateCode='"&pfl_statetoship&"', shippingCity='"&pfl_citytoship&"', shippingCountryCode='"&pfl_countrytoship&"', shippingZip='"&pfl_ziptoship&"', ShippingFullName='"&pfl_firstname &" "&pfl_lastname&"' WHERE idOrder="& PFLpcOrderId
		'//PAYPAL LOGGING START
		If scPPLogging = "1" Then
			OutputFile.WriteLine query
		End If
		'//PAYPAL LOGGING END
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		'//Auth-only orders we will save in a new table
		query = "INSERT INTO pcPay_PFL_Authorize (idOrder, orderDate,paySource, amount, paymentmethod, transtype, authcode,  captured, fraudcode) VALUES ("& PFLpcOrderId &", '"&date()&"', 'PFL', "&pfl_amt&", '"&pfl_method&"', '"&pfl_type&"', '"&pfl_pnref&"',0, '"&ppa_result&"');"
		'//PAYPAL LOGGING START
		If scPPLogging = "1" Then
			OutputFile.WriteLine query
		End If
		'//PAYPAL LOGGING END
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		set rs=nothing
		call closedb()
		'//gwReturn.asp
		session("GWAuthCode")=pfl_pnref
		session("GWTransId")=pfl_ppref
		session("GWTransType")=pfl_type
		session("AVSCode")=pfl_avsdata
		
		'//PAYPAL LOGGING START
		If scPPLogging = "1" Then
			OutputFile.WriteLine "Let's redirect!"
			OutputFile.WriteBlankLines(2)
			OutputFile.Close
			Set fs = nothing
		End If
		'//PAYPAL LOGGING END
		
		fraudmode = ""
		if cstr(pfl_result) = "126" then
			fraudmode = "review"
			Session("FraudCode") = fraudmode
		end if
		
		err.number=0
		err.clear
		
		session("GWTransType")=pfl_type

		%>
		<script language="javascript" type="text/javascript">window.parent.location.href='gwReturn.asp?s=true&gw=PFLink&fraudmode=<%=fraudmode%>';</script>
		<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwpflEB.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
		<%
		response.End()
	else
		if lcase(pfl_cancelflag) = "true" then
			'Customer cancelled payment."
			'//PAYPAL LOGGING START
			If scPPLogging = "1" Then
				OutputFile.WriteLine "Line 222"
				OutputFile.WriteBlankLines(1)
				pfl_message = "Result: Customer Canceled Payment"
				OutputFile.WriteLine "Result: " & pfl_message
				OutputFile.WriteBlankLines(2)
				OutputFile.Close
				Set fs = nothing
			End If
			'//PAYPAL LOGGING END
			err.number=0
			err.clear
			%>
			<script language="javascript" type="text/javascript">window.parent.location.href='<%=PFLURL&"gwpflEB.asp?Message="&pfl_message%>';</script>
        	<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=PFLURL&"gwpflEB.asp?Message="&pfl_message%>">here </a>to continue.</noscript>
			<%
			response.End()
		else
			pfl_message = "The payment could not be completed for the following reasons<br><ul><li>" & pfl_responsemessage &"</li></ul>"
			session("pfl_message") = pfl_message
			
			'//PAYPAL LOGGING START
			If scPPLogging = "1" Then
				OutputFile.WriteLine "Line 244"
				OutputFile.WriteBlankLines(1)
				OutputFile.WriteLine "Result Code: " & pfl_result
				OutputFile.WriteLine "Message Displayed: " & pfl_message
				OutputFile.WriteLine "Let's redirect: "&RedirectURLA
				OutputFile.WriteBlankLines(2)
				OutputFile.Close
				Set fs = nothing
			End If
			'//PAYPAL LOGGING END
			
			RedirectURLA = PFLURL&"gwpflEB.asp?Message=session"
			err.number=0
			err.clear
			%>
			<script language="javascript" type="text/javascript">window.parent.location.href='<%=RedirectURLA%>';</script>
        	<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=RedirectURLA%>">here </a>to continue.</noscript>
			<%
			response.End()
		end if
	end if
else
	'RESPMSG=Parameter+list+format+error%3A&RESULT=-6
	if pfl_prefpsmsg&""<>"" Then
		pfl_responsemessage = pfl_prefpsmsg
	end if
	pfl_message = "The payment could not be completed for the following reasons<br><ul><li>" & pfl_responsemessage &"</li></ul>"
	session("pfl_message") = pfl_message
	RedirectURLA = PPAURL&"gwpflEB.asp?Message=session"
	
	'//PAYPAL LOGGING START
	If scPPLogging = "1" Then
		OutputFile.WriteLine "Line 276"
		OutputFile.WriteBlankLines(1)
		OutputFile.WriteLine "Result Code: " & pfl_result
		OutputFile.WriteLine "Message Displayed: " & pfl_responsemessage
		OutputFile.WriteLine "Let's redirect: "&RedirectURLA
		OutputFile.WriteBlankLines(2)
		OutputFile.Close
		Set fs = nothing
	End If
	'//PAYPAL LOGGING END
	err.number=0
	err.clear
	%>
	<script language="javascript" type="text/javascript">window.parent.location.href='<%=RedirectURLA%>';</script>
	<noscript>Your browser does not have JavaScript enabled. Please click <a href="<%=RedirectURLA%>">here </a>to continue.</noscript>
	<%
	response.End()
end if %>
</body></html>