<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'=====================================================================================
'= Cardinal Commerce (http://www.cardinalcommerce.com)
'= pcPay_Cent_XMLFunctions.asp
'= General Functions used to perform a XML Integration
'= Version 8.0
'===================================================================================== %>
<% dim pcPay_Cent_MessageVersion, pcPay_Cent_TransactionURL, pcPay_Cent_MerchantID, pcPay_Cent_ProcessorId, pcPay_Cent_TermURL

call opendb()

if gateway="WPP" then

	query="SELECT pcPay_Centinel.pcPay_Cent_TransactionURL, pcPay_Centinel.pcPay_Cent_ProcessorId, pcPay_Centinel.pcPay_Cent_MerchantID, pcPay_Centinel.pcPay_Cent_Active, pcPay_Centinel.pcPay_Cent_Password FROM pcPay_Centinel WHERE (((pcPay_Centinel.pcPay_Cent_ID)=1));"

else

	query="SELECT pcPay_Centinel.pcPay_Cent_TransactionURL, pcPay_Centinel.pcPay_Cent_ProcessorId, pcPay_Centinel.pcPay_Cent_MerchantID, pcPay_Centinel.pcPay_Cent_Active FROM pcPay_Centinel WHERE (((pcPay_Centinel.pcPay_Cent_ID)=1));"


end if

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
pcPay_Cent_TransactionURL	= rs("pcPay_Cent_TransactionURL")
pcPay_Cent_ProcessorId		= rs("pcPay_Cent_ProcessorId")
pcPay_Cent_MerchantID		= rs("pcPay_Cent_MerchantID")
if gateway="WPP" then
	pcPay_Cent_Password		= rs("pcPay_Cent_Password")
end if
pcPay_Cent_Active=rs("pcPay_Cent_Active")
set rs=nothing
call closedb()
if scSSL="0" or scSSL="" then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/pcPay_Cent_Ecverifier.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
else
	tempURL=replace(( scSslURL&"/"&scPcFolder&"/pc/pcPay_Cent_Ecverifier.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
end if

pcPay_Cent_TermURL = tempURL
pcPay_Cent_MessageVersion	= "1.7"
'///////////////////////////////////////////////
'// EDIT YOUR CARDINAL COMMERCE PASSWORD
if gateway="WPP" then
	pcPay_Cent_TransactionPwd = pcPay_Cent_Password
else
	pcPay_Cent_TransactionPwd = "tek10"
end if
'///////////////////////////////////////////////

dim resolveTimeout, connectTimeout, sendTimeout, receiveTimeout

resolveTimeout	= 5000
connectTimeout	= 5000
sendTimeout		= 5000
receiveTimeout	= 10000
'1000ms = 1 sec

Function sendMessage(strMsg, istrRedirectPage)
	on error resume next
	Dim requestMessage, objXMLSender

	' Timeout Values represented as Milliseconds, receiveTimeout is the most important
	' The receiveTimeout will allow the checkout process to continue in the event a timely response
	' is not received from CardinalCommerce
	
	requestMessage = "cmpi_msg=" & Server.URLEncode(strMsg)
	
	Set objXMLSender = Server.CreateObject("MSXML2.ServerXMLHTTP"&scXML)
	
	objXMLSender.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	
	objXMLSender.open "POST", pcPay_Cent_TransactionURL, False
			
	objXMLSender.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXMLSender.send requestMessage
	If Err.Number <> 0 Then
		
		' Check for specific Error Codes, if you required enhanced error handling
		If Err.Number = "-2147012894" Then
			Session("Centinel_ErrorNo") = "X100"
			Session("Centinel_ErrorDesc") = "Communication Timeout"
			Else If Err.Number = "-2147012889" Then
				Session("Centinel_ErrorNo") = "X200"
				Session("Centinel_ErrorDesc") = "The server name or address could not be resolved"
			Else 
				Session("Centinel_ErrorNo") = "X300"
				Session("Centinel_ErrorDesc") = "Error Sending/Receiving XML Message (General)"
			End If
		End If
	End If
	sendMessage = objXMLSender.responseText
End Function
			
 
Function redirectBasic(istrRedirectPage, istrError)
	strRedirect = istrRedirectPage & "?strError=" & server.urlencode(istrError)
	FormatErrorRedirectBasic = strRedirect
End Function


Function generateXMLTag(strTagName, strValue)
	Dim tmpValue, strReturn
	tmpValue = Replace(strValue, "&", "&amp;")
	tmpValue = Replace(tmpValue, "<", "&lt;")
	strReturn = "<" & strTagName & ">" & tmpValue & "</" & strTagName & ">"
	generateXMLTag = strReturn
End Function
%>