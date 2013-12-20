<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<%
dim connTemp, rstemp
call openDB() 
	
query="SELECT pcPay_TripleDeal.pcPay_TD_MerchantName, pcPay_TripleDeal.pcPay_TD_MerchantPassword, pcPay_TripleDeal.pcPay_TD_Profile, pcPay_TripleDeal.pcPay_TD_ClientLang, pcPay_TripleDeal.pcPay_TD_PayPeriod, pcPay_TripleDeal.pcPay_TD_TestMode FROM pcPay_TripleDeal WHERE (((pcPay_TripleDeal.pcPay_TD_ID)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcPay_TD_MerchantName=rs("pcPay_TD_MerchantName")
pcPay_TD_MerchantName=enDeCrypt(pcPay_TD_MerchantName, scCrypPass)
pcPay_TD_MerchantPassword=rs("pcPay_TD_MerchantPassword")
pcPay_TD_MerchantPassword=enDeCrypt(pcPay_TD_MerchantPassword, scCrypPass)
pcPay_TD_Profile=rs("pcPay_TD_Profile")
pcPay_TD_ClientLang=rs("pcPay_TD_ClientLang")
pcPay_TD_PayPeriod=rs("pcPay_TD_PayPeriod")
pcPay_TD_TestMode=rs("pcPay_TD_TestMode")
	
set rs=nothing
call closedb()

Dim objXMLHTTP, xml

'Send the request to the TripleDeal processor.
stext="command=status_payment_cluster"
stext=stext &"&merchant_name="&pcPay_TD_MerchantName
stext=stext &"&merchant_password="&pcPay_TD_MerchantPassword
stext=stext &"&payment_cluster_key="&session("keyValue")
stext=stext &"&report_type=xml_std"

'Send the transaction info as part of the querystring
set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
if pcPay_TD_TestMode="1" then
	xml.open "POST", "https://test.tripledeal.com/ps/com.tripledeal.paymentservice.servlets.PaymentService?"& stext & "", false
else
	xml.open "POST", "https://www.tripledeal.com/ps/com.tripledeal.paymentservice.servlets.PaymentService?"& stext & "", false
end if

xml.send ""
strStatus = xml.Status

'store the response
strRetVal = xml.responseText
	
Set TDXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
TDXMLdoc.async = false 
if TDXMLdoc.loadXML(strRetVal) then ' if loading from a string
	set objLst=TDXMLdoc.getElementsByTagName("status_payment_cluster")
	for i = 0 to (objLst.length - 1)
		varFlag=0
		for j=0 to ((objLst.item(i).childNodes.length)-1)
			response.write objLst.item(i).childNodes(j).nodeName&"<BR>"
			If objLst.item(i).childNodes(j).nodeName="status" then
				for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
					if objLst.item(i).childNodes(j).childNodes(k).nodeName="payment_cluster_process" then
						session("GWTransType") = objLst.item(i).childNodes(j).childNodes(k).Text
					end if
				next
				response.redirect "gwReturn.asp?s=true&gw=TD"
			end if
		next
	next
end if
Set xml = Nothing
set TDXMLdoc = Nothing

'Check the ErrorCode to make sure that the component was able to talk to the authorization network
If (strStatus <> 200) Then
	Response.Write "An error occurred during processing. Please try again later."
Else

End If
%>
<!--#include file="footer.asp"-->