<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<%
dim conntemp, query, rs
call opendb()
query="SELECT pcPay_Moneris_StoreId, pcPay_Moneris_Key, pcPay_Moneris_TransType, pcPay_Moneris_Lang, pcPay_Moneris_Testmode,pcPay_Moneris_Meth FROM pcPay_Moneris Where pcPay_Moneris_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcPay_Moneris_StoreId=rs("pcPay_Moneris_StoreId")
pcPay_Moneris_StoreId=enDeCrypt(pcPay_Moneris_StoreId, scCrypPass)
pcPay_Moneris_Key=rs("pcPay_Moneris_Key")
pcPay_Moneris_Key=enDeCrypt(pcPay_Moneris_Key, scCrypPass)
pcPay_Moneris_TransType=rs("pcPay_Moneris_TransType")
pcPay_Moneris_Lang=rs("pcPay_Moneris_Lang")
pcPay_Moneris_Testmode=rs("pcPay_Moneris_Testmode")
pcPay_Moneris_Meth = rs("pcPay_Moneris_Meth")

set rs=nothing
call closedb()

response_order_id=request("response_order_id")
response_code=request("response_code")
date_stamp=request("date_stamp")
time_stamp=request("time_stamp")
bank_approval_code=request("bank_approval_code")
result=request("result")
trans_name=request("trans_name")
cardholder=request("cardholder")
charge_total=request("charge_total")
card=request("card")
f4l4=request("f4l4")
message=request("message")
iso_code=request("iso_code")
bank_transaction_id=request("bank_transaction_id")
transactionKey=request("transactionKey")
expiry_date=request("expiry_date")

if result="0" then
	session("idOrderSaved")=""
	dim tempURL
	If scSSL="" OR scSSL="0" Then
		tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
		tempURL=replace(tempURL,"https:/","https://")
		tempURL=replace(tempURL,"http:/","http://") 
	Else
		tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
		tempURL=replace(tempURL,"https:/","https://")
		tempURL=replace(tempURL,"http:/","http://")
	End If
	
if session("IDEBIT_ISSCONF") <> "" then 
	   
	    response.redirect "msgb.asp?message="&server.URLEncode("<b>Errors&nbsp;</b>:&nbsp;"&message&"<br><br><a href="""&tempURL&"?psslurl=gwMonerisInterac.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
        response.end
else
    	response.redirect "msgb.asp?message="&server.URLEncode("<b>Errors&nbsp;</b>:&nbsp;"&message&"<br><br><a href="""&tempURL&"?psslurl=gwMoneris2.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
        response.end
	
end if
	
end if
session("moneris_transaction Key")=transactionKey
session("moneris_result")=result
session("moneris_response_order_id")=response_order_id
session("moneris_response_code")=response_code
session("moneris_date_stamp")=date_stamp
session("moneris_time_stamp")=time_stamp
session("moneris_bank_approval_code")=bank_approval_code
session("moneris_trans_name")=trans_name
session("moneris_cardholder")=cardholder
session("moneris_charge_total")=charge_total
session("moneris_card")=card
session("moneris_f4l4")=f4l4
session("moneris_expiry_date")=expiry_datey
session("moneris_message")=message
session("moneris_iso_code")=iso_code
session("moneris_bank_transaction_id")=bank_transaction_id

Dim objXMLHTTP, xml

'Build XML string
stext="ps_store_id="&pcPay_Moneris_StoreId
stext=stext & "&hpp_key="&pcPay_Moneris_Key
stext=stext & "&transactionKey=" & transactionKey
stext=stext & "&txn_num=" & request("txn_num")
stext=stext & "&eci=" & request("eci")

if pcPay_Moneris_TestMode="1" then
	strHostURL="https://esqa.moneris.com/HPPDP/verifyTxn.php"
else
	strHostURL="https://www3.moneris.com/HPPDP/verifyTxn.php"
end if

dim resolveTimeout, connectTimeout, sendTimeout, receiveTimeout

resolveTimeout	= 5000
connectTimeout	= 5000
sendTimeout		= 5000
receiveTimeout	= 10000
	
'Send the transaction info as part of the querystring
set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
xml.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
if pcPay_Moneris_Meth ="1"  then 
	xml.open "POST", strHostURL &"", false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send(stext)
 Else
    xml.open "GET", strHostURL &"?" &stext & "", false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send "" 
End if  

strRetVal = xml.responseText
response.write strRetVal
response.end
%>
<!--#include file="footer.asp"-->