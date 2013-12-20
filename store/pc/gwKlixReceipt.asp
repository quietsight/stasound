<%@ Language=VBScript %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="header.asp"-->
<%'THIS FILE RECEIVES THE RESPONSE FROM viaKLIX AND FORWARDS IT %>
<% ssl_result=request.QueryString("ssl_result")
ssl_result_message=request.QueryString("ssl_result_message")
ssl_approval_code=request.QueryString("ssl_approval_code")
ssl_txn_id=request.QueryString("ssl_txn_id")
ssl_cvv2_response=request.QueryString("ssl_cvv2_response")
ssl_avs_response=request.QueryString("ssl_avs_response")
ssl_invoice_number=request.QueryString("ssl_invoice_number")
ssl_amount=request.QueryString("ssl_amount")
ssl_card_number=request.QueryString("ssl_card_number")
ssl_customer_code=request.QueryString("ssl_customer_code")

If int(ssl_result)=0 then
	'response.redirect "thankyou_klix.asp?idOrder="&ssl_invoice_number
	tordnum=(int(TransactionID)-scpre)
	session("GWAuthCode")=ssl_approval_code
	session("GWTransId")=ssl_txn_id

	Response.redirect "gwReturn.asp?s=true&gw=KLIX"
Else
	dim tempURL
	If scSSL="" OR scSSL="0" then 
		tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
		tempURL=replace(tempURL,"https:/","https://")
		tempURL=replace(tempURL,"http:/","http://") 
	else
		tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
		tempURL=replace(tempURL,"https:/","https://")
		tempURL=replace(tempURL,"http:/","http://")
	end if 

	response.redirect "msgb.asp?message="&server.URLEncode("<b>Error:&nbsp;&nbsp;"&ssl_result&"&nbsp;&nbsp;"&lcase(ssl_result_message)&"<br><br><a href="""&tempURL&"?psslurl=gwklix.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
	response.end
end if
%>
<!--#include file="footer.asp"-->