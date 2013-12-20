<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<% 
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

strDeclinedRedirect = "msgb.asp?message="&server.URLEncode("<b>Errors&nbsp;</b>:&nbsp;"&session("Message")&"<br><br><a href="""&tempURL&"?psslurl=gwMonerisUS.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")

 	strStatus=replace(getUserInput(request("message"),0)," ","")

response.write = strStatus
select case ucase(strStatus)
	case "VALID-APPROVED"
		response.write "The transaction was approved and successfully validated."
		tempOrdId=getUserInput(request("cust_id"),0)

		'save all in table
		dim conntemp, rs, query
		call opendb()
		query="INSERT INTO pcPay_OrdersMoneris (pcPay_MOrder_OrderID,pcPay_MOrder_TransKey, pcPay_MOrder_Result,pcPay_MOrder_responseId,pcPay_MOrder_responseCode,pcPay_MOrder_DateStamp,pcPay_MOrder_TimeStamp,pcPay_MOrder_Bankcode,pcPay_MOrder_Transname,pcPay_MOrder_cardholder,pcPay_MOrder_total,pcPay_MOrder_card,pcPay_MOrder_f4l4,pcPay_MOrder_expDate,pcPay_MOrder_message,pcPay_MOrder_ISOcode,pcPay_MOrder_TransId) VALUES ("&tempOrdId&",'"&session("moneris_transaction Key")&"','"&session("moneris_result")&"','"&session("moneris_response_order_id")&"','"&session("moneris_response_code")&"','"&session("moneris_date_stamp")&"','"&session("moneris_time_stamp")&"','"&session("moneris_bank_approval_code")&"','"&session("moneris_trans_name")&"','"&session("moneris_cardholder")&"','"&session("moneris_charge_total")&"','"&session("moneris_card")&"','"&session("moneris_f4l4")&"','"&session("moneris_expiry_date")&"','"&session("moneris_message")&"','"&session("moneris_iso_code")&"','"&session("moneris_bank_transaction_id")&"');"

		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		set rs=nothing
		call closedb()

		session("GWAuthCode")=session("moneris_bank_approval_code")
		session("GWTransId")=session("moneris_bank_transaction_id")
		session("GWTransType")=session("moneris_trans_name")
		
		response.Redirect "gwReturn.asp?s=true&gw=Moneris2"
	case "VALID-DECLINED"
		session("Message") = "The transaction was declined and successfully validated."
		'back to payment page
		response.Redirect(strDeclinedRedirect)
	case "INVALID"
		session("Message") = "No reference to this transactionKey, validation failed."
		'back to payment page
		response.Redirect(strDeclinedRedirect)
	case "INVALID-RECONFIRMED"
		session("Message") = "An attempt has already been made with this transaction key, validation failed."
		'back to payment page
		response.Redirect(strDeclinedRedirect)
	case "INVALID-BAD_SOURCE"
		session("Message") = "The Referring URL is not correct, validation failed."
		'back to payment page
		response.Redirect(strDeclinedRedirect)
end select
%>
<!--#include file="footer.asp"-->