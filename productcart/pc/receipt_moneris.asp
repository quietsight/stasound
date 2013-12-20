<%@ Language=VBScript %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<%
'THIS FILE RECEIVES THE RESPONSE FROM Moneris AND FORWARDS IT

response_code=request.QueryString("response_code")
if ucase(response_code)="NULL" or response_code="" then
	response_code=-1
end if
iso_code=request.QueryString("iso_code")
message=request.QueryString("message")
rvar_transid=request.QueryString("rvar_transid")
If int(response_code)=>00 AND int(response_code)<=029 then
	session("rvar")=rvar_transid
	response.redirect "gwReturn.asp?rvar=Y"
else
	trans_name=request.QueryString("trans_name")
	result=request.QueryString("result")
	rvar_idbsession=request.QueryString("rvar_idbsession")
	rvar_irandomkey=request.QueryString("rvar_irandomkey")
	card=request.QueryString("card")
	expiry_date=request.QueryString("expiry_date")
	response.redirect "sorry_moneris.asp?monerispmsg="&message&"&idbsession="&rvar_idbsession&"&randomkey="&rvar_irandomkey&"&card="&card&"&expiry_date="&expiry_date
end if
%>