<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<%
session("AllowUsing")="1"
session("usingM")=request("idnews")
'Start SDBA
if request("pagetype")<>"" then
	response.redirect "sds_newsWizStep1.asp?pagetype=" & request("pagetype")
else
	response.redirect "newsWizStep1.asp"
end if
'End SDBA
%>