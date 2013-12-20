<%@ Language=VBScript %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<%
'THIS FILE RECEIVES THE RESPONSE FROM AUTHORIZE.NET AND FORWARDS IT %>
<HTML>
<HEAD>
</HEAD>
<body onLoad="document.frmCC.submit();">
<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwReturn.asp"),"//","/")
tempURL=replace(tempURL,"http:/","http://")
tempURL=replace(tempURL,"https:/","https://") %>
<form name="frmCC" method=POST action="<%=tempURL%>">
<input type=hidden name="authorize" value="1">
<input type=hidden name="x_response_code" value="<%=request("x_response_code")%>"><br>
<input type=hidden name="x_response_subcode" value="<%=request("x_response_subcode")%>"><br>
<input type=hidden name="x_response_reason_code" value="<%=request("x_response_reason_code")%>"><br>
<input type=hidden name="x_response_reason_text" value="<%=request("x_response_reason_text")%>"><br>
<input type=hidden name="x_auth_code" value="<%=request("x_auth_code")%>"><br>
<input type=hidden name="x_avs_code" value="<%=request("x_avs_code")%>"><br>
<input type=hidden name="x_trans_id" value="<%=request("x_trans_id")%>"><br>
<input type=hidden name="x_invoice_num" value="<%=request("x_invoice_num")%>"><br>
<input type=hidden name="x_amount" value="<%=request("x_amount")%>"><br>
<input type=hidden name="x_cust_id" value="<%=request("x_cust_id")%>"><br>
<input type=hidden name="x_Email" value="<%=request("x_Email")%>"><br>
<input type=hidden name="x_Card_Type" value="<%=request("x_Card_Type")%>"><br>
<input type=hidden name="x_city" value="<%=request("x_city")%>"><br>
<input type=hidden name="x_phone" value="<%=request("x_phone")%>"><br>
<input type=hidden name="x_first_name" value="<%=request("x_first_name")%>"><br>
<input type=hidden name="x_last_name" value="<%=request("x_last_name")%>"><br>
<input type=hidden name="x_address" value="<%=request("x_address")%>"><br>
<input type=hidden name="x_company" value="<%=request("x_company")%>"><br>
<input type=hidden name="x_state" value="<%=request("x_state")%>"><br>
<input type=hidden name="x_country" value="<%=request("x_country")%>"><br>
<input type=hidden name="x_zip" value="<%=request("x_zip")%>"><br>
<input type=hidden name="x_type" value="<%=request("x_type")%>"><br>
<input type=hidden name="x_cvv2_resp_code" value="<%=request("x_cvv2_resp_code")%>"><br>
<input type=hidden name="x_avs_code" value="<%=request("x_avs_code")%>"><br>
<input type=hidden name="x_userID" value="<%=request("x_userID")%>"><br>
<input type=hidden name="randomkey" value="<%=request("randomkey")%>"><br>

</form>

</BODY>
</HTML>


