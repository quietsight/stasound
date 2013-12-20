<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

' Call back URL for WorldPay
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->

<%
idOrder=request("MC_OrderID")
billingname=request("name")
billingaddress=request("address")
billingzip=request("postcode")
billingcountry=request("country")
billingcity=request("city")
billingstate=request("state")
billingphone=request("tel")
billingemail=request("email")
pc_amount=request("amount")
pc_status=request("transStatus")
pc_transId=request("transId")

If scSSL = "1" Then
	pcvWPTmpURL = scSSLUrl
Else
	pcvWPTmpURL = scStoreURL
End If

if pc_status="C" then
	tempURL=replace((pcvWPTmpURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
	tempURL=replace(tempURL,"http:/","http://")
	tempURL=replace(tempURL,"https:/","https://") %>
	<center>
        <table width=95% bgcolor=#F5F5F5 cellpadding=3 cellspacing=0 border=0>
            <tr>
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">You have chosen to cancel your order. If you would like to return back to our site to shop for other items <a href="<%=tempURL%>"><b>CLICK HERE</b></a>.</font></td>
          </tr>
        </table>
    </center>
<% else 
	tempURL=replace((pcvWPTmpURL&"/"&scPcFolder&"/pc/gwWp.asp"),"//","/")
	tempURL=replace(tempURL,"http:/","http://")
	tempURL=replace(tempURL,"https:/","https://")
	%>
    <center>
<table width=95% bgcolor=#F5F5F5 cellpadding=3 cellspacing=0 border=0>
	<tr>
		<td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">YOU MUST CLICK ON '<a href="<%=tempURL%>?status=<%=pc_status%>&idorder=<%=idOrder%>&pc_amount=<%=pc_amount%>"><b>COMPLETE MY ORDER</b></a>' TO COMPLETE YOUR ORDER. You will be taken back to our store and your order status will be updated. Otherwise our store will not be notified that your payment was processed successfully and we will not  be able to ship your order. </font></td>
	</tr>
</table>
</center>
<% end if %>

