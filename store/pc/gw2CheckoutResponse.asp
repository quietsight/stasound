<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
'////////////////////////////////////////////////////
'// 2Checkout
'////////////////////////////////////////////////////
CartOrderID = request("cart_order_id")
order_number = request("order_number")
direct_post = request("direct_post")
if CartOrderID<>"" AND direct_post="" then		
	pcv_strQuery="?"
	pcv_strQuery=pcv_strQuery&"cart_order_id="&CartOrderID
	pcv_strQuery=pcv_strQuery&"&order_number="&order_number
	pcv_strQuery=pcv_strQuery&"&direct_post=False"	
	if scSSL="1" AND scIntSSLPage="1" then
		tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwreturn.asp"),"//","/")
	 else
		tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwreturn.asp"),"//","/")
	end if
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
	tempURL=tempURL & pcv_strQuery	
	%>
	<script type="text/javascript">
	<!--
	window.location = "<%=tempURL%>"
	//-->
	</script>
	<%
	response.End()	
end if
'////////////////////////////////////////////////////
'// 2Checkout
'////////////////////////////////////////////////////
%>