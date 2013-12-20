<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
Dim pcv_strRecentProducts
pcv_strRecentProducts=Request("action")
if pcv_strRecentProducts<>"" then
	Response.Cookies("pcfront_visitedPrdsCP")=""
	Response.Cookies("pcfront_visitedPrdsCP").Expires=Date()
end if
%>