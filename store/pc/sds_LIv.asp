<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%
'Get Path Info
pcv_filePath=Request.ServerVariables("PATH_INFO")
do while instr(pcv_filePath,"/")>0
	pcv_filePath=mid(pcv_filePath,instr(pcv_filePath,"/")+1,len(pcv_filePath))
loop

pcv_Query=Request.ServerVariables("QUERY_STRING")

if pcv_Query<>"" then
pcv_filePath=pcv_filePath & "?" & pcv_Query
end if

if (Session("pc_idsds")="0") or (Session("pc_idsds")="") then
 response.redirect "sds_Login.asp?redirectUrl="&Server.URLEncode(pcv_filePath)
end if
%>