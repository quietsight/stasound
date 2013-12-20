<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<html>
<title>Upload Images</title>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="4" align="center">
<tr> 
  <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">Upload 
    Images to Your Store Catalog</font></b></font></td>
</tr>
<tr> 
  <td height="10" colspan="2"></td>
</tr>
<tr> 
  <td colspan="2">
    <p>&nbsp;</p>
    <p align="center"><font face="Arial, Helvetica, sans-serif" size="2"><font size="4">Image(s) uploaded successfully</font></font></p>
    <p>&nbsp;</p>
	<%
	if (session("cp_fi")<>"") AND (request("fn")<>"") then
	%>
	<script>opener.document.<%=session("cp_fi")%>.value="<%=request("fn")%>";</script>
	<%session("cp_fi")=""
	end if%>
  </td>
</tr>
<tr> 
  <td colspan="2"> 
    <div align="left"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="button" value="Close Window" onClick="javascript:window.close();">
        </font></p>
    </div>
  </td>
</tr>
</table>
</body>
</html>