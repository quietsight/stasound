<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Upload Images" %>
<% Section="products" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->

<html>
<title>Upload Images</title>
	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form method="post" enctype="multipart/form-data" action="taxupl_popup.asp">
        <table width="100%" border="0" cellspacing="0" cellpadding="8" align="center">
          <tr> 
            <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b>Upload Tax Rate File</b></font></td>
          </tr>
          <tr> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td height="18" colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"> Select your file by using the &quot;Browse&quot; button. Then click on &quot;Upload&quot;. Your file will be automatically uploaded to the &quot;<b>productcart/pc/tax</b>&quot; folder on your Web server.</font></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"></td>
          </tr>
          <tr> 
            <td width="9%"><font face="Arial, Helvetica, sans-serif" size="2">File:</font></td>
            <td width="91%"><font face="Arial, Helvetica, sans-serif" size="2"><input type="file" name="one" size="30"></font></td>
          </tr>
          <tr> 
            <td colspan="2" height="15"></td>
          </tr>
          <tr> 
            <td colspan="2" align="center">
                  <input type="submit" name="Submit" value="Upload">
                  <input type="button" value="Close Window" onClick="javascript:window.close();">
            </td>
          </tr>
        </table>
</form>
</body>
</html>