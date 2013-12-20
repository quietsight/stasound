<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=0%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<html>
<head>
<title>Product Weight Check</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<form name="form1" method="post" class="pcForms">
<table class="pcCPcontent">
  	<tr>
    	<td>
			<h1>Product Weight Check</h1><br>
			<%  ' Do Product Weight Check
				Dim conntemp, query, rs
				call opendb()
				query="SELECT idProduct, sku, description FROM products "
				query = query & "where weight<1 and active=-1 "
				query = query & "ORDER BY idProduct;"	
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=connTemp.execute(query)
				if err.number <> 0 then
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in pcTSUtility: "&Err.Description) 
				end if
				prdStr=""
				if NOT rs.eof then %>

					<p><strong>When products do not have a weight assigned, the store cannot properly 
						calculate shipping charges for any shipping service that is based on the 
						order weight. The following products do not have a weight assigned:</strong></p>
					<p>
						<ul>

				<%	do while not rs.eof 
						idProduct=rs("idProduct")
						prdStr=rs("description")&" ("&rs("sku")&")"	%>
	
						<li><a href="Javascript:;;" onClick="opener.location.href='FindProductType.asp?id=<%=idProduct%>';self.close();"><%response.write prdStr%></a></li>

				<%
						rs.moveNext
					loop
					set rs=nothing
					call closedb()
				%>
						</ul>
					</p>
			<% else %>
                <table>
                  <tr> 
                    <td width="5%" valign="top"><img src="images/pc_checkmark_sm.gif" width="18" height="18"></td>
                    <td width="95%"><strong>All products have weights assigned.</strong></td>
                  </tr>
                </table>
			<% end if %>
		</td>
	</tr>
  	<tr>
    	<td colspan="2" align="center">
			<input type="button" name="Back" value="Close Window" onClick="self.close();">
		</td>
	</tr>
</table>
</form>
</div>
</body>
</html>