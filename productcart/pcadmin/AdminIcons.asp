<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Store Icons" %>
<% Section="layout" %>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->

<% dim query, conntemp, rs
'on error resume next
call openDb()
query="SELECT * FROM icons WHERE id=1"
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=conntemp.execute(query)

if err.number <> 0 then
    response.write "Error in AdminIcons: "&Err.Description
		set rs=nothing
		call closeDb()
end If 
erroricon=rs("erroricon")
requiredicon=rs("requiredicon")
errorfieldicon=rs("errorfieldicon")
previousicon=rs("previousicon")
nexticon=rs("nexticon")
zoom=rs("zoom")
discount=rs("discount")
arrowUp=rs("arrowUp")
arrowDown=rs("arrowDown")
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" enctype="multipart/form-data" action="iconupl.asp" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="3" class="pcCPspacer">
		<!--#include file="pcv4_showMessage.asp"-->
	</td>
</tr>
<tr> 
	<th colspan="2">Upload New Icons&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=435')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
	<th width="27%" align="center">Current Icons</th>
</tr>
<tr>
	<td colspan="3" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="22%">Error:</td>
<td width="51%"><input class=ibtng type="file" name="erroricon" size="30"></td>
<td width="27%" align="center"><img src="../pc/<%=erroricon%>"></td>
</tr>
<tr> 
<td>Required Field:</td>
<td><input class=ibtng type="file" name="requiredicon" size="30"></td>
<td align="center"><img src="../pc/<%=requiredicon%>"></td>
</tr>
<tr> 
<td>Error on Field: </td>
<td><input class=ibtng type="file" name="errorfieldicon" size="30"></td>
<td align="center"><img src="../pc/<%=errorfieldicon%>"></td>
</tr>
<tr> 
<td>Previous (page navigation): </td>
<td><input class=ibtng type="file" name="previousicon" size="30"></td>
<td align="center"><img src="../pc/<%=previousicon%>"></td>
</tr>
<tr> 
<td>Next (page navigation):</td>
<td><input class=ibtng type="file" name="nexticon" size="30"></td>
<td align="center"><img src="../pc/<%=nexticon%>"></td>
</tr>
<tr> 
<td>Zoom:</td>
<td><input name="zoom" type="file" class=ibtng id="zoom" size="30"></td>
<td align="center"><img src="../pc/<%=zoom%>"></td>
</tr>
<tr> 
<td>Discount:</td>
<td><input name="discount" type="file" class=ibtng id="discount" size="30"></td>
<td align="center"><img src="../pc/<%=discount%>"></td>
</tr>
<tr> 
<td>Up Arrow:</td>
<td><input name="arrowUp" type="file" class=ibtng id="arrowUp" size="30"></td>
<td align="center"><img src="../pc/<%=arrowUp%>"></td>
</tr>
<tr> 
<td>Down Arrow:</td>
<td><input name="arrowDown" type="file" class=ibtng id="arrowDown" size="30"></td>
<td align="center"><img src="../pc/<%=arrowDown%>"></td>
</tr>
<tr>
	<td colspan="3" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="3"><hr></td>
</tr>
<tr> 
<td colspan="3" align="center"> 
<input name="Submit" type="submit" class="submit2" value="Update">
&nbsp;                           
<input name="default" type="button" onClick="document.location.href='setIconDefault.asp'" value="Set back to default settings">
&nbsp;
<input type="button" name="Button" value="Back" onClick="javascript:history.back()">
</td>
</tr>            
</table>
</form>
<!--#include file="AdminFooter.asp"-->