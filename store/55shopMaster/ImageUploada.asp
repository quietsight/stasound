<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Upload Images"
pageIcon="pcv4_icon_upload.png"
Section="products" 
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<!--#include file="AdminHeader.asp"-->
<script>
function chgWin(file,window) {
		msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500,status=yes');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="post" enctype="multipart/form-data" name="hForm" action="imageupl.asp" class="pcForms">
<input type="hidden" name="smallImageUrl" value="">
<table class="pcCPcontent" style="width:auto;">
<tr> 
<td colspan="2">
All images are uploaded to the &quot;<b><%=scPcFolder%>/pc/catalog</b>&quot; folder on your Web server (<a href="javascript:chgWin('../pc/imageDir.asp?ffid=smallImageUrl&fid=hForm&ref=ImageUpload','window2')">Manage Uploaded Images</a>).<br>Of course, you may also use your favorite FTP program to upload images to the same location.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="20%" align="right">Image 1:</td>
<td width="80%"><input class=ibtng type="file" name="one" size="30"></td>
</tr>
<tr> 
<td align="right">Image 2:</td>
<td><input class=ibtng type="file" name="two" size="30"></td>
</tr>
<tr> 
<td align="right">Image 3:</td>
<td><input class=ibtng type="file" name="three" size="30"></td>
</tr>
<tr> 
<td align="right">Image 4:</td>
<td><input class=ibtng type="file" name="four" size="30"></td>
</tr>
<tr> 
<td align="right">Image 5:</td>
<td><input class=ibtng type="file" name="five" size="30"></td>
</tr>
<tr> 
<td align="right">Image 6:</td>
<td><input class=ibtng type="file" name="six" size="30"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr> 
<td colspan="2" align="center"> 
<input type="submit" name="Submit" value="Upload" class="submit2">&nbsp;
<input type="button" name="back" value="Back" onClick="javascript:history.back()">
</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->