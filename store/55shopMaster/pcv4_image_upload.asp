<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Upload Images" %>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim query, conntemp, rs
on error resume next
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent" style="height: 400px;">
<tr>
	<td>
  		<iframe frameborder="0" width="100%" height="390" src="../htmleditor/assetmanager/assetmanager.asp?ffilter=image">
        <p>Your browser does not support iframes. Please <a href="ImageUploada.asp">use this file</a> to upload images.</p>
        </iframe>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->