<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/validation.asp" -->
<% response.Buffer=true %><!--#include file="header.asp"-->
<%
pcv_strSamplePath=Request("path")
%>
<div id="pcMain">
    <p>Your widget should appear directly below.</p>
    <script language="javascript">
		idaffiliate="";
	</script>
	<script type="text/javascript" src="<%=pcv_strSamplePath%>"></script> 
	<br />
    Can't see your preview? Try the following:
    <br />
    <ul>
    	<li>Make sure you have uploaded all the E-Commerce Widget files.</li>
    	<li>Ensure that the category you selected contains products.</li>
    </ul>
</div>    
<!--#include file="footer.asp"-->