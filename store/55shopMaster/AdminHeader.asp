<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

' If your store is using a dedicated SSL certificate (e.g. https://www.yourstore.com)
' you can use the following code to force all Control Panel pages to load securely
' using the HTTPS protocol. Remove the apostrophe from the beginning of each of the following
' 8 lines of code to use this feature. This code will not work with shared SSL certificates.

'If (Request.ServerVariables("HTTPS") = "off") Then
'    Dim xredir__, xqstr__
'    xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
'               Request.ServerVariables("SCRIPT_NAME")
'    xqstr__ = Request.ServerVariables("QUERY_STRING")
'    if xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
'    Response.redirect xredir__
'End if
Session.LCID = 1033
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
if err.number <> 0 then
	response.redirect "dbError.asp"
	response.End()
end if
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
Dim pcv_strAdminPrefix
pcv_strAdminPrefix="1"
%>
<!--#include file="smallRecentProducts-header.asp"-->
<html>
<head>
<title>ProductCart v4 - Control Panel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="description" content="ProductCart asp shopping cart software is published by NetSource Commerce. ProductCart's Control Panel allows you to manage every aspect of your ecommerce store. For more information and for technical support, please visit NetSource Commerce at http://www.earlyimpact.com">
<link href="../includes/spry/SpryMenuBarHorizontal-CP.css" rel="stylesheet" type="text/css" />
<% if pcSpryCP="PP" then %>
    <link href="../includes/spry/SpryCollapsiblePanelPP.css" rel="stylesheet" type="text/css" />
<% else %>
    <link href="../includes/spry/SpryCollapsiblePanel-CP.css" rel="stylesheet" type="text/css" />
<% end if %>
<script src="../includes/spry/SpryMenuBar.js" type="text/javascript"></script>
<script src="../includes/spry/SpryCollapsiblePanel.js" type="text/javascript"></script> 
<script src="../includes/javascripts/pcControlPanelFunctions.js" type="text/javascript"></script>
<script src="http://code.jquery.com/jquery-1.4.4.js"></script>
<!--#include file="inc_header.asp" -->
<script type="text/javascript">
var Spry; if (!Spry) Spry = {}; if (!Spry.Utils) Spry.Utils = {};

// We need an unload listener so we can store the data when the user leaves the page
// SpryDOMUtils.js only provides us with a load listener so we create this function

Spry.Utils.addUnLoadListener = function(handler /* function */)
{ 
	if (typeof window.addEventListener != 'undefined')
		window.addEventListener('unload', handler, false);
	else if (typeof document.addEventListener != 'undefined')
		document.addEventListener('unload', handler, false);
	else if (typeof window.attachEvent != 'undefined')		
		window.attachEvent('onunload', handler);
};

Spry.Utils.Cookie = function(type /* string*/, name /* string */, value /* string or number */, options /* object */){
	var	name = name + '=';
		
	if(type == 'create'){
		// start cookie string creation
		var str = name + value;
		
		// check if we have options to add
		if(options){
			// convert days and create an expire setting
			if(options.days){
				var date = new Date();
				str += '; expires=' + (date.setTime(date.getTime() + (options.days * 86400000 /* one day 24 hours x 60 min x 60 seconds x 1000 miliseconds */))).toGMTString();
			}
			// possible path settings
			if(options.path){
				str += '; path=' + options.path				
			}
			// allow cookies to be set per domain
			if(options.domain){
				str += '; domain=' + options.domain;
			}
		} else {
			// always set the path to /
			str += '; path=/';
		}
		// set the cookie
		document.cookie = str;
	} else if(type == 'read'){
		var c = document.cookie.split(';'),
			str = name,
			i = 0,
			length = c.length;
		
		// loop through our cookies
		for(; i < length; i++){
			while(c[i].charAt(0) == ' ')
				c[i] = c[i].substring(1,c[i].length);
				if(c[i].indexOf(str) == 0){
					return c[i].substring(str.length,c[i].length);
				}
		}
		return false;
	} else {
		// remove the cookie, this is done by settings an empty cookie with a negative date
		Spry.Utils.Cookie('create',name,null,{days:-1});
	}
};
</script>

<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.CollapsiblePanel {
	width: 750;
}
.pcPanelTitle1 {
	font-size:14px;
	background-color:#EEE;
	font-weight:bold;
}
.pcPanelDesc {
	font-size: 12px;
	background-color: #EEE;
}
.CollapsiblePanelTab1 {
	background-color:#fff;
	border:dotted;
	border-width:thin;
	font-family:Verdana, Geneva, sans-serif;
}
.pcPanelItalic {
	font-style:italic;
	color:#F60;
	font-weight:bold;
}
.pcSubmenuHeader {
	font-family:Verdana, Geneva, sans-serif;
	font-size:12px;
	font-weight:bold;
}
.pcSubmenuContent {
	font-family:Verdana, Geneva, sans-serif;
	font-size:11px;
	font-weight:normal;
	text-align:center;
}
-->
</style>

</head>
<body style="background-image: url(images/pcv4_template_back.png); background-repeat: repeat-x;">
<script language="javascript" type="text/javascript" src="../includes/pcjscolorchooser.js"></script>
<div id="pcCPmain">
	<div id="pcCPheader">
    	<div id="pcCPstoreName">
		<% '// Prepare and show company name
		Dim pcvStrCompanyName
		pcvStrCompanyName=scCompanyName
		if Len(pcvStrCompanyName)>34 then
		 pcvStrCompanyName=Left(pcvStrCompanyName,31) & "..."
		end if
		if pcvStrCompanyName="" or IsNull(pcvStrCompanyName) then
			pcvStrCompanyName="ProductCart v4"
		end if
		response.write pcvStrCompanyName
		%>
        </div>
        
        <div id="pcCPheaderNav">
            
            <a href="../pc/default.asp" target="_blank"><img src="images/cp11/cp11-storefront.png" width="14" height="14" alt="Storefront"> Storefront</a> [<% if scStoreOff="0" then %><span style="color: #090;">OPEN</span><% else %><span style="color: #F30;">CLOSED</span><% end if %>]
            <a href="http://wiki.earlyimpact.com" target="_blank"><img src="images/cp11/cp11-docs.png" width="14" height="14" alt="ProductCart WIKI"> Wiki</a>
            <a href="http://blog.earlyimpact.com" target="_blank"><img src="images/cp11/cp11-blog.png" width="14" height="14" alt="NetSource Commerce Blog"> Blog</a>
            <a href="http://twitter.com/productcart" target="_blank"><img src="images/cp11/cp11-twitter.png" width="14" height="14" alt="ProductCart on Twitter"> Twitter</a>
            <a href="sitemap.asp"><img src="images/cp11/cp11-internet.png" width="14" height="14" alt="Site Map"> Map</a>
            <a href="about.asp"><img src="images/cp11/cp11-company.png" width="14" height="14" alt="About ProductCart"> About</a>
            
		</div>
        
        <div id="pcCPversion">
            ProductCart <strong>v<%=scVersion&scSubVersion%><% if scSP<>"" and scSP<>"0" then Response.Write(" SP " & scSP) end if %><% if PPD="1" then Response.Write(" PPD") end if %></strong>
    	</div>
        
        <div id="pcCPtopNav">
            <!--#include file="pcv4_navigation_links.asp"--> 
        </div>
    </div>
    
    <div id="pcCPmainArea">
        <div id="pcCPmainLeft">
			<h1 style="background-image:url(images/<%=pageIcon%>); background-position: 700px 0px; background-repeat:no-repeat;"><%=pageTitle%></h1>
            