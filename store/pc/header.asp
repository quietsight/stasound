<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="include-metatags.asp"-->
<html>
<head>
<%Session.LCID = 1033
if pcv_PageName<>"" then%>
<title><%=pcv_PageName%></title>
<%end if%>
<%GenerateMetaTags()%>
<%Response.Buffer=True%> 
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcHeaderFooter11.css" />
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<link type="text/css" rel="stylesheet" href="pcBTO.css" />
<!-- Stylesheets -->  
    <link href="/../style.css" rel="stylesheet" type="text/css">
<!-- TypeKit -->    
    <script type="text/javascript" src="//use.typekit.net/vfd8tfq.js"></script>
    <script type="text/javascript">try{Typekit.load();}catch(e){}</script>
<!--#include file="inc_header.asp" -->
</head>
<body>
<!-- HEADER -->
    <div class="wrapper">   
        <div id="header">   
            <a href="index.html" />
                <div class="logo">

                </div>
            </a>

<div class="clear"></div>

        </div><!--header-->

<!-- NAV -->    
        <div class="nav">
            <div class="navblock">
                <p>
                    Superior Products for 
                    <a href="http://stasound.com/index.html#dogsection" class="navlink">DOGS</a> |
                    <a href="http://stasound.com/index.html#catsection" class="navlink">CATS</a> |
                    <a href="http://stasound.com/index.html#horsesection" class="navlink">HORSES</a> |
                    <a href="http://stasound.com/index.html#livestocksection" class="navlink">LIVESTOCK SHOW ANIMALS</a>
                </p>
            </div><!--navblock-->
        </div><!--nav-->

<!-- MAINCONTENT -->

        <div class="shadowtop"></div>

        <div id="maincontent">
            <div class="contentblock">
<!--#include file="inc_catsmenu.asp" -->