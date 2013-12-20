<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle = "Reward Points - General Settings" %>
<% Section = "specials" %>
<!--#include file="Adminheader.asp"-->
<table class="pcCPcontent">
    <tr> 
        <td>
        <ul class="pcListIcon">
            <li><a href="RpSettings.asp">General Settings</a><br />
                Use this link to set general settings for the program such as the program name, the conversion rate, whether or not customer earn points or referrals, and more.
            </li>
            <li style="padding-top: 10px;"><a href="RpProducts.asp">Assign Points to Multiple Products</a><br />
                You can assign points to products one at a time when modifying or adding a product, or you can use this link to assign points to all products at once.
            </li>
            <li style="padding-top: 10px;"><a href="http://wiki.earlyimpact.com/productcart/marketing-reward_points" target="_blank">How It Works</a></li>
        </ul>
        </td>
    </tr>
</table>
<!--#include file="Adminfooter.asp"-->