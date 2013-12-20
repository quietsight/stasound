<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Download Exported File" %>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<% 
strFile=request.querystring("File")
fileType=request.querystring("type")
%>

<!--#include file="AdminHeader.asp"-->
    <table class="pcCPcontent">
        <tr>
            <td>
            <h2>Download Your File</h2>
            <p><a href="<%=strFile%>.<%=fileType%>"><img src="images/DownLoad.gif" border="0"></a></p>
            <p style="margin: 10px 0 10px 0;">To ensure that your file downloads correctly, right click on the icon above and choose &quot;<b>Save Target As...</b>&quot; from your menu.</p>
        </tr>
    </table>
<!--#include file="AdminFooter.asp"-->
