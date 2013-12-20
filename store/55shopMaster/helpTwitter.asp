<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "ProductCart Twitter Updates - Latest 10 Messages" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>
    	<td>
			<script src="http://widgets.twimg.com/j/2/widget.js"></script>
            <script>
            new TWTR.Widget({
              version: 2,
              type: 'profile',
              rpp: 10,
              interval: 6000,
              width: 800,
              height: 300,
              theme: {
                shell: {
                  background: '#f1f1f1',
                  color: '#555555'
                },
                tweets: {
                  background: '#ffffff',
                  color: '#777777',
                  links: '#0790eb'
                }
              },
              features: {
                scrollbar: false,
                loop: false,
                live: false,
                hashtags: true,
                timestamp: true,
                avatars: false,
                behavior: 'all'
              }
            }).render().setUser('productcart').start();
            </script>
        </td>
    </tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->