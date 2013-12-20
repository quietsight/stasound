<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Dim pageTitle, pageName, pageIcon, Section
pageTitle="ProductCart <-> SubscriptionBridge Integration"
pageName="sb_Default.asp"
pageIcon=""
Section="SB" 
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="sb_inc.asp"-->
<% Dim connTemp,query,rs

pcv_pageType=request("pagetype")

'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
	<%if request("msg")<>"" then%>
	<tr>
		<td>
			<%
            Select Case request("msg")
                Case "1": msg = "Your ProductCart-powered store has been successfully registered with your SubscriptionBridge account. You can now exchange data between the two systems."
				msgType=1
            End Select
            %>
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
		</td>
	</tr>
	<%end if%>
	<tr>
		<td><a href="http://www.subscriptionbridge.com" target="_blank"><img src="SubscriptionBridge/images/subscription_logo.jpg" alt="SubscriptionBridge Management System" width="300" height="53" border="0" align="right" style="margin-left: 30px;" /></a>The integration between ProductCart and <a href="http://www.subscriptionbridge.com" target="_blank">SubscritionBridge</a> allows you to <strong>sell subscription-based products and services</strong> right through your store, leveraging SubscriptionBridge's subscription management features, recurring payment support, and its many customer service tools.
		</td>
	</tr>
	<%if tmp_setup=0 then%>
	<tr>
		<td>
			<div class="pcCPmessage">
				Click "Activate SubscriptionBridge Account" below to get started.
			</div>
		</td>
	</tr>
	<%end if%>
	<tr>
		<td>
		<ul>
		<%if tmp_setup=1 then%>
			<li><a href="sb_CreatePackages.asp">Generate Package Link</a></li>
            <li><a href="sb_ViewPackages.asp">View/Modify Package Links</a></li>
            <%
			call opendb()
			query="SELECT orders.idOrder FROM orders, productsordered WHERE orders.OrderStatus>1  And orders.idOrder = ProductsOrdered.idOrder and ProductsOrdered.pcSubscription_ID >0  ORDER BY ProductsOrdered.idProductOrdered DESC"
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=conntemp.execute(query)
			if NOT rstemp.eof then
				%>
                <li><a href="sb_ViewSubs.asp?idmain=0">Subscription Report</a></li>
            	<% 
			end if 
			set rstemp = nothing
			call closedb()
			%>
			<li style="margin-top: 15px;"><a href="sb_Settings.asp" onclick="javascript:pcf_Open_SubscriptionBridge();">Settings</a></li>
			<li><a href="sb_manageAcc.asp">Enter API Credentials</a></li>		
		<%else%>
			<li><a href="sb_manageAcc.asp">Activate SubscriptionBridge Integration</a></li>
            <li><i>Settings</i></li>
			<li><i>Create Subscription Package Link</i></li>
            <li><i>View/Modify Package Links</i></li>
		<%end if%>
			<li style="margin-top: 15px;"><a href="http://wiki.subscriptionbridge.com/cartintegration/productcart/" target="_blank">User Guide</a></li>
			<li><a href="http://www.earlyimpact.com/subscriptionbridge/support.asp" target="_blank">Technical Support</a></li>
		</ul>
		</td>
	</tr>
</table>
<%Response.write(pcf_ModalWindow("Loading...","SubscriptionBridge", 300))%>
<!--#include file="AdminFooter.asp"-->