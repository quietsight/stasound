<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pcv_RevName=request("nav")
if pcv_RevName="" then
	pcv_RevName="2"
end if
pageTitle="View Customer Reviews"
%>
<% section="reviews" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="AdminHeader.asp"-->
<% Dim rs, connTemp, query
pcv_IDProduct=getUserInput(request("IDProduct"),10)
pcv_IDReview=getUserInput(request("IDReview"),10)
call opendb()
%>
<!--#include file="../pc/prv_getsettings.asp"-->
<!--#include file="prv_incfunctions.asp"-->
<%
query="SELECT description FROM products WHERE idproduct=" & pcv_IDProduct
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

pcv_PrdName=rs("description")
set rs=nothing

%>
	<% ' START show message, if any %>
		<!--#include file="pcv4_showMessage.asp"-->
	<% 	' END show message %>
	<h2>Product Name: <strong><%=pcv_PrdName%></strong>&nbsp;|&nbsp;Review ID#: <%=pcv_IDReview%></h2>
    
		<input type="hidden" name="IDProduct" value="<%=pcv_IDProduct%>">
		<input type="hidden" name="IDReview" value="<%=pcv_IDReview%>">
		<input type="hidden" name="nav" value="<%=pcv_RevName%>">
		<table class="pcCPcontent">

			<tr>
				<td width="30%" align="right" valign="top">&nbsp;
				</td>
				<td>
				<br>
				<input type="button" name="Back" value="Back" onclick="location='prv_EditReview.asp?IDReview=<% = pcv_IDReview %>&IDProduct=<%=pcv_IDProduct%>&nav=<%=pcv_RevName%>'"></td>
			</tr>
		</table>
<!--#include file="AdminFooter.asp"-->