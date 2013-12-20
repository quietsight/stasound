<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="XML Tools Manager" %>
<% section="layout" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/pcXMLsettings.asp"-->
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,rs,query
call opendb()

tmpExportAdmin=0
query="SELECT pcXP_ExportAdmin FROM pcXMLPartners WHERE pcXP_ExportAdmin=1;"
set rs=connTemp.execute(query)
if not rs.eof then
	tmpExportAdmin=1
end if
set rs=nothing

call closedb()

'****************************
'* Store name and version
'****************************
%>       
<table class="pcCPcontent">
	<tr>
		<td>
        <h2>ProductCart XML Tools v<%=scXMLVersion%><% if PPD="1" then Response.Write(" PPD") end if %>: Welcome to XML Tools Manager</h2>
			<div class="pcCPsectionTitle">Manage XML Tools</div>
			<ul>
				<li><a href="AdminXMLSettings.asp">General Settings</a></li>
				<li><a href="AdminManageXMLPartner.asp">Manage Partners</a></li>
				<li><a href="AdminManageXMLIPs.asp">Allowed IP Addresses</a></li>
			</ul>
			<div class="pcCPsectionTitle">Export to XML</div>
			<ul>
				<li><a href="<%if tmpExportAdmin=0 then%>javascript:alert('This function requires an Export Admin Account. Please update a XML Partner to Export Admin.');<%else%>XMLExportOrdFile.asp<%end if%>">Orders</a></li>
				<li><a href="<%if tmpExportAdmin=0 then%>javascript:alert('This function requires an Export Admin Account. Please update a XML Partner to Export Admin.');<%else%>XMLExportCustFile.asp<%end if%>">Customers</a></li>
				<li><a href="<%if tmpExportAdmin=0 then%>javascript:alert('This function requires an Export Admin Account. Please update a XML Partner to Export Admin.');<%else%>XMLExportPrdFile.asp<%end if%>">Products</a></li>
			</ul>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->