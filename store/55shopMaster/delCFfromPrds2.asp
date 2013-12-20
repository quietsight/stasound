<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove custom field from products" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim query, conntemp, rstemp

if (request("action")="apply") and (request("idcustom")<>"") then

	idcustom=mid(request("idcustom"),2,len(request("idcustom")))  
	  
	prdlist=request("prdlist")
	
	pcArr=split(prdlist,",")
	intCount=ubound(pcArr)
	 
	RSu=0
	RFa=0
	
	call openDb()

	for i=0 to intCount
		if trim(pcArr(i))<>"" then
			if Left(request("idcustom"),1)="C" then
				query="UPDATE products SET xfield1=0,x1req=0 WHERE idproduct=" & trim(pcArr(i)) & " and xfield1=" & idcustom
				Set rstemp=conntemp.execute(query)
				query="UPDATE products SET xfield2=0,x2req=0 WHERE idproduct=" & trim(pcArr(i)) & " and xfield2=" & idcustom
				Set rstemp=conntemp.execute(query)
				query="UPDATE products SET xfield3=0,x3req=0 WHERE idproduct=" & trim(pcArr(i)) & " and xfield3=" & idcustom
				Set rstemp=conntemp.execute(query)
			else
				query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & trim(pcArr(i)) & " AND (idSearchData IN (SELECT DISTINCT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & "))"
				Set rstemp=conntemp.execute(query)
			end if
			RSu=RSu+1	
		end if
	next 
	
	set rstemp = nothing
	call closedb()
end if
%>
<!--#include file="AdminHeader.asp"-->

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<div class="pcCPmessageSuccess">
    The selected custom field was deleted from <b><%=RSu%></b> products. <a href="ManageCFields.asp">Manage Custom Fields</a>.
</div>               
<!--#include file="AdminFooter.asp"-->