<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add custom field to categories" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<%
Dim connTemp

call openDb()

if (request("action")="apply") and (session("admin_customtype")>"0") then
  
	CN1=""
	CC1=""
	
	CFieldType=session("admin_customtype")
	
	if CFieldType=1 then
		CN1=session("admin_idcustom")
		CC1=session("admin_skeyword")
	end if
	
	RSu=0
	RFa=0

	If (request("catlist")<>"") and (request("catlist")<>",") then

		catlist=split(request("catlist"),",")
		For i=lbound(catlist) to ubound(catlist)
			id=catlist(i)
			IF (id<>"0") AND (id<>"") THEN

				query="DELETE FROM pcSearchFields_Categories WHERE idCategory=" & id & " AND idSearchData IN (SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & CN1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing

				query="INSERT INTO pcSearchFields_Categories (idCategory,idSearchData) VALUES (" & id & "," & CN1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
				
				call updCatEditedDate(id,"")
				
				RSu=RSu+1

			END IF
		next
	End if 'have catlist

End if 'action=apply

	session("admin_idxfield")=0
	session("admin_xreq")=0
	session("admin_customtype")=0
	session("admin_useExist")=0
	session("admin_idcustom")=0
	session("admin_skeyword")=""
%>
<!--#include file="AdminHeader.asp"-->
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
			<div class="pcCPmessageSuccess">The selected custom field was added to <%=RSu%> categories.
			<%if RFa>0 then%>
				<br /><br />
				<%=RFa%> of the selected categories could not be updated because they already had the maximum allowed number of search or input fields assigned to them.
			<%end if%>	
            <br><br>
            <a href="ManageCFields.asp">Manage custom fields</a>         
			</div>
<!--#include file="AdminFooter.asp"-->