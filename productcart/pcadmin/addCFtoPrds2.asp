<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add custom field to products" %>
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
Dim connTemp, rstemp, query

call openDb()

if (request("action")="apply") and (session("admin_customtype")>"0") then
  
  CN1=""
  CC1=""
  
  CFieldType=session("admin_customtype")
 
  if CFieldType=1 then
  CN1=session("admin_idcustom")
  CC1=session("admin_skeyword")
  end if
  
  if CFieldType=2 then
  CN1=session("admin_idxfield")
  CC1a=session("admin_xreq")
  end if
  
 RSu=0
 RFa=0

If (request("prdlist")<>"") and (request("prdlist")<>",") then
	prdlist=split(request("prdlist"),",")
	For i=lbound(prdlist) to ubound(prdlist)
		id=prdlist(i)
		IF (id<>"0") AND (id<>"") THEN
			IF CFieldType=2 THEN
			
			query="SELECT * FROM products WHERE idproduct=" & id
			set rstemp=Server.CreateObject("ADODB.Recordset")
			Set rstemp=conntemp.execute(query)
			mytest=false
			vtri=0
			
			if (rstemp("xfield1")="") or (rstemp("xfield1")="0") then
				mytest=true
				vtri=1
			end if
			if mytest=false then
				if (rstemp("xfield2")="") or (rstemp("xfield2")="0") then
					mytest=true
					vtri=2
				end if
			end if
			if mytest=false then
				if (rstemp("xfield3")="") or (rstemp("xfield3")="0") then
					mytest=true
					vtri=3
				end if
			end if
  
			if mytest=false then
				RFa=RFa+1
			else
				LN1="xfield" & vtri
				LN2="x" & vtri & "req"
				query="UPDATE products SET " & LN1 & "="& CN1 &", " & LN2 & "=" & CC1a & "  WHERE idproduct="& id
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
				RSu=RSu+1	
			end if
			ELSE 'CFieldType=1
				query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & id & " AND idSearchData IN (SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & CN1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing

				query="INSERT INTO pcSearchFields_Products (idproduct,idSearchData) VALUES (" & id & "," & CC1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
				
				RSu=RSu+1
			END IF
			
			call updPrdEditedDate(id)
			
		END IF
	next
End if 'have prdlist

End if 'action=apply

call closedb()

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

    <div class="pcCPmessageSuccess">
    The selected custom field was added to <%=RSu%> products.
    <%if RFa>0 then%>
        <br /><br />
        <%=RFa%> of the selected products could not be updated because they already had the maximum allowed number of search or input fields assigned to them.
    <%end if%>
    <br /><br />
    <a href="ManageCFields.asp">Manage custom fields</a>
    <% if RSu=1 and not RFa >0 then%>
        &nbsp;|&nbsp;
        <a href="adminCustom.asp?idproduct=<%=prdlist(0)%>">View custom fields for this product &gt;&gt;</a>
    <% end if %>
    </div>
<!--#include file="AdminFooter.asp"-->