<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Copy custom fields to other products" %>
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
Dim rsOrd, connTemp, strSQL, pid,rstemp

call openDb()

CustFieldCopy=session("CustFieldCopy")
idproduct=session("ACidproduct")

if request("action")="apply" then

  mySQL="SELECT * FROM products WHERE idproduct=" & idproduct
  Set rstemp=conntemp.execute(mySQL)
  
  CN1=""
  CC1=""
  
  CFieldType=0
 
  if instr(CustFieldCopy,"custom")>0 then
	  CC1=replace(CustFieldCopy,"custom","")
	  query="SELECT idSearchField FROM pcSearchData WHERE idSearchData=" & CC1 & ";"
	  set rsQ=connTemp.execute(query)
	  CN1=0
	  if not rsQ.eof then
	  	CN1=rsQ("idSearchField")
	  end if
	  set rsQ=nothing
	  CFieldType=1
  end if
  
  if instr(CustFieldCopy,"xfield")>0 then
  CN1=rstemp(CustFieldCopy)
  CC1=rstemp(replace(CustFieldCopy,"xfield","x") & "req")
  CFieldType=2
  end if
  
 RSu=0
 RFa=0

If (request("prdlist")<>"") and (request("prdlist")<>",") then
	prdlist=split(request("prdlist"),",")
	For i=lbound(prdlist) to ubound(prdlist)
		id=prdlist(i)
		IF (id<>"0") AND (id<>"") THEN
			query="SELECT * FROM products WHERE idproduct=" & id
			Set rstemp=conntemp.execute(query)
			mytest=false
			vtri=0
			  
			IF CFieldType=2 THEN
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
					CC1a=CC1 
					query="UPDATE products SET " & LN1 & "="& CN1 &", " & LN2 & "=" & CC1a & "  WHERE idproduct="& id
					Set rstemp=conntemp.execute(query)
					set rstemp=nothing
					RSu=RSu+1
				end if
			ELSE
				query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & id & " AND idSearchData IN (SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & CN1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing

				query="INSERT INTO pcSearchFields_Products (idproduct,idSearchData) VALUES (" & id & "," & CC1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
				
				RSu=RSu+1
			END IF
			
			call updPrdEditedDate(id)
			
		END IF 'have id product
	Next
End if

call closedb()

end if 'action=apply
%>
<!--#include file="AdminHeader.asp"-->

	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>

			<div class="pcCPmessageSuccess">
            	The selected custom field was copied to <%=RSu%> products.
				<%if RFa>0 then%>
                    <br>
                    <%=RFa%> of the selected products could not be updated because they already had the maximum allowed number of search or input fields assigned to them.
                <%end if%>
            	<div><a href="AdminCustom.asp?idproduct=<%=idproduct%>">Return to product's Custom Fields page</a>.</div>
            </div>                   
<!--#include file="AdminFooter.asp"-->