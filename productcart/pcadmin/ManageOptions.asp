<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Product Options" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>
    	<td colspan="2">
        	<!--#include file="pcv4_showMessage.asp"-->
        	<div class="cpOtherLinks"><a href="instOptGrpa.asp">Add New Option Group</a> | <a href="ApplyOptionsMulti1.asp">Copy Options from one Product to N other Products</a></div>
        </td>
	</tr>
                    
<%
	Dim rs, rstemp, connTemp, query, pid
	call openDb()

	' gets group assignments
	query="SELECT * FROM OptionsGroups WHERE idOptionGroup>1 ORDER BY OptionGroupDesc ASC"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	
	if rs.EOF then
		set rs=nothing
		call closeDb()
%>      
      <tr> 
        <td colspan="2"><div class="pcCPmessage">No option groups found</div></td>
      </tr>
      <tr>
        <td colspan="2" class="pcCPspacer"></td>
      </tr>                
<% 

	Else 
		Do While NOT rs.EOF %>         
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td width="60%"><a href="modOptGrpa.asp?idOptionGroup=<%=rs("idOptionGroup")%>"><%=rs("OptionGroupDesc")%></a></td>
			<td width="40%" nowrap class="cpLinksList">
				<a href="modOptGrpa.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Manage Attributes</a> | <a href="javascript:if (confirm('You are about to remove this option group from your database. Are you sure you want to complete this action?')) location='delOptGrpb.asp?idOptionGroup=<%= rs("idOptionGroup") %>'">Delete Group</a> | Products: <a href="AssignMultiOptions.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Assign To</a>
				<%
					query="SELECT pcProductsOptions.idProduct FROM pcProductsOptions INNER JOIN products ON pcProductsOptions.idProduct = products.idProduct WHERE pcProductsOptions.idOptionGroup = " & rs("idOptionGroup") & " AND products.removed = 0"
					set rstemp=Server.CreateObject("ADODB.Recordset")
					set rstemp=connTemp.execute(query)
					if not rstemp.eof then
				%>
                    &nbsp;:&nbsp;<a href="ManageOptionsProducts.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Used By</a>
                    &nbsp;:&nbsp;<a href="RevMultiOptions.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Remove From</a>
				<%
					end if
					set rstemp=nothing	
				%>
			</td>
		</tr>
<% 
		rs.MoveNext
    Loop
		set rs=nothing
		call closeDb()
    End If
%>     
</table>
<br /><br />
<!--#include file="AdminFooter.asp"-->