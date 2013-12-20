<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Option Group Product Assignments" %>
<% section="products" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
  	Dim rs, connTemp, strSql
	Dim il_idOpiontGroup, il_strListName, il_strMode

'--- Initialize required variables ---
	il_strMode=""

	If Request.QueryString("idOptionGroup") <> "" Then
		il_idOpiontGroup=Request.QueryString("idOptionGroup")
	ElseIf Request.Form("idOptionGroup") <> "" Then
		il_idOpiontGroup=Request.Form("idOptionGroup")
	Else
		Response.Redirect("AdminOptions.asp")
	End If

	If Request.QueryString("mode") <> "" Then
		il_strMode=Request.QueryString("mode")
	ElseIf Request.Form("mode") <> "" Then
		il_strMode=Request.Form("mode")
	Else
		il_strMode="view"
	End If
'---

  Set rs=Server.CreateObject("ADODB.Recordset")
  call openDb()

  strSQL="SELECT optionGroupDesc, idOptionGroup FROM optionsGroups WHERE idOptionGroup=" & il_idOpiontGroup & " ORDER BY optionGroupDesc"
  rs.Open strSQL, connTemp, adOpenStatic, adLockReadOnly

  If Err Then
    TrapError Err.Description
  Else
    il_strListName=rs("optionGroupDesc")
  End If
	
  rs.Close
%>
<!--#include file="AdminHeader.asp"-->
<form method="POST" action="actionGrpOptions.asp" name="myForm" class="pcForms">
<input type="hidden" name="mode" value="<%= il_strMode %>">
<input type="hidden" name="lid" value="<%= il_idOpiontGroup %>">   
    <table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr>
            <td colspan="3"><h2>Products assigned to the Option Group: <strong><%=il_strListName %></strong></h2></td>
        </tr> 
<% 
strSQL="SELECT DISTINCT products.description, products.sku, options_optionsGroups.idProduct, options_optionsGroups.idOptionGroup FROM options_optionsGroups INNER JOIN products ON options_optionsGroups.idProduct=products.idProduct WHERE (((options_optionsGroups.idOptionGroup)="&il_idOpiontGroup&"));"
rs.Open strSQL, connTemp, adOpenStatic, adLockReadOnly
If Err Then
	rs.Close
	connTemp.Close
ElseIf rs.EOF OR rs.BOF Then
%>                       
	<tr> 
		<td colspan="3" align="left"><div class="pcCPmessage">No Products Found. <a href="manageOptions.asp">Manage Option Groups</a>.</div></td>
	</tr>                
<%
Else
	Do While NOT rs.EOF 
%>                            
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
		<td width="7%" align="center"><input type="checkbox" name="optionDescrip" value="<%=rs("idProduct")%>"></td>
		<td colspan="2" width="93%"><a href="FindProductType.asp?id=<%=rs("idProduct")%>"><%= rs("description") %> - <%= rs("sku") %></a></td>
	</tr>                            
<%
rs.MoveNext
Loop
rs.Close
%> 
	<tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>               
	<tr> 
		<td align="center"> 
			<input type="checkbox" value="ON" onclick="javascript:checkTheBoxesB();">
		</td>
		<td colspan="2" class="cpLinksList">Select All Products</td>
	</tr>								
	<tr>
		<td colspan="3" align="center">&nbsp;</td>
	</tr>
	<tr> 
		<td colspan="3" align="center"> 
			<input type="submit" value="Remove Checked" name="submit" class="submit2">&nbsp;
			<input type="button" name="back" value="Back" onClick="location.href='manageOptions.asp'">
		</td>
	</tr>
<% End If %>
                          
</table>
</form>
<%
	Set rs=Nothing
	call closeDb()
%>
<!--#include file="AdminFooter.asp"-->