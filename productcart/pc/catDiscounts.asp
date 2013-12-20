<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% dim query, conntemp, rsTemp, pIdCategory, query2, rsTemp2, pDescription

pIdCategory=request.QueryString("idCategory")

if trim(pIdCategory)="" or not validNum(pIdCategory) then
   response.redirect "msg.asp?message=86"
end if

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

call opendb()

'query="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idCategory="& pIdCategory &" ORDER BY num"
query="SELECT pcCD_quantityfrom,pcCD_quantityUntil,pcCD_percentage,pcCD_discountPerWUnit,pcCD_discountPerUnit FROM pcCatDiscounts WHERE pcCD_idcategory="& pIdCategory &" ORDER BY pcCD_num"

set rsTemp=Server.CreateObject("ADODB.Recordset")
set rsTemp=conntemp.execute(query)

'query="SELECT description FROM products WHERE idCategory="& pIdCategory
query="SELECT categoryDesc,idcategory FROM categories WHERE idcategory="&pIdCategory
Set rsTemp2=Server.CreateObject("ADODB.Recordset")
Set rsTemp2=conntemp.execute(query)
pDescription=rsTemp2("categoryDesc")
Set rsTemp2 = nothing

if err.number <> 0 then
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in catDiscounts, line 22: "&err.description) 
end if
%> 
<html>
<head>
<title>Category Discounts</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 0;">
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td colspan="2">
		<h2><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_5")%><%=pDescription%></h2>
		</td>
	</tr>
	<tr>
		<td colspan="2" style="padding: 0px 5px 0px 5px;">
		    <%response.write dictLanguage.Item(Session("language")&"_pricebreaks_6")%>
		</td>
	</tr>
	<tr> 
		<th width="70%"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_1")%></td>
		<th width="30%"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_2")%>&nbsp;
		<img src="<%=rsIconObj("discount")%>" border="0"></td>
	</tr>
  <% do until rstemp.eof %>
	<tr>
		<td style="padding: 5px 5px 0px 10px;">
			<% if rstemp("pcCD_quantityFrom")=rstemp("pcCD_quantityUntil") then %>
				<%=rstemp("pcCD_quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
			<% else %>
				<%=rstemp("pcCD_quantityFrom")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_3")&"&nbsp;"&rstemp("pcCD_quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
		<% end if %>
		</td>
		<td style="padding-top: 5px;">
			<% If (request.querystring("Type")="1") or (session("CustomerType")="1") Then %>
				<% If rstemp("pcCD_percentage")="0" then %>
				<%=scCurSign & money(rstemp("pcCD_discountPerWUnit"))%> 
				<% else %>
				<%=rstemp("pcCD_discountPerWUnit")%>%
				<% End If %>
			<% else %>
				<% If rstemp("pcCD_percentage")="0" then %>
				<%=scCurSign & money(rstemp("pcCD_discountPerUnit"))%> 
				<% else %>
				<%=rstemp("pcCD_discountPerUnit")%>%
				<% End If %>
			<% end If %>
		</td>
	</tr>
	<% rstemp.moveNext
	loop
	set rsTemp = nothing
	call closeDb()
	%>
	<tr> 
		<td colspan="2" align="right" style="padding-top: 20px;">
		<input type="image" value="Close Window" src="images/close.gif" width="32" height="25" onClick="self.close()">
    </td>
	</tr>
</table>
</div>
</body>
</html>
<%conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing%>