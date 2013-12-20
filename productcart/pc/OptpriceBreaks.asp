<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<% 
dim mySQL, conntemp, rsTemp, pIdProduct, pDescription, pPrice, pDetails, pListPrice, pLgimageURL, pImageUrl, pWeight, pcv_strProductsArray

call openDB()

'// Load icons
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
set rsIconObj = server.CreateObject("ADODB.Recordset")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

categoryDescName = getUserInput(request.QueryString("cd"),0)
pcv_strProductsArray = getUserInput(request.QueryString("SIArray"),0)

'// Trim the last comma so we can use this feature with one item
if instr(pcv_strProductsArray,",")>0 then
	xStringLength = len(pcv_strProductsArray)
	if xStringLength>0 then
		pcv_strProductsArray = left(pcv_strProductsArray,(xStringLength-1))
	end if
end if
ProductArray = Split(pcv_strProductsArray,",")
%>
<html>
<head>
<title><%response.write dictLanguage.Item(Session("language")&"_showcart_20")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<script>
	function WinResize()
	{
	var showScroll=0;
		if (/Firefox[\/\s](\d+\.\d+)/.test(navigator.userAgent)){
			wH=document.body.scrollHeight+100;
			wW=document.body.scrollWidth+20;
		}
			else
		{
			wH=document.body.scrollHeight+80;
			wW=document.body.scrollWidth+20;
		}
	if (wH>550)
	{
		showScroll=1;
		wH=550;
	}
	if (wW>650)
	{
		showScroll=1;
		wW=650;
	}
	
	window.resizeTo(wW,wH);
	if (showScroll==1) document.body.scroll="yes";
		
	}
</script>
</head>
<body style="margin: 0;" onload="javascript:WinResize()">
	<div id="pcMain">
	<%    
	for i = lbound(ProductArray) to UBound(ProductArray)
	pIdProduct=ProductArray(i)
	IF validNum(pIdProduct) THEN
	
		query="SELECT description FROM products WHERE idProduct="& pIdProduct
		Set rsTemp=Server.CreateObject("ADODB.Recordset")
		Set rsTemp=conntemp.execute(query)
		
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsTemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
	
		pDescription=rsTemp("description")
		Set rsTemp = nothing
	
		query="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" ORDER BY num"
		Set rsTemp=Server.CreateObject("ADODB.Recordset")
		set rsTemp=conntemp.execute(query)
	
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsTemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		
		if not rstemp.eof then
	%> 
				<table class="pcMainTable">
					<tr>
						<td colspan="2">
						<h2><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_5")%><%=pDescription%></h2>
						</td>
					</tr>
					<tr> 
						<th width="70%"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_1")%></td>
						<th width="30%"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_2")%>&nbsp;<img src="<%=rsIconObj("discount")%>"></td>
					</tr>
					<%
					do until rstemp.eof
					%>
						<tr>
							<td style="padding: 5px 5px 0px 10px;">
								<% if rstemp("quantityFrom")=rstemp("quantityUntil") then %>
								<%=rstemp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
								<% else %>
								<%=rstemp("quantityFrom")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_3")&"&nbsp;"&rstemp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
						<% end if %>
						</td>
						<td style="padding-top: 5px;">
							<% If (request.querystring("Type")="1") or (session("CustomerType")="1") Then %>
								<% If rstemp("percentage")="0" then %>
								<%=scCurSign & money(rstemp("discountPerWUnit"))%> 
								<% else %>
								<%=rstemp("discountPerWUnit")%>%
								<% End If %>
							<% else %>
								<% If rstemp("percentage")="0" then %>
								<%=scCurSign & money(rstemp("discountPerUnit"))%> 
								<% else %>
								<%=rstemp("discountPerUnit")%>%
								<% End If %>
						 <% end If %>
							</td>
						</tr>
				<%
					rstemp.moveNext
					loop
				%>
				<tr>
					<td colspan="2" class="pcSpacer"></td>
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>
				<tr>
					<td colspan="2" class="pcSpacer"></td>
				</tr>
			</table>
		<%
			end if
			
		set rsTemp = nothing
		
		END IF
		
		next
		%>
		
		<table class="pcMainTable">
			<tr> 
				<td colspan="2" align="right" style="padding: 10px;">
				<input type="image" src="images/close.gif" onClick="self.close()" alt="<%=dictLanguage.Item(Session("language")&"_AddressBook_5")%>">
				</td>
			</tr>
		</table>
	</div>
</body>
</html>
<%
call closeDb()
conlayout.Close
Set conlayout=nothing
Set rsIconObj = nothing 
%>