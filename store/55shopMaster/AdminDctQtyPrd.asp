<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Add Quantity Discounts (Tiered Pricing)" %>
<% Section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->

<!--#include file="AdminHeader.asp"-->
<% 
Dim rs, connTemp, query, idproduct, discountdesc, percentage, baseproductonly, discountPerUnit1, discountPerWUnit1, quantityfrom1

call opendb()

CanNotRun=0

pIDProduct=Request("idproduct")
if pIDProduct="" OR IsNull(pIDProduct) then
	pIDProduct=0
end if

query="SELECT DISTINCT idproduct FROM pcPrdPromotions WHERE idproduct=" & pIDProduct & ";"
set rs=connTemp.execute(query)

if not rs.eof then
CanNotRun=1%>
<table class="pcCPcontent">
       <tr>
        	<td colspan="3">
				<div class="pcCPmessage">You cannot add quantity discounts to this product because a promotion has been created for it. <a href="ModPromotionPrd.asp?idproduct=<%=pIdProduct%>&iMode=start">Review it</a></div>
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" name="back" value=" Product Quantity Discounts " onclick="location='viewDisca.asp';" class="ibtnGrey">
				&nbsp;&nbsp;<input type="button" name="back2" value=" Back to Main menu " onclick="location='menu.asp';" class="ibtnGrey">
			</td>
		</tr>		
</table>
<%
end if
set rs=nothing

IF CanNotRun=0 THEN

sMode=Request.Form("Submit")

If sMode<>"" Then
	'save all inputs in temporary session state
	Session("adminidproduct")=Request("idproduct")
	Session("admindiscountdesc")=Request("discountdesc")
	Session("adminpercentage")=Request("percentage")
	Session("adminbaseproductonly")=Request("baseproductonly")	
	Session("admindiscountPerUnit1")=replacecomma((Request("discountPerUnit1")))
	Session("admindiscountPerWUnit1")=replacecomma((Request("discountPerWUnit1")))
	Session("adminquantityfrom1")=replacecomma((Request("quantityfrom1")))
	Session("adminquantityUntil1")=replacecomma((Request("quantityuntil1")))
	
	if NOT isNumeric(Session("admindiscountPerUnit1")) OR Session("admindiscountPerUnit1")="" then
		Session("admindiscountPerUnit1")=0
	end if
	if NOT isNumeric(Session("admindiscountPerWUnit1")) OR Session("admindiscountPerWUnit1")="" then
		Session("admindiscountPerWUnit1")=0
	end if

	'check to make sure there is a percentage identifier
	if Session("adminpercentage")="" then
		msg="You must specify a whether this is a percentage or absolute discount."
		response.redirect "AdminDctQtyPrd.asp?msg="&msg
	end if
	
	'check to make sure there is a baseproduct identifier
	if Session("adminbaseproductonly")="" then
		msg="You must specify a whether this is a dicount calculated on the product price only or the product price and options price together."
		response.redirect "AdminDctQtyPrd.asp?msg="&msg
	end if

	'check to make sure there are no overlaps
	if Session("adminquantityfrom1")="" then
		Session("adminquantityfrom1")=0
	end if
	if Session("adminquantityUntil1")="" then
		Session("adminquantityUntil1")=99999
	end if
	

	if Session("adminquantityfrom1") <> "" AND Session("adminquantityUntil1") <> "" AND Session("admindiscountPerUnit1")="" then
		msg="You must specify a discount price."
		response.redirect "AdminDctQtyPrd.asp?msg="&msg
	end if

	'make sure the from < until
	if int(Session("adminquantityfrom1"))>int(Session("adminquantityUntil1")) then
		msg="Your quantity ""To"" must be greater then your quantity ""From""."
		response.redirect "AdminDctQtyPrd.asp?msg="&msg
	end if

	If (money(Session("admindiscountPerUnit1")) > 0 OR money(Session("admindiscountPerWUnit1"))>0) AND Session("adminquantityfrom1") <> "" AND Session("adminquantityUntil1") <> "" Then
		call openDb()
		'check to see if this num already exists in db for this product
		query="SELECT * FROM discountsPerQuantity WHERE idproduct="&Session("adminidproduct")&" AND num=1"
		Set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if NOT rs.eof then
			set rs=nothing
			call closedb()
			response.redirect "ModDctQtyPrd.asp?idproduct="& Session("adminidproduct")
		end if
		query="INSERT INTO discountsPerQuantity (idcategory,discountPerUnit,discountPerWUnit,discountdesc,quantityfrom, quantityuntil,idProduct,num,percentage,baseproductonly) VALUES (0,"&Session("admindiscountPerUnit1")&","&Session("admindiscountPerWUnit1")&",'"&Session("admindiscountdesc")&"',"&Session("adminquantityfrom1")&","&Session("adminquantityUntil1")&","&Session("adminidproduct")&",1,"& Session("adminpercentage") &","&Session("adminbaseproductonly")&");"
		Set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		Set rs=Nothing
		call closedb()
	Else
		msg="The amounts entered below are not in the correct format."
		response.redirect "AdminDctQtyPrd.asp?msg="&msg
	End If
	idproduct=Session("adminidproduct")
	Session("adminidproduct")=""
	Session("admindiscountdesc")=""
	Session("adminpercentage")=""
	Session("adminbaseproductonly")=""
	Session("admindiscountPerUnit1")=""
	Session("adminquantityfrom1")=""
	Session("adminquantityUntil1")=""
	
	response.redirect "ModDctQtyPrd.asp?idproduct="&idProduct
End If

idproduct=request("idproduct")
	if trim(idproduct)="" then
		idproduct=Session("adminidproduct")
	end if
	
if NOT isNumeric(Session("admindiscountPerUnit1")) OR Session("admindiscountPerUnit1")="" then
	Session("admindiscountPerUnit1")=0
end if
if NOT isNumeric(Session("admindiscountPerWUnit1")) OR Session("admindiscountPerWUnit1")="" then
	Session("admindiscountPerWUnit1")=0
end if

if Session("adminquantityfrom1")="" then
	Session("adminquantityfrom1")=1
end if
if Session("adminquantityUntil1")="" then
	Session("adminquantityUntil1")=99999
end if

discountdesc=Session("admindiscountdesc")
percentage=Session("adminpercentage")
baseproductonly=Session("adminbaseproductonly")
discountPerUnit1=Session("admindiscountPerUnit1")
discountPerWUnit1=Session("admindiscountPerWUnit1")
quantityfrom1=Session("adminquantityfrom1")
quantityUntil1=Session("adminquantityUntil1")
%>

<form method="POST" action="AdminDctQtyPrd.asp" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="5">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
            The discount can be a dollar amount off the &quot;online price&quot; or a percentage off the same price. Select the type of discount below, then enter the &quot;from&quot; and &quot;to&quot; quantity values, and specify the discount value. Refer to the ProductCart <a href="http://www.earlyimpact.com/productcart/support/" target="_blank">User Guide</a> for more information.</td>
		</tr>
		<tr>
			<td colspan="5" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td height="22" colspan="5">  
				<%
				call opendb()
				query="SELECT description, sku, price,configOnly, serviceSpec, btoBprice FROM Products WHERE idproduct="& idproduct
				Set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
					strDescription=rs("description")
					strSKU=rs("sku")
					pcv_dblProductPrice=Cdbl(rs("price"))
					pcv_IntConfigOnly=rs("configOnly")
					pcv_serviceSpec=rs("serviceSpec")
					pcv_dblProductWPrice=Cdbl(rs("btoBprice"))
				set rs=nothing
				call closedb()
				%>
				<h2><a href="FindProductType.asp?id=<%=idProduct%>"><%=strDescription%></a> - Sku: <%=strSKU%></h2>
						<div style="padding-top: 10px;">
							Online Price: <b><%=scCurSign%><%=money(pcv_dblProductPrice)%></b>
							<br>Wholesale Price: <b><%=scCurSign%><%=money(pcv_dblProductWPrice)%></b></br> 
						</div>
				<input type="hidden" name="idproduct" value="<%=idProduct%>">
				<input type="hidden" name="idDiscountPerQuantity" value="<%=idDiscountPerQuantity%>">
				<input type="hidden" name="discountdesc" size="40" value="PD">
			</td>
		</tr>
		<tr> 
			<td colspan="5">Will the discount be based on dollars or percentage?
				<% if percentage="0" then %>
					<input type="radio" name="percentage" value="0" checked class="clearBorder">&nbsp;<%=scCurSign%>  
					<input type="radio" name="percentage" value="-1" class="clearBorder">&nbsp;% 
				<%else %>
					<input type="radio" name="percentage" value="0" class="clearBorder">&nbsp;<%=scCurSign%>  
					<input type="radio" name="percentage" value="-1" checked class="clearBorder">&nbsp;% 
				<% end if %>
			</td>
		</tr>
		<% if pcv_serviceSpec=True then %>
			<input type="hidden" name="baseproductonly" value="-1" checked class="clearBorder">
		<% else %>
		<tr> 
			<td colspan="5"> 
			<% if baseproductonly="-1" then %>
				<input type="radio" name="baseproductonly" value="-1" checked class="clearBorder">
			<% else 
				if baseproductonly="" then %>
					<input type="radio" name="baseproductonly" value="-1" checked class="clearBorder">
				<% else %>
					<input type="radio" name="baseproductonly" value="-1" <%if pcv_IntConfigOnly <> 0 then%>checked<%end if%> class="clearBorder">
				<% end if %>
			<% end if %>
			<%if pcv_IntConfigOnly <> 0 then%>
				Apply discount to base price
			<%else%>
				Apply discount to base price only (product options not included)
			<%end if%></td>
		</tr>
    
		<tr> 
			<td colspan="5"> 
				<%if pcv_IntConfigOnly=0 then%>        
					<% if baseproductonly="0" then %>
						<input type="radio" name="baseproductonly" value="0" checked class="clearBorder">
					<% else %>
						<input type="radio" name="baseproductonly" value="0" class="clearBorder">
					<% end if %>
						Apply discount to base price + options prices (if any)
				<%end if%>
			</td>
		</tr>
	<% end if %>  

	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
    
	<tr> 
		<th width="64">&nbsp;</th>
		<th width="99">From</th>
		<th width="85">To</th>
		<th width="92"><%=scCurSign%> or % (retail)</th>
		<th width="154"><%=scCurSign%> or % (wholesale)</th>
	</tr>
	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td width="64" height="29"><b>Quantity: 
		</b></td>
		<td width="99" height="29">
		<input name="quantityFrom1" size="6" value="<%=quantityFrom1%>">
		</td>
		<td width="85" height="29"> 
		<input name="quantityUntil1" size="6" value="<%=quantityUntil1%>">
		</td>
		<td width="92" height="29"> 
		<input name="discountPerUnit1" size="10" value="<%=money(discountPerUnit1)%>">
		</td>
		<td width="154" height="29"> 
		<input name="discountPerWUnit1" size="10" value="<%=money(discountPerWUnit1)%>">
		</td>
	</tr>
	<tr> 
		<td colspan="5">&nbsp;</td>
	</tr>

	<tr> 
		<td colspan="5" align="center">
			<input type="submit" name="Submit" value="Add First Tier" class="submit2">&nbsp;
		</td>
	</tr>
</table>
</form>
<%END IF 'CanNotRun%>
<!--#include file="AdminFooter.asp"-->