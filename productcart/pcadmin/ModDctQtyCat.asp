<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Modify Category-based Quantity Discounts (Tiered Pricing)" %>
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
Dim rs, connTemp, query, idcategory, discountdesc, percentage, baseproductonly, discountPerUnit1, discountPerWUnit1, quantityfrom1, 	nquantityfrom, nquantityUntil, ndiscountPerUnit

call opendb()

CanNotRun=0

pIDCategory=Request("idcategory")
if pIDCategory="" OR IsNull(pIDCategory) then
	pIDCategory=0
end if

query="SELECT DISTINCT products.idproduct FROM products INNER JOIN pcPrdPromotions ON products.idproduct=pcPrdPromotions.idproduct WHERE products.removed=0 AND pcPrdPromotions.idproduct IN (SELECT DISTINCT idproduct FROM categories_products WHERE idCategory=" & pIDCategory & ");"
set rs=connTemp.execute(query)

if not rs.eof then
CanNotRun=1%>
<table class="pcCPcontent">
       <tr>
        	<td colspan="3">
				<div class="pcCPmessage">You cannot add quantity discounts to this category because one or more of the products already had promotions assigned to them. <a href="PromotionPrdSrc.asp">Review it</a></div>
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" name="back" value=" Category Quantity Discounts " onclick="location='viewCatDisc.asp';" class="ibtnGrey">
				&nbsp;&nbsp;<input type="button" name="back2" value=" Back to Main menu " onclick="location='menu.asp';" class="ibtnGrey">
			</td>
		</tr>		
</table>
<%
end if
set rs=nothing

IF CanNotRun=0 THEN

'~~~~~~~~~~~~~~ Delete last tier only ~~~~~~~~~~~~~~~~~~~~~~~
if request("sAction")="D" then
	intId=request("Id")
	idcategory=request("idCategory")
	call openDb()
	
	query="DELETE FROM pcCatDiscounts WHERE pcCD_idDiscount="&intId&";"
	Set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query) 
	
	set rs=nothing
	call closedb()
	
	response.redirect "ModDctQtyCat.asp?idcategory="&idcategory
end if
'~~~~~~~~~~~~~~ Delete last tier only ~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~ DELETE ~~~~~~~~~~~~~~~~~~~~~~~
dMode=Request.QueryString("Delete")
if dMode<>"" then
	call openDb()
	Session("adminidcategory")=Request.QueryString("idcategory")
	idcategory=Session("adminidcategory")
	Set rs=Server.CreateObject("ADODB.Recordset")
	query="DELETE FROM pcCatDiscounts WHERE pcCD_idcategory="&idcategory 
	set rs=conntemp.execute(query)
	Set rs=nothing
	
	Session("adminidcategory")=""
	Session("admindiscountdesc")=""
	Session("admindiscountPerUnit1")=""
	Session("adminquantityfrom1")=""
	Session("adminidDiscountPerQuantity1")=""
	call closedb()
	response.redirect "viewCatDisc.asp"
end if
'~~~~~~~~~~~~~~ DELETE ~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~ UPDATE W/O ADDING ~~~~~~~~~~~~~~~~~~~~~~~
uMode=Request.Form("SubmitUPD")
If uMode<>"" Then
	call opendb()
	ndiscountdesc=Request("discountdesc")
	npercentage=Request("percentage")
	nbaseproductonly=Request("baseproductonly")
	idcategory=Request("idcategory")
	query="UPDATE pcCatDiscounts SET pcCD_percentage="&npercentage&", pcCD_baseproductonly="&nbaseproductonly&" WHERE pcCD_idcategory="&idcategory
	Set rs=server.CreateObject("ADODB.RecordSet")
	Set rs=conntemp.execute(query)
	Set rs=Nothing
	call closeDb()
	response.redirect "ModDctQtyCat.asp?idcategory="&idcategory
End If

'~~~~~~~~~~~~~ ADD ~~~~~~~~~~~~~~~~~~
sMode=Request.Form("Submit")
If sMode<>"" Then
	iNextNum=request("iNextNum")
	iPrevNum=request("iPrevNum")
	idcategory=Request("idcategory")
	ndiscountdesc=Request("discountdesc")
	npercentage=Request("percentage")
	nbaseproductonly=Request("baseproductonly")	
	ndiscountPerUnit=replacecomma(Request("discountPerUnitAdd"&iNextNum))	
	ndiscountPerWUnit=replacecomma(Request("discountPerWUnitAdd"&iNextNum))
	if ndiscountPerUnit="" then
		ndiscountPerUnit="0"
	end if
	if ndiscountPerWUnit="" then
		ndiscountPerWUnit="0"
	end if
	idr=request("ID"&iPrevNum)
	iPriority=int(iPrevNum)+1
	nquantityFrom=int(Request("quantityUntil"&idr))+1
	nquantityUntil=Request("quantityuntilAdd"&iNextNum)
	nidDiscountPerQuantity=Request("idDiscountPerQuantityAdd"&iNextNum)
	
	'check to make sure there are no overlaps
	if nquantityfrom = "" OR nquantityUntil = "" OR ndiscountPerUnit="" then
		msg="Both the 'To' and 'Retail Price/Percentage' fields are required."
		response.redirect "ModDctQtyCat.asp?idcategory="&idcategory&"&msg="&msg
	end if
	'check to make sure there are no overlaps
	if nquantityfrom <> "" AND nquantityUntil <> "" AND ndiscountPerUnit="" then
		msg="You must specify a discount price for each tier"
		response.redirect "ModDctQtyCat.asp?idcategory="&idcategory&"&msg="&msg
	end if
	'make sure the from < until
	if int(nquantityfrom)>int(nquantityUntil) then
		msg="Your quantity 'To' must be greater then the 'To' in the previous Tier."
		response.redirect "ModDctQtyCat.asp?idcategory="&idcategory&"&msg="&msg
	end if
	
	If (money(ndiscountPerUnit) > 0 OR money(ndiscountPerWUnit)>0) AND nquantityfrom <> "" AND nquantityUntil <> "" AND nidDiscountPerQuantity="" Then
		call opendb()
		query="INSERT INTO pcCatDiscounts (pcCD_idcategory,pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_quantityuntil,pcCD_quantityfrom,pcCD_num,pcCD_percentage,pcCD_baseproductonly) VALUES ("&idcategory&","&ndiscountPerUnit&","&ndiscountPerWUnit&","&nquantityuntil&","&nquantityfrom&","&iPriority&","&npercentage&","&nbaseproductonly&");"
		Set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		Set rs=Nothing
		call closedb()
	End If
	response.redirect "ModDctQtyCat.asp?idcategory="&idcategory
End If
'~~~~~~~~~~~~~~ ADD ~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~ SHOW ADMIN ~~~~~~~~~~~~~~~~~~~~~~~
idcategory=request("idcategory")
call opendb()
query="SELECT * FROM pcCatDiscounts WHERE pcCD_idcategory="&idcategory&" ORDER BY pcCD_num"
Set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 

	if rs.eof then
		set rs=nothing
		call closedb()
		response.Redirect "AdminDctQtyCat.asp?idcategory="&idcategory
	end if

	Session("adminpercentage")=rs("pcCD_percentage")
	Session("adminbaseproductonly")=rs("pcCD_baseproductonly")
%>
<form method="POST" action="ModDctQtyCat.asp" class="pcForms">		
		<table class="pcCPcontent">
			<tr>
				<td colspan="5">	 
				<% 'get category info
				query="SELECT categoryDesc,idcategory FROM categories WHERE idcategory="&idcategory
				set rsPrdObj=server.CreateObject("ADODB.RecordSet")
				set rsPrdObj=conntemp.execute(query)%>
                <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
				<h2>Category Name: <strong><%=rsPrdObj("categoryDesc")%></strong> - ID#: <%=rsPrdObj("idcategory")%></h2>
				<% set rsPrdObj=nothing %>
				<input type="hidden" name="discountdesc" value="PD">
				<input type="hidden" name="idcategory" value="<%=idcategory%>">
				</td>
			</tr>			

			<tr> 
				<td colspan="5">
					<div style="padding: 10px; margin: 10px 0 10px 0; border: 1px solid #e1e1e1;">
						<div style="padding-bottom: 10px;">
							Discount based on:
							<% if Session("adminpercentage")="" then %>
								<input type="radio" name="percentage" value="0" class="clearBorder"><%=scCurSign%> 
								<input type="radio" name="percentage" value="-1" class="clearBorder">% 
							<% else %>
								<% if Session("adminpercentage")="0" then %>
									<input type="radio" name="percentage" value="0" checked class="clearBorder"><%=scCurSign%> 
									<input type="radio" name="percentage" value="-1" class="clearBorder">% 
								<%else %>
									<input type="radio" name="percentage" value="0" class="clearBorder"><%=scCurSign%> 
									<input type="radio" name="percentage" value="-1" checked class="clearBorder">% 
								<% end if %>
							<% end if %>
						</div>
							<% if pServiceSpec=True then %>
								<input type="hidden" name="baseproductonly" value="-1" checked class="clearBorder">
							<% else %>
								<div style="padding-bottom: 5px;">
										<% if Session("adminbaseproductonly")="-1" then %>
											<input type="radio" name="baseproductonly" value="-1" checked class="clearBorder">
										<% else 
												if Session("adminbaseproductonly")="" then %>
												<input type="radio" name="baseproductonly" value="-1" checked class="clearBorder">
											<% else %>
												<input type="radio" name="baseproductonly" value="-1" class="clearBorder">				
											<% end if %>		
										<% end if %>
										<%if configonly <> 0 then%>
											Apply discount to base price
										<%else%>
											Apply discount to base price only (product options not included)
										<%end if%>
								</div>
								<div style="padding-bottom: 10px;">
									<%if configonly<>true then%>					
										<% if Session("adminbaseproductonly")="0" then %>
											<input type="radio" name="baseproductonly" value="0" checked class="clearBorder">
										<% else %>
											<input type="radio" name="baseproductonly" value="0" class="clearBorder">
										<% end if %>
										Apply discount to base price + options prices (if any)
									<%end if%>
								</div>
							<% end if %>
							<input name="SubmitUPD" type="submit" id="SubmitUPD" value="Update" class="submit2">
						</div>
				</td>
			</tr>
			<tr>
				<td colspan="5"></td>
		</tr>			
		<tr> 
			<td colspan="5">
					<table class="pcCPcontent">
								<tr>										
									<th width="16%">Disc. Tiers</th>
									<th width="14%">From</th>
									<th width="15%">To</th>
									<th width="23%"><%=scCurSign%> or	% (retail)</th>
									<th width="27%" colspan="2"><%=scCurSign%> or	% (wholesale)</th>
								</tr>

								<% 
								iDCnt=0
								do until rs.eof
									r=rs("pcCD_num")
									discountPerUnit=rs("pcCD_discountPerUnit")
									discountPerWUnit=rs("pcCD_discountPerWUnit")
									quantityfrom=rs("pcCD_quantityfrom")
									quantityUntil=rs("pcCD_quantityUntil")
									discountPerUnit=rs("pcCD_discountPerUnit")
									idDiscountPerQuantity=rs("pcCD_idDiscount")
									%>										
								<tr>
									<td>
									<% 
									if iDCnt=0 then
										response.write "Product Quantity:"
									else 
										response.write "&nbsp;"
									end if 
									%>
									</td>
									<td>
									<%=quantityFrom%>
									<input type="hidden" name="idDiscountPerQuantity" value="<%=idDiscountPerQuantity%>">
									<input type="hidden" name="ID<%=r%>" value="<%=idDiscountPerQuantity%>">
									<input type="hidden" name="quantityFrom<%=idDiscountPerQuantity%>" value="<%=quantityFrom%>">
									</td>
									<td>
									<%=quantityUntil%>
									<input type="hidden" name="quantityUntil<%=idDiscountPerQuantity%>" value="<%=quantityUntil%>">
									</td>
									<td>
									<%=money(discountPerUnit)%> 
									<input type="hidden" name="discountPerUnit<%=idDiscountPerQuantity%>" value="<%=discountPerUnit%>">
									</td>
									<td>
									<%=money(discountPerWUnit)%>
									<input type="hidden" name="discountPerWUnit<%=idDiscountPerQuantity%>" value="<%=discountPerWUnit%>">
									</td>
									<td align="right"><a href="ModAllDctQtyCat.asp?idcategory=<%=idcategory%>">Edit</a></td>
								</tr>										
								<% 
								iDCnt=iDCnt + 1
								rs.movenext
								loop
								Set rs=Nothing
								call closedb()
								iPrevNum=r
								iNextNum=r+1
								%>
								<tr>										
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>
									<input name="quantityUntilAdd<%=iNextNum%>" type="text" size="6">
									<input type="hidden" name="iNextNum" value="<%=iNextNum%>">
									<input type="hidden" name="iPrevNum" value="<%=iPrevNum%>">
									</td>
									<td><input name="discountPerUnitAdd<%=iNextNum%>" type="text" size="6"></td>
									<td><input name="discountPerWUnitAdd<%=iNextNum%>" type="text" size="6"></td>
									<td align="right"><% if iDCnt>1 then %><a href="ModDctQtyCat.asp?Id=<%=idDiscountPerQuantity%>&idCategory=<%=idCategory%>&sAction=D">Delete Last Tier</a><% end if %></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="5" class="pcCPspacer"></td>
		</tr>
		<tr> 
		<td colspan="5" align="center">
			<input type="submit" name="Submit" value="Add New Tier" class="submit2">&nbsp;
			<input type="button" name="Delete" value="Delete discount" onClick="javascript:if (confirm('You are about to permanantly delete this discount from the database. Are you sure you want to complete this action?')) location='moddctQtyCat.asp?Delete=Yes&idcategory=<%=idcategory%>'">
		</td>
		</tr>
		<tr>
			<td colspan="5"><hr></td>
		</tr>
		<tr> 
			<td colspan="5" align="center">
			<input type="button" name="Apply" value="Apply to Other Categories" onClick="location.href='ApplyDctToCats.asp?idcategory=<%=idcategory%>'">&nbsp;
			<input type="button" value="Locate Another Category" onClick="location.href='viewCatDisc.asp'">
			</td>
		</tr>
		<tr>
			<td colspan="5" class="pcCPspacer"></td>
		</tr>
	</table>
</form>
<%END IF 'CanNotRun%>
<!--#include file="AdminFooter.asp"-->