<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Modify Quantity Discounts (Tiered Pricing)" %>
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
Dim rs, connTemp, query, idproduct, discountdesc, percentage, baseproductonly, discountPerUnit1, discountPerWUnit1, quantityfrom1, 	nquantityfrom, nquantityUntil, ndiscountPerUnit

call opendb()

CanNotRun=0

pIDProduct=Request("idproduct")
if pIDProduct="" OR IsNull(pIDProduct) then
	pIDProduct=0
end if

'~~~~~~~~~~~~~~ Delete last tier only ~~~~~~~~~~~~~~~~~~~~~~~
if request("sAction")="D" then
	intId=request("Id")
	idproduct=request("idProduct")
	call openDb()
	
	query="DELETE FROM discountsPerQuantity WHERE idDiscountPerQuantity="&intId&";"
	Set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query) 
	
	set rs=nothing
	call closedb()
	
	response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct
end if
'~~~~~~~~~~~~~~ Delete last tier only ~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~ DELETE ~~~~~~~~~~~~~~~~~~~~~~~
dMode=Request.QueryString("Delete")
if dMode<>"" then
	call openDb()
	Session("adminidproduct")=Request.QueryString("idproduct")
	idproduct=Session("adminidproduct")
	Set rs=Server.CreateObject("ADODB.Recordset")
	query="DELETE FROM discountsPerQuantity WHERE idProduct="&idProduct 
	set rs=conntemp.execute(query)
	Set rs=nothing
	
	Session("adminidproduct")=""
	Session("admindiscountdesc")=""
	Session("admindiscountPerUnit1")=""
	Session("adminquantityfrom1")=""
	Session("adminidDiscountPerQuantity1")=""
	call closedb()
	response.redirect "modifyProduct.asp?idproduct="&idProduct
end if
'~~~~~~~~~~~~~~ DELETE ~~~~~~~~~~~~~~~~~~~~~~~

'// Check for conflict with Product Promotions

query="SELECT DISTINCT idproduct FROM pcPrdPromotions WHERE idproduct=" & pIDProduct & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	CanNotRun=1%>
	<table class="pcCPcontent">
		   <tr>
				<td colspan="3">
					<div class="pcCPmessage">You cannot add quantity discounts to this product because it has a promotion assigned to it. <a href="ModPromotionPrd.asp?idproduct=<%=pIdProduct%>&iMode=start">Review the promotion</a>.</div>
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

'~~~~~~~~~~~~~~ UPDATE W/O ADDING ~~~~~~~~~~~~~~~~~~~~~~~
uMode=Request.Form("SubmitUPD")
If uMode<>"" Then
	call opendb()
	ndiscountdesc=Request("discountdesc")
	npercentage=Request("percentage")
	nbaseproductonly=Request("baseproductonly")
	idproduct=Request("idproduct")
	query="UPDATE discountsPerQuantity SET percentage="&npercentage&", baseproductonly="&nbaseproductonly&" WHERE idProduct="&idProduct
	Set rs=server.CreateObject("ADODB.RecordSet")
	Set rs=conntemp.execute(query)
	Set rs=Nothing
	call closeDb()
	response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct
End If

'~~~~~~~~~~~~~ ADD ~~~~~~~~~~~~~~~~~~
sMode=Request.Form("Submit")
If sMode<>"" Then
	iNextNum=request("iNextNum")
	iPrevNum=request("iPrevNum")
	idproduct=Request("idproduct")
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
		response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct&"&msg="&msg
	end if
	'check to make sure there are no overlaps
	if nquantityfrom <> "" AND nquantityUntil <> "" AND ndiscountPerUnit="" then
		msg="You must specify a discount price for each tier."
		response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct&"&msg="&msg
	end if
	'make sure the from < until
	if int(nquantityfrom)>int(nquantityUntil) then
		msg="Your quantity 'To' must be greater then the 'To' in the previous Tier."
		response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct&"&msg="&msg
	end if
	
	If (money(ndiscountPerUnit) > 0 OR money(ndiscountPerWUnit)>0) AND nquantityfrom <> "" AND nquantityUntil <> "" AND nidDiscountPerQuantity="" Then
		call opendb()
		query="INSERT INTO discountsPerQuantity (idproduct,idcategory,discountDesc,discountPerUnit,discountPerWUnit,quantityuntil,quantityfrom,num,percentage,baseproductonly) VALUES ("&idproduct&",0,'PD',"&ndiscountPerUnit&","&ndiscountPerWUnit&","&nquantityuntil&","&nquantityfrom&","&iPriority&","&npercentage&","&nbaseproductonly&");"
		Set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		Set rs=Nothing
		call closedb()
	End If
	response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct
End If
'~~~~~~~~~~~~~~ ADD ~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~ SHOW ADMIN ~~~~~~~~~~~~~~~~~~~~~~~
idproduct=request("idproduct")
call opendb()
query="SELECT discountdesc,percentage,baseproductonly,num,discountPerUnit,discountPerWUnit,quantityfrom,quantityUntil,discountPerUnit,idDiscountPerQuantity  FROM discountsPerQuantity WHERE idproduct="&idproduct&" AND discountdesc='PD' ORDER BY num;"
Set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 

	if rs.eof then
		set rs=nothing
		call closedb()
		response.Redirect "AdminDctQtyPrd.asp?idproduct="&idproduct
	end if

	discountdesc=rs("discountdesc")
	Session("adminpercentage")=rs("percentage")
	Session("adminbaseproductonly")=rs("baseproductonly")
	%>
	<form method="POST" action="ModDctQtyPrd.asp" class="pcForms">
		<table class="pcCPcontent">
			<tr> 
				<td colspan="5">
					<% 'get product info
					query="SELECT description,serviceSpec,sku,configonly,price,BtoBPrice FROM products WHERE idproduct="&idproduct
					set rsPrdObj=server.CreateObject("ADODB.RecordSet")
					set rsPrdObj=conntemp.execute(query)
					strDescription=rsPrdObj("description")
					pServiceSpec=rsPrdObj("serviceSpec")
					StrSKU=rsPrdObj("sku")
					configonly=rsPrdObj("configonly")
					pcv_dblProductPrice=Cdbl(rsPrdObj("price"))
					pcv_dblProductWPrice=Cdbl(rsPrdObj("btoBprice"))
					set rsPrdObj=nothing
					%>
                    
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>

					<h2><a href="FindProductType.asp?id=<%=idProduct%>"><%=strDescription%></a> - Sku: <%=strSKU%></h2>
						<div style="padding-top: 10px;">
							Online Price: <b><%=scCurSign%><%=money(pcv_dblProductPrice)%></b>
							<br>Wholesale Price: <b><%=scCurSign%><%=money(pcv_dblProductWPrice)%></b></br> 
						</div>
					<input type="hidden" name="discountdesc" value="PD">
					<input type="hidden" name="idproduct" value="<%=idproduct%>">
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
							r=rs("num")
							discountPerUnit=rs("discountPerUnit")
							discountPerWUnit=rs("discountPerWUnit")
							quantityfrom=rs("quantityfrom")
							quantityUntil=rs("quantityUntil")
							discountPerUnit=rs("discountPerUnit")
							idDiscountPerQuantity=rs("idDiscountPerQuantity")
							%>
							<tr>
								<td>
									<%
									if iDCnt=0 then
										response.write "Quantity:"
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
								<td align="right"><a href="ModAllDctQtyPrd.asp?idProduct=<%=idProduct%>">Edit</a></td>
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
							<td align="right"><% if iDCnt>1 then %><a href="ModDctQtyPrd.asp?Id=<%=idDiscountPerQuantity%>&idProduct=<%=idProduct%>&sAction=D">Delete Last Tier</a><% end if %></td>
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
				<input type="button" name="Delete" value="Delete discount" onClick="javascript:if (confirm('You are about to permanantly delete this discount from the database. Are you sure you want to complete this action?')) location='moddctQtyPrd.asp?Delete=Yes&idproduct=<%=idProduct%>'">
			</td>
			</tr>
			<tr>
				<td colspan="5"><hr></td>
			</tr>
			<tr> 
				<td colspan="5" align="center">
				<input type="button" name="Apply" value="Apply to Other Products" onClick="location.href='ApplyDctToPrds.asp?idproduct=<%=idProduct%>'">&nbsp;
				<input type="button" value="Locate Another Product" onClick="location.href='viewDisca.asp'">
				</td>
			</tr>
			<tr>
				<td colspan="5" class="pcCPspacer"></td>
			</tr>
		</table>
	</form>
<%END IF 'CanNotRun%>
<!--#include file="AdminFooter.asp"-->