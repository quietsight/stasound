<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
pageTitle="Purge All Inactive Products and Related Orders from Database"
Server.ScriptTimeout = 5400
%>
<!--#include file="AdminHeader.asp"-->
<% 
on error resume next
If request("submit")<>"" then

	dim query, conntemp, rs, i
	call openDb()
	
	' get products
	query="SELECT products.idProduct FROM products WHERE products.active = 0;"
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conntemp.execute(query)
	if rs.eof then
		set rs=nothing
		call closeDB()
		response.redirect "PurgeInactiveProducts.asp?msg=2"
	end if
	
	pcProductsArray = rs.getRows()
	Dim pcProudctsCount
	pcProudctsCount = ubound(pcProductsArray,2)
	i = 0
	pcProudctsDeletedCount = 0
	FOR i=0 TO pcProudctsCount
	
			pidproduct=pcProductsArray(0,i)
			
			' delete from taxPrd
			query="DELETE FROM taxPrd WHERE idProduct=" &pidproduct
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
		
			' delete product from configSpec_products
			query="DELETE FROM configSpec_products WHERE configProduct=" &pIdProduct
			set rs=conntemp.execute(query)
		
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from configSpec_categories
			query="DELETE FROM configSpec_categories WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from cs_relationships
			query="DELETE FROM cs_relationships WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from categories_products
			query="DELETE FROM categories_products WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from wishlist
			query="DELETE FROM wishList WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from options_optionsGroups
			query="DELETE FROM options_optionsGroups WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from configSessions	
			query="DELETE FROM configSessions WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from discountsPerQuantity	
			query="DELETE FROM discountsPerQuantity WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			' delete product from ProductsOrdered
			query="SELECT idOrder FROM ProductsOrdered WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			do until rs.eof
				tempIdOrder=rs("idOrder")
				query="DELETE FROM ProductsOrdered WHERE idOrder=" &tempIdOrder
				set rs2=Server.CreateObject("ADODB.Recordset")
				set rs2=conntemp.execute(query)
					if err.number <> 0 then
						errDesc = err.description 
						set rs=nothing
						call closeDb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error 10a purging product on PurgeInactiveProducts.asp. Details: " & errDesc) 
					end If
					
				query="DELETE FROM creditCards WHERE idOrder=" &tempIdOrder
				set rs2=conntemp.execute(query)
					if err.number <> 0 then
						errDesc = err.description 
						set rs=nothing
						call closeDb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error 10b purging product on PurgeInactiveProducts.asp. Details: " & errDesc) 
					end If
					
				query="DELETE FROM offlinepayments WHERE idOrder=" &tempIdOrder
				set rs2=conntemp.execute(query)
					if err.number <> 0 then
						errDesc = err.description 
						set rs=nothing
						call closeDb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error 10c purging product on PurgeInactiveProducts.asp. Details: " & errDesc) 
					end If
					
				query="DELETE FROM Orders WHERE idOrder=" &tempIdOrder
				set rs2=conntemp.execute(query)
					if err.number <> 0 then
						errDesc = err.description 
						set rs=nothing
						call closeDb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error 10d purging product on PurgeInactiveProducts.asp. Details: " & errDesc) 
					end If
					
				set rs2=nothing
			rs.movenext
			loop
			
			' delete product from products table
			query="DELETE FROM products WHERE idProduct=" &pIdProduct
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				pcErrorNumber = err.number
				pcErrorDescription = err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Inactive Products: " & pcErrorNumber & " - " & pcErrorDescription) 
			end If
			
			pcProudctsDeletedCount = pcProudctsDeletedCount + 1
	NEXT
	
	set rs=nothing
	call closeDb()
	response.redirect "PurgeInactiveProducts.asp?msg=1&n="&pcProudctsDeletedCount
	
ELSE 

			msg=request.QueryString("msg")
			if validNum(msg) then
				pcProudctsDeletedCount=request.QueryString("n")
				if msg="2" then
					response.write "<div class=pcCPmessage>There are no inactive products to remove. <a href=purgeProducts.asp>Other options</a>.</div>"
				else
					response.write "<div class=pcCPmessageSuccess>"& pcProudctsDeletedCount & " inactive products were permanently removed from the database. <a href=menu.asp>Continue</a>.</div>"
				end if
			else
				
				call openDB()
				query="SELECT products.idProduct FROM products WHERE products.active = 0;"
				set rs = Server.CreateObject("ADODB.Recordset")
				set rs = conntemp.execute(query)
				productCount=0
				if rs.eof then
					productCount=0
				else
					do while NOT rs.eof
						productCount = productCount + 1
					rs.movenext
					loop
				end if
				set rs = nothing
				call closeDB()

%>

			<form action="PurgeInactiveProducts.asp" method="post" name="form" id="form" class="pcForms">
				<table class="pcCPcontent">
					<tr> 
						<td>
                          <div class="pcCPmessage" style="color: #F00; font-weight: bold; font-size: 18px;">WARNING: This is a dangerous feature.</div>
                          <p>This feature allows you to completely remove <strong>all inactive products</strong> from your database. When you take this action you will <u>completely <strong>remove all inactive products</strong></u> and <u><strong>all orders containing those products</strong></u>. This action is permanent and cannot be undone.</p>                          
                          <p style="margin-top: 8px;">There are <strong><%=productCount%></strong> inactive products in the database. Here are the <% if productCount>4 then%> first 5 <%end if%>inactive products.</p>
                          
							<% '// Start Inactive Products Preview
								if productCount > 0 then
								
								call openDB()
								err.clear
								err.number=0
				
								query="SELECT idProduct, sku, description FROM products WHERE products.removed <> 0 ORDER BY idProduct;"	
								set rs2=Server.CreateObject("ADODB.Recordset")     
								rs2.PageSize=5
				
								rs2.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
								'// Page Count
								iProductsPageCount=rs2.PageCount
								if err.number <> 0 then
									set rs2=nothing
									call closedb()
									response.redirect "techErr.asp?error="& Server.Urlencode("Error in PurgeRemovedProducts.asp: "&Err.Description) 
								end if
								prdStr=""
								idProduct=0
								wCount = 0
								if NOT rs2.eof then %>
										<ul>
				
										<%	do while not rs2.eof and wCount < rs2.PageSize
											idProduct=rs2("idProduct")
											prdStr=rs2("description")&" ("&rs2("sku")&")"	%>
						
											<li><a href="FindProductType.asp?id=<%=idProduct%>" target="_blank"><%response.write prdStr%></a></li>
											<%
											wCount=wCount+1
											rs2.moveNext
										loop
										set rs2=nothing					
										%>
										</ul>
									</p>
							<% 	 end if 
							  	call closedb()
								end if
                            '// End Inactive Products Preview
							%>
 
                          <hr>
                          
                          
                          <p style="margin-top: 8px;">Other options:</p>
                          <ul>
                            <li><a href="LocateProducts.asp?cptype=0">Locate a product and delete it</a> (<em>the product remains in the database, but it's hidden</em>)</li>
                            <li><a href="PurgeProducts.asp">Permanently remove a previously deleted product</a></li>
                            <li><a href="PurgeRemovedProducts.asp">Permanently remove all previously deleted products</a></li>
                            <li><a href="PurgeAllProducts.asp">Permanently remove all products</a></li>
                          </ul>
						</td>
					</tr>
					<tr>
						<td class="pcCPspacer"></td>
					</tr>			
					<tr> 
						<td align="center">
						<input type="submit" value="Remove ALL Inactive Products" name="submit" class="submit2" OnClick="return confirm('You are about to permanently remove all INACTIVE products from your database. Orders that contain those products will be removed as well. Are you sure that you wish to continue? This action cannot be undone. Click OK to continue or Cancel to abort.');" />&nbsp;
						<input type="button" value="Back" onClick="javascript:history.back()" />
						</td>
					</tr>
					<tr>
						<td class="pcCPspacer"></td>
					</tr>	
				</table>
		</form>
		
		<%
		end if
end if %>
<!--#include file="AdminFooter.asp"-->