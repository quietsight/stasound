<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<%
on error resume next
dim f, mySQL, conntemp, rstemp
call opendb()
%>
<% pageTitle="Generate Store Links" %>
<% section="layout" %>
<!--#include file="AdminHeader.asp"-->
            
<%
sMode=request.Form("submit1")
if sMode <> "" then
	sMode="1"
	idproduct=request.Form("product")
end If

cMode=request.Form("submit2")
if cMode <> "" then
	cMode="1"
	idcategory=request.Form("CategoryLinkId")
	query="SELECT categoryDesc FROM categories WHERE idCategory="&idcategory
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	categoryName=rs("categoryDesc")
	set rs=nothing
end If
%>

	<form method="post" name="addCateg" action="genLinksa.asp" class="pcForms">
		<table class="pcCPcontent">
			<tr> 
				<td colspan="2"><img src="images/edit2.gif" width="25" height="23" align="left" hspace="5">Click on this button and the link will be copied to your clipboard so that you can use it in your favorite HTML editor.</td>
			</tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<th colspan="2">Generate Product Links</th>
			</tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td colspan="2">Use this feature to obtain the link to the product details page for the selected product. Select a prduct:</td>
			</tr>
			
			<%
			query="SELECT idproduct,description FROM products WHERE removed=0 AND active=-1 AND configOnly=0 ORDER BY description ASC"
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=conntemp.execute(query)
			if err.number <> 0 then
				set rstemp=nothing
				call closeDb()
				response.redirect "techErr.asp?error="&Server.Urlencode("Error generating product drop-down")
			end If
			
			if not rstemp.EOF then
				prdArray = rstemp.getRows()
				set rstemp = nothing
				intCount=ubound(prdArray,2)
			%>
					<tr> 
						<td colspan="2">
							<select name="product">
							<% for i=0 to intCount%>
								<option value="<%=prdArray(0,i)%>" <% if sMode="1" And Cint(idproduct)=Cint(prdArray(0,i)) then%>selected<%end if%>><%=prdArray(1,i)%></option>
							<%
									if sMode="1" And Cint(idproduct)=Cint(prdArray(0,i)) then
										pDescription = prdArray(1,i)
									end if
							%>
							<% next %>
							</select>
						</td>
					</tr>
					<tr> 
						<td colspan="2" class="normal">  
							<input name="submit1" type="submit" value="Generate Link">
						</td>
					</tr>
					<% If sMode="1" then %>
					<tr> 
						<td colspan="2"><b>Link for <%=pDescription%>:</b></td>
					</tr>
					<tr> 
						<td>
							<a class="highlighttext" href="javascript:HighlightAll('addCateg.link1')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
						</td>
						<td>
							<%
							tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/viewPrd.asp?idproduct="&idproduct),"//","/")
							tempURL=replace(tempURL,"http:/","http://")
							%>
							<input type="text" name="link1" size="65" value="<%=tempURL%>"> <a href="<%=tempURL%>" target="_blank">View</a>
						</td>
					</tr>
					<% 	if scSeoURLs=1 then %>
					<tr> 
						<td colspan="2"><b>Search Engine Friendly link for <%=pDescription%>:</b></td>
					</tr>
					<tr> 
						<td>
							<a class="highlighttext" href="javascript:HighlightAll('addCateg.link2')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
						</td>
						<td>
							<%
							'// SEO Links
							'// Build Navigation Product Link
							'// Get the first category that the product has been assigned to, filtering out hidden categories
								query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& IDproduct &" AND categories.iBTOhide<>1 AND categories.pccats_RetailHide<>1"
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=conntemp.execute(query)
								if not rs.EOF then
									pIdCategory=rs("idCategory")
								else
									pIdCategory=1
								end if
								set rs=nothing
							pcStrPrdLink=pDescription & "-" & pIdCategory & "p" & IDproduct & ".htm"
							pcStrPrdLink=removeChars(pcStrPrdLink)
							'//
							tempURL2=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrPrdLink),"//","/")
							tempURL2=replace(tempURL2,"http:/","http://")
							%>
							<input type="text" name="link2" size="65" value="<%=tempURL2%>"> <a href="<%=tempURL2%>" target="_blank">View</a>
						</td>
					</tr>
					<% end if %>
				<% end If
				else %>
				<tr> 
					<td colspan="2">
					<div class="pcCPmessage">There are currently no products in your store.</div>
					</td>
				</tr>
				<%
				end if
				%>
			</table>
		</form>
		<br>
		<form method="post" name="addCatLink" action="genLinksa.asp" class="pcForms">
		<table class="pcCPcontent">
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<th colspan="2">Generate Category Links</th>
			</tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td colspan="2">Use this feature to obtain the link to the category details page for the selected category.</td>
			</tr>
			<tr>
				<td colspan="2">
					<%
					cat_DropDownName="CategoryLinkId"
					cat_Type="0"
					cat_DropDownSize="1"
					cat_MultiSelect="0"
					cat_ExcBTOHide="0"
					cat_StoreFront="0"
					cat_ShowParent="1"
					cat_DefaultItem=""
					cat_SelectedItems="1,"
					cat_ExcItems=""
					%>
					<!--#include file="../includes/pcCategoriesList.asp"-->
					<%call pcs_CatList()%>
				</td>
			</tr>
			
					<tr> 
						<td colspan="2" class="normal">  
							<input name="submit2" type="submit" value="Generate Link">
						</td>
					</tr>
					<% If cMode="1" then %>
					<tr> 
						<td colspan="2"><b>Category link for "<%=categoryName%>":</b></td>
					</tr>
					<tr> 
						<td width="6%">
							<a class="highlighttext" href="javascript:HighlightAll('addCatLink.link20')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
						</td>
						<td width="94%">
							<%
							tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/viewcategories.asp?idCategory="&idcategory),"//","/")
							tempURL=replace(tempURL,"http:/","http://")
							%>
							<input type="text" name="link20" size="65" value="<%=tempURL%>"> <a href="<%=tempURL%>" target="_blank">View</a>
						</td>
					</tr>
					<% 	if scSeoURLs=1 then %>
					<tr> 
						<td colspan="2"><b>Search Engine Friendly Category link for "<%=categoryName%>":</b></td>
					</tr>
					<tr> 
						<td width="6%">
							<a class="highlighttext" href="javascript:HighlightAll('addCatLink.link21')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
						</td>
						<td width="94%">
							<%
							'// SEO Links
							'// Build Navigation Category Link
							pcStrCatLink=categoryName & "-c" & idcategory & ".htm"
							pcStrCatLink=removeChars(pcStrCatLink)
							'//
							tempURL2=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrCatLink),"//","/")
							tempURL2=replace(tempURL2,"http:/","http://")
							%>
							<input type="text" name="link21" size="65" value="<%=tempURL2%>"> <a href="<%=tempURL2%>" target="_blank">View</a>
						</td>
					</tr>
					<% end if %>
					<tr> 
						<td colspan="2">Note: if you <a href="<%=tempURL%>" target="_blank">test this link</a> in the storefront and receive a message that says: "This is not a valid category", this typically means that the category has been setup as a <em>hidden category</em> (<a href="modCata.asp?idcategory=<%=idcategory%>">edit the category</a>).</td>
					</tr>
				<% end If %>
			</table>
		</form>

		<br>
		<form name="links" method="post" action="" class="pcForms">
		<table class="pcCPcontent">
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<th colspan="2" class="pcCPsectionTitle">Links to Popular Storefront Pages</th>
			</tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td colspan="2">Default Home Page:</td>
			</tr>
			<tr> 
				<td width="2%"><a class="highlighttext" href="javascript:HighlightAll('links.link2')"><img src="images/pcIconClone.jpg"></a></td>
				<td width="98%">
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/home.asp"),"//","/")
						tempURL=replace(tempURL,"http:/","http://") %>
					<input type="text" name="link2" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Specials:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link3')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/showspecials.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link3" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">New Arrivals:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link8')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<%
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/shownewarrivals.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
					<input type="text" name="link8" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Best Sellers:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link9')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/showbestsellers.asp"),"//","/")
						tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link9" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Show Recently Reviewed Products:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link30')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/showrecentlyreviewed.asp"),"//","/")
						tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link30" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">View Cart:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link4')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<%
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/viewcart.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
					<input type="text" name="link4" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Contact Us form:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link13')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<%
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/contact.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
					<input type="text" name="link13" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Browse by Category:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link5')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/viewcategories.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link5" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Browse <a href="cmsManage.asp" target="_blank">Content Pages</a>:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link31')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/viewpages.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link31" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Advanced Search Page:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link7')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/search.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link7" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr>
				<td colspan="2">Gift Registry Search Page:</td>
			</tr>
			<tr>
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link12')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/ggg_srcGR.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link12" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Customer Login:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link6')"><img src="images/pcIconClone.jpg"></a> 
				</td>
				<td>
				<%
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/Checkout.asp?cmode=1"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
					<input type="text" name="link6" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Drop-Shipper Login:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link20')"><img src="images/pcIconClone.jpg"></a> 
				</td>
				<td>
				<%
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/sds_Login.asp"),"//","/")
					tempURL=replace(tempURL,"http:/","http://") %>
					<input type="text" name="link20" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">New Affiliate Sign Up Form:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link10')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/NewAffa.asp"),"//","/")
						tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link10" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2">Affiliate Login:</td>
			</tr>
			<tr> 
				<td><a class="highlighttext" href="javascript:HighlightAll('links.link11')"><img src="images/pcIconClone.jpg"></a></td>
				<td>
				<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/AffiliateLogin.asp"),"//","/")
						tempURL=replace(tempURL,"http:/","http://") %>
				<input type="text" name="link11" size="65" value="<%=tempURL%>">
				</td>
			</tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
		</table>
	</form>
<% call closedb()%><!--#include file="AdminFooter.asp"-->