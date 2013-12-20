<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

if validNum(pFeaturedCategory) then
	' Get data about the featured subcategory
	query="SELECT categoryDesc, [image], largeimage, SDesc FROM categories WHERE idCategory=" &pFeaturedCategory&";"
	SET rsTemp=Server.CreateObject("ADODB.RecordSet")
	SET rsTemp=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsTemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	pcStrCategoryDesc=replace(rsTemp("categoryDesc"), """", "&quot;")
	pcStrCategoryDesc=replace(pcStrCategoryDesc, "&amp;", "&")
	pImage=rsTemp("image")
	plargeImage=rsTemp("largeimage")
	if pFeaturedCategoryImage=0 then
		pFeaturedCatImage=pImage
		else
		pFeaturedCatImage=plargeImage
	end if
	pcStrCategorySDesc=rsTemp("SDesc")							
	set rsTemp=nothing
	'// Call SEO Routine
	pcGenerateSeoLinks
	'//
%>
		<tr>
			<td<% if pIntCategoryColumns>1 then %> colspan="<%=pIntCategoryColumns%>"<% end if %>>
			<p><%=dictLanguage.Item(Session("language")&"_viewCategories_4")%>&quot;<%=pCategoryName%>&quot;<%=dictLanguage.Item(Session("language")&"_viewCategories_5")%></p>
			</td>
		</tr>
		<tr>
			<td<% if pIntCategoryColumns>1 then %> colspan="<%=pIntCategoryColumns%>"<% end if %>>
				<table class="pcShowCategoryP">
					<tr>
						<td class="pcShowCategoryImage">
							<%if pFeaturedCatImage<>"" then%>
								<a href="<%=pcStrFeaturedCatLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="3" then%>onmouseover="javascript:document.getCatPre.idcategory.value='<%=pFeaturedCategory%>'; sav_CatPrecallxml='1'; return runPreCatXML('cat_<%=pFeaturedCategory%>');" onmouseout="javascript: sav_CatPrecallxml=''; hidetip();"<%end if%>><img src="catalog/<%=pFeaturedCatImage%>" alt="<%=pcStrCategoryDesc%>"></a>
							<%end if%>
						</td>
					<td class="pcShowCategoryInfoP">
						<p>
							<a href="<%=pcStrFeaturedCatLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="3" then%>onmouseover="javascript:document.getCatPre.idcategory.value='<%=pFeaturedCategory%>'; sav_CatPrecallxml='1'; return runPreCatXML('cat_<%=pFeaturedCategory%>');" onmouseout="javascript: sav_CatPrecallxml=''; hidetip();"<%end if%>><%=pcStrCategoryDesc%></a>
			                <!-- Load category discount icon -->
			                <%intIdCategory=pFeaturedCategory%>
			                <!--#Include File="pcShowCatDiscIcon.asp" -->
						</p>
						<%		
						' Show short category description
						if not pcStrCategorySDesc="" then%>
						<p>
							<%=pcStrCategorySDesc%>
						</p>
						<%end if%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
<% end if %>