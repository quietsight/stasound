<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>

<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"-->
<!--#include FILE="../includes/SearchConstants.asp"-->
<!--#include FILE="prv_incFunctions.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<%
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "viewCategories.asp"

'*******************************
' START: Check store on/off, start PC session, check affiliate ID
'*******************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*******************************
' END: Check store on/off, start PC session, check affiliate ID
'*******************************

dim query, conntemp, rstemp
dim pTempIntSubCategory
call opendb()
%>
<!--#include FILE="prv_getSettings.asp"-->
<%

pTempIntSubCategory=session("idCategoryRedirect")
if pTempIntSubCategory = "" then
	pTempIntSubCategory=getUserInput(request("idCategory"),10)
end if

'// Validate Category ID
	if not validNum(pTempIntSubCategory) then
		pTempIntSubCategory=""
	end if
	if pTempIntSubCategory="" or pTempIntSubCategory="0" then
		pTempIntSubCategory=1
	end if
intIdCategory=pTempIntSubCategory
'// Wholesale-only categories
If Session("customerType")=1 Then
	pcv_strTemp=""
else
	pcv_strTemp=" AND pccats_RetailHide<>1"
end if

'*******************************
' START Display Settings
'*******************************

pFeaturedCategory=0
pFeaturedCategoryImage=0

If validNum(pTempIntSubCategory) and pTempIntSubCategory<>1 then
	query="SELECT pcCats_SubCategoryView, pcCats_CategoryColumns, pcCats_CategoryRows, pcCats_PageStyle, pcCats_ProductColumns, pcCats_ProductRows, pcCats_FeaturedCategory, pcCats_FeaturedCategoryImage FROM categories WHERE (((idCategory)="&pTempIntSubCategory&")" & pcv_strTemp &");"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	if rs.EOF then
		set rs=nothing
		call closeDb()
		response.redirect "msg.asp?message=86"
	end if	
	
	Dim pIntSubCategoryView
	Dim pIntCategoryColumns
	Dim pIntCategoryRows
	Dim pIntProductColumns
	Dim pIntProductRows
	
	pIntSubCategoryView=rs("pcCats_SubCategoryView")
	pIntCategoryColumns=rs("pcCats_CategoryColumns")
	pIntCategoryRows=rs("pcCats_CategoryRows")
	pStrPageStyle=rs("pcCats_PageStyle")
	pIntProductColumns=rs("pcCats_ProductColumns")
	pIntProductRows=rs("pcCats_ProductRows")
	pFeaturedCategory=rs("pcCats_FeaturedCategory")
	pFeaturedCategoryImage=rs("pcCats_FeaturedCategoryImage")
	
	set rs=nothing
End if
	
' START Load category-specific values. If empty, use storewide settings

' How sub-categories are displayed
' 	0 = in a list, with images
'	1 = in a list, without images
'	2 = drop-down
'	3 = default
'	4 = thumbnail only
if NOT validNum(pIntSubCategoryView) OR pIntSubCategoryView=3 then
	 pIntSubCategoryView=scCatImages
end if

' How many per row: number of columns
if NOT validNum(pIntCategoryColumns) OR pIntCategoryColumns=0 then
	pIntCategoryColumns=scCatRow
end if

' How many rows per page
if NOT validNum(pIntCategoryRows) OR pIntCategoryRows=0 then
	pIntCategoryRows=scCatRowsPerPage
end if

' How many products per row
if NOT validNum(pIntProductColumns) OR pIntProductColumns=0 then
	pIntProductColumns=scPrdRow
end if

' How many rows per page
if NOT validNum(pIntProductRows) OR pIntProductRows=0 then
	pIntProductRows=scPrdRowsPerPage
end if

' END Load category-specific values


' OVERRIDE page style: check to see if a querystring or a form is sending the page style.
Dim pcPageStyle, strSeoQueryString

pcPageStyle = LCase(getUserInput(Request("pageStyle"),1))

'// Check querystring saved to session by 404.asp
if pcPageStyle = "" then
	strSeoQueryString=lcase(session("strSeoQueryString"))
	if strSeoQueryString<>"" then
		if InStr(strSeoQueryString,"pagestyle")>0 then
			pcPageStyle=left(replace(strSeoQueryString,"pagestyle=",""),1)
		end if
	end if
end if

'// Category Level Settings
if pcPageStyle = "" then
	pcPageStyle = pStrPageStyle
end if

'// Global Settings
if isNULL(pcPageStyle) OR trim(pcPageStyle) = "" then
	pcPageStyle = LCase(bType)
end if

if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

' OTHER display settings
' These variables show/hide information when products are shown with Page Style = L or M
Dim pShowSKU, pShowSmallImg
pShowSKU = scShowSKU ' If 0, then the SKU is hidden
pShowSmallImg = scShowSmallImg ' If 0, then the small image is not shown
' Note: the size of the small image is set via the pcStorefront.css stylesheet

'*******************************
' END Display Settings
'*******************************


if pFeaturedCategory<>0 then
	pcv_strTemp=pcv_strTemp&" AND idCategory<>"&pFeaturedCategory & " "
end if

dim pIdCategory, pCategoryDesc, pcStrViewAll

rMode=server.HTMLEncode(request.querystring("mode"))
if rMode="" then
	iPageSize=(pIntProductColumns*pIntProductRows)
	iCatPageSize=(pIntCategoryColumns*pIntCategoryRows)
	If Request("page")="" Then
		iPageCurrent=1
	Else
		iPageCurrent=CInt(Request("page"))
	End If
end if

'// View All
pcStrViewAll = Lcase(getUserInput(Request("viewAll"),3))
if pcStrViewAll = "yes" then
	iPageSize = 9999
end if	

if NOT validNum(iPageSize) OR iPageSize=0 then
	iPageSize=5
end if

pIdCategory=session("idCategoryRedirect")
mIdCategory=session("idCategoryRedirect")
if pIdCategory="" then
	pIdCategory=getUserInput(request.querystring("idCategory"),10)
	mIdCategory=getUserInput(request.querystring("idCategory"),10)
	'// Validate Category ID
	if not validNum(pIdCategory) then
		pIdCategory=""          
	end if
	if not validNum(mIdCategory) then
		mIdCategory=""          
	end if
	
	if pIdCategory="" then
		pIdCategory=1
		mIdCategory=1
	end if
end if
session("idCategoryRedirect")=""

'*******************************
' get category tree array
'*******************************
if pIdCategory<>1 then %>
	<!--#include file="pcBreadCrumbs.asp"-->
<% end if

'*******************************
' End get category tree array
'*******************************

'*******************************
' Get sub-categories array
'*******************************
Dim intSubCatExist
Dim iCategoriesPageCount
intSubCatExist=0

IF pIdCategory=1 THEN
	scCatTotal=(pIntCategoryColumns*pIntCategoryRows)
	if pIntSubCategoryView="2" then
		scCatTotal=999999
	end if
	iCategoriesPageSize=scCatTotal
	if pcStrViewAll = "yes" then
		iCategoriesPageSize = 9999
	end if
	
	Dim pcInt_CategoriesPage
	pcInt_CategoriesPage=getUserInput(request("CategoriesPage"),10)
	if not validNum(pcInt_CategoriesPage) then
		iCategoriesPageCurrent=1
	Else
		iCategoriesPageCurrent=Cint(pcInt_CategoriesPage)
	End If

	query = "SELECT idCategory,categoryDesc,[image],idParentCategory,SDesc,HideDesc FROM Categories WHERE idParentCategory=1 AND idCategory<>1 AND iBTOhide=0 " & pcv_strTemp & " ORDER BY priority, categoryDesc ASC;"
	SET rs=Server.CreateObject("ADODB.RecordSet")

	rs.PageSize=iCategoriesPageSize
	pcv_strPageSize=iCategoriesPageSize
	rs.CacheSize=iCategoriesPageSize
		
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
	'// Page Count
	iCategoriesPageCount=rs.PageCount
	If Cint(iCategoriesPageCurrent) > Cint(iCategoriesPageCount) Then iCategoriesPageCurrent=Cint(iCategoriesPageCount)
	If Cint(iCategoriesPageCurrent) < 1 Then iCategoriesPageCurrent=1	
ELSE
	
	scCatTotal=(pIntCategoryColumns*pIntCategoryRows)
	if pIntSubCategoryView="2" then
		scCatTotal=999999
	end if
	iCategoriesPageSize=scCatTotal
	if pcStrViewAll = "yes" then
		iCategoriesPageSize = 9999
	end if
	
	pcInt_CategoriesPage=getUserInput(request("CategoriesPage"),10)
	if not validNum(pcInt_CategoriesPage) then
		iCategoriesPageCurrent=1
	Else
		iCategoriesPageCurrent=Cint(pcInt_CategoriesPage)
	End If
	
	query = "SELECT idCategory, categoryDesc FROM Categories WHERE idParentCategory = " & pIdCategory & " AND idCategory<>1 AND iBTOhide=0 " & pcv_strTemp & " ORDER BY priority, categoryDesc ASC;"
	SET rs=Server.CreateObject("ADODB.RecordSet")

	rs.PageSize=iCategoriesPageSize
	pcv_strPageSize=iCategoriesPageSize
	rs.CacheSize=iCategoriesPageSize
		
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
	'// Page Count
	iCategoriesPageCount=rs.PageCount
	If Cint(iCategoriesPageCurrent) > Cint(iCategoriesPageCount) Then iCategoriesPageCurrent=Cint(iCategoriesPageCount)
	If Cint(iCategoriesPageCurrent) < 1 Then iCategoriesPageCurrent=1	
END IF

If NOT rs.EOF Then
	rs.AbsolutePage=iCategoriesPageCurrent
	intSubCatExist=1
	SubCatArray=rs.GetRows()
	intSubCatCount=ubound(SubCatArray,2)
End If

SET rs=nothing
'*******************************
' End get sub-categories array
'*******************************
%>

<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<!--#include file="inc_PrdCatTip.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<% if pIdCategory=1 then %>
	<tr>
		<td> 
			<h1><% response.write dictLanguage.Item(Session("language")&"_titles_9")%></h1>
		</td>
	</tr>
	<% else
		'*******************************
		' Show current category info
		'*******************************
		' Show BreadCrumbs - current category name and location - If subcategory z %>
		<tr>
			<td>
				<h1><%=pCategoryName%></h1>
                <%
				' Display promotion message if any
					pcs_CategoryPromotionMsg
				' End display promotion
	
				'// SEO-S
				pMainCategoryName=pCategoryName
				'// SEO-E
				%>
				<div class="pcPageNav">
				<% 
				response.write dictLanguage.Item(Session("language")&"_viewCat_P_2")
				response.write strBreadCrumb
				intIdCategory=pIdCategory

				'// Load category discount icon
				%>
                 <!--#Include File="pcShowCatDiscIcon.asp" -->
				</div>
			</td>
		</tr>
		<% ' End Show BreadCrumbs
	end if
	
	' Show large category image
	if pLargeImage<>"" then %>
		<tr>
			<td align="center">
				<img src="catalog/<%=pLargeImage%>" alt="<%=pCategoryName%>" vspace="5">
			</td>
		</tr>
 	<% end if
	' End show large category image
	
	' Start Show long category description
		if (LDesc<>"") and (HideDesc<>"1") then %>
		<tr>
			<td>
				<div class="pcPageDesc"><%=LDesc%></div>
			</td>
		</tr>
 	<% end if
	' End Show Categories Description
		
	'*******************************
	' Show subcategories, if any
	'*******************************
	if intSubCatExist=1 then %>
	<tr>
		<td>
		<% if pIdCategory<>1 then %>
			<h3><%=dictLanguage.Item(Session("language")&"_viewCategories_2")%>&quot;<%=pCategoryName%>&quot;</h3>
		<% end if %>
		<table class="pcShowContent">
			<% ' FIRST subcategory display option = Drop-down
			if pIntSubCategoryView="2" then %>
				<% if pFeaturedCategory<>0 then %>
						<!--#include file="pcShowCategoryFeatured.asp" -->
				<% end if %>
				<tr>
					<td>
					<p>
						<form class="pcForms">
						<% if trim(pCategoryName)<>"" then %>
							<%=dictLanguage.Item(Session("language")&"_viewCategories_3")%>&quot;<%=pCategoryName%>&quot;:&nbsp;
						<% else %>
							<%=dictLanguage.Item(Session("language")&"_viewCategories_6")%>
						<% end if %>
							<select onChange="window.location.href=this.options[selectedIndex].value" name="CatDropSelect">
								<option>Browse Subcategories</option>
								<% 	if pIdCategory=1 then
									pcv_mc=0 
									Do While (pcv_mc < iCategoriesPageSize) And (pcv_mc < (intSubCatCount+1))
										intIdCategory=SubCatArray(0,pcv_mc)
										pcStrCategoryDesc=SubCatArray(1,pcv_mc)
										'// Call SEO Routine
										pcGenerateSeoLinks
										'//
										query="SELECT categories_products.idProduct FROM categories_products WHERE categories_products.idCategory = " & intIdCategory
										
										%>
										<option value="<%=pcStrCatLink%>"><%=pcStrCategoryDesc%></option>
										<% 	pcv_mc=pcv_mc+1
									Loop
								else
									For pcv_mc=0 to intSubCatCount
									intIdCategory=SubCatArray(0,pcv_mc)
									pcStrCategoryDesc=SubCatArray(1,pcv_mc)
									'// Call SEO Routine
									pcGenerateSeoLinks
									'//							
									query="SELECT categories_products.idProduct FROM  categories_products WHERE categories_products.idCategory = " & intIdCategory
									
									%>
									<option value="<%=pcStrCatLink%>"><%=pcStrCategoryDesc%></option>
								<% Next
							 End if %>
							</select>
						</form>
						</p>
					</td>
				</tr>
			<% end if 
			' SECOND & THIRD subcategory display options
			if pIntSubCategoryView<>"2" then
				if pFeaturedCategory<>0 then
				'// Call SEO Routine
				pcGenerateSeoLinks
				'//
			%>
					<!--#include file="pcShowCategoryFeatured.asp" -->
			<% end if %>
				
			<tr>
			<%
				iCurOGNum=0
				if pIdCategory=1 then
					pcv_mc=0 
					Do While pcv_mc < iCategoriesPageSize And pcv_mc<intSubCatCount+1
						intIdCategory=SubCatArray(0,pcv_mc)
						strCategoryDesc=SubCatArray(1,pcv_mc)
						pcStrCategoryDesc=SubCatArray(1,pcv_mc)
						' SECOND display option: rich display
						' Thumbnail only view
						if pIntSubCategoryView=4 then %>
								<td onmouseover="this.className='pcShowCategoryBgHover'" onmouseout="this.className='pcShowCategoryBg'" class="pcShowCategoryBg"> 
									<!--#include file="pcShowCategoryT.asp" -->
								</td>
                        <%
						elseif pIntSubCategoryView="0" then
							if pIntCategoryColumns > 1 then %>
								<td onmouseover="this.className='pcShowCategoryBgHover'" onmouseout="this.className='pcShowCategoryBg'" class="pcShowCategoryBg"> 
									<!--#include file="pcShowCategoryH.asp" -->
								</td>
							<% else %>
								<td onmouseover="this.className='pcShowCategoryBgHover'" onmouseout="this.className='pcShowCategoryBg'" class="pcShowCategoryBg"> 
									<!--#include file="pcShowCategoryP.asp" -->
								</td>
							<% end if
						else
						'// Show categories as text links only
						'// Call SEO Routine
						pcGenerateSeoLinks
						'//
						%>
							<td>
								<a href="<%=pcStrCatLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="3" then%>onmouseover="javascript:document.getCatPre.idcategory.value='<%=intIdCategory%>'; sav_CatPrecallxml='1'; return runPreCatXML('cat_<%=intIdCategory%>');" onmouseout="javascript: sav_CatPrecallxml=''; hidetip();"<%end if%>><%=pcStrCategoryDesc%></a>
								<% '// Load category discount icon %>
		                        <!--#Include File="pcShowCatDiscIcon.asp" -->
							</td>
						<% end if
						iCurOGNum = iCurOGNum + 1
						pcv_mc=pcv_mc+1
						If ( iCurOGNum = pIntCategoryColumns ) Then
							iCurOGNum = 0
							Response.Write "</tr><tr>"
						End If
					Loop
				else
					pcv_mc=0 
					Do While pcv_mc < iCategoriesPageSize And pcv_mc<intSubCatCount+1
						intIdCategory=SubCatArray(0,pcv_mc)
						strCategoryDesc=SubCatArray(1,pcv_mc)
						pcStrCategoryDesc=SubCatArray(1,pcv_mc)
						' SECOND display option: rich display
						' Thumbnail only view
						if pIntSubCategoryView=4 then %>
								<td onmouseover="this.className='pcShowCategoryBgHover'" onmouseout="this.className='pcShowCategoryBg'" class="pcShowCategoryBg"> 
									<!--#include file="pcShowCategoryT.asp" -->
								</td>
                        <%
						elseif pIntSubCategoryView="0" then
							if pIntCategoryColumns > 1 then %>
								<td onmouseover="this.className='pcShowCategoryBgHover'" onmouseout="this.className='pcShowCategoryBg'" class="pcShowCategoryBg"> 
									<!--#include file="pcShowCategoryH.asp" -->
								</td>
							<% else %>
								<td onmouseover="this.className='pcShowCategoryBgHover'" onmouseout="this.className='pcShowCategoryBg'" class="pcShowCategoryBg"> 
									<!--#include file="pcShowCategoryP.asp" -->
								</td>
							<% end if
						else 
						'// Call SEO Routine
						pcGenerateSeoLinks
						'//
						%>
							<td>
								<a href="<%=pcStrCatLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="3" then%>onmouseover="javascript:document.getCatPre.idcategory.value='<%=intIdCategory%>'; sav_CatPrecallxml='1'; return runPreCatXML('cat_<%=intIdCategory%>');" onmouseout="javascript: sav_CatPrecallxml=''; hidetip();"<%end if%>><%=strCategoryDesc%></a>
								<% '// Load category discount icon %>
		                        <!--#Include File="pcShowCatDiscIcon.asp" -->
							</td>
						<% end if
						iCurOGNum = iCurOGNum + 1
						pcv_mc=pcv_mc+1
						If ( iCurOGNum = pIntCategoryColumns ) Then
							iCurOGNum = 0
							Response.Write "</tr><tr>"
						End If
					Loop
				end if
				response.Write "</tr>"
				%>
				<tr>
					<td colspan="<%=pIntCategoryColumns%>"><% call PageCategoriesNav(iCategoriesPageCount) %></td>
				</tr>
			<% End If %>
			</table>
		</td>
	</tr>
	<% End If
	'*******************************
	' END show subcategories
	'*******************************
		
	'*******************************
	' START show products
	'*******************************
	
	'Query order	
	Dim UONum, pcIntProductOrder
	query="SELECT POrder FROM categories_products WHERE idCategory="& pIdCategory &";"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	UONum=0
	do while not rs.eof
		pcIntProductOrder=rs("POrder")
		if not validNum(pcIntProductOrder) then pcIntProductOrder=0
		if pcIntProductOrder>0 then
			UONum=UONum+CLng(pcIntProductOrder)
		end if
		rs.MoveNext
	loop
	SET rs=nothing
	
	'Decide Order By
	Dim ProdSort 
	ProdSort=trim(getUserInput(request("prodsort"),2))
	if NOT validNum(ProdSort) then
		ProdSort=""
	end if
	if ProdSort="" then
		if UONum>0 then
			ProdSort="19"
		else
			ProdSort=PCOrd
	end if
	end if

	select case ProdSort
		Case "19": query1 = " ORDER BY categories_products.POrder Asc"
		Case "0": query1 = " ORDER BY products.SKU Asc"
		Case "1": query1 = " ORDER BY products.description Asc" 	
		Case "2": 
			If Session("customerType")=1 then
				if Ucase(scDB)="SQL" then
					query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
				else
					query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) DESC"
				end if
			else
				if Ucase(scDB)="SQL" then
					query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
				else
					query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) DESC"
				end if
			End if
		Case "3":
			If Session("customerType")=1 then
				if Ucase(scDB)="SQL" then
					query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
				else
					query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) ASC"
				end if
			else
				if Ucase(scDB)="SQL" then
					query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
				else
					query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) ASC"
				end if
			End if	
	end select
	
	'////////////////////////////////////////////////////////////////
	'// START: Category Seach Fields 
	'////////////////////////////////////////////////////////////////
	If SRCH_CSFON = "1" Then  

		pcs_CSFSetVariables()
		pcv_strCSFieldQuery = pcf_CSFieldQuery()

		'// Get the list of products that are currently available
		'// Only run this block if the include is in the footer
		if len(pcv_strCValues)>0 AND len(pcv_strCSFilters)=0 then

			tmpStrEx3=""
			pcv_HavingCount2 = 0
			tmpSValues3=split(pcv_strCValues,"||")
			For k=lbound(tmpSValues3) to ubound(tmpSValues3)	
				if tmpSValues3(k)<>"" then
					if pcv_HavingCount2=0 then
						tmpStrEx3 = tmpStrEx3 & ""& tmpSValues3(k)
					else
						tmpStrEx3 = tmpStrEx3 & ","& tmpSValues3(k)
					end if 					
					pcv_HavingCount2 = pcv_HavingCount2 + 1										
				end if
			Next
			
			queryCSF = "SELECT pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "FROM pcSearchFields_Products "
			queryCSF = queryCSF & "INNER JOIN products ON products.idProduct=pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "INNER JOIN categories_products ON products.idProduct=categories_products.idProduct "
			queryCSF = queryCSF & "WHERE pcSearchFields_Products.idSearchData in (" & tmpStrEx3 & ") "
			queryCSF = queryCSF & "AND categories_products.idCategory="& mIdCategory &" AND active=-1 AND configOnly=0 AND removed=0 "
			queryCSF = queryCSF & "GROUP BY pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "HAVING COUNT(DISTINCT pcSearchFields_Products.idSearchData) = " & pcv_HavingCount2

			set rsCSF=Server.CreateObject("ADODB.Recordset")  
			set rsCSF=connTemp.execute(queryCSF)
			if NOT rsCSF.eof then
				ProductIdArray = pcf_ColumnToArray(rsCSF.getRows(),0)
				CartProductIdString = Join(ProductIdArray,",")
				pcv_strCSFilters = " AND (products.idProduct In ("& CartProductIdString &"))"
			else 
				pcv_strCSFilters = " AND (products.idProduct In (0))"
			end if
			set rsCSF = nothing

		end if
		err.clear '// clear an error if they turn the feature on without a widget and ignore all warnings
	End If
	'////////////////////////////////////////////////////////////////
	'// END: Category Seach Fields
	'////////////////////////////////////////////////////////////////
	
	
	'// Query Products of current category
	query="SELECT products.idProduct, products.sku, products.description, products.price, products.listhidden, products.listprice, products.serviceSpec, products.bToBPrice, products.smallImageUrl,products.noprices,products.stock, products.noStock,products.pcprod_HideBTOPrice, POrder,products.FormQuantity,products.pcProd_BackOrder FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& mIdCategory&" AND active=-1 AND configOnly=0 and removed=0 " & pcv_strCSFilters & query1
	set rs=Server.CreateObject("ADODB.Recordset")   
	rs.CacheSize=iPageSize
	rs.PageSize=iPageSize
	pcv_strPageSize=iPageSize
		
	rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	dim iPageCount, pcv_intProductCount
	iPageCount=rs.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=1
	
	if NOT rs.eof then
		rs.AbsolutePage=Cint(iPageCurrent)
		pcArray_Products = rs.getRows()
		pcv_intProductCount = UBound(pcArray_Products,2)+1
	end if

	set rs = nothing

	'*******************************
	' Start: Set variables for "M" display
	'*******************************
	if pcPageStyle = "m" then
		'Check if customers are allowed to order products
		dim iShow
		iShow=0
		If scOrderlevel=0 then
			iShow=1
		end if
		If scOrderlevel=1 AND session("customerType")="1" then
			iShow=1
		End if
		
		Dim pCnt, pAddtoCart, pAllCnt
		'reset count variables
		pCnt=Cint(0)
		pAllCnt=Cint(0)

		'// Run through the products to count all products, products with options, and BTO products
		do while (pCnt < pcv_intProductCount)		
			
			pidrelation=pcArray_Products(0,pCnt) '// rsCount("idProduct")
			pserviceSpec=pcArray_Products(6,pCnt) '// rsCount("serviceSpec")	
			pStock=pcArray_Products(10,pCnt) '// rsCount("stock")
			pNoStock=pcArray_Products(11,pCnt) '// rsCount("noStock")
			pcv_intBackOrder=pcArray_Products(15,pCnt) '// rs("pcProd_BackOrder")
			
			pCnt=pCnt+1
			
			' Check which items will have multi qty enabled,
			pcv_SkipCheckMinQty=-1 
			If pcf_AddToCart(pidrelation)=False Then
				pAllCnt=pAllCnt+1
			End If	
			
		loop
		
		pcv_SkipCheckMinQty=0
			
		' If all items on the page are either BTO or have options,
		' do not show the quantity column or the Add to Cart button.						
		if cint(pAllCnt) <> cint(pCnt) then 
			pAddtoCart = 1
		end if
	end if	
	'*******************************
	' End: Set variables for "M" display
	'*******************************

	if pcv_intProductCount<1 then ' START IF-1: check if there are no products in this category...
		if intSubCatExist <> 1 then ' ... and there are no sub-categories, then show a message
	%>
		<tr>
			<td>
				<p><%=dictLanguage.Item(Session("language")&"_viewCat_P_4")%></p>
			</td>
		</tr>
	<% end if
	
	else ' ELSE IF-1: there are products or sub-categories
			
			if intSubCatExist = 1 then
			' If there are products AND subcategories, then products are considered
			' "featured" products within the category and are shown above the subcategories
		 	%>
		 		<tr>
					<td><hr></td>
				</tr>
				<tr>
					<td>
						<% if pIdCategory<>1 then %>
						<h3><%=dictLanguage.Item(Session("language")&"_viewCategories_1")%>&quot;<%=pCategoryName%>&quot;</h3>
						<% else %>
						<h3><%=dictLanguage.Item(Session("language")&"_viewCategories_1b")%></h3>
						<% end if %>
					</td>
				</tr>
		<%	end if ' The category contains products, but not subcategories
	
		'if SORT BY drop-down does not exist, show page nav still %>
		<tr>
			<td>
				<table style="width:100%; margin:0; padding:0; border-collapse:collapse;">
					<tr>	
						<td align="left" valign="top">
							<% call PageNav(iPagecount)%>
						</td>
						<%
						'=================================
						'show SORT BY drop-down
						'=================================
						if HideSortPro<>"1" then %>
							<td align="right" valign="middle" style="padding-left: 10px;">
								<form action="viewCategories.asp?pageStyle=<%=pcPageStyle%>&idcategory=<%=pidcategory%><%=pcv_strCSFieldQuery%>" method="post" class="pcForms">
									<%=dictLanguage.Item(Session("language")&"_viewCatOrder_5")%>
									<select name="prodSort" onChange="javascript:if (this.value != '') {this.form.submit();}">
									<%if UONum>0 then%>          
										<option value="19" <%if ProdSort="19" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_6")%></option>
									<%end if%>                        
										<option value="0"<%if ProdSort="0" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_1")%></option>
										<option value="1"<%if ProdSort="1" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_2")%></option>
										<option value="2"<%if ProdSort="2" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_3")%></option>
										<option value="3"<%if ProdSort="3" then%> selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_4")%></option>
									</select>
								</form>
							</td>
						<% end if 
						'=================================
						'end SORT BY drop-down
						'=================================
						%>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<% if pcPageStyle = "m" then %>
					<form action="instPrd.asp" method="post" name="m" id="m" class="pcForms">
				<% end if %>
				<table class="pcShowProducts">
					<% i=0
					'*******************************
					' Add table headers for display
					' styles L and M
					'*******************************
					%>
					<% if pcPageStyle = "l" then	%>
						<tr class="pcShowProductsLheader">
							<% if pShowSmallImg <> 0 then %>
								<td>&nbsp;</td>
							<% end if %>
							<td width="70%"><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %></td>
							<% if pShowSku <> 0 then %>
								<td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %></td>
							<% end if %>
							<td width="15%"><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_10") %></td>
						</tr>
					<% elseif pcPageStyle = "m" then %>
						<tr class="pcShowProductsMheader">
							<td colspan="<%if iShow=1 then%>5<%else%>4<%end if%>">
								<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_12") %>
							</td>
						</tr>
						<tr>
						<tr class="pcShowProductsMheader">
							<% if iShow=1 then %> 
								<% if pAddtoCart = 1 then %>
									<td>
										<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_7") %>
									</td>
								<% end if %>
							<% end if %>
							<% if pShowSmallImg <> 0 then %>
								<td>&nbsp;</td>
							<% end if %>
							<% if pShowSku <> 0 then %>
							<td>
								<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %>
							</td>
							<% end if %>
							<td width="70%">
								<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %>
							</td>
							<td width="15%" align="center">
								<% If session("customerType")="1" then
									response.write dictLanguage.Item(Session("language")&"_viewCat_P_11")
								 else
									response.write dictLanguage.Item(Session("language")&"_viewCat_P_10")
								end if %>
							</td>
						</tr>
					<% else %>
						<tr>
					<% end if
					'*******************************
					' End table headers

					if pcPageStyle = "m" then
						pCnt=Cint(0)
						pSQty=0
						pAllCnt=Cint(0)
					end if

					tCnt=Cint(0)
					
					do while (tCnt < pcv_intProductCount) and (count < pcv_strPageSize)									

						pidProduct=pcArray_Products(0,tCnt) '// rs("idProduct")
						pSku=pcArray_Products(1,tCnt) '// rs("sku")
						pDescription=pcArray_Products(2,tCnt) '// rs("description")   
						pPrice=pcArray_Products(3,tCnt) '// rs("price")
						pListHidden=pcArray_Products(4,tCnt) '// rs("listhidden")
						pListPrice=pcArray_Products(5,tCnt) '// rs("listprice")						   
						pserviceSpec=pcArray_Products(6,tCnt) '// rs("serviceSpec")
						pBtoBPrice=pcArray_Products(7,tCnt) '// rs("bToBPrice")   
						pSmallImageUrl=pcArray_Products(8,tCnt) '// rs("smallImageUrl")   
						pnoprices=pcArray_Products(9,tCnt) '// rs("noprices")
						if isNULL(pnoprices) OR pnoprices="" then
							pnoprices=0
						end if
						pStock=pcArray_Products(10,tCnt) '// rs("stock")
						pNoStock=pcArray_Products(11,tCnt) '// rs("noStock")
						pcv_intHideBTOPrice=pcArray_Products(12,tCnt) '// rs("pcprod_HideBTOPrice")
						if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
							pcv_intHideBTOPrice="0"
						end if
						if pnoprices=2 then
							pcv_intHideBTOPrice=1
						end if
						pFormQuantity=pcArray_Products(14,tCnt) '// rs("FormQuantity")
						pcv_intBackOrder=pcArray_Products(15,tCnt) '// rs("pcProd_BackOrder")
						pidrelation=pcArray_Products(0,tCnt) '// rs("idProduct")						
												
						'// Get sDesc
						query="SELECT sDesc FROM products WHERE idProduct="&pidrelation&";"
						set rsDescObj=server.CreateObject("ADODB.RecordSet")
						set rsDescObj=conntemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rsDescObj=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						If NOT rsDescObj.EOF Then
							psDesc=rsDescObj("sDesc")
						Else
							psDesc=""
						End If
						set rsDescObj=nothing
						
						if pcPageStyle = "m" then
							pCnt=pCnt+1
						end if
						tCnt=tCnt+1
						%>
						<!--#include file="pcGetPrdPrices.asp"-->
						<%   
						
						if pnoprices=0 then				
							' check for discount per quantity
							query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct
							if session("CustomerType")<>"1" then
								query=query & " and discountPerUnit<>0"
							else
								query=query & " and discountPerWUnit<>0"
							end if
							Dim rsDisc
							set rsDisc=Server.CreateObject("ADODB.Recordset")
							set rsDisc=conntemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsDisc=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
							
							Dim pDiscountPerQuantity
							if not rsDisc.eof then
								pDiscountPerQuantity=-1
							else
								pDiscountPerQuantity=0
							end if
							set rsDisc = nothing
						end if %>

						<% '*******************************
						' Show product information
						' depending on the page style
						'*******************************
							
						' FIRST STYLE - Show products horizontally, with images
							if pcPageStyle = "h" then %>
								<td onmouseover="this.className='pcShowProductBgHover'" onmouseout="this.className='pcShowProductBg'" class="pcShowProductBg"> 
									<!--#include file="pcShowProductH.asp" -->
								</td>
								<% i=i + 1
								If i > (pIntProductColumns-1) then 
									if (tCnt < pcv_strPageSize) then
									response.write "</TR><TR>"
									else
									response.write "</TR>"
									end if
									i=0
								End If
							end if
					
							' SECOND STYLE - Show products vertically, with images 
							if pcPageStyle = "p" then	%>
								<td> 
									<!--#include file="pcShowProductP.asp" -->
								</td>
							</tr>
							<% end if
					
							' THIRD STYLE - Show a list of products, with a small image 
							if pcPageStyle = "l" then	%>
									<!--#include file="pcShowProductL.asp" -->
							<% end if
							
							' FOURTH STYLE - Show a list of products, with multiple add to cart 
							if pcPageStyle = "m" then	%>
									<!--#include file="pcShowProductM.asp" -->
							<% end if %>

							<%	
							count=count + 1						
						loop
						%>

				</table>
				<%' If page style is M, show the Add to Cart button when
				' products can be added to the cart from this page.	
				if pcPageStyle = "m" then %>
					<input type="hidden" name="pCnt" value="<%=pCnt%>">
					<% if iShow=1 and clng(pSQty)<>0 then %>
							<div style="padding: 10px 0 10px 0;">
								<input name="submit" type="image" src="<%=rslayout("addtocart")%>" id="submit">
							</div>
					<% end if %>
					</form>
				<% end if %>
			</td>
		</tr>
		
	<% 
	end if ' END IF-1
	call closeDb() 
	%>
	<tr>
		<td><% call PageNav(iPagecount) %></td>
	</tr>
</table>
<!--#include file="atc_viewprd.asp"-->
</div>
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Category Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CategoryPromotionMsg
Dim rs,query,tmpStr
	call opendb()
	query="SELECT pcCatPro_PromoMsg FROM pcCatPromotions WHERE idcategory=" & pIdCategory & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpStr=rs("pcCatPro_PromoMsg")
		set rs=nothing
		' Display long product description if there is a short description
		if tmpStr <> "" then %>
            <div class="pcPromoMessage">
				<%=tmpStr%>
            </div>
		<%
		end if
	end if
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Category Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>

<% 
'====================
' Page Navigation
'==================== 

Sub PageNav(thepagecount)
	'// SEO-S
	intIdCategory=mIdCategory
	pcStrCategoryDesc=pMainCategoryName
	'// Call SEO Routine
	pcGenerateSeoLinks
	'// SEO-E
	iRecSize=10
	If thepagecount>1 then %>
		<div class="pcPageNav">
		<%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & thepagecount)%>
		&nbsp;-&nbsp;
        <% if thepagecount>iRecSize then %>
			<% if cint(iPageCurrent)>iRecSize then %>
                <a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&page=1&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_15")%></a>&nbsp;
            <% end if %>
			<% if cint(iPageCurrent)>1 then
                if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
                    iPagePrev=cint(iPageCurrent)-1
                else
                    iPagePrev=iRecSize
                end if %>
                <a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&page=<%=cint(iPageCurrent)-iPagePrev%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17a")%><%=iPagePrev%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
			<% end if %>
			<% 
			if cint(iPageCurrent)+1>1 then
                intPageNumber=cint(iPageCurrent)
            else
                intPageNumber=1
            end if
        else
			intPageNumber=1
		end if
		
		if (cint(thepagecount)-cint(iPageCurrent))<iRecSize then
			iPageNext=cint(thepagecount)-cint(iPageCurrent)
		else
			iPageNext=iRecSize
		end if
	
		For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
			If Cint(pageNumber)=Cint(iPageCurrent) Then %>
				<strong><%=pageNumber%></strong> 
			<% Else %>
				<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&page=<%=pageNumber%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=pageNumber%></a>
			<% End If 
		Next
		
		if (cint(iPageNext)+cint(iPageCurrent))=thepagecount then
		else
			if thepagecount>(cint(iPageCurrent) + (iRecSize-1)) then %>
				&nbsp;<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&page=<%=cint(intPageNumber)+iPageNext%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17")%><%=iPageNext%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
			<% end if
		
			if cint(thepagecount)>iRecSize AND (cint(iPageCurrent)<>cint(thepagecount)) then %>
				<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&page=<%=cint(thepagecount)%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_16")%></a>
			<% end if 
		end if %>
        &nbsp;<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&page=<%=cint(thepagecount)%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>&viewAll=yes" onClick="pcf_Open_viewAll();"><%=dictLanguage.Item(Session("language")&"_viewCategories_21")%></a>
		</div>
	<% end if
end Sub

Sub PageCategoriesNav(thepagecount)
	'// SEO-S
	intIdCategory=mIdCategory
	pcStrCategoryDesc=pMainCategoryName
	'// Call SEO Routine
	pcGenerateSeoLinks
	'// SEO-E

	iRecSize=10
	If thepagecount>1 then %>
		<div class="pcPageNav">
		<%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iCategoriesPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & thepagecount)%>
		&nbsp;-&nbsp;
        <% if thepagecount>iRecSize then %>
			<% if cint(iCategoriesPageCurrent)>iRecSize then %>
                <a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&page=1&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_15")%></a>&nbsp;
            <% end if %>
            <% if cint(iCategoriesPageCurrent)>1 then
                if cint(iCategoriesPageCurrent)<iRecSize AND cint(iCategoriesPageCurrent)<iRecSize then
                    iPagePrev=cint(iCategoriesPageCurrent)-1
                else
                    iPagePrev=iRecSize
                end if %>
                	&nbsp;<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&CategoriesPage=<%=cint(iCategoriesPageCurrent)-iPagePrev%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17a")%><%=iPagePrev%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
            <% end if
			if cint(iCategoriesPageCurrent)+1>1 then
				intPageNumber=cint(iCategoriesPageCurrent)
			else
				intPageNumber=1
			end if
		else
			intPageNumber=1
		end if
		if (cint(thepagecount)-cint(iCategoriesPageCurrent))<iRecSize then
			iPageNext=cint(thepagecount)-cint(iCategoriesPageCurrent)
		else
			iPageNext=iRecSize
		end if
	
		For pageNumber=intPageNumber To (cint(iCategoriesPageCurrent) + (iPageNext))
			If Cint(pageNumber)=Cint(iCategoriesPageCurrent) Then %>
				<strong><%=pageNumber%></strong> 
			<% Else %>
				<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&CategoriesPage=<%=pageNumber%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=pageNumber%></a>
			<% End If 
		Next
		
		if (cint(iPageNext)+cint(iCategoriesPageCurrent))=thepagecount then
		else
			if thepagecount>(cint(iCategoriesPageCurrent) + (iRecSize-1)) then %>
				<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&CategoriesPage=<%=cint(intPageNumber)+iPageNext%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_17")%><%=iPageNext%><%=dictLanguage.Item(Session("language")&"_viewCategories_18")%></a>&nbsp;
			<% end if
		
			if cint(thepagecount)>iRecSize AND (cint(iCategoriesPageCurrent)<>cint(thepagecount)) then %>
				<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&CategoriesPage=<%=cint(thepagecount)%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_viewCategories_16")%></a>
			<% end if 
		end if %>
        &nbsp;<a href="<%=pcStrCatLink2%>?pageStyle=<%=pcPageStyle%>&ProdSort=<%=ProdSort%>&idCategory=<%=mIdCategory%><%=pcv_strCSFieldQuery%>&viewAll=yes" onClick="pcf_Open_viewAll();"><%=dictLanguage.Item(Session("language")&"_viewCategories_21")%></a>
		</div>
	<% end if
end Sub

'====================
' END Page Navigation
'==================== 
%>
<%
Response.Write(pcf_InitializePrototype())
response.Write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_viewCategories_22"), "viewAll", 200))
%>
<!--#include file="footer.asp"-->