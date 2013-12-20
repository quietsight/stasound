<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<% 
dim query, conntemp, rs
call opendb()

Sub DupPrds(sourcecat,targetcat)
	Dim query,rs,pcArr,i,intCount
	
	query="SELECT idproduct,POrder FROM categories_products WHERE idcategory=" & sourcecat & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
			query="INSERT INTO categories_products (idproduct,idcategory,POrder) VALUES (" & pcArr(0,i) & "," & targetcat & "," & pcArr(1,i) & ");"
			set rs=connTemp.execute(query)
			set rs=nothing
		Next
	end if
	set rs=nothing
	
End Sub


Sub DupCats(sourcecat,targetcat)

	Dim subfeatcat,query,rs,rsQ,query1,pcArr,i,intCount,tmpquery1,tmpquery2,dd,iCols,newcat
	
	query="SELECT pcCats_FeaturedCategory FROM categories WHERE idcategory=" & sourcecat & " AND pcCats_FeaturedCategory>0;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		subfeatcat=rs("pcCats_FeaturedCategory")
	else
		subfeatcat=0
	end if
	set rs=nothing
	
	query="SELECT * FROM categories WHERE idParentCategory=" & sourcecat & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
		iCols = rs.Fields.Count
		tmpquery1=""
		for dd=1 to iCols-1
			if Ucase(rs.Fields.Item(dd).Name)<>Ucase("pcCats_CreatedDate") AND Ucase(rs.Fields.Item(dd).Name)<>Ucase("pcCats_EditedDate") then
				if tmpquery1<>"" then
					tmpquery1=tmpquery1 & ","
				end if
				tmpquery1=tmpquery1 & rs.Fields.Item(dd).Name
			end if
		next
		for i=0 to intCount
			tmpquery2=""
			for dd=2 to iCols-1
				if Ucase(rs.Fields.Item(dd).Name)<>Ucase("pcCats_CreatedDate") AND Ucase(rs.Fields.Item(dd).Name)<>Ucase("pcCats_EditedDate") then
					if tmpquery2<>"" then
						tmpquery2=tmpquery2 & ","
					end if
					if rs.Fields.Item(dd).Type="202" OR rs.Fields.Item(dd).Type="203" then
						tmpquery2=tmpquery2 & "'"
					end if
					if pcArr(dd,i)="False" then
						pcArr(dd,i)=0
					else
						if pcArr(dd,i)="True" then
							pcArr(dd,i)=-1
						end if
					end if
					if rs.Fields.Item(dd).Type<>"202" AND rs.Fields.Item(dd).Type<>"203" AND (pcArr(dd,i)="" OR IsNull(pcArr(dd,i))) then
						pcArr(dd,i)=0
					end if
					tmpquery2=tmpquery2 & pcArr(dd,i)
					if rs.Fields.Item(dd).Type="202" OR rs.Fields.Item(dd).Type="203" then
						tmpquery2=tmpquery2 & "'"
					end if
				end if
			next
			query1="INSERT INTO categories (" & tmpquery1 & ") VALUES (" & targetcat & "," & tmpquery2 & ");"
			set rsQ=connTemp.execute(query1)
			
			query1="SELECT idcategory FROM categories WHERE idParentCategory=" & targetcat & " AND categoryDesc like '" & pcArr(3,i) & "';"
			set rsQ=connTemp.execute(query1)
			newcat=rsQ("idcategory")
			set rsQ=nothing
			
			call updCatCreatedDate(newcat,"")
			
			if clng(pcArr(0,i))=clng(subfeatcat) AND clng(subfeatcat)>0 then
				query1="UPDATE categories SET pcCats_FeaturedCategory=" & newcat & " WHERE idcategory=" & targetcat & ";"
				set rsQ=connTemp.execute(query1)
				set rsQ=nothing
			end if
			
			if request("runDupPrds")="1" then
				call DupPrds(pcArr(0,i),newcat)
			end if
			
			call DupCats(pcArr(0,i),newcat)
			
		next
		set rs=nothing
	end if
	set rs=nothing

End Sub

' form parameter
SDesc=replace(request.form("SDesc"),"'","''")
LDesc=replace(request.form("LDesc"),"'","''")
HideDesc=request.form("HideDesc")

if NOT validNum(HideDesc) then HideDesc=0

pCategoryDesc=replace(request.form("categoryDesc"),"'","''")
pCategoryDesc=replace(pCategoryDesc,"&amp;","&")
pCategoryDesc=replace(pCategoryDesc,"&","&amp;")
pImage=request.form("image")
if pImage="" then
	pImage="no_image.gif"
end if
plargeImage=request.form("largeimage")
if plargeImage="" then
	plargeImage="no_image.gif"
end if

pIntSubCategoryView=request.form("intSubCategoryView")
pIntCategoryColumns=request.form("intCategoryColumns")
pIntCategoryRows=request.form("intCategoryRows")
pStrPageStyle=request.form("strPageStyle")
pIntProductColumns=request.form("intProductColumns")
pIntProductRows=request.form("intProductRows")
pIntFeaturedCategory=request.form("intFeaturedCategory")
pIntFeaturedCategoryImage=request.form("intFeaturedCategoryImage")
if NOT validNum(pIntSubCategoryView) then pIntSubCategoryView=0
if NOT validNum(pIntCategoryColumns) then pIntCategoryColumns=0
if NOT validNum(pIntCategoryRows) then pIntCategoryRows=0
if NOT validNum(pIntProductColumns) then pIntProductColumns=0
if NOT validNum(pIntProductRows) then pIntProductRows=0
if NOT validNum(pIntFeaturedCategory) then pIntFeaturedCategory=0
if NOT validNum(pIntFeaturedCategoryImage) then pIntFeaturedCategoryImage=0
if NOT validNum(HideDesc) then HideDesc=0
if NOT validNum(pcv_intRetailHide) then pcv_intRetailHide=0

iBTOhide=request.form("iBTOhide")
if NOT validNum(iBTOhide) then iBTOhide=0
	pcv_intRetailHide=request.form("RetailHide")
if NOT validNum(pcv_intRetailHide) then pcv_intRetailHide=0

pIdParentCategory=request.form("idparentCategory")
reqstr=request.form("reqstr")

'//Retrieve Category Level Product Display Setting
pcv_StrCatDisplayLayout=getUserInput(request.Form("CatDisplayLayout"),4)
if pcv_StrCatDisplayLayout="D" then pcv_StrCatDisplayLayout=""

'//Retrieve new Meta Tag related fields
pcv_StrCatMetaTitle=getUserInput(request.Form("CatMetaTitle"), 0)
pcv_StrCatMetaDesc=getUserInput(request.Form("CatMetaDesc"), 0)
pcv_StrCatMetaKeywords=getUserInput(request.Form("CatMetaKeywords"), 0)

' identify tier of parent category and set tier + 1
query="SELECT tier,iBTOhide FROM categories WHERE idCategory="& pIdParentCategory
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
ptier=rs("tier")+1
pcv_ParentiBTOhide=rs("iBTOhide")
if IsNull(pcv_ParentiBTOhide) or pcv_ParentiBTOhide="" then
	pcv_ParentiBTOhide=0
end if
set rs=nothing

if pcv_ParentiBTOhide="1" then
	iBTOhide=pcv_ParentiBTOhide
end if

' insert category in to db
query="INSERT INTO categories (SDesc, LDesc, HideDesc, tier, idParentCategory, categoryDesc, [image], largeimage, iBTOhide, pcCats_RetailHide, pcCats_SubCategoryView,  pcCats_CategoryColumns, pcCats_CategoryRows, pcCats_PageStyle, pcCats_ProductColumns, pcCats_ProductRows, pcCats_FeaturedCategory, pcCats_FeaturedCategoryImage, pcCats_DisplayLayout, pcCats_MetaTitle, pcCats_MetaDesc, pcCats_MetaKeywords) VALUES ('" & SDesc & "', '" & LDesc & "'," & HideDesc & ", " &pTier&", "&pIdParentCategory&", '"&pcategoryDesc&"', '"&pImage&"', '"&plargeImage&"', "&iBTOhide&", " & pcv_intRetailHide & ", "&pIntSubCategoryView&", "&pIntCategoryColumns&", "&pIntCategoryRows&", '"&pStrPageStyle&"', "&pIntProductColumns&", "&pIntProductRows&", "&pIntFeaturedCategory&", "&pIntFeaturedCategoryImage&", '"&pcv_StrCatDisplayLayout&"', '"&pcv_StrCatMetaTitle&"', '"&pcv_StrCatMetaDesc&"', '"&pcv_StrCatMetaKeywords&"');"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number <> 0 then
	set rs=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error adding a new category.") 
end If

' get newly created category ID
query="SELECT idcategory, idParentCategory FROM categories ORDER BY idcategory DESC;"
set rs=conntemp.execute(query)
pIdCategory=rs("idcategory")
pIdParent=rs("idParentCategory")
set rs=nothing

'Duplicate CAT products and sub-categories
Dim pcIntDubCatFrom
pcIntDubCatFrom = request("preID")
if validNum(pcIntDubCatFrom) then
	if request("runDupSubCats")="1" then
		call DupCats(pcIntDubCatFrom,pIdCategory)
	else
		if request("runDupPrds")="1" then
			call DupPrds(pcIntDubCatFrom,pIdCategory)
		end if
	end if
end if	

' get "top" category ID
if validNum(pIdParent) then
	query="SELECT idParentCategory FROM categories WHERE idcategory=" & pIdParent
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=connTemp.execute(query)
	pIdTop=rs("idParentCategory")
	SET rs=nothing
end if

'Update Category Tree XML Cache
%>
<!--#include file="inc_genCatXML.asp"-->

<% pageTitle="Add New Category" %>
<% section="products" %>

<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr> 
    <td><div class="pcCPmessageSuccess">
      <p>Category added successfully.</p>
      <p>&nbsp;</p>
      <p>If you are using &quot;Static Navigation&quot; in your storefront, remember to update the navigation files using the <a href="genCatNavigation.asp" target="_blank">Generate Navigation</a> feature. </p>
    </div></td>
	</tr>
	<tr>
		<td> 
      		<% if reqstr<>"" then %>
			<ul>
       			<li><a href="instCata.asp?reqstr=<%=reqstr%>">Add another category</a></li>
				<li><a href="<%=reqstr%>">Continue BTO configuration</a></li>
			</ul>
			<% else %> 
      			<ul class="pcListIcon">
				<li><a href="editCategories.asp?nav=&lid=<%=pIdCategory%>">Add products to this category</a></li>
				<li><a href="modCata.asp?idcategory=<%=pIdCategory%>&top=<%=pIdTop%>&parent=<%=pIdParent%>">Edit the category again</a></li>
				<li style="padding-top: 10px;"><a href="genCatNavigation.asp" target="_blank">Update storefront navigation</a></li>				
                <li><a href="../pc/viewcategories.asp?idcategory=<%=pIdCategory%>" target="_blank">View in the storefront</a></li>
        		<li style="padding-top: 10px;"><a href="instCata.asp">Add another category</a></li>
                <li><a href="adddupcat.asp?idcategory=<%=pIdCategory%>&top=<%=pIdTop%>&parent=<%=pIdParent%>">Clone this category</a></li>
				<li><a href="manageCategories.asp?prdType=1">Manage categories</a></li>
				</ul>
			<% end if %>
		</td>
	</tr>   
</table>
<%
call closeDb
%>
<!--#include file="AdminFooter.asp"-->