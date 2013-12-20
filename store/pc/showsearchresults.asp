<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<%
Dim pcv_strHideCatSearch
pcv_strHideCatSearch = False '// Set to "True" to disable category search

'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "showSearchResults.asp"

'*******************************
' Query results
'*******************************
dim query, connTemp, rs, pSearchSKU, pKeywords, pPriceFrom, pPriceUntil, pIdCategory, pIdSupplier, pWithStock, strSearch, pCValues, tKeywords, tIncludeSKU, IDBrand, strPrdOrd, strOrderBy

pSearchSKU=getUserInput(request.querystring("SKU"),150)
pKeywords=getUserInput(request.querystring("keyWord"),100)
pCValues=getUserInput(request.querystring("SearchValues"),0)
tKeywords=pKeywords
tIncludeSKU=getUserInput(request.querystring("includeSKU"),10)
	if tIncludeSKU = "" then
		tIncludeSKU = "true"
	end if
pPriceFrom=getUserInput(request.querystring("priceFrom"),20)
if Instr(pPriceFrom,",")>Instr(pPriceFrom,".") then
	pPriceFrom=replace(pPriceFrom,",",".")
end if
if NOT isNumeric(pPriceFrom) then
	pPriceFrom=0
end if
pPriceUntil=getUserInput(request.querystring("priceUntil"),20)
if Instr(pPriceUntil,",")>Instr(pPriceUntil,".") then
	pPriceUntil=replace(pPriceUntil,",",".")
end if
if NOT isNumeric(pPriceUntil) then
	pPriceUntil=999999999
end if
pIdCategory=getUserInput(request.querystring("idCategory"),4)
pIdSupplier=getUserInput(request.querystring("idSupplier"),4)
if NOT validNum(pIdSupplier) or trim(pIdSupplier)="" then
	pIdSupplier=0
end if
pWithStock=getUserInput(request.querystring("withStock"),2)	
IDBrand=getUserInput(request.querystring("IDBrand"),20)
if NOT validNum(IDBrand) or trim(IDBrand)="" then
	IDBrand=0
end if
incSale=getUserInput(request("incSale"),4)
if NOT validNum(incSale) or trim(incSale)="" then
	incSale=0
end if
tmpIDSale=getUserInput(request("IDSale"),4)
if NOT validNum(tmpIDSale) or trim(tmpIDSale)="" then
	tmpIDSale=0
end if
strPrdOrd=getUserInput(request.querystring("order"),4)
	if NOT validNum(strPrdOrd) or trim(strPrdOrd)="" then strPrdOrd=PCOrd
	if NOT validNum(strPrdOrd) or trim(strPrdOrd)="" then strPrdOrd=1
	Select Case strPrdOrd
		Case "0": strOrderBy="A.sku ASC, A.idproduct DESC"
		Case "1": strOrderBy="A.description ASC"
		Case "2":
			If Session("customerType")=1 then
				if Ucase(scDB)="SQL" then
					strOrderBy = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) DESC"
				else
					strOrderBy = "(iif(iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),iif(IsNull(A.pcProd_BTODefaultPrice),0,A.pcProd_BTODefaultPrice),A.pcProd_BTODefaultWPrice)=0,iif(A.btoBPrice=0,A.Price,A.btoBPrice),iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),A.pcProd_BTODefaultPrice,A.pcProd_BTODefaultWPrice))) DESC"
				end if
			else
				if Ucase(scDB)="SQL" then
					strOrderBy = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) DESC"
				else
					strOrderBy = "(iif((A.pcProd_BTODefaultPrice=0) OR (IsNull(A.pcProd_BTODefaultPrice)),A.Price,A.pcProd_BTODefaultPrice)) DESC"
				end if
			End if
		Case "3":
			If Session("customerType")=1 then
				if Ucase(scDB)="SQL" then
					strOrderBy = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) ASC"
				else
					strOrderBy = "(iif(iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),iif(IsNull(A.pcProd_BTODefaultPrice),0,A.pcProd_BTODefaultPrice),A.pcProd_BTODefaultWPrice)=0,iif(A.btoBPrice=0,A.Price,A.btoBPrice),iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),A.pcProd_BTODefaultPrice,A.pcProd_BTODefaultWPrice))) ASC"
				end if
			else
				if Ucase(scDB)="SQL" then
					strOrderBy = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) ASC"
				else
					strOrderBy = "(iif((A.pcProd_BTODefaultPrice=0) OR (IsNull(A.pcProd_BTODefaultPrice)),A.Price,A.pcProd_BTODefaultPrice)) ASC"
				end if
			End if
		Case Else: strOrderBy="A.description ASC"
	End Select
	strORD=strPrdOrd

intExact=getUserInput(request.querystring("exact"),4)
if NOT validNum(intExact) or trim(intExact)="" then
	intExact=0
end if

'*******************************
' START - Don't allow empty searches
'*******************************
Dim pcIntNullSearch
if pKeywords="" AND pIdCategory="" then
	pcIntNullSearch=1
end if
if NOT validNum(pIdCategory) or trim(pIdCategory)="" then
	pIdCategory=0
end if

if pIdCategory="0" AND (pIdSupplier="" OR pIdSupplier="0") AND pPriceFrom="0" AND pPriceUntil="999999999" AND pSearchSKU="" AND IDBrand="0" AND pKeywords="" AND (pCValues="" OR pCValues="0" OR pCValues="||") AND trim(pWithStock)="" then
	pcIntNullSearch=1
end if

'// Let price-based searches go through
if (pPriceFrom<>"0" OR pPriceUntil<>"999999999") then
	pcIntNullSearch=0
end if

'// Let brand-based searches go through
if IDBrand<>"0" then
	pcIntNullSearch=0
end if

if incSale<>"0" then
	pcIntNullSearch=0
end if

'// Let custom search field queries go through
if (pCValues<>"0" AND pCValues<>"" AND pCValues<>"||") then
	pcIntNullSearch=0
end if

if pcIntNullSearch=1 then
	response.redirect "search.asp"
end if

'*******************************
' END - Don't allow empty searches
'*******************************

%>
<!--#include file="pcStartSession.asp"-->
<%
call opendb()
%>
<!--#include FILE="prv_getSettings.asp"-->
<%

'*******************************
' GET page style
'*******************************
	' Load the page style: check to see if a querystring
	' or a form is sending the page style.
	Dim pcPageStyle
	pcPageStyle = LCase(Request.QueryString("pageStyle"))
		if pcPageStyle = "" then
			pcPageStyle = LCase(Request.Form("pageStyle"))
		end if

		if pcPageStyle = "" then
			pcPageStyle = LCase(bType)
		end if

		if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
			pcPageStyle = LCase(bType)
		end if


'//===========================
'// BRAND Information - Start
'//===========================

if IDBrand>0 then

	query="SELECT BrandName, pcBrands_Description, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg FROM Brands WHERE pcBrands_Active=1 AND idBrand="&IDBrand
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		set rstemp=nothing
		call closeDb()
		response.redirect "msg.asp?message=85"       
	End if
	pcvBrandName=pcf_PrintCharacters(rstemp("BrandName"))
	pcvBrandsDescription=pcf_PrintCharacters(rstemp("pcBrands_Description"))
	pcvIntBrandsParent=rstemp("pcBrands_Parent")
	pcvBrandLogoLg=rstemp("pcBrands_BrandLogoLg")

	pcv_DefaultTitle=rstemp("pcBrands_MetaTitle")
	if isNull(pcv_DefaultTitle) or trim(pcv_DefaultTitle)="" then
		pcv_DefaultTitle=ClearHTMLTags2(parentBrandName,0)
	end if
	pcv_DefaultTitle = pcv_DefaultTitle & " - " & scCompanyName
	pcv_DefaultDescription=rstemp("pcBrands_MetaDesc")
	pcv_DefaultKeywords=rstemp("pcBrands_MetaKeywords")
	
	set rstemp=nothing

	if not validNum(parentIntBrandsParent) then parentIntBrandsParent=0
	
end if

'//===========================
'// BRAND Information - Start
'//===========================

		
' OTHER display settings
' These variables show/hide information when products are shown with Page Style = L or M
Dim pShowSKU, pShowSmallImg
pShowSKU = scShowSKU ' If 0, then the SKU is hidden
pShowSmallImg = scShowSmallImg ' If 0, then the small image is not shown
' Note: the size of the small image is set via the pcStorefront.css stylesheet

%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<%
'*******************************
' Set page size and get current page
'*******************************
	Dim pcv_ViewAllVar, newCount
	pcv_ViewAllVar=getUserInput(request("VA"),1)
	if NOT isNumeric(pcv_ViewAllVar) OR pcv_ViewAllVar="" then
		pcv_ViewAllVar=0
	end if
	newCount=0
	
	dim iPageSize
	iPageSize=getUserInput(request("resultCnt"),10)
	if iPageSize="" then
		iPageSize=getUserInput(request("iPageSize"),0)
	end if
	if NOT validNum(iPageSize) then
	iPageSize=(scPrdRow*scPrdRowsPerPage)
	end if
	
	dim iPageCurrent
	if request.queryString("iPageCurrent")="" then
		iPageCurrent=1 
	else
		iPageCurrent=server.HTMLEncode(request.querystring("iPageCurrent"))
		if NOT validNum(iPageCurrent) then
			iPageCurrent=1
		end if
	end if


'*******************************
' Create Search Query
'*******************************
Dim strSQL, tmpSQL, tmpSQL2, tmp_StrQuery, pcv_strMaxResults

tmp_StrQuery=""
if session("customerCategory")="" or session("customerCategory")=0 then
	If session("customerType")=1 then
		tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultWPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultWPrice<=" &pPriceUntil&")"
	else
		tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultPrice<=" &pPriceUntil&")"
	end if
else
	tmp_StrQuery="(A.serviceSpec<>0 AND A.idproduct IN (SELECT DISTINCT idproduct FROM pcBTODefaultPriceCats WHERE pcBTODefaultPriceCats.idCustomerCategory=" & session("customerCategory") & " AND pcBTODefaultPriceCats.pcBDPC_Price>="&pPriceFrom&" AND pcBTODefaultPriceCats.pcBDPC_Price<=" &pPriceUntil&"))"
end if

if scDB="Access" then
	zSQL="A.sDesc"
else
	zSQL="cast(A.sDesc as varchar(8000)) sDesc"
end if

pcv_strMaxResults=SRCH_MAX
If pcv_strMaxResults>"0" Then
	pcv_strLimitPhrase="TOP " & pcv_strMaxResults
Else
	pcv_strLimitPhrase=""
End If

strSQL= "SELECT "& pcv_strLimitPhrase &" A.idProduct, A.sku, A.description, A.price, A.listHidden, A.listPrice, A.serviceSpec, A.bToBPrice, A.smallImageUrl, A.noprices, A.stock, A.noStock, A.pcprod_HideBTOPrice, A.pcProd_BackOrder, A.FormQuantity, A.pcProd_BackOrder, A.pcProd_BTODefaultPrice, "& zSQL &" " 
strSQL=strSQL& "FROM products A "
strSQL=strSQL& " WHERE (A.active=-1 AND A.removed=0 AND A.idProduct IN (" 

	'// START: Category Sub-Query
	strSQL=strSQL& "SELECT B.idProduct FROM categories_products B INNER JOIN categories C ON "
	strSQL=strSQL & "C.idCategory=B.idCategory WHERE C.iBTOhide=0 "
	if pIdCategory<>"0" then
		if (schideCategory = "1") OR (SRCH_SUBS = "1") then
			Dim TmpCatList
			TmpCatList=""
			call pcs_GetSubCats(pIdCategory) '// get sub cats
			TmpCatList = pIdCategory&TmpCatList
			if len(TmpCatList)>0 then
				strSQL=strSQL & " AND B.idCategory IN ("& TmpCatList &")" '// include sub cats
			else
				strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
			end if
		else
			strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
		end if
	end if
	if session("CustomerType")<>"1" then
		strSQL=strSQL & " AND C.pccats_RetailHide=0"
	end if
	'// END: Category Sub-Query

strSQL=strSQL& ") AND (" & tmp_StrQuery & " OR (A.serviceSpec=0 AND A.configOnly=0 AND A.price>="&pPriceFrom&" AND A.price<=" &pPriceUntil&")) " 

if UCase(scDB)="SQL" then
	if (incSale>"0") then
		if tmpIDSale="0" then
			strSQL=strSQL & " AND A.pcSC_ID>0"
		else
			strSQL=strSQL & " AND A.pcSC_ID=" & tmpIDSale
		end if
	end if
end if

if len(pSearchSKU)>0 then
	strSQL=strSQL & " AND A.sku like '%"&pSearchSKU&"%'"
end if

if pIdSupplier<>"0" then
	strSQL=strSQL & " AND A.idSupplier=" &pIdSupplier
end if

if pWithStock="-1" then
	strSQL=strSQL & " AND (A.stock>0 OR A.noStock<>0)" 
end if

if (IDBrand&""<>"") and (IDBrand&""<>"0") then
	strSQL=strSQL & " AND A.IDBrand=" & IDBrand
end if
pKeywords=replace(pKeywords,"''''","''")
TestWord=""
if intExact<>"1" then
	if Instr(pKeywords," AND ")>0 then
		keywordArray=split(pKeywords," AND ")
		TestWord=" AND "
	else
		if Instr(pKeywords," and ")>0 then
			keywordArray=split(pKeywords," and ")
			TestWord=" AND "
		else
			if Instr(pKeywords,",")>0 then
				keywordArray=split(pKeywords,",")
				TestWord=" OR "
			else
				if (Instr(pKeywords," OR ")>0) then
					keywordArray=split(pKeywords," OR ")
					TestWord=" OR "
				else
					if (Instr(pKeywords," or ")>0) then
						keywordArray=split(pKeywords," or ")
						TestWord=" OR "
					else
						if (Instr(pKeywords," ")>0) then
							keywordArray=split(pKeywords," ")
							TestWord=" AND "
						else
							keywordArray=split(pKeywords,"***")	
							TestWord=" OR "
						end if
					end if
				end if
			end if
		end if
	end if
else
	pKeywords=trim(pKeywords)
	if pKeywords<>"" then
		if scDB="SQL" then
			pKeywords="'" & pKeywords & "'***'%[^a-zA-z0-9]" & pKeywords & "[^a-zA-z0-9]%'***'" & pKeywords & "[^a-zA-z0-9]%'***'%[^a-zA-z0-9]" & pKeywords & "'"
		else
			pKeywords="'" & pKeywords & "'***'%[!a-zA-z0-9]" & pKeywords & "[!a-zA-z0-9]%'***'" & pKeywords & "[!a-zA-z0-9]%'***'%[!a-zA-z0-9]" & pKeywords & "'"
		end if
	end if
	keywordArray=split(pKeywords,"***")	
	TestWord=" OR "
end if

tmpStrEx=""
if pCValues<>"" AND pCValues<>"0" then
	tmpSValues=split(pCValues,"||")
	For k=lbound(tmpSValues) to ubound(tmpSValues)
		if tmpSValues(k)<>"" then
			sfquery=""
			sfquery = "SELECT pcSearchFields_Products.idproduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData=" & tmpSValues(k)
			set rsSearchFields=Server.CreateObject("ADODB.Recordset")
			set rsSearchFields=connTemp.execute(sfquery)
			If NOT rsSearchFields.eof Then
				SearchFieldArray = pcf_ColumnToArray(rsSearchFields.getRows(),0)
				SearchFieldString = Join(SearchFieldArray,",")		
				If len(SearchFieldString)>0 Then
					tmpStrEx=tmpStrEx & " AND A.idproduct IN ("& SearchFieldString &")"
				End If
			Else
				tmpStrEx=tmpStrEx & " AND A.idproduct IN (0)"				
			End If
			set rsSearchFields = nothing
		end if
	Next
end if

'////////////////////////////////////////////////////////////////
'// START: Category Seach Fields 
'////////////////////////////////////////////////////////////////
If SRCH_CSFRON = "1" Then 

	pcv_strCSFilters=""
	pcs_CSFSetVariables()
	pcv_strCSFieldQuery = pcf_CSFieldQuery()
	if len(pcv_strCValues)>0 then
		queryCSF = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData>0 " 
		tmpSValues3=split(pcv_strCValues,"||")
		For k=lbound(tmpSValues3) to ubound(tmpSValues3)
			if tmpSValues3(k)<>"" then
				SubQuery = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData = " & tmpSValues3(k) & ""
				set rsSubQuery=Server.CreateObject("ADODB.Recordset")  
				set rsSubQuery=connTemp.execute(SubQuery)
				If NOT rsSubQuery.eof Then
					ProductIdArray = pcf_ColumnToArray(rsSubQuery.getRows(),0)
					ProductIdString = Join(ProductIdArray,",")
					tmpStrEx3=tmpStrEx3 & " AND pcSearchFields_Products.idProduct IN "
					tmpStrEx3=tmpStrEx3 & "(" & ProductIdString & ")"
				End If
				set rsSubQuery = nothing				
			end if
		Next
		queryCSF = queryCSF & tmpStrEx3	
		set rsCSF=Server.CreateObject("ADODB.Recordset")  
		set rsCSF=connTemp.execute(queryCSF)
		if NOT rsCSF.eof then
			ProductIdArray = pcf_ColumnToArray(rsCSF.getRows(),0)
			ProductIdString = Join(ProductIdArray,",")
			pcv_strCSFilters = " AND (A.idProduct In ("& ProductIdString &"))"
		else 
			pcv_strCSFilters = " AND (A.idProduct In (0))"
		end if
		set rsCSF = nothing
	end if
	err.clear
End If
tmpStrEx = tmpStrEx & pcv_strCSFilters
'////////////////////////////////////////////////////////////////
'// END: Category Seach Fields
'////////////////////////////////////////////////////////////////

IF intExact<>"1" THEN

	if pKeywords<>"" then
	
		strSQl=strSql & " AND ("
		
		tmpSQL="(A.details LIKE "
		tmpSQL2="(A.description LIKE "
		tmpSQL3="(A.sDesc LIKE "
		tmpSQL5="(A.pcProd_MetaKeywords LIKE "
		if tIncludeSKU="true" then
			tmpSQL4="(A.SKU LIKE "
		end if
		Dim Pos
		Pos=0
		For L=LBound(keywordArray) to UBound(keywordArray)
			if trim(keywordArray(L))<>"" then
			Pos=Pos+1
			if Pos>1 Then
				tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
				tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
				tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
				tmpSQL5=tmpSQL5 & TestWord & " A.pcProd_MetaKeywords LIKE "
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
				end if
			end if
				tmpSQL=tmpSQL  & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL2=tmpSQL2 & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL3=tmpSQL3 & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL5=tmpSQL5 & "'%" & trim(keywordArray(L)) & "%'"
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & "'%" & trim(keywordArray(L)) & "%'"
				end if
			end if
		Next
		tmpSQL=tmpSQL & ")"
		tmpSQL2=tmpSQL2 & ")"
		tmpSQL3=tmpSQL3 & ")"
		tmpSQL5=tmpSQL5 & ")"
		if tIncludeSKU="true" then
			tmpSQL4=tmpSQL4 & ")"
		end if
		
		strSQL=strSQL & tmpSQL
		strSQL=strSQL & " OR " & tmpSQL2
		strSQL=strSQL & " OR " & tmpSQL5
		if tIncludeSKU="true" then
			strSQL=strSQL & " OR " & tmpSQL3
			strSQL=strSQL & " OR " & tmpSQL4 & ")"
		else	
			strSQL=strSQL & " OR " & tmpSQL3 & ")"
		end if
		strSQL=strSQL& ")" & tmpStrEx
		query=strSQL & " ORDER BY " & strOrderBy
	else
		strSQL=strSQL& ")" & tmpStrEx
		query=strSQL & " ORDER BY " & strOrderBy
	end if

ELSE 'Exact=1

	if pKeywords<>"" then
	
		strSQl=strSql & " AND ("
		
		tmpSQL="(A.details LIKE "
		tmpSQL2="(A.description LIKE "
		tmpSQL3="(A.sDesc LIKE "
		tmpSQL5="(A.pcProd_MetaKeywords LIKE "
		if tIncludeSKU="true" then
			tmpSQL4="(A.SKU LIKE "
		end if
		Pos=0
		For L=LBound(keywordArray) to UBound(keywordArray)
			if trim(keywordArray(L))<>"" then
			Pos=Pos+1
			if Pos>1 Then
				tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
				tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
				tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
				tmpSQL5=tmpSQL5 & TestWord & " A.pcProd_MetaKeywords LIKE "
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
				end if
			end if
				tmpSQL=tmpSQL & trim(keywordArray(L))
				tmpSQL2=tmpSQL2 & trim(keywordArray(L))
				tmpSQL3=tmpSQL3 & trim(keywordArray(L))
				tmpSQL5=tmpSQL5 & trim(keywordArray(L))
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & trim(keywordArray(L))
				end if
			end if
		Next
		tmpSQL=tmpSQL & ")"
		tmpSQL2=tmpSQL2 & ")"
		tmpSQL3=tmpSQL3 & ")"
		tmpSQL5=tmpSQL5 & ")"
		if tIncludeSKU="true" then
			tmpSQL4=tmpSQL4 & ")"
		end if
		
		strSQL=strSQL & tmpSQL
		strSQL=strSQL & " OR " & tmpSQL2
		strSQL=strSQL & " OR " & tmpSQL5
		if tIncludeSKU="true" then
			strSQL=strSQL & " OR " & tmpSQL3
			strSQL=strSQL & " OR " & tmpSQL4 & ")"
		else	
			strSQL=strSQL & " OR " & tmpSQL3 & ")"
		end if
		strSQL=strSQL& ")" & tmpStrEx
		query=strSQL & " ORDER BY " & strOrderBy
	else
		strSQL=strSQL& ")" & tmpStrEx
		query=strSQL & " ORDER BY " & strOrderBy
	end if
END IF 'Exact

totalrecords=0
session("pcstore_prdlist")=""
session("pcstore_newsrc")="OK"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.CacheSize=iPageSize
rs.PageSize=iPageSize

rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
if err.number <> 0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closeDb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end If

dim iPageCount, count
if NOT rs.eof then		
	
	iPageCount=rs.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then Cint(iPageCurrent)=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
	rs.AbsolutePage=Cint(iPageCurrent)
	pcv_strPageSize=iPageSize
	pcArray_Products = rs.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
	
	
	'// Next and Previous Buttons
	if session("pcstore_prdlist")="" then
		session("pcstore_prdlist")="*****"
	end if
	For prdList = 0 to UBound(pcArray_Products,2)
		session("pcstore_prdlist")=session("pcstore_prdlist") & pcArray_Products(0,prdList) & "*****"
	Next
	totalrecords = rs.recordcount

else
	set rs = nothing
	call closeDb()
	if request("fp")="bnd" then
		response.redirect "msg.asp?message=90"
	else
		response.redirect "msg.asp?message=3"                
	end if        
end if
set rs = nothing

'*******************************
' Set variables for "M" display
'*******************************
	if pcPageStyle = "m" then
		'Check if customers are allowed to order products
		dim iShow
		iShow=0
		If scOrderlevel=0 then ' Anybody can order
			iShow=1
		end if
		If scOrderlevel=1 AND session("customerType")="1" then ' Only wholesale customers can order
			iShow=1
		End if

		Dim pCnt,pAddtoCart,pAllCnt
		'reset count variables
		pCnt=Cint(0)
		pAllCnt=Cint(0)

		'// Loop until the total number of products to show
		if pcv_ViewAllVar=0 then
			newCount = pcv_strPageSize
		else
			newCount = 999999
		end if			

		'// Run through the products to count all products, products with options, and BTO products
		do while (pCnt < pcv_intProductCount) and (pCnt < newCount)		
			
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

call closeDb()
%>
<!--#include file="inc_PrdCatTip.asp"-->

<div id="pcMain">
	<table class="pcMainTable">
		<% if IDBrand=0 then %>
		<tr>
			<td colspan="2">
				<h1><%response.write dictLanguage.Item(Session("language")&"_showSearchResults_1")%></h1>
			</td>
		</tr>  
        <% else %>
        <tr>
            <td colspan="2"> 
                <h1><%=pcvBrandName%></h1>
                <% if pcvBrandLogoLg<>"" then %>
                    <div style="width: 100%; text-align: center;"><img src="catalog/<%=pcvBrandLogoLg%>" alt="<%=ClearHTMLTags2(pcvBrandName,0)%>"></div>
                <% end if %>
                <% if pcvBrandsDescription<> "" then %>
                    <div class="pcPageDesc"><%=pcvBrandsDescription%></div>
                <% end if %>
            </td>
        </tr>
        <% end if %>    

		<%if pIdCategory="0" AND pcv_strHideCatSearch=False then%>

				<!--#include file="inc_srcPrdsCAT.asp"-->

		<%end if%>
		<%strORD=strPrdOrd%>
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr>
			<td class="pcSectionTitle">
            <% if pcv_strLimitPhrase="" then %>
				<%=dictLanguage.Item(Session("language")&"_advSrcb_3")%><%=totalrecords%> 
            <% else %>
				<%=dictLanguage.Item(Session("language")&"_advSrca_24")%> <%=totalrecords%> <%=dictLanguage.Item(Session("language")&"_advSrca_25")%>
            <% end if %>
            - <a href="search.asp"><%response.write dictLanguage.Item(Session("language")&"_ShowSearch_1")%></a>
			</td><td class="pcSectionTitle">
			<%if HideSortPro<>"1" then
			tKeywords=replace(tKeywords,"''''","''") %>
				<div class="pcSortProducts" style="z-index: 20;">
				<form name="neworder" class="pcForms">  
				<%=dictLanguage.Item(Session("language")&"_advSrca_16")%>     
				<select  name="order" onchange="javascript: if (document.neworder.order.value!='') location='showSearchResults.asp?VA=<%=pcv_ViewAllVar%>&SearchValues=<%=pCValues%>&exact=<%=intExact%>&iPageSize=<%=iPageSize%>&iPageCurrent=<%=iPageCurrent%>&pageStyle=<%=pcPageStyle%>&keyword=<%=replace(tKeywords,"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&SKU=<%=pSearchSKU%><%=pcv_strCSFieldQuery%>&order=' + document.neworder.order.value;">
					<option	value=""></option>
					<option value="0" <%if strPrdOrd="0" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_18")%></option>
					<option value="1" <%if strPrdOrd="1" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_19")%></option>
					<option value="3" <%if strPrdOrd="3" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_20")%></option>
					<option value="2" <%if strPrdOrd="2" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_advSrca_21")%></option>
				</select>
				</form></div><% end if %>
			</td> 
    </tr>
	<%
    if pCValues<>"" AND pCValues<>"0" then
		call openDb()
        %>
        <tr>
            <td colspan="2">
                <div style="padding-top:2px">
                    <%
                    tmpSValues=split(pCValues,"||")
                    For k=lbound(tmpSValues) to ubound(tmpSValues)
                        if tmpSValues(k)<>"" then
                            sfquery=""
							sfquery = "SELECT pcSearchFields.pcSearchFieldName, pcSearchData.pcSearchDataName FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchFields_Products.idSearchData = pcSearchData.idSearchData) ON pcSearchData.idSearchField = pcSearchFields.idSearchField  WHERE pcSearchFields_Products.idSearchData=" & tmpSValues(k)
                            set rsSearchFields=Server.CreateObject("ADODB.Recordset")
                            set rsSearchFields=connTemp.execute(sfquery)
                            If NOT rsSearchFields.eof Then
								pcv_strSearchDataName=rsSearchFields("pcSearchDataName")
								pcv_strSearchFieldName=rsSearchFields("pcSearchFieldName")
								pTempCValues = replace(pCValues,tmpSValues(k),"")
								pTempCValues = replace(pTempCValues,"||||","")
                            	%>
                            	<%=pcv_strSearchFieldName%>: <%=pcv_strSearchDataName%> <a href="<%=pcStrPageName%>?ProdSort=<%=ProdSort%>&iPageCurrent=<%=cint(iPageCurrent)%>&iPageSize=<%=iPageSize%>&PageStyle=<%=pcPageStyle%>&SearchValues=<%=pTempCValues%>&exact=<%=intExact%>&keyword=<%=replace(replace(tKeywords,"''''","''"),"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&order=<%=strORD%>&SKU=<%=pSearchSKU%><%=pcv_strCSFieldQuery%>"><img src="images/minus.jpg" border="0" hspace="2"></a>
                            	<%
								if k<(ubound(tmpSValues)-1) then
									response.Write("&nbsp;|&nbsp;")
								end if
                            End If
                            set rsSearchFields = nothing
                        end if
                    Next
                    %>
                </div>
        	</td>
       	</tr>
        <%
		call closeDb()
    end if
    %>
	<tr>
		<td colspan="2">
			<!-- Insert page navigation -->
			<% if pcv_ViewAllVar=0 then %>
				<!--#Include File="pcPageNavigation.asp"-->
			<% end if %>

			<% if pcPageStyle = "m" then %>
				<form action="instPrd.asp" method="post" name="m" id="m" class="pcForms">
			<% end if %>

			<table class="pcShowProducts">
				<%
				'*******************************
				' Add table headers for display
				' styles L and M
				'*******************************
				%>
				<% if pcPageStyle = "l" then	%>
                    <tr class="pcShowProductsLheader">
						<% if pShowSmallImg <> 0 then %>
                            <td>&nbsp;</td><% end if %>
                            <td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %></td><% if pShowSku <> 0 then %>
                            <td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %></td><% end if %>
                            <td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_10") %></td></tr>
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
                                        <td width="8%">
                                            <% response.write dictLanguage.Item(Session("language")&"_viewCat_P_7") %>
                                        </td>
                                    <% end if %>
                                <% end if %>
								<% if pShowSmallImg <> 0 then %>
									<td width="8%">&nbsp;</td>
								<% end if %>
								<% if pShowSku <> 0 then %>
									<td width="11%">
										<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %>
									</td>
								<% end if %>
                                <td width="47%">
                                    <% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %>
                                </td>
                                <td width="16%" align="center">
                                    <% If session("customerType")="1" then
                                            response.write dictLanguage.Item(Session("language")&"_viewCat_P_11")
                                         else
                                            response.write dictLanguage.Item(Session("language")&"_viewCat_P_10")
                                    end if %>
                                </td>
                            </tr>
					<% else %>
						<tr>
					<% end if %>
					<%
                    call openDb()
                    '*******************************
                    ' End table headers
                    '*******************************
    
                    '*******************************
                    ' Load product information
                    ' Loop through the products
                    '*******************************
                    
                    'Set the product count to zero
                    count=0

					if pcPageStyle = "m" then
						pCnt=Cint(0)
						pSQty=0
						pAllCnt=Cint(0)
					end if
					dim i, iRecordsShown
					i=0
					iRecordsShown=0
				
					if pcv_ViewAllVar=0 then
						newCount = pcv_strPageSize
					else
						newCount = 32767
					end if

					'// Loop until the total number of products to show	
					do while (tCnt < pcv_intProductCount) and (tCnt < cint(newCount))
						
						pidProduct=trim(pcArray_Products(0,tCnt)) '// rs("idProduct")				
						if pidProduct <> tempidProduct then
							tempidProduct=pidProduct
	
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
							psDesc=rsDescObj("sDesc")
							set rsDescObj=nothing
							
							if pcPageStyle = "m" then
								pCnt=pCnt+1
							end if
							tCnt=tCnt+1
							%>
							<!--#include file="pcGetPrdPrices.asp"-->
                            <%
                            '*******************************
                            ' Show product information
                            ' depending on the page style
                            '*******************************
                                        
                            ' FIRST STYLE - Show products horizontally, with images
                            if pcPageStyle = "h" then	%>
                                <td> 
                                    <!--#include file="pcShowProductH.asp" -->
                                </td><% i=i + 1
                                If i > (scPrdRow-1) then 
                                    response.write "</TR><TR>"
                                    i=0
                                End If
                            end if

							' SECOND STYLE - Show products vertically, with images 
							if pcPageStyle = "p" then	
								%> <td><!--#include file="pcShowProductP.asp" --></td></tr> <% 
							end if
								
							' THIRD STYLE - Show a list of products, with a small image 
							if pcPageStyle = "l" then	
								%> <!--#include file="pcShowProductL.asp" --> <% 
							end if
								
							' FOURTH STYLE - Show a list of products, with multiple add to cart 
							if pcPageStyle = "m" then	
								%> <!--#include file="pcShowProductM.asp" --> <% 
							end if
						 
							'*******************************
							' End show product information
							'*******************************
	
						end if ' End "if pidProduct <> tempidProduct"
	
						iRecordsShown=iRecordsShown + 1
						count=count + 1
					loop
					%>
                </table>
				<% 
                ' If page style is M, show the Add to Cart button when
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
                <!-- Insert page navigation -->
                <% if pcv_ViewAllVar=0 then %>
                    <!--#Include File="pcPageNavigation.asp"-->
			  	<% end if %>
				<%	  
                set rs=Nothing
                set iPageCurrent=Nothing
                call closeDb()
                %>
			</td>
    	</tr>
	</table>
    <!--#include file="atc_viewprd.asp"-->
</div>
<!--#include file="footer.asp"-->