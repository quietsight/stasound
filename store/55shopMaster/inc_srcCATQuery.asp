<%
iPageSize=getUserInput(request("resultCnt"),10)
if iPageSize="" then
	iPageSize=10
end if
if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=server.HTMLEncode(request("iPageCurrent"))
end if

Function CreateQuery(Desc,keynum)
Dim m
Dim tmpStr,keywordArray,keylink,keydesc

	tmpStr=""

	Select Case keynum
		Case 1: keydesc="categoryDesc"
		Case 2: keydesc="SDesc"
		Case 3: keydesc="LDesc"
	End Select

	if Instr(Desc," AND ")>0 then
		keywordArray=split(Desc," AND ")
		keylink=" AND "
	else
	if Instr(Desc,",")>0 then
		keywordArray=split(Desc,",")
		keylink=" OR "
	else
		if Instr(Desc," OR ")>0 then
			keywordArray=split(Desc," OR ")
			keylink=" OR "
		else
			keywordArray=split(Desc,"***")
			keylink=" OR "
		end if
	end if
	end if

			
	For m=lbound(keywordArray) to ubound(keywordArray)
	if trim(keywordArray(m))<>"" then
		if tmpStr<>"" then
		tmpStr=tmpStr & keylink
		end if
		tmpStr=tmpStr & "(" & keydesc & " like '%"&trim(keywordArray(m))&"%')"
	end if
	Next
	
	if tmpStr<>"" then
		tmpStr="(" & tmpStr & ")"
	else
		tmpStr="(" & keydesc & " like '%"&Desc&"%')"
	end if

CreateQuery=tmpStr
End Function

strORD=getUserInput(request("order"),4)
if NOT isNumeric(strORD) then
	strORD=1
end if

if strORD<>"" then
	Select Case StrORD
		Case "1": strORD1="categories.categoryDesc ASC"
		Case "2": strORD1="categories.categoryDesc ASC"
		Case "3": strORD1="categories.categoryDesc DESC"
	End Select
Else
	strORD="1"
	strORD1="categories.categoryDesc ASC"
End If

' create sql statement
	query1=""
	query2=""
	if request("key1")<>"" then
		tmpKey=request("key1")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,1)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key2")<>"" then
		tmpKey=request("key2")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,2)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key3")<>"" then
		tmpKey=request("key3")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,3)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	'Start - Filter Category By Discount
	query2=""
	if request("CatDiscType") <> "" AND request("CatDiscType")<>"0" then
		if request("CatDiscType")="1" then 'Categories With Quantity Discount
			query2 = " categories.idCategory IN "
		elseif request("CatDiscType")="2" then 'Categories Without Quantity Discount
			query2 = " categories.idCategory NOT IN "
		end if
		if query2<>"" then
			query2 = query2 & " ( SELECT pcCD_idcategory FROM pcCatDiscounts ) "
			if query1<>"" then
				query1=query1 & " AND "
			end if
			query1=query1 & query2
			if request("CatDiscType")="2" then
				query1=query1 & " AND (categories.idcategory NOT IN (SELECT DISTINCT idcategory FROM pcCatPromotions))"
			end if
		end if
	else
		if request("CatDiscType")="0" AND request("CatPromoType")="" then
			if query1<>"" then
				query1=query1 & " AND "
			end if
			query1=query1 & " (categories.idcategory NOT IN (SELECT DISTINCT idcategory FROM pcCatPromotions))"
		end if
	end if
	'End - Filter Category By Discount
	
	'Start - Filter Category By Promotion
	query2=""
	if request("CatPromoType") <> "" AND request("CatPromoType")<>"0" then
		if request("CatPromoType")="1" then 'Categories With Quantity Discount
			query2 = " categories.idCategory IN "
		elseif request("CatPromoType")="2" then 'Categories Without Quantity Discount
			query2 = " categories.idCategory NOT IN "
		end if
		if query2<>"" then
			query2 = query2 & " ( SELECT idcategory FROM pcCatPromotions ) "
			if query1<>"" then
				query1=query1 & " AND "
			end if
			query1=query1 & query2
			if request("CatPromoType")="2" then
				query1=query1 & " AND (categories.idcategory NOT IN (SELECT pcCD_idcategory FROM pcCatDiscounts))"
			end if
		end if
	else
		if request("CatDiscType")="" AND request("CatPromoType")="0" then
			if query1<>"" then
				query1=query1 & " AND "
			end if
			query1=query1 & " (categories.idcategory NOT IN (SELECT pcCD_idcategory FROM pcCatDiscounts))"
		end if
	end if
	'End - Filter Category By Promotion
	
	src_IncNotShDesc=getUserInput(request("src_IncNotShDesc"),0)
	if src_IncNotShDesc="" then
		src_IncNotShDesc=0
	end if
	src_IncNotDisplay=getUserInput(request("src_IncNotDisplay"),0)
	if src_IncNotDisplay="" then
		src_IncNotDisplay="0"
	end if
	src_IncNotFRetail=getUserInput(request("src_IncNotFRetail"),0)
	if src_IncNotFRetail="" then
		src_IncNotFRetail="0"
	end if
	src_ParentOnly=getUserInput(request("src_ParentOnly"),0)
	if src_ParentOnly="" then
	src_ParentOnly="0"
	end if
	
	query3=""
	
	if src_IncNotShDesc="1" then
		query3=" AND categories.HideDesc<>0"
	end if
	
	if src_IncNotShDesc="2" then
		query3=" AND categories.HideDesc=0"
	end if
	
	if src_IncNotDisplay="1" then
		query3=query3 & " AND categories.iBTOhide<>0"
	end if
	
	if src_IncNotDisplay="2" then
		query3=query3 & " AND categories.iBTOhide=0"
	end if
	
	if src_IncNotFRetail="1" then
		query3=query3 & " AND categories.pccats_RetailHide<>0"
	end if
	
	if src_IncNotFRetail="2" then
		query3=query3 & " AND categories.pccats_RetailHide=0"
	end if
	
	if request("CatPromoType") <> "" then
		query="SELECT DISTINCT idcategory,categoryDesc,idParentCategory,iBTOhide FROM categories " & session("srcCat_from") & " WHERE categories.idcategory<>1 "
	else
		if src_ParentOnly="1" then
			query="SELECT DISTINCT categories.idcategory,categories.categorydesc,categories.idParentCategory,categories.iBTOhide FROM categories " & session("srcCat_from") & " WHERE ((categories.idCategory) Not in (SELECT DISTINCT categories_products.idCategory FROM categories_products)) AND categories.idcategory<>1 "
		else
			if src_ParentOnly="2" then
				query="SELECT DISTINCT categories.idcategory,categories.categorydesc,categories.idParentCategory,categories.iBTOhide FROM categories " & session("srcCat_from") & " WHERE idCategory>1 AND idCategory NOT IN (SELECT A.idCategory from categories A, categories B WHERE A.idcategory=B.idparentcategory) AND categories.idcategory<>1 AND categories.idcategory IN (SELECT DISTINCT categories_products.idcategory FROM categories_products GROUP by (categories_products.idcategory))"
			else
				query="SELECT DISTINCT idcategory,categoryDesc,idParentCategory,iBTOhide FROM categories " & session("srcCat_from") & " WHERE categories.idcategory<>1 "
			end if
		end if
	end if
	if query1<>"" then
	query=query & " AND " & query1
	end if
	query=query & " " & query3 & session("srcCat_where")

	query=query&" ORDER BY "& strORD1
%>