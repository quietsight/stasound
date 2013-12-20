<!--#include file="../includes/SearchConstants.asp"-->
<%
IF (request("action")="newsrc") or (request("act")="newsrc") THEN
	iPageSize=getUserInput(request("resultCnt"),10)
	if iPageSize="" then
		iPageSize=getUserInput(request("iPageSize"),0)
	end if
	if (iPageSize="") then
		iPageSize=10
	end if
	if (not IsNumeric(iPageSize)) then
		iPageSize=10
	end if
	iPageCurrent=request("iPageCurrent")
	if iPageCurrent="" then
		iPageCurrent=1 
	end if
	pSKU=getUserInput(request("SKU"),150)
	pKeywords=getUserInput(request("keyWord"),100)
	pCValues=getUserInput(request("SearchValues"),0)
	tIncludeSKU=getUserInput(request("includeSKU"),10)
	pPriceFrom=getUserInput(request("priceFrom"),20)
	if trim(pPriceFrom)="" then
		pPriceFrom=0
	end if
	if NOT isNumeric(pPriceFrom) then
		pPriceFrom=0
	end if
	if Instr(pPriceFrom,",")>Instr(pPriceFrom,".") then
		pPriceFrom=replace(pPriceFrom,",",".")
	end if
	pPriceUntil=getUserInput(request("priceUntil"),20)
	if trim(pPriceUntil)="" then
		pPriceUntil=9999999
	end if
	if NOT isNumeric(pPriceUntil) then
		pPriceUntil=9999999
	end if
	if Instr(pPriceUntil,",")>Instr(pPriceUntil,".") then
		pPriceUntil=replace(pPriceUntil,",",".")
	end if
	pIdCategory=getUserInput(request("idCategory"),4)
	if NOT isNumeric(pIdCategory) then
		pIdCategory=0
	end if
	pWithStock=getUserInput(request("withStock"),2)
	if pWithStock="" then
		pWithStock=0
	end if
	
	pcustomfield=getUserInput(request("customfield"),0)
	
	IDBrand=getUserInput(request("IDBrand"),20)
	if NOT isNumeric(IDBrand) then
		IDBrand=0
	end if
	strORD=getUserInput(request("order"),4)
	if NOT isNumeric(strORD) then
		strORD=1
	end if
	pInactive=getUserInput(request("pInactive"),0)
	
	pcIntNotForSale=getUserInput(request("notforsale"),4)
	
	form_exact=request("exact")
	
	src_IncNormal=getUserInput(request("src_IncNormal"),0)
	src_IncBTO=getUserInput(request("src_IncBTO"),0)
	src_IncItem=getUserInput(request("src_IncItem"),0)
	src_SM=getUserInput(request("src_IncSM"),0)
	src_IncDown=getUserInput(request("src_IncDown"),0)
	src_IncGC=getUserInput(request("src_IncGC"),0)
	src_Special=getUserInput(request("src_Special"),0)
	src_Featured=getUserInput(request("src_Featured"),0)
	src_DiscType=getUserInput(request("src_DiscType"),0)
	src_DiscType=getUserInput(request("src_DiscType"),0)
	src_PromoType=getUserInput(request("src_PromoType"),0)
	
	if src_IncNormal="" then
		src_IncNormal="0"
	end if
	
	if src_IncBTO="" then
		src_IncBTO="0"
	end if
	
	if src_IncItem="" then
		src_IncItem="0"
	end if
	
	if src_Special="" then
		src_Special="0"
	end if

	if src_Featured="" then
		src_Featured="0"
	end if
	
	if (src_IncBTO="0") AND (src_IncItem="0") AND (src_IncDown="") AND (src_IncGC="") then
		src_IncNormal="1"
	end if

	'Start SDBA
	src_PageType=getUserInput(request("src_PageType"),0)
	src_IDSDS=getUserInput(request("src_IDSDS"),0)
	src_IsDropShipper=getUserInput(request("src_IsDropShipper"),0)
	src_sdsAssign=getUserInput(request("src_sdsAssign"),0)
	src_sdsStockAlarm=getUserInput(request("src_sdsStockAlarm"),0)
	
	src_StockLevel=getUserInput(request("stocklevel"),0)
	
	if src_PageType="" then
		src_PageType="0"
	end if
	
	if src_IDSDS="" then
		src_IDSDS="0"
	end if
	
	if src_IsDropShipper="" then
		src_IsDropShipper="0"
	end if
	
	if src_sdsAssign="" then
		src_sdsAssign="0"
	end if
	
	if src_sdsStockAlarm="" then
		src_sdsStockAlarm="0"
	end if
	'End SDBA
	
	session("cp_lct_form_iPageSize")=iPageSize
	session("cp_lct_form_iPageCurrent")=iPageCurrent
	session("cp_lct_form_sku")=pSKU
	session("cp_lct_form_keyWord")=pKeywords
	session("cp_lct_form_SearchValues")=pCValues
	session("cp_lct_form_priceFrom")=pPriceFrom
	session("cp_lct_form_priceUntil")=pPriceUntil
	session("cp_lct_form_idcategory")=pIdCategory
	session("cp_lct_form_withstock")=pWithStock
	session("cp_lct_form_stocklevel")=src_StockLevel
	session("cp_lct_form_customfield")=pcustomfield
	session("cp_lct_form_IDBrand")=IDBrand
	session("cp_lct_form_order")=strORD
	session("cp_lct_form_pinactive")=pInactive
	session("cp_lct_form_notforsale")=pcIntNotForSale
	session("cp_lct_form_exact")=form_exact
	session("cp_lct_src_IncNormal")=src_IncNormal
	session("cp_lct_src_IncBTO")=src_IncBTO
	session("cp_lct_src_IncItem")=src_IncItem
	session("cp_lct_src_SM")=src_SM
	session("cp_lct_src_IncDown")=src_IncDown
	session("cp_lct_src_IncGC")=src_IncGC
	session("cp_lct_src_Special")=src_Special
	session("cp_lct_src_Featured")=src_Featured
	session("cp_lct_src_DiscType")=src_DiscType
	session("cp_lct_src_PromoType")=src_PromoType
	'Start SDBA
	session("cp_lct_src_PageType")=src_PageType
	session("cp_lct_src_IDSDS")=src_IDSDS
	session("cp_lct_src_sdsAssign")=src_IsDropShipper
	session("cp_lct_src_sdsAssign")=src_sdsAssign
	session("cp_lct_src_sdsStockAlarm")=src_sdsStockAlarm
	'End SDBA

ELSE

	iPageSize=session("cp_lct_form_iPageSize")
	iPageCurrent=session("cp_lct_form_iPageCurrent")
	pSKU=session("cp_lct_form_sku")
	pKeywords=session("cp_lct_form_keyWord")
	pCValues=session("cp_lct_form_SearchValues")
	pPriceFrom=session("cp_lct_form_priceFrom")
	if Instr(pPriceFrom,",")>Instr(pPriceFrom,".") then
		pPriceFrom=replace(pPriceFrom,",",".")
	end if
	pPriceUntil=session("cp_lct_form_priceUntil")
	if Instr(pPriceUntil,",")>Instr(pPriceUntil,".") then
		pPriceUntil=replace(pPriceUntil,",",".")
	end if
	pIdCategory=session("cp_lct_form_idcategory")
	pWithStock=session("cp_lct_form_withstock")
	src_StockLevel=session("cp_lct_form_stocklevel")
	pcustomfield=session("cp_lct_form_customfield")
	IDBrand=session("cp_lct_form_IDBrand")
	strORD=session("cp_lct_form_order")
	pInactive=session("cp_lct_form_pinactive")
	pcIntNotForSale=session("cp_lct_form_notforsale")
	form_exact=session("cp_lct_form_exact")
	src_IncNormal=session("cp_lct_src_IncNormal")
	src_IncBTO=session("cp_lct_src_IncBTO")
	src_IncItem=session("cp_lct_src_IncItem")
	src_SM=session("cp_lct_src_SM")
	src_IncDown=session("cp_lct_src_IncDown")
	src_IncGC=session("cp_lct_src_IncGC")
	src_Special=session("cp_lct_src_Special")
	src_Featured=session("cp_lct_src_Featured")
	src_DiscType=session("cp_lct_src_DiscType")
	src_PromoType=session("cp_lct_src_PromoType")
	'Start SDBA
	src_PageType=session("cp_lct_src_PageType")
	src_IDSDS=session("cp_lct_src_IDSDS")
	src_IsDropShipper=session("cp_lct_src_sdsAssign")
	src_sdsAssign=session("cp_lct_src_sdsAssign")
	src_sdsStockAlarm=session("cp_lct_src_sdsStockAlarm")
	'End SDBA

END IF

if session("cp_lct_src_DiscType")<>"" OR session("cp_lct_src_PromoType")<>"" then
	session("srcprd_DiscArea")="1"
else
	session("srcprd_DiscArea")=""
end if
	
' create sql statement
Dim strSQL, tmpSQL, tmpSQL2

if strORD<>"" then
	Select Case StrORD
		Case "1": strORD1="products.idproduct ASC"
		Case "2": strORD1="products.sku ASC"
		Case "3": strORD1="products.sku DESC"
		Case "4": strORD1="products.description ASC"
		Case "5": strORD1="products.description DESC"
		Case "6": strORD1="products.stock ASC"
		Case "7": strORD1="products.stock DESC"
	End Select
Else
	strORD="1"
	strORD1="products.idproduct ASC"
End If
	
Dim PrdTypeStr

PrdTypeStr=""

if (src_IncBTO="0") and (src_IncItem="0") and (src_IncNormal="0") then
	if (src_IncDown="1") then
		PrdTypeStr="(products.Downloadable=1)"
	end if
	if (src_IncGC="1") then
		if PrdTypeStr<>"" then
			PrdTypeStr=PrdTypeStr & " OR "
		end if
		PrdTypeStr=PrdTypeStr & "(products.pcprod_GC=1)"
	end if
	PrdTypeStr= " AND (" & PrdTypeStr & ") "
end if

if (src_IncBTO="1") and (src_IncItem="0") and (src_IncNormal="0") then
	if (src_IncDown="1") then
		PrdTypeStr="(products.Downloadable=1)"
	end if
	if (src_IncGC="1") then
		if PrdTypeStr<>"" then
			PrdTypeStr=PrdTypeStr & " OR "
		end if
		PrdTypeStr=PrdTypeStr & "(products.pcprod_GC=1)"
	end if
	if PrdTypeStr<>"" then
		PrdTypeStr=PrdTypeStr & " OR "
	end if
	PrdTypeStr=PrdTypeStr & "(serviceSpec<>0)"
	PrdTypeStr= " AND (" & PrdTypeStr & ") "
end if

if (src_IncBTO="0") and (src_IncItem="1") and (src_IncNormal="0") then
	if (src_IncDown="1") then
		PrdTypeStr="(products.Downloadable=1)"
	end if
	if (src_IncGC="1") then
		if PrdTypeStr<>"" then
			PrdTypeStr=PrdTypeStr & " OR "
		end if
		PrdTypeStr=PrdTypeStr & "(products.pcprod_GC=1)"
	end if
	if PrdTypeStr<>"" then
		PrdTypeStr=PrdTypeStr & " OR "
	end if
	PrdTypeStr=PrdTypeStr & "(configOnly<>0)"
	PrdTypeStr= " AND (" & PrdTypeStr & ") "
end if

if (src_IncBTO="1") and (src_IncItem="1") and (src_IncNormal="0") then
	if (src_IncDown="1") then
		PrdTypeStr="(products.Downloadable=1)"
	end if
	if (src_IncGC="1") then
		if PrdTypeStr<>"" then
			PrdTypeStr=PrdTypeStr & " OR "
		end if
		PrdTypeStr=PrdTypeStr & "(products.pcprod_GC=1)"
	end if
	if PrdTypeStr<>"" then
		PrdTypeStr=PrdTypeStr & " OR "
	end if
	PrdTypeStr=PrdTypeStr & "((serviceSpec<>0) OR (configOnly<>0))"
	PrdTypeStr= " AND (" & PrdTypeStr & ") "
end if

if (src_IncBTO="0") and (src_IncItem="1") and (src_IncNormal="1") then
	PrdTypeStr=" AND serviceSpec=0 "
end if

if (src_IncBTO="1") and (src_IncItem="0") and (src_IncNormal="1") then
	PrdTypeStr=" AND configOnly=0 "
end if

if (src_IncBTO="0") and (src_IncItem="0") and (src_IncNormal="1") then
	PrdTypeStr=" AND ((serviceSpec=0) AND (configOnly=0)) "
end if


'Start SDBA
if (src_IDSDS<>"0") and (src_PageType="0") then
	if src_sdsAssign="1" then
		PrdTypeStr=PrdTypeStr & " AND products.pcSupplier_ID=0"
	else
		PrdTypeStr=PrdTypeStr & " AND products.pcSupplier_ID=" & src_IDSDS
	end if
end if
tmp_addFrom=""
if (src_IDSDS<>"0") and (src_PageType="1") then
	if src_sdsAssign="1" then
		PrdTypeStr=PrdTypeStr & " AND products.pcDropShipper_ID=0"
	else
		PrdTypeStr=PrdTypeStr & " AND products.pcDropShipper_ID=" & src_IDSDS & " AND pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & src_IsDropShipper
		tmp_addFrom=",pcDropShippersSuppliers"
	end if
end if
if (src_sdsStockAlarm="1") then
	PrdTypeStr=PrdTypeStr & " AND products.stock<products.pcProd_ReorderLevel"
end if
'End SDBA

if pIdCategory<>"0" then
	tmpSQL1=",categories_products "
else
	tmpSQL1=""
end if

tmpStrEx=""
if src_DiscType="1" then
	tmpSQL1=tmpSQL1 & ",discountsPerQuantity "
	tmpStrEx=" AND (products.idproduct=discountsPerQuantity.idproduct)"
else
	if src_DiscType="2" then
		tmpStrEx=" AND (products.idproduct NOT IN (SELECT DISTINCT idproduct FROM discountsPerQuantity)) AND (products.idproduct NOT IN (SELECT DISTINCT idproduct FROM pcPrdPromotions))"
	else
		if session("srcprd_DiscArea")="1" AND (src_PromoType="") then
			tmpStrEx=" AND (products.idproduct NOT IN (SELECT DISTINCT idproduct FROM pcPrdPromotions))"
		end if
	end if
end if

if src_PromoType="1" then
	tmpSQL1=tmpSQL1 & ",pcPrdPromotions "
	tmpStrEx=" AND (products.idproduct=pcPrdPromotions.idproduct)"
else
	if src_PromoType="2" then
		tmpStrEx=" AND (products.idproduct NOT IN (SELECT DISTINCT idproduct FROM discountsPerQuantity)) AND (products.idproduct NOT IN (SELECT DISTINCT idproduct FROM pcPrdPromotions))"
	else
		if src_PromoType="0" then
			tmpStrEx=" AND (products.idproduct NOT IN (SELECT DISTINCT idproduct FROM discountsPerQuantity))"
		end if
	end if
end if

pcv_strMaxResults=SRCH_MAX
If pcv_strMaxResults>"0" Then
	pcv_strLimitPhrase="TOP " & pcv_strMaxResults
Else
	pcv_strLimitPhrase=""
End If

strSQL="SELECT DISTINCT " & pcv_strLimitPhrase & " products.idProduct, products.description, products.active, products.smallImageUrl, products.sku, products.serviceSpec,products.configonly,products.stock,products.pcProd_ReorderLevel,products.price,products.cost FROM products " & tmpSQL1 & session("srcprd_from") & tmp_addFrom & " WHERE products.price>="&pPriceFrom&" And products.price<=" &pPriceUntil&" AND products.removed=0"

if len(pSKU)>0 then
	strSQL=strSQL & " AND products.sku like '%"&pSKU&"%'"
end if

if pIdCategory<>"0" then
	strSQL=strSQL & " AND products.idProduct=categories_products.idProduct AND categories_products.idCategory=" &pIdCategory   
end if

if pWithStock="-1" then
	strSQL=strSQL & " AND (stock>0 OR noStock<>0)" 
end if

if pWithStock="2" then
	strSQL=strSQL & " AND (stock<=0 AND noStock=0)" 
end if

if src_StockLevel<>"" then
	strSQL=strSQL & " AND (stock<" & src_StockLevel & ")" 
end if

if (IDBrand&""<>"") and (IDBrand&""<>"0") then
	strSQL=strSQL & " AND IDBrand=" & IDBrand
end if

if pInactive="-1" then
else
   strSQL=strSQL & " AND products.active=-1" 
end if

if pcIntNotForSale="-1" then
   strSQL=strSQL & " AND products.formQuantity=0" 
elseif pcIntNotForSale="2" then
   strSQL=strSQL & " AND products.formQuantity=-1" 
else
end if

if src_Special="1" then
   strSQL=strSQL & " AND products.hotdeal<>0" 
end if

if src_Special="2" then
   strSQL=strSQL & " AND products.hotdeal=0" 
end if

if src_Featured="1" then
   strSQL=strSQL & " AND products.showInHome<>0" 
end if

if src_Featured="2" then
   strSQL=strSQL & " AND products.showInHome=0" 
end if

TestWord=""
if form_exact<>"1" then
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

if pCValues<>"" AND pCValues<>"0" then
	tmpSValues=split(pCValues,"||")
	For k=lbound(tmpSValues) to ubound(tmpSValues)
		if tmpSValues(k)<>"" then
			tmpStrEx=tmpStrEx & " AND products.idproduct IN (SELECT idproduct FROM pcSearchFields_Products WHERE idSearchData=" & tmpSValues(k) & ")"
		end if
	Next
end if

IF form_exact<>"1" THEN

if pKeywords<>"" then

	strSQl=strSql & " AND ("
	
	tmpSQL="(details LIKE "
	tmpSQL2="(description LIKE "
	tmpSQL3="(sDesc LIKE "
	tmpSQL5="(pcProd_MetaKeywords LIKE "
	if tIncludeSKU="true" then
		tmpSQL4="(SKU LIKE "
	end if
	Dim Pos
	Pos=0
	For L=LBound(keywordArray) to UBound(keywordArray)
		if trim(keywordArray(L))<>"" then
		Pos=Pos+1
		if Pos>1 Then
			tmpSQL=tmpSQL  & TestWord & " details LIKE "
			tmpSQL2=tmpSQL2 & TestWord & " description LIKE "
			tmpSQL3=tmpSQL3 & TestWord & " sDesc LIKE "
			tmpSQL5=tmpSQL5 & TestWord & " pcProd_MetaKeywords LIKE "
			if tIncludeSKU="true" then
				tmpSQL4=tmpSQL4 & TestWord & " SKU LIKE "
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
	query=strSQL & PrdTypeStr & tmpStrEx & session("srcprd_where") & " ORDER BY " & strORD1
else
	query=strSQL & PrdTypeStr & tmpStrEx & session("srcprd_where") & " ORDER BY " & strORD1
end if

ELSE 'Exact=1

if pKeywords<>"" then

	strSQl=strSql & " AND ("
	
	tmpSQL="(details LIKE "
	tmpSQL2="(description LIKE "
	tmpSQL3="(sDesc LIKE "
	tmpSQL5="(pcProd_MetaKeywords LIKE "
	if tIncludeSKU="true" then
		tmpSQL4="(SKU LIKE "
	end if
	Pos=0
	For L=LBound(keywordArray) to UBound(keywordArray)
		if trim(keywordArray(L))<>"" then
		Pos=Pos+1
		if Pos>1 Then
			tmpSQL=tmpSQL  & TestWord & " details LIKE "
			tmpSQL2=tmpSQL2 & TestWord & " description LIKE "
			tmpSQL3=tmpSQL3 & TestWord & " sDesc LIKE "
			tmpSQL5=tmpSQL5 & TestWord & " pcProd_MetaKeywords LIKE "
			if tIncludeSKU="true" then
				tmpSQL4=tmpSQL4 & TestWord & " SKU LIKE "
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
	query=strSQL & PrdTypeStr & tmpStrEx & session("srcprd_where") & " ORDER BY " & strORD1
else
	query=strSQL & PrdTypeStr & tmpStrEx & session("srcprd_where") & " ORDER BY " & strORD1
end if

END IF 'Exact
%>
