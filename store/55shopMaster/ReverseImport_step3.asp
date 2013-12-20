<% pageTitle = "Reverse Import Wizard - Export Results" %>
<% section = "products"
Server.ScriptTimeout = 5400%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, rstemp, rs, connTemp
Dim pcArr,i,tmp_query,intCount,pcv_HaveRecords
Dim fs,A,strFile
Dim fseparator
Dim valueArr

pcv_intExportSize = session("cp_ExportSize")
fseparator=session("cp_revImport_cseparator")
if session("cp_revImport_prdlist")="" then
	response.redirect "ReverseImport_step1.asp"
end if

call opendb()

set DataBuilderObj = new StringBuilder
pcv_HaveRecords=0

Function RmvHTMLWhiteSpace(tmpValue)
	Dim tmp1,re,colMatch,objMatch
	tmp1=tmpValue
	Set re = New RegExp

	With re
	  .Pattern = "(\r\n[\s]+)"
	  .Global = True
	End With 

	Set colMatch = re.Execute(tmp1)
	For each objMatch in colMatch
		tmp1=replace(tmp1,objMatch.Value," ")
	Next
	RmvHTMLWhiteSpace=tmp1
End Function

Function GenFileName()
	dim fname
	fname="File-"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime)) & "-"
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

Function getAttrList(tmp_IDProduct,tmp_IDOption)
	Dim query,rs2

	query =	"SELECT options_optionsGroups.InActive, options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice, options_optionsGroups.sortOrder "
	query = query & "FROM options_optionsGroups "
	query = query & "INNER JOIN options "
	query = query & "ON options_optionsGroups.idOption = options.idOption "
	query = query & "WHERE options_optionsGroups.idOptionGroup=" & tmp_IDOption &" "
	query = query & "AND options_optionsGroups.idProduct=" & tmp_IDProduct &" "
	query = query & "ORDER BY options_optionsGroups.sortOrder;"	
	set rs2=server.createobject("adodb.recordset")
	set rs2=conntemp.execute(query)
	set StringBuilderObj = new StringBuilder
	do while not rs2.eof
		pcv_OptInActive=rs2("InActive")
		if IsNull(pcv_OptInActive) or pcv_OptInActive="" then
			pcv_OptInActive="0"
		end if
		if pcv_OptInActive="0" then
			pcv_OptActive="1"
		else
			pcv_OptActive="0"
		end if
		StringBuilderObj.append replace(rs2("optionDescrip"),"""","""""") & "*" & rs2("price") & "*" & rs2("Wprice") & "*" & rs2("sortOrder") & "*" & pcv_OptActive & "**"
		rs2.MoveNext
	loop
	set rs2=nothing
	getAttrList=StringBuilderObj.toString()
	set StringBuilderObj = nothing
End Function

Function getPrdOptions(pcv_IDProduct)
	Dim query,rs1,Count,m

	query="SELECT pcProductsOptions.idOptionGroup,optionsGroups.OptionGroupDesc,pcProductsOptions.pcProdOpt_Required,pcProductsOptions.pcProdOpt_order FROM optionsGroups INNER JOIN pcProductsOptions ON optionsGroups.idOptionGroup=pcProductsOptions.idOptionGroup WHERE pcProductsOptions.idProduct=" & pcv_IDProduct & ";"
	set rs1=connTemp.execute(query)
	if not rs1.eof then
		set StringBuilderObj = new StringBuilder
		Count=0
		Do while (not rs1.eof) and (Count<5)
			Count=Count+1
			pcv_IDOption=rs1("idOptionGroup")
			StringBuilderObj.append """" & replace(rs1("OptionGroupDesc"),"""","""""") & """" & fseparator & """" & getAttrList(pcv_IDProduct,pcv_IDOption) & """" & fseparator & """" & rs1("pcProdOpt_Required") & """" & fseparator & """" & rs1("pcProdOpt_order") & """" & fseparator
			rs1.MoveNext
		Loop
		For m=Count+1 to 5
			StringBuilderObj.append """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator
		Next
		getPrdOptions=StringBuilderObj.toString()
	else
		getPrdOptions="""""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator
	end if
	set rs1=nothing
	set StringBuilderObj = nothing
End Function

Function getIsSupplier(pcv_IDProduct)
	Dim query,rs1
	query="SELECT pcDS_IsDropShipper FROM pcDropShippersSuppliers WHERE idProduct=" & pcv_IDProduct & ";"
	set rs1=connTemp.execute(query)
	if not rs1.eof then
		getIsSupplier="""" & rs1("pcDS_IsDropShipper") & """" & fseparator
	else
		getIsSupplier="""0""" & fseparator
	end if
	set rs1=nothing	
End Function

Function getParentName(pcv_IDParent)
	Dim query,rs2
	query="SELECT categoryDesc FROM categories WHERE idCategory=" & pcv_IDParent & ";"
	set rs2=connTemp.execute(query)
	if not rs2.eof then
		getParentName=replace(rs2("categoryDesc"),"""","""""")
	else
		getParentName=""
	end if
	set rs2=nothing	
End Function

Function getCATInfor(pcv_IDProduct)
	Dim query,rs1,Count,m
	query="SELECT categories.categoryDesc,categories.SDesc,categories.LDesc,categories.[image],categories.largeimage,categories.idParentCategory FROM categories INNER JOIN categories_products ON categories.idCategory=categories_products.idCategory WHERE categories_products.idProduct=" & pcv_IDProduct & ";"
	set rs1=connTemp.execute(query)
	if not rs1.eof then
		set StringBuilderObj = new StringBuilder
		Count=0
		Do while (not rs1.eof) and (Count<3)
			Count=Count+1
			tmp_CatName=rs1("categoryDesc")
			if tmp_CatName<>"" then
				tmp_CatName=replace(tmp_CatName,"""","""""")
			end if
			tmp_CatSDesc=rs1("SDesc")
			if tmp_CatSDesc<>"" then
				tmp_CatSDesc=replace(tmp_CatSDesc,"""","""""")
			end if
			tmp_CatLDesc=rs1("LDesc")
			if tmp_CatLDesc<>"" then
				tmp_CatLDesc=replace(tmp_CatLDesc,"""","""""")
			end if
			StringBuilderObj.append """" & tmp_CatName & """" & fseparator & """" & tmp_CatSDesc & """" & fseparator & """" & tmp_CatLDesc & """" & fseparator & """" & rs1("image") & """" & fseparator & """" & rs1("largeimage") & """" & fseparator & """" & getParentName(rs1("idParentCategory")) & """" & fseparator
			rs1.MoveNext
		Loop
		For m=Count+1 to 3
			StringBuilderObj.append """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator
		Next
		getCATInfor=StringBuilderObj.toString()
		set StringBuilderObj = nothing
	else
		getCATInfor="""""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator
	end if
	set rs1=nothing
End Function

Function getGCInfor(pcv_IDProduct,pcv_GiftCert)
	Dim query,rs1	
	IF pcv_GiftCert="0" THEN
		getGCInfor="""" & pcv_GiftCert & """" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator
	ELSE
		query="SELECT pcGC_Exp,pcGC_EOnly,pcGC_CodeGen,pcGC_ExpDate,pcGC_ExpDays,pcGC_GenFile FROM pcGC WHERE pcGC_IDProduct=" & pcv_IDProduct & ";"
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			getGCInfor="""" & pcv_GiftCert & """" & fseparator & """" & rs1("pcGC_Exp") & """" & fseparator & """" & rs1("pcGC_EOnly") & """" & fseparator & """" & rs1("pcGC_CodeGen") & """" & fseparator & """" & rs1("pcGC_ExpDate") & """" & fseparator & """" & rs1("pcGC_ExpDays") & """" & fseparator & """" & rs1("pcGC_GenFile") & """" & fseparator
		else
			getGCInfor="""" & pcv_GiftCert & """" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator
		end if
		set rs1=nothing
	END IF
End Function

Function getDownloadInfor(pcv_IDProduct,pcv_Downloadable)
	Dim query,rs1	
	IF pcv_Downloadable="0" THEN
		getDownloadInfor="""" & pcv_Downloadable & """" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator
	ELSE
		query="SELECT ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail FROM DProducts WHERE IDProduct=" & pcv_IDProduct & ";"
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			set StringBuilderObj = new StringBuilder
			StringBuilderObj.append """" & pcv_Downloadable & """" & fseparator & """" & rs1("ProductURL") & """" & fseparator & """" & rs1("URLExpire") & """" & fseparator & """" & rs1("ExpireDays") & """" & fseparator & """" & rs1("License") & """" & fseparator & """" & rs1("LocalLG") & """" & fseparator & """" & rs1("RemoteLG") & """" & fseparator & """" & rs1("LicenseLabel1") & """" & fseparator & """" & rs1("LicenseLabel2") & """" & fseparator & """" & rs1("LicenseLabel3") & """" & fseparator & """" & rs1("LicenseLabel4") & """" & fseparator & """" & rs1("LicenseLabel5") & """" & fseparator & """"
			tmp_AddToMail=rs1("AddToMail")
			if tmp_AddToMail<>"" then
				tmp_AddToMail=replace(tmp_AddToMail,"""","""""")
			end if
			StringBuilderObj.append tmp_AddToMail & """" & fseparator
			getDownloadInfor=StringBuilderObj.toString()
			set StringBuilderObj = nothing
		else
			getDownloadInfor="""" & pcv_Downloadable & """" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & """""" & fseparator & ""
		end if
		set rs1=nothing
	END IF	
End Function

Function getBrandInfor(pcv_IDBrand)
	Dim query,rs1	
	IF pcv_IDBrand="0" THEN
		getBrandInfor="""""" & fseparator & """""" & fseparator
	ELSE
		query="SELECT BrandName,BrandLogo FROM Brands WHERE IdBrand=" & pcv_IDBrand & ";"
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			getBrandInfor="""" & rs1("BrandName") & """" & fseparator & """" & rs1("BrandLogo") & """" & fseparator
		else
			getBrandInfor="""""" & fseparator & """""" & fseparator
		end if
		set rs1=nothing
	END IF
End Function

'***** Generate HeadLine *****
set HeaderBuilderObj = new StringBuilder
if request("C1")="1" then
	HeaderBuilderObj.append """SKU""" & fseparator
end if
if request("C2")="1" then
	HeaderBuilderObj.append """Name""" & fseparator
end if
if request("C3")="1" then
	HeaderBuilderObj.append """Description""" & fseparator
end if
if request("C4")="1" then
	HeaderBuilderObj.append """Short Description""" & fseparator
end if
if request("C5")="1" then
	HeaderBuilderObj.append """Product Type""" & fseparator
end if
if request("C6")="1" then
	HeaderBuilderObj.append """Online Price""" & fseparator
end if
if request("C7")="1" then
	HeaderBuilderObj.append """List Price""" & fseparator
end if
if request("C8")="1" then
	HeaderBuilderObj.append """Wholesale Price""" & fseparator
end if
if request("C9")="1" then
	HeaderBuilderObj.append """Weight""" & fseparator
end if
if request("C10")="1" then
	HeaderBuilderObj.append """Stock""" & fseparator
end if
if request("C11")="1" then
	HeaderBuilderObj.append """Category Name""" & fseparator & """Short Category Description""" & fseparator & """Long Category Description""" & fseparator & """Category Small Image""" & fseparator & """Category Large Image""" & fseparator & """Parent Category""" & fseparator & """Additional Category 1""" & fseparator & """Short Category Description 1""" & fseparator & """Long Category Description 1""" & fseparator & """Category Small Image 1""" & fseparator & """Category Large Image 1""" & fseparator & """Parent Category 1""" & fseparator & """Additional Category 2""" & fseparator & """Short Category Description 2""" & fseparator & """Long Category Description 2""" & fseparator & """Category Small Image 2""" & fseparator & """Category Large Image 2""" & fseparator & """Parent Category 2""" & fseparator
end if
if request("C12")="1" then
	HeaderBuilderObj.append """Brand Name""" & fseparator & """Brand Logo""" & fseparator
end if
if request("C13")="1" then
	HeaderBuilderObj.append """Thumbnail Image""" & fseparator
end if
if request("C14")="1" then
	HeaderBuilderObj.append """General Image""" & fseparator
end if
if request("C15")="1" then
	HeaderBuilderObj.append """Detail view Image""" & fseparator
end if
if request("C16")="1" then
	HeaderBuilderObj.append """Active""" & fseparator
end if
if request("C17")="1" then
	HeaderBuilderObj.append """Show savings""" & fseparator
end if
if request("C18")="1" then
	HeaderBuilderObj.append """Special""" & fseparator
end if
if request("C46")="1" then
	HeaderBuilderObj.append """Featured""" & fseparator
end if
if request("C19")="1" then
	HeaderBuilderObj.append """Option 1""" & fseparator & """Attributes 1""" & fseparator & """Option 1 Required""" & fseparator & """Option 1 Order""" & fseparator & """Option 2""" & fseparator & """Attributes 2""" & fseparator & """Option 2 Required""" & fseparator & """Option 2 Order""" & fseparator & """Option 3""" & fseparator & """Attributes 3""" & fseparator & """Option 3 Required""" & fseparator & """Option 3 Order""" & fseparator & """Option 4""" & fseparator & """Attributes 4""" & fseparator & """Option 4 Required""" & fseparator & """Option 4 Order""" & fseparator & """Option 5""" & fseparator & """Attributes 5""" & fseparator & """Option 5 Required""" & fseparator & """Option 5 Order""" & fseparator
end if
if request("C20")="1" then
	HeaderBuilderObj.append """Reward Points""" & fseparator
end if
if request("C21")="1" then
	HeaderBuilderObj.append """Non-taxable""" & fseparator
end if
if request("C22")="1" then
	HeaderBuilderObj.append """No shipping charge""" & fseparator
end if
if request("C23")="1" then
	HeaderBuilderObj.append """Not for sale""" & fseparator
end if
if request("C24")="1" then
	HeaderBuilderObj.append """Not for sale copy""" & fseparator
end if
if request("C25")="1" then
	HeaderBuilderObj.append """Disregard stock""" & fseparator
end if
if request("C26")="1" then
	HeaderBuilderObj.append """Display No Shipping Text""" & fseparator
end if
if request("C27")="1" then
	HeaderBuilderObj.append """Minimum Quantity customers can buy""" & fseparator
end if
if request("C28")="1" then
	HeaderBuilderObj.append """Force purchase of multiples of minimum""" & fseparator
end if
if request("C29")="1" then
	HeaderBuilderObj.append """Oversized Product Details""" & fseparator
end if
if request("C30")="1" then
	HeaderBuilderObj.append """Product Cost""" & fseparator
end if
if request("C31")="1" then
	HeaderBuilderObj.append """Back-Order""" & fseparator
end if
if request("C32")="1" then
	HeaderBuilderObj.append """Ship within N Days""" & fseparator
end if
if request("C33")="1" then
	HeaderBuilderObj.append """Low inventory notification""" & fseparator
end if
if request("C34")="1" then
	HeaderBuilderObj.append """Reorder Level""" & fseparator
end if
if request("C35")="1" then
	HeaderBuilderObj.append """Is Drop-shipped""" & fseparator
end if
if request("C36")="1" then
	HeaderBuilderObj.append """Supplier ID""" & fseparator
end if
if request("C37")="1" then
	HeaderBuilderObj.append """Drop-Shipper ID""" & fseparator
end if
if request("C38")="1" then
	HeaderBuilderObj.append """Drop-Shipper is also a Supplier""" & fseparator
end if
if request("C39")="1" then
	HeaderBuilderObj.append """Meta Tags - Title""" & fseparator
	HeaderBuilderObj.append """Meta Tags - Description""" & fseparator
	HeaderBuilderObj.append """Meta Tags - Keywords""" & fseparator
end if
if request("C40")="1" then
	HeaderBuilderObj.append """Downloadable Product""" & fseparator & """Downloadable File Location""" & fseparator & """Make Download URL expire""" & fseparator & """URL Expiration in Days""" & fseparator & """Use License Generator""" & fseparator & """Local Generator""" & fseparator & """Remote Generator""" & fseparator & """License Field Label (1)""" & fseparator & """License Field Label (2)""" & fseparator & """License Field Label (3)""" & fseparator & """License Field Label (4)""" & fseparator & """License Field Label (5)""" & fseparator & """Additional copy""" & fseparator
end if
if request("C41")="1" then
	HeaderBuilderObj.append """Gift Certificate""" & fseparator & """Gift Certificate Expiration""" & fseparator & """Electronic Only (Gift Certificate)""" & fseparator & """Use Generator (Gift Certificate)""" & fseparator & """Expiration Date (Gift Certificate)""" & fseparator & """Expire N days (Gift Certificate)""" & fseparator & """Custom Generator Filename (Gift Certificate)""" & fseparator
end if
if request("C42")="1" then
	HeaderBuilderObj.append """Hide BTO Prices""" & fseparator
end if
if request("C43")="1" then
	HeaderBuilderObj.append """Hide Default Configuration""" & fseparator
end if
if request("C44")="1" then
	HeaderBuilderObj.append """Disallow Purchasing""" & fseparator
end if
if request("C45")="1" then
	HeaderBuilderObj.append """Skip Product Details Page""" & fseparator
end if
if request("C49")="1" then
	HeaderBuilderObj.append """Units to make 1 lb""" & fseparator
end if
if request("C50")="1" then
	HeaderBuilderObj.append """First Unit Surcharge""" & fseparator
end if
if request("C51")="1" then
	HeaderBuilderObj.append """Additional Unit(s) Surcharge""" & fseparator
end if
if request("C52")="1" then
	HeaderBuilderObj.append """Product Notes""" & fseparator
end if
if request("C53")="1" then
	HeaderBuilderObj.append """Enable Image Magnifier""" & fseparator
end if
if request("C54")="1" then
	HeaderBuilderObj.append """Page Layout""" & fseparator
end if
if request("C55")="1" then
	HeaderBuilderObj.append """Hide SKU on the product details page""" & fseparator
end if
if request("C56")="1" then
	HeaderBuilderObj.append """Google Product Category""" & fseparator
end if
if request("C57")="1" then
	HeaderBuilderObj.append """Google Shopping - Gender"",""Google Shopping - Age"",""Google Shopping - Color"",""Google Shopping - Size"",""Google Shopping - Pattern"",""Google Shopping - Material""" & fseparator
end if
'***** End of Generate HeadLine *****

'***** Create SQL Query *****
query="SELECT sku, description, serviceSpec, configOnly, price, listPrice, bToBPrice, weight, stock, IDBrand, smallImageUrl, imageUrl, largeImageURL, active, listHidden, hotDeal, iRewardPoints, notax, noshipping, formQuantity, emailText, noStock, noshippingtext, pcprod_minimumqty, pcprod_qtyvalidate, OverSizeSpec, cost, pcProd_BackOrder, pcProd_ShipNDays, pcProd_NotifyStock, pcProd_ReorderLevel, pcProd_IsDropShipped, pcSupplier_ID, pcDropShipper_ID, Downloadable, pcprod_GC, pcprod_hidebtoprice, pcprod_HideDefConfig, NoPrices, pcProd_SkipDetailsPage, idproduct, ShowInHome, pcprod_QtyToPound, pcProd_Surcharge1, pcProd_Surcharge2, pcPrd_MojoZoom, pcProd_HideSKU, products.pcProd_GoogleCat,products.pcProd_GoogleGender,products.pcProd_GoogleAge,products.pcProd_GoogleColor,products.pcProd_GoogleSize,products.pcProd_GooglePattern,products.pcProd_GoogleMaterial, pcProd_DisplayLayout, pcProd_PrdNotes, pcProd_MetaTitle, pcProd_MetaDesc, pcProd_MetaKeywords, details, sDesc FROM Products WHERE removed=0 "
'Last: 60

if session("cp_revImport_prdlist")<>"ALL" then
	pcArr=split(session("cp_revImport_prdlist"),",")
	tmp_query=""
	For i=lbound(pcArr) to ubound(pcArr)
		if trim(pcArr(i))<>"" then
			if tmp_query<>"" then
				tmp_query=tmp_query & ","
			end if
			tmp_query=tmp_query & trim(pcArr(i))
		end if
	Next
	if tmp_query<>"" then
		tmp_query=" AND (idproduct IN (" & tmp_query & "))"
	end if
	query=query & tmp_query
end if
'***** End of Create SQL Query *****

if session("cp_revImport_pagecurrent")<>"" then
	Set rs=Server.CreateObject("ADODB.Recordset")

	iPageSize=pcv_intExportSize
	
	rs.CacheSize=iPageSize
	rs.PageSize=iPageSize
	
	rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

	rs.AbsolutePage=session("cp_revImport_pagecurrent")

else
	Set rs=Server.CreateObject("ADODB.Recordset")

	iPageSize=pcv_intExportSize
	
	rs.CacheSize=iPageSize
	rs.PageSize=iPageSize
	
	rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

	rs.AbsolutePage=1
	
end if

maxRecords=100000

if not rs.eof then
	if session("cp_revImport_pagecurrent")<>"" then
		pcArr=rs.GetRows(iPageSize)
	else
		pcArr=rs.GetRows(maxRecords)
	end if
	intCount=ubound(pcArr,2)
	pcv_HaveRecords=1
end if
set rs=nothing

IF pcv_HaveRecords=1 THEN
	
	SearchFieldCount=0
	if request("CSearchFields")="1" then
		tmpPrd=""
		For i=0 to intCount
			if tmpPrd<>"" then
				tmpPrd=tmpPrd & ","
			end if
			tmpPrd=tmpPrd & pcArr(40,i)
		Next
		tmpPrd="(" & tmpPrd & ")"
		query="SELECT DISTINCT pcSearchData.idSearchField FROM pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData WHERE pcSearchFields_Products.idProduct IN " & tmpPrd & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			tmpArr=rsQ.getRows()
			SearchFieldCount=ubound(tmpArr,2)+1
			tmpQ=""
			For i=0 to SearchFieldCount-1
				if tmpQ<>"" then
					tmpQ=tmpQ & ","
				end if
				tmpQ=tmpQ & tmpArr(0,i)
			Next
			if tmpQ<>"" then
				tmpQ="(" & tmpQ & ")"
			end if
		end if
		set rsQ=nothing
		if SearchFieldCount>0 then
			query="SELECT DISTINCT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName FROM pcSearchFields WHERE pcSearchFields.idSearchField IN " & tmpQ & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpArr=rsQ.getRows()
				For i=0 to ubound(tmpArr,2)
					HeaderBuilderObj.append """" & tmpArr(1,i) & """" & fseparator
				Next
				ReDim valueArr(ubound(tmpArr,2))
			end if
			set rsQ=nothing
		end if
	end if
	For i=0 to intCount

'***** Generate Data Lines *****
if request("C1")="1" then
	DataBuilderObj.append """" & pcArr(0,i) & """" & fseparator
end if
if request("C2")="1" then
	tmp_PrdName=pcArr(1,i)
	if tmp_PrdName<>"" then
		tmp_PrdName=replace(tmp_PrdName,"""","""""")
	end if
	DataBuilderObj.append """" & tmp_PrdName & """" & fseparator
end if
if request("C3")="1" then
	tmp_PrdDesc=pcArr(59,i)
	if tmp_PrdDesc<>"" then
		if fseparator="," then
			tmp_PrdDesc=RmvHTMLWhiteSpace(replace(tmp_PrdDesc,"""",""""""))
			tmp_PrdDesc=replace(tmp_PrdDesc,"&quot;","""""")
		else
			tmp_PrdDesc=RmvHTMLWhiteSpace(replace(tmp_PrdDesc,"&quot;",""""))
		end if
		tmp_PrdDesc=replace(tmp_PrdDesc,vbCrLf,"")
		tmp_PrdDesc=replace(tmp_PrdDesc,vbCr,"")
		tmp_PrdDesc=replace(tmp_PrdDesc,vbLf,"")
	end if
	DataBuilderObj.append """" & tmp_PrdDesc & """" & fseparator
end if
if request("C4")="1" then
	tmp_PrdSDesc=pcArr(60,i)
	if tmp_PrdSDesc<>"" then
		if fseparator="," then
			tmp_PrdSDesc=RmvHTMLWhiteSpace(replace(tmp_PrdSDesc,"""",""""""))
			tmp_PrdSDesc=replace(tmp_PrdSDesc,"&quot;","""""")
		else
			tmp_PrdSDesc=RmvHTMLWhiteSpace(replace(tmp_PrdSDesc,"&quot;",""""))
		end if
		tmp_PrdSDesc=replace(tmp_PrdSDesc,vbCrLf,"")
		tmp_PrdSDesc=replace(tmp_PrdSDesc,vbCr,"")
		tmp_PrdSDesc=replace(tmp_PrdSDesc,vbLf,"")
	end if
	DataBuilderObj.append """" & tmp_PrdSDesc & """" & fseparator
end if
if request("C5")="1" then
	if pcArr(2,i)<>0 then
		DataBuilderObj.append """BTO""" & fseparator
	else
		if pcArr(3,i)<>0 then
			DataBuilderObj.append """ITEM""" & fseparator
		else
			DataBuilderObj.append """STANDARD""" & fseparator
		end if
	end if
end if
if request("C6")="1" then
	DataBuilderObj.append """" & pcArr(4,i) & """" & fseparator
end if
if request("C7")="1" then
	DataBuilderObj.append """" & pcArr(5,i) & """" & fseparator
end if
if request("C8")="1" then
	DataBuilderObj.append """" & pcArr(6,i) & """" & fseparator
end if
if request("C9")="1" then
	DataBuilderObj.append """" & pcArr(7,i) & """" & fseparator
end if
if request("C10")="1" then
	DataBuilderObj.append """" & pcArr(8,i) & """" & fseparator
end if
if request("C11")="1" then
	DataBuilderObj.append getCATInfor(pcArr(40,i))
end if
if request("C12")="1" then
	tmp_IDBrand=pcArr(9,i)
	if IsNull(tmp_IDBrand) or tmp_IDBrand="" then
		tmp_IDBrand="0"
	end if
	DataBuilderObj.append getBrandInfor(tmp_IDBrand)
end if
if request("C13")="1" then
	DataBuilderObj.append """" & pcArr(10,i) & """" & fseparator
end if
if request("C14")="1" then
	DataBuilderObj.append """" & pcArr(11,i) & """" & fseparator
end if
if request("C15")="1" then
	DataBuilderObj.append """" & pcArr(12,i) & """" & fseparator
end if
if request("C16")="1" then
	DataBuilderObj.append """" & clng(pcArr(13,i)) & """" & fseparator
end if
if request("C17")="1" then
	DataBuilderObj.append """" & clng(pcArr(14,i)) & """" & fseparator
end if
if request("C18")="1" then
	DataBuilderObj.append """" & clng(pcArr(15,i)) & """" & fseparator
end if
if request("C46")="1" then
	DataBuilderObj.append """" & clng(pcArr(41,i)) & """" & fseparator
end if
if request("C19")="1" then
	DataBuilderObj.append getPrdOptions(pcArr(40,i))
end if
if request("C20")="1" then
	DataBuilderObj.append """" & pcArr(16,i) & """" & fseparator
end if
if request("C21")="1" then
	DataBuilderObj.append """" & clng(pcArr(17,i)) & """" & fseparator
end if
if request("C22")="1" then
	DataBuilderObj.append """" & clng(pcArr(18,i)) & """" & fseparator
end if
if request("C23")="1" then
	DataBuilderObj.append """" & clng(pcArr(19,i)) & """" & fseparator
end if
if request("C24")="1" then
	tmp_NFSmessage=pcArr(20,i)
	if tmp_NFSmessage<>"" then
		tmp_NFSmessage=replace(tmp_NFSmessage,"""","""""")
	end if
	DataBuilderObj.append """" & tmp_NFSmessage & """" & fseparator
end if
if request("C25")="1" then
	DataBuilderObj.append """" & clng(pcArr(21,i)) & """" & fseparator
end if
if request("C26")="1" then
	DataBuilderObj.append """" & clng(pcArr(22,i)) & """" & fseparator
end if
if request("C27")="1" then
	DataBuilderObj.append """" & pcArr(23,i) & """" & fseparator
end if
if request("C28")="1" then
	DataBuilderObj.append """" & pcArr(24,i) & """" & fseparator
end if
if request("C29")="1" then
	tmp_OverSize=pcArr(25,i)
	if tmp_OverSize<>"" then
		tmp_OverSize=replace(tmp_OverSize,"||","x")
	end if
	DataBuilderObj.append """" & tmp_OverSize & """" & fseparator
end if
if request("C30")="1" then
	DataBuilderObj.append """" & pcArr(26,i) & """" & fseparator
end if
if request("C31")="1" then
	DataBuilderObj.append """" & pcArr(27,i) & """" & fseparator
end if
if request("C32")="1" then
	DataBuilderObj.append """" & pcArr(28,i) & """" & fseparator
end if
if request("C33")="1" then
	DataBuilderObj.append """" & pcArr(29,i) & """" & fseparator
end if
if request("C34")="1" then
	DataBuilderObj.append """" & pcArr(30,i) & """" & fseparator
end if
if request("C35")="1" then
	DataBuilderObj.append """" & pcArr(31,i) & """" & fseparator
end if
if request("C36")="1" then
	DataBuilderObj.append """" & pcArr(32,i) & """" & fseparator
end if
if request("C37")="1" then
	DataBuilderObj.append """" & pcArr(33,i) & """" & fseparator
end if
if request("C38")="1" then
	DataBuilderObj.append getIsSupplier(pcArr(40,i))
end if
if request("C39")="1" then
	tmp_MtTitle=pcArr(56,i)
	if tmp_MtTitle<>"" then
		tmp_MtTitle=replace(tmp_MtTitle,"""","""""")
	end if
	DataBuilderObj.append """" & tmp_MtTitle & """" & fseparator
	tmp_MtDesc=pcArr(57,i)
	if tmp_MtDesc<>"" then
		tmp_MtDesc=replace(tmp_MtDesc,"""","""""")
	end if
	DataBuilderObj.append """" & tmp_MtDesc & """" & fseparator
	tmp_MtKey=pcArr(58,i)
	if tmp_MtKey<>"" then
		tmp_MtKey=replace(tmp_MtKey,"""","""""")
	end if
	DataBuilderObj.append """" & tmp_MtKey & """" & fseparator
end if
if request("C40")="1" then
	tmp_Downloadable=pcArr(34,i)
	if IsNull(tmp_Downloadable) or tmp_Downloadable="" then
		tmp_Downloadable="0"
	end if
	DataBuilderObj.append getDownloadInfor(pcArr(40,i),tmp_Downloadable)
end if
if request("C41")="1" then
	tmp_GiftCert=pcArr(35,i)
	if IsNull(tmp_GiftCert) or tmp_GiftCert="" then
		tmp_GiftCert="0"
	end if
	DataBuilderObj.append getGCInfor(pcArr(40,i),tmp_GiftCert)
end if
if request("C42")="1" then
	DataBuilderObj.append """" & pcArr(36,i) & """" & fseparator
end if
if request("C43")="1" then
	DataBuilderObj.append """" & pcArr(37,i) & """" & fseparator
end if
if request("C44")="1" then
	DataBuilderObj.append """" & pcArr(38,i) & """" & fseparator
end if
if request("C45")="1" then
	DataBuilderObj.append """" & pcArr(39,i) & """" & fseparator
end if
if request("C49")="1" then
	DataBuilderObj.append """" & pcArr(42,i) & """" & fseparator
end if
if request("C50")="1" then
	DataBuilderObj.append """" & pcArr(43,i) & """" & fseparator
end if
if request("C51")="1" then
	DataBuilderObj.append """" & pcArr(44,i) & """" & fseparator
end if
if request("C52")="1" then
	tmp_PrdNotes=pcArr(55,i)
	if tmp_PrdNotes<>"" then
		if fseparator="," then
			tmp_PrdNotes=RmvHTMLWhiteSpace(replace(tmp_PrdNotes,"""",""""""))
			tmp_PrdNotes=replace(tmp_PrdNotes,"&quot;","""""")
		else
			tmp_PrdNotes=RmvHTMLWhiteSpace(replace(tmp_PrdNotes,"&quot;",""""))
		end if
		tmp_PrdNotes=replace(tmp_PrdNotes,vbCrLf,"")
		tmp_PrdNotes=replace(tmp_PrdNotes,vbCr,"")
		tmp_PrdNotes=replace(tmp_PrdNotes,vbLf,"")
	end if
	DataBuilderObj.append """" & tmp_PrdNotes & """" & fseparator
end if
if request("C53")="1" then
	DataBuilderObj.append """" & pcArr(45,i) & """" & fseparator
end if
if request("C54")="1" then
	DataBuilderObj.append """" & pcArr(54,i) & """" & fseparator
end if
if request("C55")="1" then
	DataBuilderObj.append """" & pcArr(46,i) & """" & fseparator
end if
if request("C56")="1" then
	DataBuilderObj.append """" & pcArr(47,i) & """" & fseparator
end if
if request("C57")="1" then
	DataBuilderObj.append """" & pcArr(48,i) & """" & fseparator & """" & pcArr(49,i) & """" & fseparator & """" & pcArr(50,i) & """" & fseparator & """" & pcArr(51,i) & """" & fseparator & """" & pcArr(52,i) & """" & fseparator & """" & pcArr(53,i) & """" & fseparator
end if
if request("CSearchFields")="1" AND SearchFieldCount>0 then
	For k=0 to SearchFieldCount-1
		valueArr(k)=""
	Next
	query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pcArr(40,i) & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		tmpValue=rsQ.getRows()
		set rsQ=nothing
		intCount1=ubound(tmpValue,2)
		For k=0 to intCount1
			For m=0 to SearchFieldCount-1
				if clng(tmpValue(0,k))=clng(tmpArr(0,m)) then
					valueArr(m)=tmpValue(3,k)
					exit for
				end if
			Next
		Next
	end if
	set rsQ=nothing
	For k=0 to SearchFieldCount-1
		DataBuilderObj.append """" & valueArr(k) & """" & fseparator
	Next
end if

'***** End of Generate Data Lines *****	

	DataBuilderObj.append VBCrLf
	
	Next
	
	strFile=GenFileName()
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	if session("cp_revImport_fseparator")="0" then
		tmpext=".csv"
	else
		tmpext=".txt"
	end if
	Set A=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & tmpext,True)
	A.Write(HeaderBuilderObj.toString & VBCrLf & DataBuilderObj.toString)
	A.Close
	Set A=Nothing
	Set fs=Nothing	
	
END IF 'Have product records

set DataBuilderObj = nothing
set HeaderBuilderObj = nothing
call closedb()

%>
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcSpacer"></td>
</tr>
<%IF pcv_HaveRecords=0 THEN%>
<tr>
	<td colspan="2">
		<div class="pcCPmessage">
			No Products found!
		</div>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<td colspan="2">
		<input type="button" name="back" value="Start Again" onclick="javasccript:location='ReverseImport_step1.asp';" class="ibtnGrey">
	</td>
</tr>
<%ELSE%>
<tr>
	<td colspan="2">
		<div class="pcCPmessageSuccess">
			Products were exported successfully!
		</div>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td align="center">
		<p><b>Download your file.</b></p>
		<p style="padding-top:10px;"><a href="<%=strFile & tmpext%>"><img src="images/DownLoad.gif"></a></p>
		<p style="padding-top:10px;">To ensure that your file downloads correctly, right click on the icon above and choose &quot;<b>Save Target As...</b>&quot;. If the browser attempts to save the file with a *.htm extension, change the file name so that it uses the extension *.txt.</p>
		<p style="padding-top:10px;">Also note that to open a *.txt file in MS Excel you should first start Excel, and then open the file by selecting &quot;File > Open&quot;. This way your will use the <u>MS Excel Text Import Wizard</u>, where you can specify the custom separator used in the file, if any.
		</p>
	</td>
</tr>
<%END IF%>
<% 
session("cp_revImport_prdlist")=""
session("cp_revImport_pagecurrent")=""
session("cp_revImport_cseparator")=""
session("cp_ExportSize")=""
%>
</table>
<!--#include file="AdminFooter.asp"-->