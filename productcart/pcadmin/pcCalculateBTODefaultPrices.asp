<%
Dim iAddDefaultWPrice1, iAddDefaultPrice1, pBtoBPrice1, pPrice1
Dim iAddDefaultWPrice, iAddDefaultPrice, pBtoBPrice, pPrice
Dim dblpcCC_Price,pidProduct,pnoprices,pDescription

Public Function CheckPrdPrices(tmpIDBTO,tmpIDPrd,tmpPrdPrice,tmpPrdWPrice,custType)
Dim intCC_BTO_Pricing,query,rsCCObj,pcCC_BTO_Price
Dim prdPrice,prdBtoBPrice

	prdPrice=tmpPrdPrice
	prdBtoBPrice=tmpPrdWPrice
	if prdBtoBPrice=0 then
		prdBtoBPrice=prdPrice
	end if

	intCC_BTO_Pricing=0
	if session("admin_tmp_customercategory")<>0 then
		query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("admin_tmp_customercategory") & " AND idBTOItem=" & tmpIDPrd & " AND idBTOProduct=" & tmpIDBTO & ";"
		set rsCCObj=server.CreateObject("ADODB.RecordSet")
		set rsCCObj=conntemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rsCCObj=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		if NOT rsCCObj.eof then
			intCC_BTO_Pricing=1
			pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
		else
			query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("admin_tmp_customercategory") &" AND pcCC_Pricing.idProduct=" & tmpIDPrd & ";"
			set rsCCObj=server.CreateObject("ADODB.RecordSet")
			set rsCCObj=conntemp.execute(query)
			if NOT rsCCObj.eof then
				intCC_BTO_Pricing=1
				pcCC_BTO_Price=rsCCObj("pcCC_Price")
			end if
		end if
		set rsCCObj=nothing
	end if

	'customer logged in as ATB customer based on retail price
	if session("admin_tmp_ATBCustomer")=1 AND session("admin_tmp_ATBPercentOff")=0 then
		prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("admin_tmp_ATBPercentage"))/100),2))
	end if


	'customer logged in as ATB customer based on wholesale price
	if session("admin_tmp_ATBCustomer")=1 AND session("admin_tmp_ATBPercentOff")=1 then
		prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("admin_tmp_ATBPercentage"))/100),2))
		prdPrice=Cdbl(prdBtoBPrice)
	end if

	'customer logged in as a wholesale customer
	if prdBtoBPrice>0 and custType=1 and session("admin_tmp_ATBCustomer")<>1 then
		prdPrice=Cdbl(prdBtoBPrice)
	end if

	'customer logged in as a customer type with price different then the online price
	if intCC_BTO_Pricing=1 then
		'if cdbl(pcCC_BTO_Price)>0 then
			prdPrice=Cdbl(pcCC_BTO_Price)
		'end if
	end if

	CheckPrdPrices=prdPrice

End Function

Public Function CheckParentPrices(tmpIDBTO,tmpPrdPrice,tmpPrdWPrice,custType)
Dim intCC_BTO_Pricing,query,rsCCObj,pcCC_BTO_Price
Dim prdPrice,prdBtoBPrice

	prdPrice=tmpPrdPrice
	prdBtoBPrice=tmpPrdWPrice
	if prdBtoBPrice=0 then
		prdBtoBPrice=prdPrice
	end if

	intCC_BTO_Pricing=0
	if session("admin_tmp_customercategory")<>0 then
		query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory=" & session("admin_tmp_customercategory") & " AND idProduct=" & tmpIDBTO & ";"
		set rsCCObj=server.CreateObject("ADODB.RecordSet")
		set rsCCObj=conntemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rsCCObj=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		if NOT rsCCObj.eof then
			intCC_BTO_Pricing=1
			pcCC_BTO_Price=rsCCObj("pcCC_Price")
			pcCC_BTO_Price=pcf_Round(pcCC_BTO_Price, 2)
		end if
		set rsCCObj=nothing
	end if

	'customer logged in as ATB customer based on retail price
	if session("admin_tmp_ATBCustomer")=1 AND session("admin_tmp_ATBPercentOff")=0 then
		prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("admin_tmp_ATBPercentage"))/100),2))
	end if

	'customer logged in as ATB customer based on wholesale price
	if session("admin_tmp_ATBCustomer")=1 AND session("admin_tmp_ATBPercentOff")=1 then
		prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("admin_tmp_ATBPercentage"))/100),2))
		prdPrice=Cdbl(prdBtoBPrice)
	end if

	'customer logged in as a wholesale customer
	if prdBtoBPrice>0 and custType=1 and session("admin_tmp_ATBCustomer")<>1 then
		prdPrice=Cdbl(prdBtoBPrice)
	end if

	'customer logged in as a customer type with price different then the online price
	if intCC_BTO_Pricing=1 then
		'if cdbl(pcCC_BTO_Price)>0 then
			prdPrice=Cdbl(pcCC_BTO_Price)
		'end if
	end if

	CheckParentPrices=prdPrice

End Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_GetBTOConfigPrices()
Dim query,rsBTOConfig
	query="SELECT categories.categoryDesc, products.description, products.iRewardPoints,configSpec_products.configProductCategory, configSpec_products.price, configSpec_products.Wprice, categories_products.idCategory, categories_products.idProduct, products.weight, products.pcprod_minimumqty FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
	set rsBTOConfig=server.CreateObject("ADODB.RecordSet")
	set rsBTOConfig=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsBTOConfig=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	iAddDefaultPrice=Cdbl(0)
	iAddDefaultWPrice=Cdbl(0)
	iAddDefaultPrice1=Cdbl(0)
	iAddDefaultWPrice1=Cdbl(0)

	if NOT rsBTOConfig.eof then
		pcArray_BTOConfig=rsBTOConfig.getRows()
		intBTOConfigCount=ubound(pcArray_BTOConfig,2)
	end if
	set rsBTOConfig=nothing
		Dim FirstCnt
		FirstCnt=0
	If intBTOConfigCount>0 Then

		For a=0 to intBTOConfigCount
			FirstCnt=FirstCnt+1
			strCategoryDesc=pcArray_BTOConfig(0,a) '// rsBTOConfig("categoryDesc")
			strDescription=pcArray_BTOConfig(1,a) '// rsBTOConfig("description")
			intReward=pcArray_BTOConfig(2,a) '// rsBTOConfig("iRewardPoints")
			if (intReward<>"") and (intReward<>"0") then
			else
			intReward=0
			end if
			pcv_BTORP=pcv_BTORP+clng(intReward)
			strConfigProductCategory=pcArray_BTOConfig(3,a) '// rsBTOConfig("configProductCategory")
			dblPrice=pcArray_BTOConfig(4,a) '// rsBTOConfig("price")
			dblWPrice=pcArray_BTOConfig(5,a) '// rsBTOConfig("Wprice")
			if dblWPrice=0 then
				dblWPrice=dblPrice
			end if
			intIdCategory=pcArray_BTOConfig(6,a) '// rsBTOConfig("idCategory")
			intIdProduct=pcArray_BTOConfig(7,a) '// rsBTOConfig("idProduct")
			intWeight=pcArray_BTOConfig(8,a) '// rsBTOConfig("weight")
			pcv_qty=pcArray_BTOConfig(9,a)
			if IsNull(pcv_qty) or pcv_qty="" then
				pcv_qty=0
			end if
			if clng(pcv_qty)=0 then
				pcv_qty=1
			end if

			dblPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,0)
			dblWPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,1)
			iAddDefaultPrice=Cdbl(iAddDefaultPrice+dblPrice*pcv_qty)
			iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+dblWPrice*pcv_qty)
			iAddDefaultPrice1=Cdbl(iAddDefaultPrice1+dblPrice1*pcv_qty)
			iAddDefaultWPrice1=Cdbl(iAddDefaultWPrice1+dblWPrice1*pcv_qty)

			ItemPrice=0
			if session("admin_tmp_CustomerType")=1 then
				if (dblWPrice<>0) then
					ItemPrice=dblWPrice1
				else
					ItemPrice=dblPrice1
				end if
			else
				ItemPrice=dblPrice1
			end if

		Next
	End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Get BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub CalBTODefaultPriceCat()
	if pBtoBPrice=0 then
		pBtoBPrice=pPrice
	end if

	dblpcCC_Price=0
	pPrice1=pPrice
	pBtoBPrice1=pBtoBPrice

	If pnoprices<2 Then

		pPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
		pBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)

		call pcs_GetBTOConfigPrices()

		pPrice=Cdbl(pPrice+iAddDefaultPrice)
		pBtoBPrice=Cdbl(pBtoBPrice+iAddDefaultWPrice)
		pPrice1=Cdbl(pPrice1+iAddDefaultPrice1)
		pBtoBPrice1=Cdbl(pBtoBPrice1+iAddDefaultWPrice1)

		if session("admin_tmp_customertype")=1 and pBtoBPrice1>0 then
			dblpcCC_Price=pBtoBPrice1
		else
			dblpcCC_Price=pPrice1
		end if
	End if

End Sub

Public Sub RunCalBDPC()
Dim rs,query
' get item details from db
query="SELECT description,price,bToBPrice,NoPrices FROM products WHERE products.idProduct=" & pidProduct & " AND serviceSpec<>0;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

IF NOT rs.eof THEN 'If it is BTO Product

	' charge rscordset data into local variables
	pDescription = rs("description")
	pPrice = rs("price")
	pWPrice = rs("bToBPrice")
	if pWPrice=0 then
		pWPrice=pPrice
	end if
	pBtoBPrice=pWPrice
	pnoprices = rs("NoPrices")
	set rs=nothing

	save_pPrice=pPrice
	save_pBtoBPrice=pBtoBPrice

	query="SELECT configSpec_products.price,configSpec_products.Wprice,configSpec_products.cdefault,products.pcprod_minimumqty FROM (configSpec_products INNER JOIN products ON configSpec_products.configProduct = products.idProduct) INNER JOIN categories ON configSpec_products.configProductCategory = categories.idCategory WHERE (((configSpec_products.specProduct)=" & pidProduct & ")) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort,products.description;"
	Set rs=conntemp.execute(query)

	if not rs.eof then
		If pnoprices<2 Then
			iAddDefaultPrice=Cdbl(0)
			iAddDefaultWPrice=Cdbl(0)
			pcArr=rs.getRows()
			set rs=nothing
			intCount=ubound(pcArr,2)
			For i=0 to intCount
				dblprice=pcArr(0,i)
				dblWprice=pcArr(1,i)
				pcv_noDefault=pcArr(2,i)
				pcv_qty=pcArr(3,i)
				if IsNull(pcv_qty) or pcv_qty="" then
					pcv_qty=0
				end if
				if clng(pcv_qty)=0 then
					pcv_qty=1
				end if
				if pcv_noDefault<>0 then
					if dblWprice=0 then
						dblWprice=dblprice
					end if
					iAddDefaultPrice=Cdbl(iAddDefaultPrice+dblprice*pcv_qty)
					iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+dblWprice*pcv_qty)
				end if
			Next
			pPrice=Cdbl(pPrice+iAddDefaultPrice)
			pWPrice=Cdbl(pWPrice+iAddDefaultWPrice)
		else
			pPrice=0
			pWPrice=0
		end if
	else
		If pnoprices<2 Then
		else
			pPrice=0
			pWPrice=0
		end if
	end if
	set rs=nothing

	query="UPDATE Products SET pcProd_BTODefaultPrice=" & pPrice & ",pcProd_BTODefaultWPrice=" & pWPrice & " WHERE idProduct=" & pidProduct & ";"
	Set rs=conntemp.execute(query)
	set rs=nothing

	query="SELECT idcustomerCategory, pcCC_WholesalePriv, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories;"
	Set rs=conntemp.execute(query)
	if not rs.eof then
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
			session("admin_tmp_customerCategory")=pcArr(0,i)
			session("admin_tmp_customertype")=pcArr(1,i)
			session("admin_tmp_customerCategoryType")=pcArr(2,i)
			if session("admin_tmp_customerCategoryType")="ATB" then
				session("admin_tmp_ATBCustomer")=1
				session("admin_tmp_ATBPercentage")=pcArr(3,i)
				intpcCC_ATB_Off=pcArr(4,i)
				if intpcCC_ATB_Off="Retail" then
					session("admin_tmp_ATBPercentOff")=0
				else
					session("admin_tmp_ATBPercentOff")=1
				end if
			else
				session("admin_tmp_ATBCustomer")=0
				session("admin_tmp_ATBPercentage")=0
				session("admin_tmp_ATBPercentOff")=0
			end if

			pPrice=save_pPrice
			pBtoBPrice=save_pBtoBPrice

			Call CalBTODefaultPriceCat()

			query="DELETE FROM pcBTODefaultPriceCats WHERE idproduct=" & pidProduct & " AND idcustomerCategory=" & session("admin_tmp_customerCategory") & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			query="INSERT INTO pcBTODefaultPriceCats (idproduct,idcustomerCategory,pcBDPC_Price) VALUES (" & pidProduct  & "," & session("admin_tmp_customerCategory") & "," & dblpcCC_Price & ");"
			set rs=connTemp.execute(query)
			set rs=nothing
		Next
	end if
	set rs=nothing
END IF 'If it is BTO Product
set rs=nothing

End Sub
%>