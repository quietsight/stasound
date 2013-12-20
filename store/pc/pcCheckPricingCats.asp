<%iAddDefaultWPrice1=0
iAddDefaultPrice1=0
pBtoBPrice1=0
pPrice1=0

Public Function CheckPrdPrices(tmpIDBTO,tmpIDPrd,tmpPrdPrice,tmpPrdWPrice,custType)
Dim intCC_BTO_Pricing,query,rsCCObj,pcCC_BTO_Price
Dim prdPrice,prdBtoBPrice

	prdPrice=tmpPrdPrice
	prdBtoBPrice=tmpPrdWPrice
	if prdBtoBPrice=0 then
		prdBtoBPrice=prdPrice
	end if

	intCC_BTO_Pricing=0
	if session("customercategory")<>0 then
		query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & tmpIDPrd & " AND idBTOProduct=" & tmpIDBTO & ";"
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
			query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & tmpIDPrd & ";"
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
	if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
		prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("ATBPercentage"))/100),2))
	end if
		
	
	'customer logged in as ATB customer based on wholesale price
	if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
		prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("ATBPercentage"))/100),2))
		prdPrice=Cdbl(prdBtoBPrice)
	end if
		
	'customer logged in as a wholesale customer
	if prdBtoBPrice>0 and custType=1 and session("ATBCustomer")<>1 then
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
	if session("customercategory")<>0 then
		query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idProduct=" & tmpIDBTO & ";"
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
	if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
		prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("ATBPercentage"))/100),2))
	end if
	
	'customer logged in as ATB customer based on wholesale price
	if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
		prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("ATBPercentage"))/100),2))
		prdPrice=Cdbl(prdBtoBPrice)
	end if
	
	'customer logged in as a wholesale customer
	if prdBtoBPrice>0 and custType=1 and session("ATBCustomer")<>1 then
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


Public Function GetDefaultPrice(pIdProduct,tempVarCat)
Dim query,rsTempObj,dblprice,dblWprice,intCC_BTO_Pricing,rsCCObj,pcCC_BTO_Price,tmpItemID

	'***** GET DEFAULT PRICE OF THE CAT
	query="SELECT configSpec_products.configProduct,configSpec_products.price, configSpec_products.Wprice, configSpec_products.cdefault FROM configSpec_products WHERE configSpec_products.configProductCategory="&tempVarCat&" AND configSpec_products.specProduct="&pIdProduct&" AND configSpec_products.cdefault<>0;"
	set rsTempObj=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsTempObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	pcv_strRunCCSection=0
	If NOT rsTempObj.eof then
		tmpItemID=rsTempObj("configProduct")
		dblprice=Cdbl(rsTempObj("price"))
		dblWprice=Cdbl(rsTempObj("Wprice"))
		pcv_strRunCCSection=1
	end if
	Set rsTempObj=nothing
		
	
	If pcv_strRunCCSection=1 Then	
		if dblWprice=0 then
			dblWprice=dblprice
		end if
		
		intCC_BTO_Pricing=0
		if session("customercategory")<>0 then
			query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & tmpItemID & " AND idBTOProduct=" & pIdProduct & ";" 
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
			end if
			set rsCCObj=nothing
			
			if NOT intCC_BTO_Pricing=1 then
				query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & tmpItemID & ";"
				set rsCCObj=server.CreateObject("ADODB.RecordSet")
				set rsCCObj=conntemp.execute(query)
				if NOT rsCCObj.eof then
					intCC_BTO_Pricing=1
					pcCC_BTO_Price=rsCCObj("pcCC_Price")
				end if
				set rsCCObj=nothing
			end if
			
		end if
																		
		'customer logged in as ATB customer based on retail price
		if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
			dblprice=Cdbl(dblprice)-(pcf_Round(Cdbl(dblprice)*(cdbl(session("ATBPercentage"))/100),2))
		end if

		defaultPrice= Cdbl(dblprice)
		
		'customer logged in as ATB customer based on wholesale price
		if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
			dblWprice=Cdbl(dblWprice)-(pcf_Round(Cdbl(dblWprice)*(cdbl(session("ATBPercentage"))/100),2))
			defaultPrice=Cdbl(dblWprice)
		end if
		
		'customer logged in as a wholesale customer
		if dblWprice>0 and session("customerType")=1 then
			defaultPrice=Cdbl(dblWprice)
		end if

		'customer logged in as a customer type with price different then the online price
		if intCC_BTO_Pricing=1 then
			'if pcCC_BTO_Price>0 then
				defaultPrice=Cdbl(pcCC_BTO_Price)
			'end if
		end if
		
		
	End If
	'***** END OF GET DEFAULT PRICE OF THE CAT
	GetDefaultPrice=defaultPrice

End Function


Public Function GetCDefaultPrice(pIdProduct,tempVarCat)
Dim query,rsTempObj,dblprice,dblWprice,intCC_BTO_Pricing,rsCCObj,pcCC_BTO_Price,tmpItemID

	'***** GET DEFAULT PRICE OF THE CAT
	query="SELECT configSpec_Charges.configProduct,configSpec_Charges.price, configSpec_Charges.Wprice, configSpec_Charges.cdefault FROM configSpec_Charges WHERE configSpec_Charges.configProductCategory="&tempVarCat&" AND configSpec_Charges.specProduct="&pIdProduct&" AND configSpec_Charges.multiSelect=0 AND configSpec_Charges.cdefault=-1;"
	set rsTempObj=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsTempObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	pcv_strRunCCSection=0
	If NOT rsTempObj.eof then
		tmpItemID=rsTempObj("configProduct")
		dblprice=Cdbl(rsTempObj("price"))
		dblWprice=Cdbl(rsTempObj("Wprice"))
		pcv_strRunCCSection=1
	End If
	Set rsTempObj = nothing
	
	If pcv_strRunCCSection=1 Then	
		if dblWprice=0 then
			dblWprice=dblprice
		end if
		
		intCC_BTO_Pricing=0
		if session("customercategory")<>0 then
			query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & tmpItemID & " AND idBTOProduct=" & pIdProduct & ";" 
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
			end if
			set rsCCObj=nothing
			
			If NOT intCC_BTO_Pricing=1 Then
				query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & tmpItemID & ";"
				set rsCCObj=server.CreateObject("ADODB.RecordSet")
				set rsCCObj=conntemp.execute(query)
				if NOT rsCCObj.eof then
					intCC_BTO_Pricing=1
					pcCC_BTO_Price=rsCCObj("pcCC_Price")
				end if
				set rsCCObj=nothing
			End If
			
		end if
																		
		'customer logged in as ATB customer based on retail price
		if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
			dblprice=Cdbl(dblprice)-(pcf_Round(Cdbl(dblprice)*(cdbl(session("ATBPercentage"))/100),2))
		end if

		defaultPrice= Cdbl(dblprice)
		
		'customer logged in as ATB customer based on wholesale price
		if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
			dblWprice=Cdbl(dblWprice)-(pcf_Round(Cdbl(dblWprice)*(cdbl(session("ATBPercentage"))/100),2))
			defaultPrice=Cdbl(dblWprice)
		end if
		
		'customer logged in as a wholesale customer
		if dblWprice>0 and session("customerType")=1 then
			defaultPrice=Cdbl(dblWprice)
		end if

		'customer logged in as a customer type with price different then the online price
		if intCC_BTO_Pricing=1 then
			'if pcCC_BTO_Price>0 then
				defaultPrice=Cdbl(pcCC_BTO_Price)
			'end if
		end if
		
	End If
	'***** END OF GET DEFAULT PRICE OF THE CAT
	GetCDefaultPrice=defaultPrice

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_GetBTOConfigPrices(pCnt,dtype)
Dim query,rs
	query="SELECT categories.categoryDesc, products.description, products.iRewardPoints,configSpec_products.configProductCategory, configSpec_products.price, configSpec_products.Wprice, categories_products.idCategory, categories_products.idProduct, products.weight FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	iAddDefaultPrice=Cdbl(0)
	iAddDefaultWPrice=Cdbl(0)
	iAddDefaultPrice1=Cdbl(0)
	iAddDefaultWPrice1=Cdbl(0)
		
	if NOT rs.eof then 
		Dim FirstCnt
		FirstCnt=0
		do until rs.eof
			FirstCnt=FirstCnt+1
			strCategoryDesc=rs("categoryDesc")
			strDescription=rs("description")
			intReward=rs("iRewardPoints")
			if (intReward<>"") and (intReward<>"0") then
			else
			intReward=0
			end if
			pcv_BTORP=pcv_BTORP+clng(intReward)
			strConfigProductCategory=rs("configProductCategory")
			dblPrice=rs("price")
			dblWPrice=rs("Wprice")
			intIdCategory=rs("idCategory")
			intIdProduct=rs("idProduct")
			intWeight=rs("weight")
			
			dblPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,0)
			dblWPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,1)
			
			ItemPrice=0
			if Session("CustomerType")=1 then
				if (dblWPrice<>0) then
					ItemPrice=dblWPrice1
				else
					ItemPrice=dblPrice1
				end if
			else
				ItemPrice=dblPrice1
			end if
			
			if dtype = "m" then
				response.write "<input name="""&pCnt&"||CAT"&FirstCnt&""" type=""HIDDEN"" value=""CAG"&intIdCategory&""">"
				response.write "<input type=""hidden"" name="""&pCnt&"||CAG"&intIdCategory&""" value="""&intIdProduct&"_0_"&intWeight&"_" & ItemPrice & """>"
			end if
			rs.moveNext
		loop
		if dtype = "m" then
			response.write "<input type=""hidden"" name="""&pCnt&"||FirstCnt"" value="""&FirstCnt&""">"
		end if		
	end if
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function ChkMultiSelectDef(pIdProduct,tempVarCat,itemID)
Dim query,rs
	
	query="SELECT configSpec_products.configProduct FROM configSpec_products WHERE configSpec_products.configProduct=" & itemID & " AND configSpec_products.configProductCategory="&tempVarCat&" AND configSpec_products.specProduct="&pIdProduct&" AND configSpec_products.multiSelect<>0;"
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	If rs.eof then
		ChkMultiSelectDef=2
		set rs=nothing
		exit function
	end if
	Set rs=nothing

	query="SELECT configSpec_products.configProduct FROM configSpec_products WHERE configSpec_products.configProduct=" & itemID & " AND configSpec_products.configProductCategory="&tempVarCat&" AND configSpec_products.specProduct="&pIdProduct&" AND configSpec_products.cdefault<>0 AND configSpec_products.multiSelect<>0;"
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	If not rs.eof then
		ChkMultiSelectDef=0
	else
		ChkMultiSelectDef=1
	end if
	Set rs=nothing

End Function

Public Function GetDefaultQty(pIdProduct,tempVarCat)
Dim query,rs,tmp1

	query="SELECT products.pcprod_minimumqty FROM products INNER JOIN configSpec_products ON products.idproduct=configSpec_products.configProduct WHERE configSpec_products.configProductCategory="&tempVarCat&" AND configSpec_products.specProduct="&pIdProduct&" AND configSpec_products.cdefault<>0;"
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	tmp1=1
	if not rs.eof then
		tmp1=rs("pcprod_minimumqty")
		if IsNull(tmp1) or tmp1="" then
			tmp1=1
		end if
		if tmp1="0" then
			tmp1=1
		end if
	end if
	set rs=nothing
	
	GetDefaultQty=tmp1

End Function

Public Function GetItemDefaultQty(pIdProduct)
Dim query,rs,tmp1

	query="SELECT products.pcprod_minimumqty FROM products WHERE products.idproduct="&pIdProduct&";"
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	tmp1=1
	if not rs.eof then
		tmp1=rs("pcprod_minimumqty")
		if IsNull(tmp1) or tmp1="" then
			tmp1=1
		end if
		if tmp1="0" then
			tmp1=1
		end if
	end if
	set rs=nothing
	
	GetItemDefaultQty=tmp1

End Function


Public Function NotForSaleOverride(pcvIntCustCategory)
Dim query,rsNFSO,tmpNFSO,tmpIntCustCategory

	tmpIntCustCategory=pcvIntCustCategory
	if validNum(tmpIntCustCategory) and tmpIntCustCategory>0 then
		query="SELECT pcCC_NFSoverride FROM pcCustomerCategories WHERE idCustomerCategory=" & tmpIntCustCategory
		set rsNFSO=Server.CreateObject("ADODB.Recordset")
		set rsNFSO=conntemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsNFSO=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		if rsNFSO.eof then
			tmpNFSO=0
		else
			tmpNFSO=rsNFSO("pcCC_NFSoverride")
			if IsNull(tmpNFSO) or tmpNFSO="" or tmpNFSO="0" then
				tmpNFSO=0
			else
				tmpNFSO=1
			end if
		end if
		set rsNFSO=nothing
	else
		tmpNFSO=0
	end if
	
	NotForSaleOverride=tmpNFSO

End Function
%>