<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include file="../pc/pcCheckPricingCats.asp"-->
<% 
Dim pQuantity, pxfield1, pxfield2, pxfield3, pxf1, pxf2, pxf3, pCheckIdXfield1, pCheckIdXfield2, pCheckIdXfield3
Dim xOptionGroupCount, pcv_intOptionGroupCount, pcv_strSelectedOptions, pcvstrTmpOptionGroup
Dim query, noOS, rstemp, xfieldsCnt
Dim pIdProduct, ppcCartIndex, pTotalQuantity, pOptionDescrip, pPriceToAdd, pOptionGroupDesc
Dim pPrice, pDescription, pWeight, pSku, pBtoBPrice, pDeliveringTime, pIdSupplier, pCost, pStock, pnotax, pnoshipping
Dim pOverSizeSpec, iRewardPoints, iRewardDollars, pNoStock, pcv_QtyToPound
Dim lineNumber,iNextIndex, iCapturedNext, pcArray_SelectedOptions, cCounter, ppcParentIndex
Dim conntemp
Dim pcv_strOptionsArray, xOptionsArrayCount, pcv_strOptionsPriceArray, pcv_strOptionsPriceTotal, pcv_strOptionsPriceArrayCur
Dim pcv_strProductsQuantity, pCSCount, pIsAccessory,cv_strCSRequired,pcArray_CSRequired,pcv_ParentDiscount,pcv_ChildDiscount
dim pcv_strSelectedCSProducts,pcArray_SelectedCSProducts,pcv_strCSDiscounts,pcArray_CSDiscounts,pcv_strPrdDiscounts,pcArray_PrdDiscounts
Dim pcv_ChildBundleID

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: On Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'on error resume next
pIdProduct=getUserInput(request("idproduct"),10)
pIdOrder=getUserInput(request("idorder"),20)
err.number=0
pTotalQuantity=Cint(0)
call opendb()

	query="SELECT customers.customerType,customers.idCustomerCategory FROM customers INNER JOIN orders ON customers.idcustomer=orders.idcustomer WHERE idorder=" & pIdOrder & ";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if not rs.eof then
		session("customerType")=rs("customerType")
		idcustomerCategory=rs("idcustomerCategory")
		if IsNull(idcustomerCategory) or idcustomerCategory="" then
			idcustomerCategory=0
		end if
	end if
	set rs=nothing

	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory="&idcustomerCategory&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then
		session("customerCategory")=rs("idcustomerCategory")
		strpcCC_Name=rs("pcCC_Name")
		session("customerCategoryDesc")=strpcCC_Name
		strpcCC_Description=rs("pcCC_Description")
		session("customerCategoryType")=rs("pcCC_CategoryType")
		if session("customerCategoryType")="ATB" then
			session("ATBCustomer")=1
			session("ATBPercentage")=rs("pcCC_ATB_Percentage")
			intpcCC_ATB_Off=rs("pcCC_ATB_Off")
			if intpcCC_ATB_Off="Retail" then
				session("ATBPercentOff")=0
			else
				session("ATBPercentOff")=1
			end if
		else
			session("ATBCustomer")=0
			session("ATBPercentage")=0
			session("ATBPercentOff")=0
		end if
	end if
	set rs=nothing

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: On Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Get Form Data
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pIdProduct=getUserInput(request("idproduct"),10)
pQuantity=getUserInput(request.Form("quantity"),10)
pxfield1=getUserInput(request.Form("xfield1"),0)
pxfield2=getUserInput(request.Form("xfield2"),0)
pxfield3=getUserInput(request.Form("xfield3"),0)
pxf1=getUserInput(request.Form("xf1"),10)
pxf2=getUserInput(request.Form("xf2"),10)
pxf3=getUserInput(request.Form("xf3"),10)
'--> New Product Options
pcv_intOptionGroupCount = getUserInput(request.Form("OptionGroupCount"),0)
if IsNull(pcv_intOptionGroupCount) OR pcv_intOptionGroupCount="" then
	pcv_intOptionGroupCount = 0
end if
pcv_intOptionGroupCount = cint(pcv_intOptionGroupCount)

xOptionGroupCount = 0
pcv_strSelectedOptions = ""
do until xOptionGroupCount = pcv_intOptionGroupCount	
	xOptionGroupCount = xOptionGroupCount + 1
	pcvstrTmpOptionGroup = request.Form("idOption"&xOptionGroupCount)
	if pcvstrTmpOptionGroup <> "" then			
		pcv_strSelectedOptions = pcv_strSelectedOptions & pcvstrTmpOptionGroup & chr(124)	
	end if	
loop
' trim the last pipe if there is one
xStringLength = len(pcv_strSelectedOptions)
if xStringLength>0 then
	pcv_strSelectedOptions = left(pcv_strSelectedOptions,(xStringLength-1))
end if
' if cannot get quantity get quantity 1 (from listing)
if pQuantity="" then
 pQuantity=1
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Get Form Data
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get the Options for the item
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))

pcv_strOptionsArray = ""
pcv_strOptionsPriceArray = ""
pcv_strOptionsPriceArrayCur = ""
pcv_strOptionsPriceTotal = 0
xOptionsArrayCount = 0
pPriceToAdd = 0

For cCounter = LBound(pcArray_SelectedOptions) TO UBound(pcArray_SelectedOptions)
	
	' SELECT DATA SET
	' TABLES: optionsGroups, options, options_optionsGroups
	query = 		"SELECT optionsGroups.optionGroupDesc, options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice "
	query = query & "FROM optionsGroups, options, options_optionsGroups "
	query = query & "WHERE idoptoptgrp=" & pcArray_SelectedOptions(cCounter) & " "
	query = query & "AND options_optionsGroups.idOption=options.idoption "
	query = query & "AND options_optionsGroups.idOptionGroup=optionsGroups.idoptiongroup "	
	
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	
	if rs.eof then 
		set rs=nothing
		call closeDb()
		response.redirect "msg.asp?message=42"   	  
	else
		
		xOptionsArrayCount = xOptionsArrayCount + 1
		
		pOptionDescrip=""
		pOptionGroupDesc=""
		pPriceToAdd=""
		pOptionDescrip=rs("optiondescrip")
		pOptionGroupDesc=rs("optionGroupDesc")
		
		If Session("customerType")=1 Then
			pPriceToAdd=rs("Wprice")
			If rs("Wprice")=0 then
				pPriceToAdd=rs("price")
			End If
		Else
			pPriceToAdd=rs("price")
		End If	
		
		'// Generate Our Strings
		if xOptionsArrayCount > 1 then
			pcv_strOptionsArray = pcv_strOptionsArray & chr(124)
			pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & chr(124)
			pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & chr(124)
		end if
		'// Column 4) This is the Array of Product "option groups: options"
		pcv_strOptionsArray = pcv_strOptionsArray & pOptionGroupDesc & ": " & pOptionDescrip
		'// Column 25) This is the Array of Individual Options Prices
		pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & pPriceToAdd
		'// Column 26) This is the Array of Individual Options Prices, but stored as currency "scCurSign & money(pcv_strOptionsPriceTotal) "
		pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & formatnumber(pPriceToAdd, 2)
		'// Column 5) This is the total of all option prices
		pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd
		
	end if
	
	set rs=nothing
Next
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Get the Options for the item
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'response.write pcv_strOptionsArray


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check if the product has optionals assigned
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT xfield1, xfield2, xfield3 FROM products WHERE idProduct=" &pIdProduct

set rstemp=conntemp.execute(query)

if err.number <> 0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pCheckIdXfield1=rstemp("xfield1")
pCheckIdXfield2=rstemp("xfield2")
pCheckIdXfield3=rstemp("xfield3")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check if the product has optionals assigned
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Get the Product Details
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
err.clear
noOS=0
query="SELECT OverSizeSpec FROM products"
set rstemp=conntemp.execute(query)
if err.number<>0 then
	noOS=1
end if
err.clear
query="SELECT description, price, bToBPrice, sku, weight, deliveringTime, idSupplier, cost, stock, notax, noshipping"
if noOS=0 then
	query=query&", OverSizeSpec"
end if
query=query&" FROM products WHERE idproduct=" &pIdProduct& " AND active=-1"

set rstemp=conntemp.execute(query)

if err.number <> 0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then 
	set rstemp = nothing
	call closeDb()
 	response.redirect "msg.asp?message=41"
end if

pPrice=rstemp("price")
pDescription=rstemp("description")
pWeight=rstemp("weight")
pSku=rstemp("sku")
pBtoBPrice=rstemp("bToBPrice")
pDeliveringTime=rstemp("deliveringTime")
pIdSupplier=rstemp("idSupplier")
pCost=rstemp("cost")
pStock=rstemp("stock")
pnotax=rstemp("notax")
pnoshipping=rstemp("noshipping")
if noOS=0 then
	pOverSizeSpec=rstemp("OverSizeSpec")
else
	pOverSizeSpec="NO"
end if

pPrice1=pPrice
pBtoBPrice1=pBtoBPrice

pPrice=CheckParentPrices(pidProduct,pPrice1,pBtoBPrice1,0)
pBtoBPrice=CheckParentPrices(pidProduct,pPrice1,pBtoBPrice1,1)

if session("customerType")=1 then
	pPrice=pBtoBPrice
end if

pOptionDescripA=Cstr("")
pPriceToAddA=Cint(0)
pOptionDescripB=Cstr("")
pPriceToAddB=Cint(0)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Get the Product Details
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: GET X Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
xfieldsCnt=0
if pxfield1<>"" then
	xfieldsCnt=xfieldsCnt+1
	query="SELECT xfield FROM xfields WHERE idxfield="&pxf1
        	
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if rstemp.eof then
		set rstemp = nothing
		call closeDb() 
	  	response.redirect "msg.asp?message=44"  	  
	end if
	
	pXfieldDescrip1=rstemp("xfield")
	
end if
if pxfield2<>"" then
	xfieldsCnt=xfieldsCnt+1
	query="SELECT xfield FROM xfields WHERE idxfield="&pxf2
        	
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if rstemp.eof then
		set rstemp = nothing
		call closeDb()
	  	response.redirect "msg.asp?message=45"  	  
	end if
	
	pXfieldDescrip2=rstemp("xfield")
	
end if
if pxfield3<>"" then
	xfieldsCnt=xfieldsCnt+1
	query="SELECT xfield FROM xfields WHERE idxfield="&pxf3
        	
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if rstemp.eof then
		set rstemp = nothing
		call closeDb()
	  	response.redirect "msg.asp?message=46"      	  
	end if
	
	pXfieldDescrip3=rstemp("xfield")
	
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: GET X Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Dim checkSS
' insert new basket line, is not in the cart
pTotalQuantity=pQuantity

pta_IdProduct=pIdProduct 
pta_QTY=pQuantity

' add price or BtoB price depending on customer type
pta_Price=pPrice
	
 

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: X Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 if xfieldsCnt>0 then
	xCnt=0
	pta_xfdetails=""
	if pxfield1<>"" then
		xCnt=1
		pta_xfdetails=pta_xfdetails&pXfieldDescrip1&": "&pxfield1
	end if
	if pxfield2<>"" then
		if xCnt=1 then
			pta_xfdetails=pta_xfdetails& "|"
		end if
		xCnt=1
		pta_xfdetails=pta_xfdetails&pXfieldDescrip2&": "&pxfield2
	end if
	if pxfield3<>"" then
		if xCnt=1 then
			pta_xfdetails=pta_xfdetails& "|"
		end if
		xCnt=1
		pta_xfdetails=pta_xfdetails&pXfieldDescrip3&": "&pxfield3
	end if
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: X Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Get discount per quantity
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
set rstemp=conntemp.execute(query)
if err.number <> 0 and err.number<>9 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

p_Const=0
if not rstemp.eof and err.number<>9 then
 	' there are quantity discounts defined for that quantity 
 	pDiscountPerUnit=rstemp("discountPerUnit")
 	pDiscountPerWUnit=rstemp("discountPerWUnit")
 	pPercentage=rstemp("percentage")
	pbaseproductonly=rstemp("baseproductonly")
	pOrigPrice=pta_Price

	if pPercentage="0" then 
		pta_Price=pta_Price - pDiscountPerUnit  'Price - discount per unit
		p_Const=p_Const + (pDiscountPerUnit * pta_QTY)  'running total of discounts
	else
		if pbaseproductonly="-1" then
			pta_Price=pta_Price - ((pDiscountPerUnit/100) * pOrigPrice)
			p_Const=p_Const + (((pDiscountPerUnit/100) * pOrigPrice) * pta_QTY)
		else
			pta_Price=pta_Price - ((pDiscountPerUnit/100) * (pOrigPrice+pta_PriceToAddOptions))
			p_Const=p_Const + (((pDiscountPerUnit/100) * (pOrigPrice+pta_PriceToAddOptions)) * pta_QTY)
		end if
	end if	
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Get discount per quantity
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'// Unit Pride
tempVar1=(pta_Price + pcv_strOptionsPriceTotal)

'pcPrdOrd_SelectedOptions = pcv_strSelectedOptions
'pcPrdOrd_OptionsPriceArray = pcv_strOptionsPriceArray
'pcPrdOrd_OptionsArray = pcv_strOptionsArray
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Insert Line Item
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query = 		"INSERT INTO ProductsOrdered "
query = query & "(idOrder, idProduct, quantity, unitPrice, unitCost, xfdetails, idconfigSession, QDiscounts, pcPrdOrd_SelectedOptions, pcPrdOrd_OptionsPriceArray, pcPrdOrd_OptionsArray) VALUES "
query = query & "("&pIdOrder&","&pIdProduct&","&pta_QTY&","&tempVar1&",0,'"&replace(pta_xfdetails,"'","''")&"',0," & p_Const & ",'"&replace(pcv_strSelectedOptions,"'","''")&"','"&replace(pcv_strOptionsPriceArray,"'","''")&"','"&replace(pcv_strOptionsArray,"'","''")&"');"
'response.write query
'response.end
set rstemp=conntemp.execute(query)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Insert Line Item
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

set rstemp = nothing
call closeDB()
call clearLanguage()
response.Redirect "AdminEditOrder.asp?ido="&pIdOrder&"&action=upd"
%>