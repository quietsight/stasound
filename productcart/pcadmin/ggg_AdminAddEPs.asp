<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Dim pageTitle, Section
pageTitle="Add Gift Registry Products to Order"
Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../pc/ggg_inc_chkEPPrices.asp"-->
<% err.number=0
'on error resume next
dim query, conntemp, rstemp

call opendb()

pidOrder=getUserInput(request("ido"),0)
gIDEvent=getUserInput(request("IDEvent"),0)

Count=getUserInput(request("Count"),0)

if Count<>"" then
else
	Count="0"
end if

gAdd=0

FOR dd=1 TO Count
	geID=getUserInput(request("geID" & dd),0)
	geadd=getUserInput(request("add" & dd),0) 
	if geadd="" then
		geadd="0"
	end if

	IF (geID<>"") AND (clng(geadd)>0) then
		gAdd=1

		query="SELECT pcEvProducts.pcEP_idProduct,pcEvProducts.pcEP_OptionsArray,pcEvProducts.pcEP_xdetails,pcEvProducts.pcEP_IDConfig "
		query=query&",products.description,products.sku, products.weight,products.pcprod_QtyToPound,products.emailText, products.deliveringTime, products.idSupplier, products.cost, products.stock, products.notax, products.noshipping, products.iRewardPoints FROM products,pcEvProducts WHERE pcEvProducts.pcEP_ID=" & geID & " and pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_idproduct and products.active=-1 and products.removed=0"
		set rstemp=conntemp.execute(query)
		
		IF not rstemp.eof THEN

		pidProduct=rstemp("pcEP_idProduct")
		pquantity=geadd
		pTotalQuantity=geadd

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Product Options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcv_strSelectedOptions=""
		pcv_strSelectedOptions = rstemp("pcEP_OptionsArray")
		pcv_strSelectedOptions=pcv_strSelectedOptions&""
		if pcv_strSelectedOptions="" then
			pcv_strSelectedOptions="NULL"
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Product Options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		pxfdetails=rstemp("pcEP_xdetails")
		pidconfigSession=rstemp("pcEP_IDConfig")
		if pidconfigSession="0" then
			pidconfigSession=""
		end if
		pDescription = rstemp("description")
		pSku = rstemp("sku")
		pWeight	= rstemp("weight")
		pcv_QtyToPound=rstemp("pcprod_QtyToPound")
		if pcv_QtyToPound>0 then
			pWeight=(16/pcv_QtyToPound)
			if scShipFromWeightUnit="KGS" then
				pWeight=(1000/pcv_QtyToPound)
			end if
		end if
		pEmailText = rstemp("emailText")
		pDeliveringTime	= rstemp("deliveringTime")
		pIdSupplier	= rstemp("idSupplier")
		pCost = rstemp("cost")
		pStock = rstemp("stock")
		pnotax = rstemp("notax")
		pnoshipping = rstemp("noshipping")
    iRewardPoints=rstemp("iRewardPoints")
		set rstemp=nothing
		pIdConfigSession=trim(pidconfigSession)
		
	'*************************************************************************************************
	' START: GET OPTIONS
	'*************************************************************************************************
	Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc
	Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
	Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	IF len(pcv_strSelectedOptions)>0 AND pcv_strSelectedOptions<>"NULL" THEN 
	pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))
	
	pcv_strOptionsArray = ""
	pcv_strOptionsPriceArray = ""
	pcv_strOptionsPriceArrayCur = ""
	pcv_strOptionsPriceTotal = 0
	xOptionsArrayCount = 0
	
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
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rs=nothing
			'//close any connections
			call closedb()
			'//redirect to error page
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if					
		
		if Not rs.eof then 
			
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
			pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & money(pPriceToAdd)
			'// Column 5) This is the total of all option prices
			pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd
			
		end if
		
		set rs=nothing
	Next
	END IF	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			
		punitPrice=updPrices(pidProduct,pIdConfigSession,pcv_strSelectedOptions,1)
		punitPrice=punitPrice+pcv_strOptionsPriceTotal
		
		
		
		if 	pIdConfigSession="0" then
		else
			pIdConfigSession=""
		end if		

		'-----------------------------
		'ReCalculate BTO Items Discounts

		itemsDiscounts=0
		if pIdConfigSession<>"" then 
			query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rs=conntemp.execute(query)

			stringProducts=rs("stringProducts")
			stringValues=rs("stringValues")
			stringCategories=rs("stringCategories")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			Qstring=rs("stringQuantity")
			ArrQuantity=Split(Qstring,",")
			Pstring=rs("stringPrice")
			ArrPrice=split(Pstring,",")
			set rs=nothing

			if ArrProduct(0)<>"na" then
				for j=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(j)
					set rsQ=connTemp.execute(query)
					TempDiscount=0
					do while not rsQ.eof
		 				QFrom=rsQ("quantityFrom")
						QTo=rsQ("quantityUntil")
						DUnit=rsQ("discountperUnit")
						QPercent=rsQ("percentage")
						DWUnit=rsQ("discountperWUnit")
						if (DWUnit=0) and (DUnit>0) then
							DWUnit=DUnit
						end if
						

						TempD1=0
						if (clng(ArrQuantity(j)*tQTY)>=clng(QFrom)) and (clng(ArrQuantity(j)*tQTY)<=clng(QTo)) then
							if QPercent="-1" then
								if session("customerType")=1 then
									TempD1=ArrQuantity(j)*tQTY*ArrPrice(j)*0.01*DWUnit
								else
									TempD1=ArrQuantity(j)*tQTY*ArrPrice(j)*0.01*DUnit
								end if
							else
								if session("customerType")=1 then
									TempD1=ArrQuantity(j)*tQTY*DWUnit
								else
									TempD1=ArrQuantity(j)*tQTY*DUnit
								end if
							end if
						end if
						TempDiscount=TempDiscount+TempD1
						rsQ.movenext
					loop
					set rsQ=nothing
					itemsDiscounts=ItemsDiscounts+TempDiscount
				next			
			end if 'Have BTO Items
		end if 'Have ConfigSession

		'End ReCulculate BTO Items Discounts		
		'------------------------------------

		if 	pIdConfigSession<>"" then
		else
			pIdConfigSession="0"
		end if

		' get discount per quantity

		query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
		set rstemp1=conntemp.execute(query)

		pOrigPrice = punitPrice
	
		QDiscounts = 0
		if not rstemp1.eof and err.number<>9 then
		 	' there are quantity discounts defined for that quantity 
		 	pDiscountPerUnit = rstemp1("discountPerUnit")
		 	pDiscountPerWUnit = rstemp1("discountPerWUnit")
		 	pPercentage = rstemp1("percentage")
		 	if session("customerType")<>1 then
		 		if pPercentage = "0" then 
					QDiscounts = QDiscounts + (pDiscountPerUnit * pTotalQuantity)
				else
					QDiscounts = QDiscounts + (((pDiscountPerUnit/100) * pOrigPrice) * pTotalQuantity)
				end if
			else
				if pPercentage = "0" then 
					QDiscounts = QDiscounts + (pDiscountPerWUnit * pTotalQuantity)
				else
					QDiscounts = QDiscounts + (((pDiscountPerWUnit/100) * pOrigPrice)* pTotalQuantity)
				end if
			end if
		end if
		set rstemp1=nothing

		query="select IDProduct from ProductsOrdered where pcPO_EPID=" & geID & " and IdOrder=" & pIDOrder
		set rs1=connTemp.execute(query)

		if not rs1.eof then
			query="update ProductsOrdered SET quantity=quantity+" & pQuantity & " where pcPO_EPID=" & geID
			set rs1=connTemp.execute(query)
			set rs1=nothing
		else
			query="insert into ProductsOrdered (IDOrder,IDProduct,Quantity,pcPrdOrd_SelectedOptions,pcPrdOrd_OptionsPriceArray,pcPrdOrd_OptionsArray,unitPrice,xfdetails,IDConfigSession,QDiscounts,ItemsDiscounts,pcPO_EPID) values (" & pIDOrder & "," & pIDProduct & "," & pQuantity & ",'" & pcv_strSelectedOptions & "','" & pcv_strOptionsPriceArray & "','" & pcv_strOptionsArray & "'," & punitPrice & ",'" & pxfdetails & "'," & pIDConfigSession & "," & QDiscounts & "," & ItemsDiscounts & "," & geID & ")"
			set rs1=connTemp.execute(query)
			set rs1=nothing
		end if

		query="update pcEvProducts set pcEP_HQty=pcEP_HQty+" & pQuantity & " where pcEP_ID=" & geID
		set rs1=connTemp.execute(query)
		set rs1=nothing
		
		END IF 'not rstemp.eof

	END IF 'Have Product (geID<>"")

Next 'From Line 38

call closedb()

response.Redirect "AdminEditOrder.asp?ido="&pIdOrder&"&action=upd"
%>