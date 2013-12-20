<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "repeatorder.asp"
' This page generates repeat orders
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="chkPrices.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<%
Response.Buffer = True


Dim query, conntemp, rstemp, pIdOrder 

err.number=0

' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function

pIdOrder=getUserInput(request("idOrder"),0)
pIdOrder1=pIdOrder

Dim ItemsDiscounts
ItemsDiscounts=0
Dim Charges
Charges=0
Dim pDefaultPrice
pDefaultPrice=0

pcCartArray=session("pcCartSession")
ppcCartIndex = Session("pcCartIndex")

call openDb()

query="SELECT ProductsOrdered.idProductOrdered,ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.pcSubscription_ID, ProductsOrdered.unitPrice, ProductsOrdered.xfdetails, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, products.btoBPrice, products.price  "
'BTO ADDON-S
If scBTO=1 then
	query=query&",ProductsOrdered.idconfigSession"
End If
'BTO ADDON-E
query=query&",products.description,products.sku, products.weight,products.pcprod_QtyToPound,products.emailText, products.deliveringTime, products.pcSupplier_ID, products.cost, products.stock, products.notax, products.noshipping, products.iRewardPoints, products.pcProd_Surcharge1, products.pcProd_Surcharge2 FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idCustomer=" &Session("idcustomer")& " AND orders.idOrder=" &pIdOrder

set rstemp=conntemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rstemp=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	call closeDb() 
 	response.redirect "msg.asp?message=35"     
end if

if request("OrderRepeat")<>"haveto" then
pcv_OrdHaveOutStock=0
pcv_OrdHaveStock=0

do while not rstemp.eof
	pidProductOrdered=rstemp("idProductOrdered")
	pidProduct=rstemp("idProduct")
	pquantity=rstemp("quantity")
	pSubscriptionID=rstemp("pcSubscription_ID")
	
	'Check if this customer is logged in with a customer category
	if session("customerCategory")<>0 then
		query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pidProduct&";"
		set rsCCObj=server.CreateObject("ADODB.RecordSet")
		set rsCCObj=conntemp.execute(query)
		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rsCCObj=nothing
			'//close any connections
			call closedb()
			'//redirect to error page
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		if NOT rsCCObj.eof then
			strcustomerCategory="YES"
			dblpcCC_Price=rsCCObj("pcCC_Price")
			dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
		else
			strcustomerCategory="NO"
		end if
		set rsCCObj=nothing
	end if
	
	tmp2=",categories,categories_products"
	tmp3=" AND categories_products.idproduct=products.idproduct AND categories.idcategory=categories_products.idcategory AND categories.iBTOhide=0 "
	tmp3=tmp3 & " AND products.idproduct IN (SELECT DISTINCT categories_products.idproduct FROM categories,categories_products WHERE categories_products.idproduct=" & pidProduct & " AND categories.idcategory=categories_products.idcategory AND (categories.iBTOhide=0"
		
	if session("idCustomer")<>0 AND session("customerType")=1 then
		tmp3=tmp3 & "))"
	else
		tmp3=tmp3 & " OR categories.pccats_RetailHide<>0)) AND categories.pccats_RetailHide=0"
	end if

	'// START v4.1 - Not For Sale override
		if NotForSaleOverride(session("customerCategory"))=1 then
			queryNFSO=""
		else
			queryNFSO=" AND products.formQuantity=0"
		end if
	'// END v4.1
	
	query="SELECT DISTINCT products.serviceSpec,products.stock,products.noStock,products.pcprod_minimumqty,products.pcprod_qtyvalidate,products.pcProd_BackOrder FROM Products" & tmp2 & " WHERE products.idproduct=" & pidProduct & " AND products.removed=0 AND products.active<>0" & queryNFSO & tmp3 & ";"
	set rs=connTemp.execute(query)
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
	
	IF not rs.eof THEN
		pserviceSpec=rs("serviceSpec")
		pStock=rs("stock")
		pNoStock=rs("noStock")
		pcv_minqty=rs("pcprod_minimumqty")
		pcv_qtyvalid=rs("pcprod_qtyvalidate")
		pcv_BackOrder=rs("pcProd_BackOrder")
	
		if (PStock<pcv_minqty) and (pcv_qtyvalid=0) then
			pStock=0
		end if
	
		if pcv_qtyvalid="1" then
			if PStock<pcv_minqty then
				pStock=0
			else
				if (PStock<pquantity) and (pStock>pcv_minqty) then
					pStock=Fix(pStock/pcv_minqty)*pcv_minqty
				end if
			end if
		end if
	
		IF (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_BackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_BackOrder=0) THEN
			pcv_OrdHaveOutStock=1
		ELSE
			pcv_OrdHaveStock=1
		END IF
	ELSE
		pcv_OrdHaveOutStock=1
	END IF
	rstemp.MoveNext
loop

if (pcv_OrdHaveOutStock=1) and (pcv_OrdHaveStock=0) then
	call closedb()
	response.redirect "msg.asp?message=132"&"&IdOrder=" & pIdOrder1
end if

if (pcv_OrdHaveOutStock=1) and (pcv_OrdHaveStock=1) then
	call closedb()
	response.redirect "msg.asp?message=133"&"&IdOrder=" & pIdOrder1
end if
rstemp.MoveFirst

end if


do while not rstemp.eof
	pidProductOrdered=rstemp("idProductOrdered")
	pidProduct=rstemp("idProduct")
	pquantity=rstemp("quantity")
	
	tmp2=",categories,categories_products"
	tmp3=" AND categories_products.idproduct=products.idproduct AND categories.idcategory=categories_products.idcategory AND categories.iBTOhide=0 "
	tmp3=tmp3 & " AND products.idproduct IN (SELECT DISTINCT categories_products.idproduct FROM categories,categories_products WHERE categories_products.idproduct=" & pidProduct & " AND categories.idcategory=categories_products.idcategory AND (categories.iBTOhide=0"
		
	if session("idCustomer")<>0 AND session("customerType")=1 then
		tmp3=tmp3 & "))"
	else
		tmp3=tmp3 & " OR categories.pccats_RetailHide<>0)) AND categories.pccats_RetailHide=0"
	end if
	
	'// START v4.1 - Not For Sale override
		if NotForSaleOverride(session("customerCategory"))=1 then
			queryNFSO=""
		else
			queryNFSO=" AND products.formQuantity=0"
		end if
	'// END v4.1
	
	query="SELECT DISTINCT Products.serviceSpec,Products.stock,Products.noStock,Products.pcprod_minimumqty,Products.pcprod_qtyvalidate,Products.pcProd_BackOrder FROM Products" & tmp2 & " WHERE Products.idproduct=" & pidProduct & " AND Products.removed=0 AND Products.active<>0" & queryNFSO & tmp3 & ";"
	set rs=connTemp.execute(query)
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
	
IF not rs.eof THEN	
	
	pserviceSpec=rs("serviceSpec")
	pStock=rs("stock")
	pNoStock=rs("noStock")
	pcv_minqty=rs("pcprod_minimumqty")
	pcv_qtyvalid=rs("pcprod_qtyvalidate")
	pcv_BackOrder=rs("pcProd_BackOrder")
	
	if (PStock<pcv_minqty) and (pcv_qtyvalid=0) then
	pStock=0
	end if
	
	if pcv_qtyvalid="1" then
		if PStock<pcv_minqty then
			pStock=0
		else
			if (PStock<pquantity) and (pStock>pcv_minqty) then
				pStock=Fix(pStock/pcv_minqty)*pcv_minqty
			end if
		end if
	end if
	
IF (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_BackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_BackOrder=0) THEN
ELSE
	if pStock=0 then
		if pcv_minqty>"0" then
			PStock=pcv_minqty
		else
			pStock=1
		end if
	end if

	IF (scOutofStockPurchase=-1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_BackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND pNoStock=0 AND pcv_BackOrder=0) THEN	
		if pStock<pquantity then
			pquantity=pStock
		end if
	END IF
	
	'// Product Options Arrays
	pcv_strSelectedOptions = rstemp("pcPrdOrd_SelectedOptions") ' Column 11
	'pcv_strOptionsPriceArray = rstemp("pcPrdOrd_OptionsPriceArray") ' Column 25
	'pcv_strOptionsArray = rstemp("pcPrdOrd_OptionsArray") ' Column 4 
	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get the Options for the item
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
	pcv_strSelectedOptions = ""
end if

pcv_strOptionsArray = ""
pcv_strOptionsPriceArray = ""
pcv_strOptionsPriceArrayCur = ""
pcv_strOptionsPriceTotal = 0
xOptionsArrayCount = 0
pPriceToAdd = 0


if len(pcv_strSelectedOptions)>0 then

	pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))
	
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
			pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & money(pPriceToAdd)
			'// Column 5) This is the total of all option prices
			pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd
			
		end if
		
		set rs=nothing
	Next

end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Get the Options for the item
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'response.write pcv_strOptionsArray
					
	WPrice=rstemp("btoBPrice")
	
	if WPrice<>"" then
	else
		WPrice="0"
	end if
	
	if (session("CustomerType")="1") and (WPrice<>"0") then
		punitPrice=WPrice
	else
		punitPrice=rstemp("Price")
	end if
	If strcustomerCategory="YES" then
		punitPrice=dblpcCC_Price
	end if

	pDefaultPrice=punitPrice

	pxfdetails=rstemp("xfdetails")
	'BTO ADDON-S
	if scBTO=1 then
		pidconfigSession=rstemp("idconfigSession")
		if pidconfigSession="0" then
			pidconfigSession=""
		end if
	End If
	'BTO ADDON-E

	pDescription = rstemp("description")
	pWeight	= rstemp("weight")
	pcv_QtyToPound=rstemp("pcprod_QtyToPound")
	if pcv_QtyToPound>0 then
		pWeight=(16/pcv_QtyToPound)
		if scShipFromWeightUnit="KGS" then
			pWeight=(1000/pcv_QtyToPound)
		end if
	end if
	pSku = rstemp("sku")
	pEmailText = rstemp("emailText")
	pDeliveringTime	= rstemp("deliveringTime")
	pIdSupplier	= rstemp("pcSupplier_ID")
	pCost = rstemp("cost")
	pStock = rstemp("stock")
	pnotax = rstemp("notax")
	pnoshipping = rstemp("noshipping")
	iRewardPoints=rstemp("iRewardPoints")
	pcv_Surcharge1 = rstemp("pcProd_Surcharge1")
	pcv_Surcharge2 = rstemp("pcProd_Surcharge2")

	pIdConfigSession=trim(pidconfigSession)


	'// Get Prices To Add
	
	



	'-- Create new Product Config Session
		
	if pIdConfigSession<>"" then
		query="select * from configSessions where IdConfigSession=" & pIdConfigSession
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
		
		IF not rs.eof THEN
			PreRecord=""
			PreRecord1=""
			pConfigKey=trim(randomNumber(9999)&randomNumber(9999))
			
			iCols = rs.Fields.Count
			for dd=1 to iCols-1
				if dd=1 then
		    	PreRecord=PreRecord & Rs.Fields.Item(dd).Name
		    else
		    	PreRecord=PreRecord & "," & Rs.Fields.Item(dd).Name
		    end if
		    IF Ucase(Rs.Fields.Item(dd).Name)="CONFIGKEY" then
					if dd=1 then
						PreRecord1=PreRecord1 & "'" & pConfigKey & "'"
					else
						PreRecord1=PreRecord1 & ",'" & pConfigKey & "'"
					end if
				ELSE
					IF Ucase(Rs.Fields.Item(dd).Name)="DTCREATED" then
						if scDB="Access" then
							myStr11="#"
						else
							myStr11="'"
						end if
						dtTodaysDate=Date()
						if SQL_Format="1" then
							dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
						else
							dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
						end if
						if dd=1 then
							PreRecord1=PreRecord1 & myStr11 & dtTodaysDate & myStr11
						else
							PreRecord1=PreRecord1 & "," & myStr11 & dtTodaysDate & myStr11
						end if
					ELSE
						IF Ucase(Rs.Fields.Item(dd).Name)="STRINGOPTIONS" then
							if dd=1 then
								PreRecord1=PreRecord1 & "' '"
							else
								PreRecord1=PreRecord1 & ",' '"
							end if
						ELSE
							FType="" & Rs.Fields.Item(dd).Type
							if (Ftype="202") or (Ftype="203") then
								PTemp=Rs.Fields.Item(dd).Value
								if PTemp<>"" then
									PTemp=replace(PTemp,"'","''")
								end if
								if dd=1 then
									PreRecord1=PreRecord1 & "'" & PTemp & "'"
								else
									PreRecord1=PreRecord1 & ",'" & PTemp & "'"
								end if
							else
								PTemp="" & Rs.Fields.Item(dd).Value
								if PTemp<>"" then
								else
									PTemp="0"
								end if
								if dd=1 then
									PreRecord1=PreRecord1 & PTemp
								else
									PreRecord1=PreRecord1 & "," & PTemp
								end if
							end if
						END IF 'stringOptions
					END IF 'DTCreated
				END IF 'Config Key
			next
			
			query="insert into configSessions (" & PreRecord & ") values (" & PreRecord1 & ")"
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
		
			query="select idConfigSession from configSessions order by idConfigSession desc"
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
			
			pIdConfigSession=rs("IdConfigSession")
			if pIdConfigSession<>"0" then
				punitPrice=updPrices(pidProduct,pIdConfigSession)
			end if
		END IF
	end if
	'-- END Create new Product Config Session
	'---------------------		

	'---------------------------
	'Calculate BTO Items Weights
	IF pIdConfigSession<>"" then
		query="SELECT stringProducts,stringCProducts,stringQuantity FROM configSessions WHERE idconfigSession=" & pIdConfigSession
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
		
		stringProducts=rs("stringProducts")
		stringCProducts=rs("stringCProducts")
		ArrProduct=Split(stringProducts, ",")
		ArrCProduct=Split(stringCProducts, ",")
		Qstring=rs("stringQuantity")
		ArrQuantity=Split(Qstring,",")
		
		CWeight=0

		if ArrProduct(0)<>"na" then
			for j=lbound(ArrProduct) to (UBound(ArrProduct)-1)
				query="SELECT weight,pcprod_QtyToPound FROM products WHERE IDProduct=" & ArrProduct(j)
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
		
				if not rs.eof then
					ItemWeight=rs("weight")
					pcv_QtyToPound=rs("pcprod_QtyToPound")
					if pcv_QtyToPound>0 then
						ItemWeight=cdbl(16/pcv_QtyToPound)
						if scShipFromWeightUnit="KGS" then
							ItemWeight=cdbl(1000/pcv_QtyToPound)
						end if
					end if
		
					CWeight=CWeight+cdbl(ItemWeight*clng(ArrQuantity(j)))
				end if
			next
		end if

		if ArrCProduct(0)<>"na" then
			for j=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
				query="SELECT weight,pcprod_QtyToPound FROM products WHERE IDProduct=" & ArrCProduct(j)
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
				
				if not rs.eof then
		
					ItemWeight=rs("weight")
					
					pcv_QtyToPound=rs("pcprod_QtyToPound")
					if pcv_QtyToPound>0 then
						ItemWeight=(16/pcv_QtyToPound)
						if scShipFromWeightUnit="KGS" then
							ItemWeight=(1000/pcv_QtyToPound)
						end if
					end if
		
					CWeight=CWeight+cdbl(ItemWeight)
		
				end if
			next
		end if	
	
	END IF 'Have ID Config Session
	'--- END Calculate BTO Items Weights	
	'---------------------------
	
	ppcCartIndex = ppcCartIndex + 1
	session("pcCartIndex")	= ppcCartIndex
	
	pcCartArray(ppcCartIndex,0) = pIdProduct 
	pcCartArray(ppcCartIndex,1) = pDescription
	pcCartArray(ppcCartIndex,2) = pQuantity
	
	' add price or BtoB price depending on customer type
	pcCartArray(ppcCartIndex,3) = punitPrice
	

	pcCartArray(ppcCartIndex,8)="" '// not in use anymore
	
	
	if len(pcv_strOptionsArray)>0 then
		 pcCartArray(ppcCartIndex,4)= pcv_strOptionsArray '// store array of product "option groups: options"
	else
		pcCartArray(ppcCartIndex,4)=""
	end if	

	pcCartArray(ppcCartIndex,5) = pcv_strOptionsPriceTotal '// Total Cost of all Options
	pcCartArray(ppcCartIndex,23)= pOverSizeSpec
	pcCartArray(ppcCartIndex,25)= pcv_strOptionsPriceArray '// Array of Individual Options Prices
	pcCartArray(ppcCartIndex,26)= pcv_strOptionsPriceArrayCur '// Array of Options Prices 
	'pcCartArray(ppcCartIndex,27)="" '// Not in use anymore - VERIFY FOR Crosssell
	'pcCartArray(ppcCartIndex,28)="" '// Not in use anymore - VERIFY FOR Crosssell	 	 	 	 
	pcCartArray(ppcCartIndex,6) = pWeight + Cweight
	pcCartArray(ppcCartIndex,7) = pSku
	pcCartArray(ppcCartIndex,9) = pDeliveringTime  
	
	' deleted mark
	pcCartArray(ppcCartIndex,10) = 0 
	
	pcCartArray(ppcCartIndex,11) = pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers
	pcCartArray(ppcCartIndex,12) = "" '// Not in use anymore
	pcCartArray(ppcCartIndex,13) = pIdSupplier
	pcCartArray(ppcCartIndex,14) = pCost
	pcCartArray(ppcCartIndex,16) = pIdConfigSession
	pcCartArray(ppcCartIndex,19) = pnotax
	pcCartArray(ppcCartIndex,20) = pnoshipping
	pcCartArray(ppcCartIndex,21)=  replace(pxfdetails,"|","<br>")
	pcCartArray(ppcCartIndex,36) = pcv_Surcharge1
	pcCartArray(ppcCartIndex,37) = pcv_Surcharge2

	'SB S
	pcCartArray(ppcCartIndex,38)= pSubscriptionID '// Subscription ID  
	'SB E

	pTotalQuantity = pQuantity
	
	'-----------------------------
	'ReCalculate BTO Items Discounts

	pcCartArray(ppcCartIndex,30)=cdbl(ItemsDiscounts)
	
	'End ReCulculate BTO Items Discounts		
	'------------------------------------
	
	'------------------------------------
	'BTO Additional Charges

	pcCartArray(ppcCartIndex,31) = Cdbl(Charges) 
	
	'End BTO Additional Charges
	'------------------------------------
	' get discount per quantity
	query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
	set rstemp1=conntemp.execute(query)
	if err.number<>0 and err.number<>9 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rstemp1=nothing
		'//close any connections
		call closedb()
		'//redirect to error page
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if Session("customerType")=1 OR session("customerCategory")<>0 then
		pcCartArray(f,18)=1
	else
		pcCartArray(f,18)=0
	end if

	pOrigPrice = pcCartArray(ppcCartIndex,3)
	pcCartArray(ppcCartIndex,17) = pOrigPrice
	if pcQDiscountType<>"1" then
	pOrigPrice=pOrigPrice+pcCartArray(ppcCartIndex,5)+(pcCartArray(ppcCartIndex,30)/pTotalQuantity)
	else
	pOrigPrice=pDefaultPrice
	end if

	pcCartArray(ppcCartIndex,15) = 0
	if not rstemp1.eof and err.number<>9 then
		' there are quantity discounts defined for that quantity 
		pDiscountPerUnit = rstemp1("discountPerUnit")
		pDiscountPerWUnit = rstemp1("discountPerWUnit")
		pPercentage = rstemp1("percentage")

		if session("customerType")<>1 then
			if pPercentage = "0" then 
				pcCartArray(ppcCartIndex,3)  = pcCartArray(ppcCartIndex,3) - pDiscountPerUnit
				pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (pDiscountPerUnit * pTotalQuantity)
			else
				pcCartArray(ppcCartIndex,3) = pcCartArray(ppcCartIndex,3) - ((pDiscountPerUnit/100) * pOrigPrice)
				pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + ((pDiscountPerUnit/100) * (pOrigPrice * pTotalQuantity))
			end if
		else
			if pPercentage = "0" then 
				pcCartArray(ppcCartIndex,3)  = pcCartArray(ppcCartIndex,3) - pDiscountPerWUnit
				pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (pDiscountPerWUnit * pTotalQuantity)
			else
				pcCartArray(ppcCartIndex,3) = pcCartArray(ppcCartIndex,3) - ((pDiscountPerWUnit/100) * pOrigPrice)
				pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + ((pDiscountPerWUnit/100) * (pOrigPrice * pTotalQuantity))
			end if
		end if
	end if

	'// Start Reward Points
	If RewardsActive = 1 Then
		pcCartArray(ppcCartIndex,22) = Clng(iRewardPoints)
	End If
	'// End Reward Points
	
	'SM-S
	if UCase(scDB)="SQL" then
		query="SELECT Products.pcSC_ID,pcSales_BackUp.pcSales_TargetPrice FROM pcSales_BackUp INNER JOIN Products ON pcSales_BackUp.pcSC_ID=Products.pcSC_ID WHERE Products.idProduct=" & pIdProduct
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			tmpSCID=rsQ("pcSC_ID")
			If IsNull(tmpSCID) then
				tmpSCID=0
			End If
			tmpTarget=rsQ("pcSales_TargetPrice")
			if IsNull(tmpTarget) then
				tmpTarget=0
			end if
			if ((clng(tmpTarget)=0) AND (session("customerCategory")=0) AND (session("customerType")<>"1")) OR ((clng(tmpTarget)=-1) AND (session("customerCategory")=0) AND (session("customerType")="1")) OR ((clng(tmpTarget)=clng(session("customerCategory"))) AND (clng(tmpTarget)>0)) then
				pcCartArray(ppcCartIndex,39)=tmpSCID '//Sale ID
			else
				pcCartArray(ppcCartIndex,39)=0
			end if
		else
			pcCartArray(ppcCartIndex,39)=0
		end if
		set rsQ=nothing
	else
		pcCartArray(ppcCartIndex,39)=0
	end if
	'SM-E
	
END IF ' ends the rs // this could be moved up to line 203
END IF ' end the rstemp
rstemp.movenext  
loop

%>
<!--#include file="inc-UpdPrdQtyDiscounts.asp"-->
<%
pcCartArray(1,18)=0
%>
<!--#include file="pcReCalPricesLogin.asp"-->
<%

session("pcCartSession") = pcCartArray

call closedb()

response.redirect "viewcart.asp"
%>