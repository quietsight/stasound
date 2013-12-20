<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "instPrd.asp"
' This page is handles adding and updating the shoppingcart array
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_checkPrdQtyCart.asp"-->
<!--#include file="inc_checkMinMul.asp"-->
<%
Response.Buffer = True

'GGG Add-on start

'Check Shopping Cart if It is used for a Gift Registry
if Session("Cust_BuyGift")<>"" then
  response.redirect "msg.asp?message=100"      
end if

'GGG Add-on end

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

'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************


'*****************************************************************************************************
' START: PAGE ON LOAD
'*****************************************************************************************************

'// check for multiple products
piCnt=request("pCnt")
if ((len(piCnt)>0) AND (validNum(piCnt))) then
	pcv_strProductsQuantity=piCnt
else
	pcv_strProductsQuantity=1
end if


'// check for cross sell products
pCSCnt=request("pCSCount")
if ((len(pCSCnt)>0) AND (validNum(pCSCnt))) then
	pcv_strProductsQuantity=(pCSCnt+1)
else
	pCSCnt=0
end if


if pcv_intFlagNoLocal="" then
	'// Set the cart array session to Local
	pcCartArray=Session("pcCartSession")
end if

' check for errors
if err.number>0 then
	response.redirect "viewPrd.asp?idproduct="&getUserInput(request("idproduct"),10)
end if

'// Set additional variables
ppcCartIndex=Session("pcCartIndex")
pTotalQuantity=Cint(0)


'// Check for bound quantity in cart
Dim pcv_BoundQty 
pcv_BoundQty = countCartRows(pcCartArray, ppcCartIndex)
if pcv_BoundQty>=scQtyLimit then
  response.redirect "msg.asp?message=39"      
end if

'SB S	
pSubscriptionID = getUserInput(request("pSubscriptionID"),0)
If pSubscriptionID = "" or pSubscriptionID ="0" Then
	pSubscriptionID = 0
End if 	
Dim pcv_sbLockCart, pcv_sbIsCartLockable
pcv_sbLockCart = findSubscription(session("pcCartSession"), Session("pcCartIndex"))		
pcv_sbIsCartLockable = IsCartLockable(session("pcCartSession"), Session("pcCartIndex"))		
'SB E

'Clear empty arrays
for f=1 to ppcCartIndex 
	if pcCartArray(f,1)="" then
		pcCartArray(f,10)=1
	end if
next

'--> open database connection
call opendb()

Function CheckMinQty(tmpID,tmpQty)
	Dim queryQ, rsQ, tmpMin, tmpMin1
	
	if tmpQty=0 then	
		CheckMinQty=0
	else
		tmpMin=tmpQty
		if tmpMin="" then
			tmpMin=0
		end if		
		queryQ="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & tmpID & ";"
		set rsQ=connTemp.execute(queryQ)	
		if not rsQ.eof then
			tmpMin1=rsQ("pcprod_minimumqty")
			if tmpMin1<>"" then
				if clng(tmpMin1)>0 AND clng(tmpMin1)>clng(tmpMin) then
					tmpMin=tmpMin1
				end if
			end if
		end if
		set rsQ=nothing		
				
		CheckMinQty=tmpMin	
	end if	
End Function

'*****************************************************************************************************
' END: PAGE ON LOAD
'*****************************************************************************************************
%>

<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td>
		
		<%
		'*****************************************************************************************************
		' 1) START: get data from viewPrd form
		'*****************************************************************************************************
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  Get the Cross Sell Product Information
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if (pCSCnt>0) then
	
			'// Cross Sell Array Size (must be zero or greater)
			pcv_intCSArraySize = (pCSCnt-1)
			if pcv_intCSArraySize < 1 then
				pcv_intCSArraySize = 0
			end if
			
			'// Clear our data
	    pcv_strSelectedCSProducts = ""
	    pcv_strPrdDiscounts = ""
	    pcv_strCSDiscounts = ""
	    pcv_strCSRequired = ""
	    pcArray_SelectedCSProducts = ""
	    pcArray_PrdDiscounts = ""
	    pcArray_CSDiscounts = ""
	    pcArray_CSRequired = ""
		
		'// Define my arrays
	    pcv_strSelectedCSProducts = getUserInput(request("pCrossSellIDs"),0)  '// All Cross Sell IDs
	    pcArray_SelectedCSProducts = Split(pcv_strSelectedCSProducts,",")
	    pcv_strCSDiscounts = getUserInput(request("pCSDiscounts"),0)  '// All Cross Sell Discounts
	    pcArray_CSDiscounts = Split(pcv_strCSDiscounts,",")
	    pcv_strPrdDiscounts = getUserInput(request("pPrdDiscounts"),0)  '// All Discounts for main Product
	    pcArray_PrdDiscounts = Split(pcv_strPrdDiscounts,",")
	    pcv_strCSRequired = getUserInput(request("pRequiredIDs"),0)  '// All Cross Sell Required flags
	    pcArray_CSRequired = Split(pcv_strCSRequired,",")

			pcv_ParentDiscount = 0
			pcv_ChildDiscount = 0
			ppcParentIndex = 0
			pcv_ChildBundleID = 0
			
			'// Bundle Discount Selection
			pcv_strSelectedCSProducts = ""
			for iAddM=0 to pcv_intCSArraySize
				'// Set the Parent/Child Discounts	
				if (request("rdobundle") = pcArray_SelectedCSProducts(iAddM)) then 
					pcv_ParentDiscount = pcArray_PrdDiscounts(iAddM)
					pcv_ChildDiscount = pcArray_CSDiscounts(iAddM)
					pcv_ChildBundleID = pcArray_SelectedCSProducts(iAddM)
				
				end if
			next        
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END:  Get the Cross Sell Products
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		tmp_start=1
		if request("from")="BTO" then
			tmp_start=2
			ppcParentIndex = ppcCartIndex
			if request("rdobundle")<>"" then
				pcCartArray(ppcCartIndex,27)=-1
			end if
		end if
		
		Dim pcv_intTotalMultiQty
		pcv_intTotalMultiQty = 0
		for iAddM=tmp_start to pcv_strProductsQuantity

			pIsAccessory=0  '// Cross Sell - Accessory flag (0=false, -1=true)
			
				IF pcv_strProductsQuantity>1 Then '// If there are multiple items to be added
				
				if pCSCnt > 0 then
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  Cross Sell
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	        	
					'// Add the Parent Product
	        		if (iAddM=1) then
		        		'--> Product ID
		       			pIdProduct=getUserInput(request("idproduct"),10)
		        		'--> Quantity
		       		 	pQuantity=getUserInput(request("quantity"),10)
	        		else
						'// Add Required Accessories (checkbox items)
						if pcArray_CSRequired(iAddM-2) <> "0" then
							Session("cs_Accessory") = pcCartArray(ppcParentIndex,1)
							pIsAccessory=-1+pcArray_CSRequired(iAddM-2)
							'--> Product ID
							pIdProduct=getUserInput(pcArray_SelectedCSProducts(iAddM-2),10)
							'--> Quantity
							pQuantity=CheckMinQty(pIdProduct,getUserInput(request("quantity"),10))
		        		elseif (request("bundle"&pcArray_SelectedCSProducts(iAddM-2)) <> "") then 
							'// Add Optional Accessories (checkbox items)
							pIsAccessory=-1+pcArray_CSRequired(iAddM-2)
							'--> Product ID
							pIdProduct=getUserInput(pcArray_SelectedCSProducts(iAddM-2),10)
							'--> Quantity
							pQuantity=CheckMinQty(pIdProduct,getUserInput(request("quantity"),10))
		        		elseif (request("rdobundle") = pcArray_SelectedCSProducts(iAddM-2)) then 
							'// Add Bundle (radio items)
							'--> Product ID
							pIdProduct=getUserInput(pcArray_SelectedCSProducts(iAddM-2),10)
							'--> Quantity
							pQuantity=CheckMinQty(pIdProduct,getUserInput(request("quantity"),10))
		        		else
							pQuantity=0
						end if
				
	        		end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  Cross Sell
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				else
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  Multi Add to Cart
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
					'--> Product ID
					pIdProduct=getUserInput(request("idproduct"&iAddM),10)
					'--> Quantity
					pQuantity=CheckMinQty(pIdProduct,getUserInput(request("QtyM"&pidProduct),10))
					
					'SB S	
					pSubscriptionID=getUserInput(request("pSubscriptionID"&iAddM),10)					
					if pSubscriptionID = "" or pSubscriptionID ="0" Then
						pSubscriptionID = 0
					End if 
					query="SELECT notax, noshipping FROM products WHERE idproduct=" & pIdProduct
	    			set rstemp=server.CreateObject("ADODB.RecordSet")
	    			set rstemp=conntemp.execute(query)
					IsTaxOrShipping = False 
					If NOT rstemp.EOF Then
						If rstemp("notax")<>-1 OR rstemp("noshipping")<>-1 Then
							IsTaxOrShipping = True
						End If
					End If
					If (pcv_sbLockCart=False AND pSubscriptionID>0 AND pcv_BoundQty>0 AND pcv_sbIsCartLockable=True) OR (pcv_sbLockCart=True AND pSubscriptionID>0) OR (pcv_sbLockCart=True AND pSubscriptionID=0 AND IsTaxOrShipping) Then
						response.redirect "msg.asp?message=306"
					End If
					'SB E

					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  Multi Add to Cart
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
				end if		    

		
				'--> if cannot get quantity get quantity 1 (from listing)
				if pQuantity="" then
					pQuantity="0"
				end if
				if (NOT validNum(pQuantity)) OR int(pQuantity)<1 then
					pQuantity=0
				end if
				'--> Check Quantity
				if int(pQuantity)>int(scAddLimit) then
					response.redirect "msg.asp?message=51"         
				end if
			
				If pQuantity>0 Then
					pcv_intTotalMultiQty = 1
				End If
			
			Else '// IF pcv_strProductsQuantity>1 Then
			
			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START:  Single Product
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'--> Product ID
				pIdProduct=getUserInput(request("idproduct"),10)
				'--> Quantity
				pQuantity=getUserInput(request("quantity"),10)
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  Single Product
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
				'// When the muli form page only has one product we need to perfrom corrections.
				' >>> Correct pIdProduct
				if pIdProduct="" then
					pIdProduct=getUserInput(request("idproduct1"),10)
				end if
				' >>> Correct pQuantity
				if pQuantity="" then
					pQuantity=getUserInput(request("QtyM"&pidProduct),10)
				end if
		
					'SB S	
					if pSubscriptionID = 0  Then
						pSubscriptionID=getUserInput(request("pSubscriptionID1"),10)						
					End if 					
					if pSubscriptionID = "" or pSubscriptionID ="0" Then
						pSubscriptionID = 0
					End if 
					query="SELECT notax, noshipping FROM products WHERE idproduct=" & pIdProduct
	    			set rstemp=server.CreateObject("ADODB.RecordSet")
	    			set rstemp=conntemp.execute(query)
					IsTaxOrShipping = False 
					If NOT rstemp.EOF Then
						If rstemp("notax")<>-1 OR rstemp("noshipping")<>-1 Then
							IsTaxOrShipping = True
						End If
					End If
					If (pcv_sbLockCart=False AND pSubscriptionID>0 AND pcv_BoundQty>0 AND pcv_sbIsCartLockable=True) OR (pcv_sbLockCart=True AND pSubscriptionID>0) OR (pcv_sbLockCart=True AND pSubscriptionID=0 AND IsTaxOrShipping) Then
						response.redirect "msg.asp?message=306"
					End If			
					'SB E
		
				'--> if cannot get quantity get quantity 1 (from listing)
				if (NOT validNum(pQuantity)) OR pQuantity="" then
					pQuantity=CheckMinQty(pIdProduct,1)
				else
					if int(pQuantity)<1 then
						pQuantity=CheckMinQty(pIdProduct,1)
					end if
				end if
				
				if int(pQuantity)>int(scAddLimit) then
					 response.redirect "msg.asp?message=40"     
				end if
				
			End If '// IF pcv_strProductsQuantity>1 Then
	
	
			if pQuantity>0 then

				'--> Check Product ID
				if trim(pIdProduct)="" or not validNum(pIdProduct) then
					response.redirect "msg.asp?message=207"
				end if			
			
				'--> New Product Options
				pcv_intOptionGroupCount = getUserInput(request("OptionGroupCount"),0)
				if IsNull(pcv_intOptionGroupCount) OR pcv_intOptionGroupCount="" then
					pcv_intOptionGroupCount = 0
				end if
				pcv_intOptionGroupCount = cint(pcv_intOptionGroupCount)
		
				xOptionGroupCount = 0
				pcv_strSelectedOptions = ""
				if iAddM=1 then
				do until xOptionGroupCount = pcv_intOptionGroupCount	
					xOptionGroupCount = xOptionGroupCount + 1
					pcvstrTmpOptionGroup = request("idOption"&xOptionGroupCount)					
					'// Validate Option ID
					if not validNum(pcvstrTmpOptionGroup) then
						pcvstrTmpOptionGroup=""
					end if
					if pcvstrTmpOptionGroup <> "" then			
						pcv_strSelectedOptions = pcv_strSelectedOptions & pcvstrTmpOptionGroup & chr(124)	
					end if	
				loop
				' trim the last pipe if there is one
				xStringLength = len(pcv_strSelectedOptions)
				if xStringLength>0 then
					pcv_strSelectedOptions = left(pcv_strSelectedOptions,(xStringLength-1))
				end if
			end if
		
	    '--> Custom input fields
	    pxfield1=getUserInput(request("xfield1"),0)
	    pxfield2=getUserInput(request("xfield2"),0)
	    pxfield3=getUserInput(request("xfield3"),0)
		
	    '--> replace line breaks to <br>
	    if pxfield1<>"" then
		    pxfield1=replace(pxfield1,vbCrlf,"<BR>")
	    end if
	    if pxfield2<>"" then
		    pxfield2=replace(pxfield2,vbCrlf,"<BR>")
	    end if
	    if pxfield3<>"" then
		    pxfield3=replace(pxfield3,vbCrlf,"<BR>")
	    end if
	    pxf1=getUserInput(request("xf1"),10)
	    pxf2=getUserInput(request("xf2"),10)
	    pxf3=getUserInput(request("xf3"),10)
	

			'*****************************************************************************************************
			' 1) END: get data from viewPrd form
			'*****************************************************************************************************
	
		
			'*****************************************************************************************************
			' 2) START: get item details
			'*****************************************************************************************************
	
	    noOS=0
		
	    query="SELECT OverSizeSpec FROM products"
	    set rstemp=server.CreateObject("ADODB.RecordSet")
	    set rstemp=conntemp.execute(query)
	    if err.number<>0 then
		    noOS=1
	    end if
	    set rstemp=nothing
	    err.clear
		
		
		'// START v4.1 - Check whether product is not for sale and Not For Sale Override
			'// Check for bundle
			if (pcCartArray(f,27) <> 0) OR (pCSCnt > 0) then
				pcv_Bundles = true
			else
				pcv_Bundles = false
			end if
			if NotForSaleOverride(session("customerCategory"))=1 or pcv_Bundles=true then
				queryNFSO=""
			else
				queryNFSO=" AND products.formQuantity=0"
			end if
		'// END v4.1
		
	    query="SELECT iRewardPoints,description, price, bToBPrice, sku, emailText, weight, deliveringTime, pcSupplier_ID, cost, stock, notax, noshipping, OverSizeSpec, noStock, pcprod_QtyToPound, pcProd_BackOrder, pcProd_ShipNDays, products.pcProd_Surcharge1, products.pcProd_Surcharge2 FROM products WHERE idproduct=" & pIdProduct & queryNFSO & " AND active=-1"
	    set rstemp=server.CreateObject("ADODB.RecordSet")
	    set rstemp=conntemp.execute(query)
	    if err.number<>0 then
		    call LogErrorToDatabase()
		    set rstemp=nothing
		    call closedb()
		    response.redirect "techErr.asp?err="&pcStrCustRefID
	    end if
			
	    if rstemp.eof then 
		    set rstemp=nothing
		    call closeDb()
		    response.redirect "msg.asp?message=41"
	    end if
		
	    iRewardPoints=rstemp("iRewardPoints")
	    iRewardDollars=pPrice * (RewardsPercent / 100)
	    pDescription=rstemp("description")
		pPrice=rstemp("price")
	    pBtoBPrice=rstemp("bToBPrice")
	    pSku=rstemp("sku")
	    pEmailText=rstemp("emailText")
	    pWeight=rstemp("weight")
	    pDeliveringTime=rstemp("deliveringTime")
		if isNULL(pDeliveringTime) OR pDeliveringTime="" then
			pDeliveringTime=0
		end if
	    pIdSupplier=rstemp("pcSupplier_ID")
	    pCost=rstemp("cost")
	    pStock=rstemp("stock")
	    pnotax=rstemp("notax")
	    pnoshipping=rstemp("noshipping")
	    pOverSizeSpec=rstemp("OverSizeSpec")
	    pNoStock=rstemp("noStock")
	    pcv_QtyToPound=rstemp("pcprod_QtyToPound")
	    if isNull(pcv_QtyToPound) OR pcv_QtyToPound="" then
		    pcv_QtyToPound = 0
	    end if
	    'Start SDBA
	    pcv_intBackOrder = rstemp("pcProd_BackOrder")
	    if isNull(pcv_intBackOrder) OR pcv_intBackOrder="" then
		    pcv_intBackOrder = 0
	    end if
	    pcv_intShipNDays = rstemp("pcProd_ShipNDays")
	    if isNull(pcv_intShipNDays) OR pcv_intShipNDays="" then
		    pcv_intShipNDays = 0
	    end if
	    'End SDBA
		pcv_Surcharge1 = rstemp("pcProd_Surcharge1")
		pcv_Surcharge2 = rstemp("pcProd_Surcharge2")
		
	    set rstemp=nothing
	
			'*****************************************************************************************************
			' 2) END: get item details
			'*****************************************************************************************************
		
		
				
			'*****************************************************************************************************
			' 3) START: GET PRODUCT OPTIONS
			'*****************************************************************************************************
	
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
				    'If rs("Wprice")=0 then
					'    pPriceToAdd=rs("price")
				    'End If
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
			    'pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & formatnumber(pPriceToAdd, 2)
					pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & money(pcv_strOptionsPriceTotal)
			    '// Column 5) This is the total of all option prices
			    pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd
				
		    end if
			
		    set rs=nothing
	    Next
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END:  Get the Options for the item
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START:  Get the Custom Fields
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	    xfieldsCnt=0
		
	    if pxfield1<>"" then
		    xfieldsCnt=xfieldsCnt+1
		    query="SELECT xfield FROM xfields WHERE idxfield="&pxf1
		    set rstemp=server.CreateObject("ADODB.RecordSet")
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
			    set rstemp=nothing
			    call closeDb() 
			    response.redirect "msg.asp?message=44"  	  
		    end if
			
		    pXfieldDescrip1=rstemp("xfield")
			
		    set rstemp=nothing
			
	    end if
		
	    if pxfield2<>"" then
		    xfieldsCnt=xfieldsCnt+1
		    query="SELECT xfield FROM xfields WHERE idxfield="&pxf2
		    set rstemp=server.CreateObject("ADODB.RecordSet")
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
		    set rstemp=nothing
			    call closeDb()
			    response.redirect "msg.asp?message=45"  	  
		    end if
			
		    pXfieldDescrip2=rstemp("xfield")
			
		    set rstemp=nothing
			
	    end if
		
	    if pxfield3<>"" then
		    xfieldsCnt=xfieldsCnt+1
		    query="SELECT xfield FROM xfields WHERE idxfield="&pxf3
		    set rstemp=server.CreateObject("ADODB.RecordSet")
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
			    set rstemp=nothing
			    call closeDb()
			    response.redirect "msg.asp?message=46"      	  
		    end if
			
		    pXfieldDescrip3=rstemp("xfield")
			
		    set rstemp=nothing
			
	    end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END:  Get the Custom Fields
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
			lineNumber=0
			iNextIndex=cint(0)
			iCapturedNext=cint(0)
	
			'*****************************************************************************************************
			' 3) END: GET PRODUCT OPTIONS
			'*****************************************************************************************************
    

		'*****************************************************************************************************
		' 4) START:  check if this is an upd
		'*****************************************************************************************************
		tIndex=0
		if request("imode")="updOrd" AND iAddM=1 then
			tIndex=request("index")		
		else
			'// START: Loop
			for f=1 to ppcCartIndex
				'// check if index is deleted and then flag
				if pcCartArray(f,10)=1 AND iCapturedNext=0 then
					iNextIndex=f
					iCapturedNext=1
				end if						
				'// START: if item is not deleted and the idProduct=idProduct of added item
				if (pcCartArray(f,10)=0) and (pcCartArray(f,0)=trim(pIdProduct)) then  
					
					
					if xfieldsCnt=0 then
					'********************************************	
					'// UPDATE CONDITIONS 
					'********************************************			
					
					'// TEST THE CONDITIONS
					'   >>> Returns true if options are involved
					if len(pcv_strOptionsArray)>0 OR len(pcCartArray(f,4))>0 then
						pcv_Options = true
					else
						pcv_Options = false
					end if

					'   >>> Returns true if bundles are involed
					if (pcCartArray(f,27) <> 0) OR (pCSCnt > 0) then
						pcv_Bundles = true
					else
						pcv_Bundles = false
					end if
					
					'11 = Array of Individual Selected Options Id Numbers
					'4 = store array of product "option groups: options"
					'8 = child id line number
					'27 = parent id line number					
					
					'1)  THERE ARE NO OPTIONS OR BUNDLES INVOLVED
						
						'// This means is a single standalone product. We can update it safely
						if (pcv_Options=false) AND (pcv_Bundles=false) then
							lineNumber=f
						end if
	
					'2)  THERE ARE OPTIONS, BUT NO BUNDLES INVOLVED
						if (pcv_Options=true) AND (pcv_Bundles=false) then
							
							'//  Check if the selected options match
							if (trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions)) then
								lineNumber=f  
							end if
							
						end if
						
					'3)  THERE ARE NO OPTIONS, BUT BUNDLES ARE INVOLVED
						if (pcv_Options=false) AND (pcv_Bundles=true) then
					
							'//  Update the parent with the same child as the item being added
							if (trim(pcCartArray(f,8))=trim(pcv_ChildBundleID)) then
								if trim(pcCartArray(f,8)) <> "" then
									lineNumber=f 								
								end if
							end if
							
							'//  Update the child with the same parent as the item being added
							if (trim(pcCartArray(f,27))=trim(ppcParentIndex)) then
								lineNumber=f  
							end if
					
						end if
						
					'4)  BOTH OPTIONS AND BUNDLES ARE INVOLVED
						if (pcv_Options=true) AND (pcv_Bundles=true) then
						
							'//  Update the parent with the same child as the item being added
							if (trim(pcCartArray(f,8))=trim(pcv_ChildBundleID)) AND (trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions)) then
								lineNumber=f 
							end if
							
							'//  Update the child with the same parent as the item being added
							if (trim(pcCartArray(f,27))=trim(ppcParentIndex)) AND (trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions))  then
								lineNumber=f 
							end if
					
						end if
						
					'5) NON-CONFIGURATION PRODUCT BEING ADDED AND CONFIGURATION EXISTS
						If pcCartArray(f,16)<>"" Then
							
							'// BTO Products "without" Configurations should never update a BTO product "with" Configurations.
							lineNumber=0 
					
						end if
						pIdConfigSession=pcCartArray(f,16)
						if pIdConfigSession<>"" then
							pGrTotal1=request("BTOTOTAL"&iAddM)
						end if	
								
					end if
					
				end if
				'// END: if item is not deleted and the idProduct=idProduct of added item
				
			next
			'// END: Loop
		end if
	
		'*****************************************************************************************************
		' 4) END:  check if this is an upd
		'*****************************************************************************************************
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Check Stock
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if (scOutofstockpurchase=-1) AND (pNoStock=0) AND (pcv_intBackOrder=0) then
			if tIndex<>"" then
				tmpIdx=tIndex
			else
				if lineNumber>0 then
					tmpIdx=lineNumber
				else
					tmpIdx=-1
				end if
			end if
			if CheckOFS(pIdProduct,pQuantity,pStock,tmpIdx)=1 then
				call closedb()
				response.Clear()
				response.redirect "msgb.asp?message="&Server.Urlencode("The quantity of "&pDescription&" that you are trying to order is greater than the quantity that we currently have in stock. We currently have "&pStock&" unit(s) in stock.<br><br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &""" border=0></a>" )
			end if
		end if
		call CheckMinMulQty(pIdProduct,pQuantity)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Check Stock
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

	
	
		'*****************************************************************************************************
		' 5) START: ADD/ MODIFY ITEMS IN CART
		'*****************************************************************************************************
		Dim checkSS
		if lineNumber=0 then
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: FULL
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
		'SB S
		If (pcv_sbLockCart=True) AND (pcv_sbIsCartLockable=True) Then
			response.redirect "msg.asp?message=305"
		End If
		'SB E	
				
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Check Stock
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		' Moved to checkCartStockLevels function in includes/productcartinc.asp		

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Check Stock
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Insert Basket Line
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pTotalQuantity=pQuantity
		
		if iCapturedNext=1 then
			ppcCartIndex=iNextIndex
		else
			if tIndex<>0 then
				ppcCartIndex=tIndex
			else
				ppcCartIndex=ppcCartIndex + 1						
				err.clear				
				checkSS=pcCartArray(ppcCartIndex, 0)
				if err.number<>0 then
					if instr(ucase(err.description, "")) then
						response.write dictLanguage.Item(Session("language")&"_instPrd_1")
						response.End()
					end if
				end if
				session("pcCartIndex")=ppcCartIndex
			end if
		end if
				
				
		'// Check for BTO items
		pIdConfigSession = ""			
		if request(iAddM&"||FirstCnt")<>"" then
			Dim Pstring, Vstring, Cstring, tempVar, tempCatarray, tempString, strArray
			Pstring = ""
			Vstring = ""
			Cstring = ""
			Cweight = 0
			FirstCnt = request(iAddM&"||FirstCnt")
			If FirstCnt<>"" then
				for i = 1 to FirstCnt
					tempVar = request(iAddM&"||CAT"&i)
					tempCatarray = split(tempVar,"G")
					tempString = request(iAddM&"||"&tempVar)
					strArray = split(tempString, "_")
					If strArray(0)<>0 then
						Cstring = Cstring & tempCatarray(1) & ","
						Pstring = Pstring & strArray(0) & ","
						Vstring = Vstring & strArray(1) & ","
						Cweight = Cweight + Clng(strArray(2))
					End If
				next
			end if
				
			pConfigKey=trim(randomNumber(9999)&randomNumber(9999))
			Dim pTodayDate
			pTodayDate=Date()
			if SQL_Format="1" then
				pTodayDate=Day(pTodayDate)&"/"&Month(pTodayDate)&"/"&Year(pTodayDate)
			else
				pTodayDate=Month(pTodayDate)&"/"&Day(pTodayDate)&"/"&Year(pTodayDate)
			end if
			err.clear
			pGrTotal1=request("BTOTOTAL"&iAddM)
			if scDB="Access" then
				query="INSERT INTO configSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,dtCreated) VALUES ("&pConfigKey &","&pIdProduct&",'"&Pstring&"','"&Vstring&"','"&Cstring&"',#"&pTodayDate&"#)"
			else
				query="INSERT INTO configSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,dtCreated) VALUES ("&pConfigKey &","&pIdProduct&",'"&Pstring&"','"&Vstring&"','"&Cstring&"','"&pTodayDate&"')"
			end if 
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
					
			if scDB="Access" then
				query="SELECT configSessions.idconfigSession FROM configSessions WHERE (((configSessions.configKey)="&pConfigKey&") AND ((configSessions.dtCreated)=#"&pTodayDate&"#));"
			else
				query="SELECT idconfigSession FROM configSessions WHERE configKey="&pConfigKey&" AND dtCreated='"&pTodayDate&"';"
			end if
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
		
			pIdConfigSession = rsConf("idconfigSession")
			set rsConf=nothing
		end if
				
		pcCartArray(ppcCartIndex,0)=pIdProduct
		pcCartArray(ppcCartIndex,1)=pDescription
		pcCartArray(ppcCartIndex,2)=pQuantity
		
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
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Insert Basket Line
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Check if this customer is logged in with a customer category
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if session("customerCategory")<>0 then
			query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
					
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			if NOT rs.eof then
				strcustomerCategory="YES"
				dblpcCC_Price=rs("pcCC_Price")
				dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
			else
				strcustomerCategory="NO"
			end if
			set rs=nothing
		end if

		if (pBtoBPrice<>0) then
			tempPrice=pBtoBPrice
		else
			tempPrice=pPrice
		end if
				
		if session("customerType")=1 then
			pcCartArray(ppcCartIndex,3)=tempPrice
		else
			pcCartArray(ppcCartIndex,3)=pPrice
		end if 
			
		if session("customerCategoryType")="ATB" then
			if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
				tempPrice=tempPrice-(pcf_Round(tempPrice*(cdbl(session("ATBPercentage"))/100),2))
				pcCartArray(ppcCartIndex,3)=tempPrice
			end if
			if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
				pPrice=pPrice-(pcf_Round(pPrice*(cdbl(session("ATBPercentage"))/100),2))
				pcCartArray(ppcCartIndex,3)=pPrice
			end if
		end if
				
		'if strcustomerCategory="YES" AND dblpcCC_Price>0 then
		if strcustomerCategory="YES" then
			pcCartArray(ppcCartIndex,3)=dblpcCC_Price
		end if

		if pIdConfigSession<>"" then
			pcCartArray(lineNumber,3)=pGrTotal1
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Check if this customer is logged in with a customer category
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Check Options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		pcCartArray(ppcCartIndex,4)="" '// store array of product "option groups: options"
		
		if ( (pCSCnt=0) OR (pCSCnt > 0 and iAddM=1) ) then
			if len(pcv_strOptionsArray)>0 then
				 pcCartArray(ppcCartIndex,4)= pcv_strOptionsArray '// store array of product "option groups: options"
			end if	    
		end if		
				
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Check Options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Populate Additional Columns
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	    
		pcCartArray(ppcCartIndex,5)= pcv_strOptionsPriceTotal '// Total Cost of all Options
		'Options Records
		pcCartArray(ppcCartIndex,23)=pOverSizeSpec
		pcCartArray(ppcCartIndex,25)=pcv_strOptionsPriceArray '// Array of Individual Options Prices
		pcCartArray(ppcCartIndex,26)= pcv_strOptionsPriceArrayCur '// Array of Options Prices stored as currency  'scCurSign & money(pPriceToAdd)
		if pcv_QtyToPound>0 then
			pWeight=(16/pcv_QtyToPound)
			if scShipFromWeightUnit="KGS" then
				pWeight=(1000/pcv_QtyToPound)
			end if
		end if
		pcCartArray(ppcCartIndex,6)=pWeight
		pcCartArray(ppcCartIndex,7)=pSku
		pcCartArray(ppcCartIndex,9)=pDeliveringTime  
		
		' deleted mark
		pcCartArray(ppcCartIndex,10)=0 
		
		pcCartArray(ppcCartIndex,11)=pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers
		pcCartArray(ppcCartIndex,13)=pIdSupplier
		pcCartArray(ppcCartIndex,14)=pCost
		pcCartArray(ppcCartIndex,16) = ""
		pcCartArray(ppcCartIndex,19)=pnotax
		pcCartArray(ppcCartIndex,20)=pnoshipping
		pcCartArray(ppcCartIndex,21)=""
		pcCartArray(ppcCartIndex,36)=pcv_Surcharge1
		pcCartArray(ppcCartIndex,37)=pcv_Surcharge2
		
		'SB S
		pcCartArray(ppcCartIndex,38)= pSubscriptionID '// Subscription ID  
		'SB E
		
		if ( ((pCSCnt=0) OR (pCSCnt > 0 and iAddM=1)) AND xfieldsCnt>0 ) then
			xCnt=0
			if pxfield1<>"" then
				xCnt=1
				pcCartArray(ppcCartIndex,21)=pcCartArray(ppcCartIndex,21)&pXfieldDescrip1&": "&pxfield1
			end if
			if pxfield2<>"" then
				if xCnt=1 then
					pcCartArray(ppcCartIndex,21)=pcCartArray(ppcCartIndex,21)& "<br>"
				end if
				xCnt=1
				pcCartArray(ppcCartIndex,21)=pcCartArray(ppcCartIndex,21)&pXfieldDescrip2&": "&pxfield2
			end if
			if pxfield3<>"" then
				if xCnt=1 then
					pcCartArray(ppcCartIndex,21)=pcCartArray(ppcCartIndex,21)& "<br>"
				end if
				xCnt=1
				pcCartArray(ppcCartIndex,21)=pcCartArray(ppcCartIndex,21)&pXfieldDescrip3&": "&pxfield3
			end if
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Populate Additional Columns
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Populate Cross Sell Columns
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	    
			
		'// If its a Master Product
		if ( iAddM=1 ) AND len(pcv_ChildBundleID)>0 then
			pcCartArray(ppcCartIndex,8) = pcv_ChildBundleID '// id of the child within a bundle
		end if
					
		'// Populate Defaults (Non-Cross)
		pcCartArray(ppcCartIndex,12)=pIsAccessory '// Required Accessory
		pcCartArray(ppcCartIndex,27)=0 '// Relationship - Parent/Child Index (-1=Parent,0=Non-Cross Sell,>0=Parent Index)
		pcCartArray(ppcCartIndex,28)=0 '// Parent/Bundle Discount
		
		'// Populate Accessory										
		if ( pCSCnt > 0 AND pIsAccessory<>0) then
			'// Cross Sell Product - (Child)
			pcCartArray(ppcCartIndex,27)=ppcParentIndex '// Relationship - Parent Index
		end if

		'// Populate Bundles									
		if ( pCSCnt > 0 AND pIsAccessory=0) then
			if ( iAddM=1 ) then
				'// Main Product - (Parent)
				pcCartArray(ppcCartIndex,27)=-1 '// Relationship - This is a Parent
				pcCartArray(ppcCartIndex,28)=pcv_ParentDiscount '// Parent Discount
				'--> Set Parent Index (Cross Sell)
				ppcParentIndex = ppcCartIndex
			else
				'// Cross Sell Product - (Child)
				pcCartArray(ppcCartIndex,27)=ppcParentIndex '// Relationship - Parent Index
				pcCartArray(ppcCartIndex,28)=pcv_ChildDiscount '// Bundle Discount    
			end if
		end if

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Populate Cross Sell Columns
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		
		'// Set current line number
		CurrentProductsIndex=ppcCartIndex
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: FULL
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	


	else
	
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: PARTIAL
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if pcCartArray(lineNumber,2)+ Int(pQuantity) <=Int(scAddLimit) then
				
			'// Set current line number
			CurrentProductsIndex=linenumber
			
			' quantity added + previous quantity is not more than allowed
			if request("imode")="updOrd" then
				pcCartArray(lineNumber,2)=Int(pQuantity)
			else
				pcCartArray(lineNumber,2)=Int(pcCartArray(lineNumber,2)) + Int(pQuantity)		
			end if
			pTotalQuantity=pcCartArray(lineNumber,2)

			'SB S	
			if pSubscriptionID > 0 and pSubInstall <> 1 Then
			    pcCartArray(lineNumber,2) = 1 
				pTotalQuantity = 1
			End if 
			'SB E


			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: 
			' Reset unit price before discounts 
			' Add price or BtoB price depending on customer type
			' Check if this customer is logged in with a customer category
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			if session("customerCategory")<>0 then
				query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
					
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
			
				if NOT rs.eof then
					strcustomerCategory="YES"
					dblpcCC_Price=rs("pcCC_Price")
					dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)  
				else
					strcustomerCategory="NO"
				end if
				set rs=nothing
			end if

			if (pBtoBPrice<>0) then
				tempPrice=pBtoBPrice
			else
				tempPrice=pPrice
			end if
				
			if session("customerType")=1 then
				pcCartArray(lineNumber,3)=tempPrice
			else
				pcCartArray(lineNumber,3)=pPrice
			end if 
			
			if session("customerCategoryType")="ATB" then
				if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
					tempPrice=tempPrice-(pcf_Round(tempPrice*(cdbl(session("ATBPercentage"))/100),2))
					pcCartArray(lineNumber,3)=tempPrice
				end if
				if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
					pPrice=pPrice-(pcf_Round(pPrice*(cdbl(session("ATBPercentage"))/100),2))
					pcCartArray(lineNumber,3)=pPrice
				end if
			end if
				
			'if strcustomerCategory="YES" AND dblpcCC_Price>0 then
			if strcustomerCategory="YES" then
				pcCartArray(lineNumber,3)=dblpcCC_Price
			end if

			if pIdConfigSession<>"" then
				pcCartArray(lineNumber,3)=pGrTotal1
			end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: 
			' Reset unit price before discounts 
			' Add price or BtoB price depending on customer type
			' Check if this customer is logged in with a customer category
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Cross Selling
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			'// If the product being updated is a master product, declare its line number.
			'// We also delcare the line number when adding a master product.
			if ( iAddM=1 ) then
				ppcParentIndex = lineNumber
			end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: Cross Selling
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
    				
		else
			response.redirect "msg.asp?message=49"         
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: PARTIAL
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
	end if	
	'*****************************************************************************************************
	' 5) END: ADD/ MODIFY ITEMS IN CART
	'*****************************************************************************************************
	
	
	'*****************************************************************************************************
	' 6) START: get discount per quantity
	'*****************************************************************************************************
	disTotalQuantity=pTotalQuantity
	
	query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &disTotalQuantity& " AND quantityUntil>=" &disTotalQuantity
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
				
	if Session("customerType")=1 OR session("customerCategory")<>0 then
		pcCartArray(f,18)=1
	else
		pcCartArray(f,18)=0
	end if
		
	pOrigPrice=pcCartArray(CurrentProductsIndex,3)
	pcCartArray(CurrentProductsIndex,15)=0
	pcCartArray(CurrentProductsIndex,17)=pOrigPrice
	if not rstemp.eof and err.number<>9 then
		' there are quantity discounts defined for that quantity 
		pDiscountPerUnit=rstemp("discountPerUnit")
		pDiscountPerWUnit=rstemp("discountPerWUnit")
		pPercentage=rstemp("percentage")
		pbaseproductonly=rstemp("baseproductonly")
		
		if session("customerType")<>1 then  'customer is a normal user
			if pPercentage="0" then 
				pcCartArray(CurrentProductsIndex,3)=pcCartArray(CurrentProductsIndex,3) - pDiscountPerUnit  'Price - discount per unit
				pcCartArray(CurrentProductsIndex,15)=pcCartArray(CurrentProductsIndex,15) + (pDiscountPerUnit * pTotalQuantity)  'running total of discounts
			else
				if pbaseproductonly="-1" then
					pcCartArray(CurrentProductsIndex,3)=pcCartArray(CurrentProductsIndex,3) - ((pDiscountPerUnit/100) * pcCartArray(CurrentProductsIndex,17))
				else
					pcCartArray(CurrentProductsIndex,3)=pcCartArray(CurrentProductsIndex,3) - ((pDiscountPerUnit/100) * (pcCartArray(CurrentProductsIndex,17)+pcCartArray(CurrentProductsIndex,5)))
				end if
				if pbaseproductonly="-1" then
					pcCartArray(CurrentProductsIndex,15)=pcCartArray(CurrentProductsIndex,15) + (((pDiscountPerUnit/100) * pOrigPrice) * pTotalQuantity)
				else
					pcCartArray(CurrentProductsIndex,15)=pcCartArray(CurrentProductsIndex,15) + (((pDiscountPerUnit/100) * (pOrigPrice+pcCartArray(CurrentProductsIndex,5))) * pTotalQuantity)
				end if
			end if
		else  'customer is a wholesale customer
			if pPercentage="0" then 
				pcCartArray(CurrentProductsIndex,3)=pcCartArray(CurrentProductsIndex,3) - pDiscountPerWUnit
				pcCartArray(CurrentProductsIndex,15)=pcCartArray(CurrentProductsIndex,15) + (pDiscountPerWUnit * pTotalQuantity)
			else
				if pbaseproductonly="-1" then
					pcCartArray(CurrentProductsIndex,3)=pcCartArray(CurrentProductsIndex,3) - ((pDiscountPerWUnit/100) * pcCartArray(CurrentProductsIndex,17))
				else
					pcCartArray(CurrentProductsIndex,3)=pcCartArray(CurrentProductsIndex,3) - ((pDiscountPerWUnit/100) * (pcCartArray(CurrentProductsIndex,17)+pcCartArray(CurrentProductsIndex,5)))
				end if
				if pbaseproductonly="-1" then
					pcCartArray(CurrentProductsIndex,15)=pcCartArray(CurrentProductsIndex,15) + (((pDiscountPerWUnit/100) * pOrigPrice)* pTotalQuantity)
				else
					pcCartArray(CurrentProductsIndex,15)=pcCartArray(CurrentProductsIndex,15) + (((pDiscountPerWUnit/100) * (pOrigPrice+pcCartArray(CurrentProductsIndex,5))) * pTotalQuantity)
				end if
			end if
		end if
	end if

	set rstemp=nothing
	
	'*****************************************************************************************************
	' 6) END: get discount per quantity
	'*****************************************************************************************************
	
	'*****************************************************************************************************
	' 7) START: RP ADDON
	'*****************************************************************************************************
	If RewardsActive <> 0 Then
		pcCartArray(CurrentProductsIndex,22)=int(iRewardPoints)
	End If
	'*****************************************************************************************************
	' 7) END: RP ADDON
	'*****************************************************************************************************	
	
	'*****************************************************************************************************
	' 9) START:  Clean up and Redirect
	'*****************************************************************************************************		

	end if '// if pQuantity>1 then

next
%>
<!--#include file="inc-UpdPrdQtyDiscounts.asp"-->
<%
session("pcCartSession")=pcCartArray

'Calculate Product Promotions - START
%>
<!--#include file="inc_CalPromotions.asp"-->
<%
'Calculate Product Promotions - END

call closeDB() 
call clearLanguage() 

' redirect to cart view
Session("pcSessionID")=Session.SessionID '// browser test session

If pcv_strProductsQuantity>1 AND pcv_intTotalMultiQty=0 Then
	response.redirect "viewCart.asp?cs=1" '// cs = Check Session. Initializes the session check script.
End If

if scATCEnabled="1" then %>
	<!--#include file="atc_instprd.asp"-->
<% else
	response.redirect "viewCart.asp?cs=1" '// cs = Check Session. Initializes the session check script.
end if

'*****************************************************************************************************
' 9) END:  Clean up and Redirect
'*****************************************************************************************************
%>		
		
		Please wait while we process your items...
		
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->

<%
' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function
%>