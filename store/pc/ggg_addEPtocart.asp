<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
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
<!--#include file="ggg_inc_chkEPPrices.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<% err.number=0
dim query, conntemp, rstemp

call opendb()

grCode=getUserInput(request("grcode"),0)
gIDEvent=getUserInput(request("IDEvent"),0)

Dim pcCartArray, pcCartIndex
pcCartArray=session("pcCartSession")
pcCartIndex=Session("pcCartIndex")

'Check Shopping Cart
if (countCartRows(pcCartArray, pcCartIndex)<>0) and (Session("Cust_BuyGift")<>"") and (Session("Cust_IDEvent")<>gIDEvent) then
  response.redirect "msg.asp?message=101"      
end if
if (countCartRows(pcCartArray, pcCartIndex)<>0) and (Session("Cust_BuyGift")="") then
  response.redirect "msg.asp?message=102"
end if

ppcCartIndex=pcCartIndex

'// GET The Number Of Items to Add
Count=getUserInput(request("Count"),0)
if Count="" then
	Count="0"
end if

if grCode="" then
	response.redirect "msg.asp?message=98"
end if

if gIDEvent="" then
	response.redirect "msg.asp?message=98"
end if

query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_Code='" & grCode & "' and pcEv_IDEvent=" & gIDEvent & " and pcEv_Active=1"
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	response.redirect "msg.asp?message=98"
end if

gAdd=0


'/////////////////////////////////////////////////////////////////////////////////////////////////////
'// START: ADD TO CART LOOP
'/////////////////////////////////////////////////////////////////////////////////////////////////////
For dd=1 to Count

	geID=getUserInput(request("geID" & dd),0)
	geadd=getUserInput(request("add" & dd),0) 
	if geadd="" then
		geadd="0"
	end if
	
	'/////////////////////////////////////////////////////////////////////////////////////////////////
	'// START:  Filter Out Products with no Quantity or ID
	'/////////////////////////////////////////////////////////////////////////////////////////////////
	IF (geID<>"") AND (clng(geadd)>0) then
	
		gAdd=1
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  Get All the Required Information
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="SELECT pcEvProducts.pcEP_idProduct, pcEvProducts.pcEP_OptionsArray, pcEvProducts.pcEP_xdetails,pcEvProducts.pcEP_IDConfig "
		query=query&",products.description,products.sku, products.weight,products.pcprod_QtyToPound,products.emailText, products.deliveringTime, products.pcSupplier_ID, products.cost, products.stock, products.notax, products.noshipping, products.iRewardPoints, noStock, pcProd_BackOrder, products.pcProd_Surcharge1, products.pcProd_Surcharge2 FROM products,pcEvProducts WHERE pcEvProducts.pcEP_ID=" & geID & " and pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_idproduct and products.removed=0"
		
		set rstemp=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		pidProduct=rstemp("pcEP_idProduct")
		pquantity=geadd
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Product Options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcv_strSelectedOptions=""
		pcv_strSelectedOptions = rstemp("pcEP_OptionsArray")
		pcv_strSelectedOptions=pcv_strSelectedOptions&""		
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
		pIdSupplier	= rstemp("pcSupplier_ID")
		pCost = rstemp("cost")
		pStock = rstemp("stock")
		pnotax = rstemp("notax")
		pnoshipping = rstemp("noshipping")
		iRewardPoints = rstemp("iRewardPoints")
		
		pNoStock=rstemp("noStock")
		if IsNull(pNoStock) or pNoStock="" then
			pNoStock=0
		end if
		
		pcv_intBackOrder=rstemp("pcProd_BackOrder")
		if IsNull(pcv_intBackOrder) or pcv_intBackOrder="" then
			pcv_intBackOrder=0
		end if
		pcv_Surcharge1 = rstemp("pcProd_Surcharge1")
		pcv_Surcharge2 = rstemp("pcProd_Surcharge2")
		
		pIdConfigSession=trim(pidconfigSession)
				
				
				
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
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END:  Get the Options for the item
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			
			punitPrice=updPrices(pidProduct,pIdConfigSession,pcv_strOptionsPriceTotal,1)
		
			if 	pIdConfigSession="0" then
			pIdConfigSession=""
			end if
		
		'---------------------------
		'Calculate BTO Items Weights
		IF pIdConfigSession<>"" then
		query="SELECT stringProducts,stringCProducts,stringQuantity FROM configSessions WHERE idconfigSession=" & pIdConfigSession
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
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
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
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
		
		
		lineNumber=0
		dim iNextIndex, iCapturedNext
		iNextIndex=cint(0)
		iCapturedNext=cint(0)
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END:  Get All the Required Information
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  Update or Add (DEV NOTE: This section is almost, but not identical to instPrd.asp)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if request.Form("imode")="updOrd" then
			tIndex=request.Form("index")
			
		else
			'// START: Loop
			for f=1 to ppcCartIndex
				'// check if index is deleted and then flag
				if pcCartArray(f,10)=1 AND iCapturedNext=0 then
					iNextIndex=f
					iCapturedNext=1
				end if	

				'// START: if item is not deleted and the idProduct=idProduct of added item
				if (pcCartArray(f,10)=0) and (pcCartArray(f,0)=int(trim(pIdProduct))) then 
					if scOutofstockpurchase=-1 AND pNoStock=0 AND pcv_intBackOrder=0 then
						iTempStockTotal=0
						for g=1 to ppcCartIndex
							if (pcCartArray(g,10)=0) and (pcCartArray(g,0)=trim(pIdProduct)) then         
								'checking stock level
								iTempStockTotal=Int(pcCartArray(g,2))+Int(iTempStockTotal)
							end if
						next
						iTempStockTotal=Int(iTempStockTotal)+Int(pQuantity)
						if Int(iTempStockTotal)>Int(pStock) then
							call closeDb()
							'Set session variables to handle error on msg.asp
							session("pcErrStrPrdDesc") = pDescription
							session("pcErrIntStock") = pStock
							response.redirect "msg.asp?message=204"						
						end if
					end if
					
					if pxfdetails="" then
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
					pcv_Bundles = false
					'11 = Array of Individual Selected Options Id Numbers
					'4 = store array of product "option groups: options"
					'8 = child id line number
					'27 = parent id line number					
					
					'1)  THERE ARE NO OPTIONS OR BUNDLES INVOLVED
						
						'// This means is a single standalone product. We can update it safely
					    if (pcv_Options=false) AND (pcv_Bundles=false) then
							lineNumber=f
							'response.write "1) " & (pcv_isBundles)
							'response.end
					    end if

					'2)  THERE ARE OPTIONS, BUT NO BUNDLES INVOLVED
					    if (pcv_Options=true) AND (pcv_Bundles=false) then
							
							'//  Check only optionals with values
							'if (trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions) AND trim(pcv_strSelectedOptions)<>"") AND (NOT pcv_isBundles) then
							'	lineNumber=f   
							'	response.write "2) " & (pcv_isBundles)
							'	response.end
							'end if	
						
							'//  Check if the selected options match
							if (trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions)) then
								lineNumber=f  
								'response.write "2) " & (pcv_isBundles)
								'response.end
							end if
							
						end if
						
					'3)  THERE ARE NO OPTIONS, BUT BUNDLES ARE INVOLVED
					    if (pcv_Options=false) AND (pcv_Bundles=true) then
						
							'//  Update the parent with the same child as the item being added
							if (trim(pcCartArray(f,8))=trim(pcv_ChildBundleID)) then
								if trim(pcCartArray(f,8)) <> "" then
									lineNumber=f 								
									'response.write "3) " & (pcv_Bundles)
									'response.end
								end if
							end if
							
							'//  Update the child with the same parent as the item being added
							if (trim(pcCartArray(f,27))=trim(ppcParentIndex)) then
								lineNumber=f  
								'response.write "3b) " & (pcv_Bundles)
								'response.end
							end if
								response.write "3b) " & (pcv_Bundles)
								response.end
						end if
						
					'4)  BOTH OPTIONS AND BUNDLES ARE INVOLVED
					    if (pcv_Options=true) AND (pcv_Bundles=true) then
						
							'//  Update the parent with the same child as the item being added
							if (trim(pcCartArray(f,8))=trim(pcv_ChildBundleID)) AND (trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions)) then
								lineNumber=f 
								'response.write "4) " & (pcv_Bundles)
								'response.end
							end if
							
							'//  Update the child with the same parent as the item being added
							if (trim(pcCartArray(f,27))=trim(ppcParentIndex)) AND (trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions))  then
								lineNumber=f 
								'response.write "4b) " & (pcv_Bundles)
								'response.end
							end if
					
						end if
						
                    end if					
				end if
				'// END: if item is not deleted and the idProduct=idProduct of added item				
			next
			'// END: Loop
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End:  Update or Add
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
	'*************************************************************************************************
	' START: ADD/ MODIFY ITEMS IN CART
	'*************************************************************************************************
		Dim checkSS
		if lineNumber=0 then
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: FULL
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Check Stock
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				if scOutofstockpurchase=-1 AND pNoStock=0 AND pcv_intBackOrder=0 then
					pTotalQuantity=pQuantity
					if Int(pQuantity)>Int(pStock) then
						call closeDb()
						'Set session variables to handle error on msg.asp
						session("pcErrStrPrdDesc") = pDescription
						session("pcErrIntStock") = pStock
						response.redirect "msg.asp?message=204"
					end if
				end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: Check Stock
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Insert Basket Line (DEV NOTE: This section is almost, but not identical to instPrd.asp. It does not have BTO)
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
		
			' add price or BtoB price depending on customer type
			pcCartArray(ppcCartIndex,3) = punitPrice
			'response.write pcCartArray(ppcCartIndex,3)
			'response.end
		
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
		
					pcCartArray(ppcCartIndex,6) = pWeight + Cweight
					pcCartArray(ppcCartIndex,7) = pSku
					pcCartArray(ppcCartIndex,9) = pDeliveringTime  
		
					' deleted mark
					pcCartArray(ppcCartIndex,10)=0 

					pcCartArray(ppcCartIndex,11)=pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers
					pcCartArray(ppcCartIndex,13) = pIdSupplier
					pcCartArray(ppcCartIndex,14) = pCost
					pcCartArray(ppcCartIndex,16) = pIdConfigSession
					pcCartArray(ppcCartIndex,19) = pnotax
					pcCartArray(ppcCartIndex,20) = pnoshipping
					pcCartArray(ppcCartIndex,21) = pxfdetails
					pcCartArray(ppcCartIndex,36) = pcv_Surcharge1
					pcCartArray(ppcCartIndex,37) = pcv_Surcharge2
					
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Populate Cross Sell Columns
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	    
				    
					'// If its a Master Product
					'if ( iAddM=1 ) AND len(pcv_ChildBundleID)>0 then
					'	pcCartArray(ppcCartIndex,8) = pcv_ChildBundleID '// id of the child within a bundle
					'end if
					
				    '// Populate Defaults (Non-Cross)
					'pcCartArray(ppcCartIndex,12)=pIsAccessory '// Required Accessory
			        'pcCartArray(ppcCartIndex,27)=0 '// Relationship - Parent/Child Index (-1=Parent,0=Non-Cross Sell,>0=Parent Index)
			        'pcCartArray(ppcCartIndex,28)=0 '// Parent/Bundle Discount
					
					'// Populate Accessory										
					'if ( pCSCnt > 0 AND pIsAccessory<>0) then
				        '// Cross Sell Product - (Child)
				   '     pcCartArray(ppcCartIndex,27)=ppcParentIndex '// Relationship - Parent Index
					'end if

					'// Populate Bundles									
					'if ( pCSCnt > 0 AND pIsAccessory=0) then
					'    if ( iAddM=1 ) then
					        '// Main Product - (Parent)
					'        pcCartArray(ppcCartIndex,27)=-1 '// Relationship - This is a Parent
					'        pcCartArray(ppcCartIndex,28)=pcv_ParentDiscount '// Parent Discount
							'--> Set Parent Index (Cross Sell)
					'		ppcParentIndex = ppcCartIndex

					'    else
					        '// Cross Sell Product - (Child)
					'        pcCartArray(ppcCartIndex,27)=ppcParentIndex '// Relationship - Parent Index
					'        pcCartArray(ppcCartIndex,28)=pcv_ChildDiscount '// Bundle Discount    
					'    end if
					'end if

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: Populate Cross Sell Columns
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			

		
		pTotalQuantity = pQuantity
		
		'-----------------------------
		'ReCalculate BTO Items Discounts
		
		itemsDiscounts=0
		if pIdConfigSession<>"" then 
			query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
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
			if ArrProduct(0)<>"na" then
				for j=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(j)
					set rsQ=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsQ=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
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
					itemsDiscounts=ItemsDiscounts+TempDiscount
				next			
			end if 'Have BTO Items
		end if 'Have ConfigSession
		
		pcCartArray(ppcCartIndex,30)=cdbl(itemsDiscounts)
		
		'End ReCulculate BTO Items Discounts		
		'------------------------------------
		
		'------------------------------------
		'BTO Additional Charges
		Charges=0
		if pIdConfigSession<>"" then 
			query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsConfigObj=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsConfigObj=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			stringCProducts=rsConfigObj("stringCProducts")
			stringCValues=rsConfigObj("stringCValues")
			stringCCategories=rsConfigObj("stringCCategories")
			ArrCProduct=Split(stringCProducts, ",")
			ArrCValue=Split(stringCValues, ",")
			ArrCCategory=Split(stringCCategories, ",")
			if ArrCProduct(0)<>"na" then
				for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
					query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
					set rsConfigObj=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsConfigObj=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if (CDbl(ArrCValue(i))>0)then
						Charges=Charges+cdbl(ArrCValue(i))
					end if
					set rsConfigObj=nothing
				next
				set rsConfigObj=nothing
			end if 
		end if
						
		pcCartArray(ppcCartIndex,31) = Cdbl(Charges)
		
		'End BTO Additional Charges
		'------------------------------------
		
		' get discount per quantity
		
		query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
		set rsQ=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsQ=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		if Session("customerType")=1 OR session("customerCategory")<>0 then
			pcCartArray(f,18)=1
		else
			pcCartArray(f,18)=0
		end if
		
		pOrigPrice = pcCartArray(ppcCartIndex,3)
		pcCartArray(ppcCartIndex,17) = pOrigPrice
			
		pcCartArray(ppcCartIndex,15) = 0
		if not rsQ.eof and err.number<>9 then
			' there are quantity discounts defined for that quantity 
			pDiscountPerUnit = rsQ("discountPerUnit")
			pDiscountPerWUnit = rsQ("discountPerWUnit")
			pPercentage = rsQ("percentage")
		
			if session("customerType")<>1 then
				if pPercentage = "0" then 
					pcCartArray(ppcCartIndex,3)  = pcCartArray(ppcCartIndex,3) - pDiscountPerUnit
					pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (pDiscountPerUnit * pTotalQuantity)
				else
					pcCartArray(ppcCartIndex,3) = pcCartArray(ppcCartIndex,3) - ((pDiscountPerUnit/100) * pcCartArray(ppcCartIndex,17))
					pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (((pDiscountPerUnit/100) * pOrigPrice) * pTotalQuantity)
				end if
			else
				if pPercentage = "0" then 
					pcCartArray(ppcCartIndex,3)  = pcCartArray(ppcCartIndex,3) - pDiscountPerWUnit
					pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (pDiscountPerWUnit * pTotalQuantity)
				else
					pcCartArray(ppcCartIndex,3) = pcCartArray(ppcCartIndex,3) - ((pDiscountPerWUnit/100) * pcCartArray(ppcCartIndex,17))
					pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (((pDiscountPerWUnit/100) * pOrigPrice)* pTotalQuantity)
				end if
			end if
		end if 
		'// Start Reward Points
		If RewardsActive <> 0 Then
			pcCartArray(ppcCartIndex,22) = Clng(iRewardPoints)
		End If
		'// End Reward Points
		
		pcCartArray(ppcCartIndex,33) = geID 'IDs of Products in the Gift Registry
		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: FULL
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		ELSE

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: PARTIAL
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			    if pcCartArray(lineNumber,2)+ Int(pQuantity) <=Int(scAddLimit) then
				    ' quantity added + previous quantity is not more than allowed
				    pcCartArray(lineNumber,2)=Int(pcCartArray(lineNumber,2)) + Int(pQuantity)
				    pTotalQuantity=pcCartArray(lineNumber,2)
    				
				    ' reset unit price before discounts 
				    ' add price or BtoB price depending on customer type
				   	'if session("customerType")=1 and pBtoBPrice>"0" then
					'    pcCartArray(lineNumber,3)=pBtoBPrice
				    'else
					'    pcCartArray(lineNumber,3)=pPrice
				    'end if
    				
				    'if pIdConfigSession<>"" then
					'    pcCartArray(lineNumber,3)=pGrTotal1
				    'end if
					
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Cross Selling
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
					'// If the product being updated is a master product, declare its line number.
					'// We also delcare the line number when adding a master product.
					'if ( iAddM=1 ) then
					'	ppcParentIndex = lineNumber
					'end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: Cross Selling
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
    				
			    else
				    response.redirect "msg.asp?message=49"         
			    end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: PARTIAL
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		
		END IF 'New Cart record
	'*************************************************************************************************
	' END: ADD/ MODIFY ITEMS IN CART
	'*************************************************************************************************
		
	END IF 'Have Product (geID<>"")
	'/////////////////////////////////////////////////////////////////////////////////////////////////
	'// END:  Filter Out Products with no Quantity or ID
	'/////////////////////////////////////////////////////////////////////////////////////////////////
	
Next
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' END: ADD TO CART LOOP
'/////////////////////////////////////////////////////////////////////////////////////////////////////


if gAdd=1 then
	session("Cust_BuyGift")="ok"
	session("Cust_IDEvent")=gIDEvent
end if

session("pcCartSession") = pcCartArray

call closedb()

response.redirect "viewcart.asp"
%>