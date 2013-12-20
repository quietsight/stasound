<% 
'// Double-check customer category assignment

If Session("idCustomer") <> 0 AND Session("idCustomer") <> "" Then
	query="SELECT idcustomerCategory FROM customers WHERE idcustomer="&session("idCustomer")&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then Session("customerCategory") = rs(0)
	set rs=nothing
End IF

'// Get customer category type to add to session and to update cart
if NOT isNULL(session("customerCategory")) and session("customerCategory")<>"" and session("customerCategory")<>0 then
	query="SELECT pcCC_Name, pcCC_Description, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory="&session("customerCategory")&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then
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
else
	session("customerCategory")=0
	session("ATBCustomer")=0
	session("ATBPercentage")=0
	session("ATBPercentOff")=0
end if

'update cart for wholesale customer, if wholesale customer did not exist til now
If (NeedReCalculate=1 OR Session("customerType")=1 OR session("customerCategory")<>0) AND pcCartArray(1,18)=0 Then
	
	'---start recalculations
	for t=1 to ppcCartIndex 'Start Loop
		if pcCartArray(t,10)=0 then 'if product not remove
			pcCartArray(t,18)=1

			' check discounts per quantity and recalculates the price
			
			pSIdProduct=pcCartArray(t,0)
			pTotalQuantity=pcCartArray(t,2)

			'Check if this customer is logged in with a customer category
			if session("customerCategory")<>0 then
				query="SELECT idCC_Price, pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pSIdProduct&";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if NOT rs.eof then
					idCC_Price=rs("idCC_Price")
					dblpcCC_Price=rs("pcCC_Price")
					dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
					'if dblpcCC_Price>0 then
						strcustomerCategory="YES"
					'else
					'	strcustomerCategory="NO"
					'end if
				else
					strcustomerCategory="NO"
				end if
				set rs=nothing
			end if
			
			if pSIdProduct<>"" then 'if product id exists
				'Check if product is a BTO product, if so, do not recalculate total
				query="SELECT serviceSpec FROM products WHERE idProduct=" &pSIdProduct
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				tserviceSpec=rstemp("serviceSpec")
				set rstemp=nothing
				
				IF tserviceSpec=false THEN 'recalculate for wholesale customer
					'====================
					' get original price
					'====================
					query="SELECT price, bToBPrice FROM products WHERE idProduct=" &pSIdProduct
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closeDb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
	 
					if rstemp.eof then
						set rstemp=nothing
						call closeDb()
						response.redirect "msg.asp?message=41"
					end if
					
					tempbToBPrice=rstemp("bToBPrice")
					tempprice=rstemp("price")
					
					if (tempbToBPrice<>0) then
						tempBPrice=tempbToBPrice
					else
						tempBPrice=tempprice
					end if
					
					if session("customerType")=1 then
						pPrice=tempBPrice
					else
						pPrice=tempprice
					end if 
					
					if session("customerCategoryType")="ATB" then
						if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
							tempBPrice=tempBPrice-(pcf_Round(tempBPrice*(cdbl(session("ATBPercentage"))/100),2))
							pPrice=tempBPrice
						end if
						if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
							tempprice=tempprice-(pcf_Round(tempprice*(cdbl(session("ATBPercentage"))/100),2))
							pPrice=tempprice
						end if						
					end if
				
					if strcustomerCategory="YES" then
						pPrice=dblpcCC_Price
					end if
					
					pcCartArray(t,3)=pPrice  '*****
					'**************************************************************************
					' START: GET OPTIONS
					'******************************************************************************
				
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START:  Get the Options for the item
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pcv_strSelectedOptions = pcCartArray(t,11)
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
							call LogErrorToDatabase()
							set rs=nothing
							call closeDb()
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
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END:  Get the Options for the item
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
			
					'***************************************************************
					' END: GET OPTIONS
					'*************************************************************
					
					pcCartArray(t,25) = pcv_strOptionsPriceArray '// Array of Individual Options Prices
					'pcCartArray(t,27)= "" '// Not in use anymore - VERIFY FOR Crosssell
					pcCartArray(t,5) = pcv_strOptionsPriceTotal '// Total Cost of all Options '*****
					'get discount per quantity
					query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" &pSIdProduct& " AND quantityFrom<=" &pcCartArray(t,2)& " AND quantityUntil>=" &pcCartArray(t,2) '*****
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
			
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closeDb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
						
					pOrigPrice=pPrice
					pcCartArray(t,17)=pOrigPrice '*****
					if not rstemp.eof then
						' there are quantity discounts defined for that quantity
						pDiscountPerUnit=rstemp("discountPerUnit")
						pDiscountPerWUnit=rstemp("discountPerWUnit")
						
						pPercentage=rstemp("percentage")
						pbaseproductonly=rstemp("baseproductonly")
						
						If Session("customerType")=1 Then
							pDiscountToUse=pDiscountPerWUnit
						Else
							pDiscountToUse=pDiscountPerUnit
						End If					
						
						'--
						if pPercentage="0" then 
							pcCartArray(t,3)=pPrice - pDiscountToUse '*****
							pcCartArray(t,15)=pDiscountToUse * pTotalQuantity '*****
						else
							if pbaseproductonly="-1" then
								pcCartArray(t,3)=pPrice - ((pDiscountToUse/100) * pOrigPrice) '*****
							else
								pcCartArray(t,3)=pPrice - ((pDiscountToUse/100) * (pOrigPrice+pcCartArray(t,5))) '*****
							end if
							
							if pbaseproductonly="-1" then
								pcCartArray(t,15)=(pDiscountToUse/100) * pOrigPrice * pTotalQuantity '*****
							else
								pcCartArray(t,15)=(pDiscountToUse/100) * (pOrigPrice+pcCartArray(t,5)) * pTotalQuantity '*****
							end if
						end if						
						'--
						
					end if 'discounts
					
				ELSE
				'---------------------
				' UPDATE BTO PRODUCT
				'---------------------
				
					'====================
					' get original price
					'====================
					query="SELECT price, bToBPrice FROM products WHERE idProduct=" &pSIdProduct
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closeDb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
	 
					if rstemp.eof then
						call closeDb()
						response.redirect "msg.asp?message=41"
					end if
					
					tempbToBPrice=rstemp("bToBPrice")
					tempprice=rstemp("price")
					
					if (tempbToBPrice<>0) then
						tempBPrice=tempbToBPrice
					else
						tempBPrice=tempprice
					end if
					
					if session("customerType")=1 then
						pPrice=tempBPrice
					else
						pPrice=tempprice
					end if 
					
					if session("customerCategoryType")="ATB" then
						if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
							tempBPrice=tempBPrice-(pcf_Round(tempBPrice*(cdbl(session("ATBPercentage"))/100),2))
							pPrice=tempBPrice
						end if
						if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
							tempprice=tempprice-(pcf_Round(tempprice*(cdbl(session("ATBPercentage"))/100),2))							
							pPrice=tempprice
						end if						
					end if
				
					if strcustomerCategory="YES" then
						pPrice=dblpcCC_Price
					end if
					
					BTODefaultPrice=pPrice
					
					'====================
					' get items price
					'====================
					AllItemPrices=0
					AllCItemPrices=0
					ItemsDiscounts=0
					IF pcCartArray(t,16)<>"0" and pcCartArray(t,16)<>"" then
						query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pcCartArray(t,16)
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
					
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					
						if not rstemp.eof then
							Pstring=rstemp("stringProducts")
							Vstring=rstemp("stringValues")
							Cstring=rstemp("stringCategories")
							Qstring=rstemp("stringQuantity")
							Pricestring=rstemp("stringPrice")
						end if
						
						set rstemp=nothing
						
						if Pstring<>"" and uCase(Pstring)<>"NA" then
						
							ArrProduct=split(Pstring,",")
							ArrValue=Split(Vstring, ",")
							ArrCategory=Split(Cstring, ",")
							ArrQuantity=Split(Qstring,",")
							ArrPrice=split(Pricestring,",")
							
							For m=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								tmpIDeQty=GetItemDefaultQty(ArrProduct(m))
								if Clng(ArrQuantity(m))<Clng(tmpIDeQty) then 'Check item Quantity
									ArrQuantity(m)=tmpIDeQty
								end if
								query="SELECT price, Wprice FROM configSpec_products WHERE specProduct="&pSIdProduct&" AND configProduct=" &ArrProduct(m) & " AND configProductCategory=" & ArrCategory(m)
								set rstemp=server.CreateObject("ADODB.RecordSet")
								set rstemp=conntemp.execute(query)
					
								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closeDb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
	 
								if NOT rstemp.eof then	 
									tmpItemWPrice_A=rstemp("Wprice")
									tmpItemPrice_A=rstemp("price")
								end if
								
								set rstemp=nothing
								
								tmpItemPrice=CheckPrdPrices(pSIdProduct,ArrProduct(m),tmpItemPrice_A,tmpItemWPrice_A,0)
								tmpItemWPrice=CheckPrdPrices(pSIdProduct,ArrProduct(m),tmpItemPrice_A,tmpItemWPrice_A,1)
								
								if Session("customerType")=1 then
									dblItemPrice=tmpItemWPrice
								else
									dblItemPrice=tmpItemPrice
								end if
								
								ArrPrice(m)=dblItemPrice
								
								tmpDefaultItemPrice=0
								
								'if cdbl(ArrValue(m))<>0 then
									tmpDefaultItemPrice=GetDefaultPrice(pSIdProduct,ArrCategory(m))*GetDefaultQty(pSIdProduct,ArrCategory(m))
									tmpDef=ChkMultiSelectDef(pSIdProduct,ArrCategory(m),ArrProduct(m))
									if tmpDef=2 then
										BTODefaultPrice=BTODefaultPrice+tmpDefaultItemPrice
										ArrValue(m)=cdbl(ArrPrice(m))-cdbl(tmpDefaultItemPrice)
									else
										if tmpDef=1 then
											tmpDefaultItemPrice=0
											ArrValue(m)=cdbl(ArrPrice(m))-cdbl(tmpDefaultItemPrice)
										else
											if tmpDef=0 then
												BTODefaultPrice=BTODefaultPrice+tmpDefaultItemPrice
												ArrValue(m)=cdbl(ArrPrice(m))-cdbl(tmpDefaultItemPrice)
											end if
										end if
									end if
								'end if
								
								AllItemPrices=AllItemPrices+cdbl(ArrPrice(m))-cdbl(tmpDefaultItemPrice)+(cdbl(ArrQuantity(m))-1)*cdbl(ArrPrice(m))
								
								'====================
								' get items discounts
								'====================
								query="SELECT quantityFrom,quantityUntil,discountperUnit,percentage,discountperWUnit FROM discountsPerQuantity WHERE IDProduct=" & ArrProduct(m)
								set rstemp=connTemp.execute(query)

								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closeDb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
 
								TempDiscount=0
								do while not rstemp.eof
					 				QFrom=rstemp("quantityFrom")
									QTo=rstemp("quantityUntil")
									DUnit=rstemp("discountperUnit")
									QPercent=rstemp("percentage")
									DWUnit=rstemp("discountperWUnit")
									if (DWUnit=0) and (DUnit>0) then
										DWUnit=DUnit
									end if
									

									TempD1=0
									if (clng(ArrQuantity(m)*pTotalQuantity)>=clng(QFrom)) and (clng(ArrQuantity(m)*pTotalQuantity)<=clng(QTo)) then
										if QPercent="-1" then
											if session("customerType")=1 then
												TempD1=ArrQuantity(m)*pTotalQuantity*ArrPrice(m)*0.01*DWUnit
											else
												TempD1=ArrQuantity(m)*pTotalQuantity*ArrPrice(m)*0.01*DUnit
											end if
										else
											if session("customerType")=1 then
												TempD1=ArrQuantity(m)*pTotalQuantity*DWUnit
											else
												TempD1=ArrQuantity(m)*pTotalQuantity*DUnit
											end if
										end if
									end if
									TempDiscount=TempDiscount+TempD1
									rstemp.movenext
								loop
								set rstemp=nothing
								ItemsDiscounts=ItemsDiscounts+TempDiscount
							Next
							
							Pstring=""
							Vstring=""
							Cstring=""
							Qstring=""
							Pricestring=""
							For m=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								IF Clng(ArrQuantity(m))>0 THEN
									Pstring=Pstring & replace(ArrProduct(m),",",".") & ","
									Vstring=Vstring & replace(ArrValue(m),",",".") & ","
									Cstring=Cstring & replace(ArrCategory(m),",",".") & ","
									Qstring=Qstring & replace(ArrQuantity(m),",",".") & ","
									Pricestring=Pricestring & replace(ArrPrice(m),",",".") & ","
								END IF
							Next
							
							query="UPDATE configSessions SET stringProducts='" & Pstring & "',stringValues='" & Vstring & "',stringCategories='" & Cstring & "',stringQuantity='" & Qstring & "',stringPrice='" & Pricestring & "' WHERE idconfigSession=" & pcCartArray(t,16)
							set rstemp=connTemp.execute(query)
							set rstemp=nothing

						end if 'Have BTO Items
						
						'Check BTO Additional Charges
						query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pcCartArray(t,16)
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
					
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closeDb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					
						if not rstemp.eof then
							PCstring=rstemp("stringCProducts")
							VCstring=rstemp("stringCValues")
							CCstring=rstemp("stringCCategories")
						end if
						
						set rstemp=nothing
						
						if PCstring<>"" and uCase(PCstring)<>"NA" then
						
							ArrProduct=split(PCstring,",")
							ArrValue=Split(VCstring, ",")
							ArrCategory=Split(CCstring, ",")
							
							For m=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								query="SELECT price, Wprice FROM configSpec_Charges WHERE specProduct="&pSIdProduct&" AND configProduct=" &ArrProduct(m)
								set rstemp=server.CreateObject("ADODB.RecordSet")
								set rstemp=conntemp.execute(query)
					
								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closeDb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
	 
								tmpItemWPrice=rstemp("Wprice")
								tmpItemPrice=rstemp("price")
								
								set rstemp=nothing
								
								tmpItemPrice=CheckPrdPrices(pSIdProduct,ArrProduct(m),tmpItemPrice,tmpItemWPrice,0)
								tmpItemWPrice=CheckPrdPrices(pSIdProduct,ArrProduct(m),tmpItemPrice,tmpItemWPrice,1)
								
								if Session("customerType")=1 then
									dblItemPrice=tmpItemWPrice
								else
									dblItemPrice=tmpItemPrice
								end if
								
								ArrValue(m)=dblItemPrice
								
								if cdbl(ArrValue(m))<>0 then
									'tmpDefaultItemPrice=GetCDefaultPrice(pSIdProduct,ArrCategory(m))
									ArrValue(m)=dblItemPrice '-cdbl(tmpDefaultItemPrice)
								end if
								
								AllCItemPrices=AllCItemPrices+ArrValue(m)
							Next
							
							PCstring=""
							VCstring=""
							CCstring=""

							For m=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								PCstring=PCstring & replace(ArrProduct(m),",",".") & ","
								VCstring=VCstring & replace(ArrValue(m),",",".") & ","
								CCstring=CCstring & replace(ArrCategory(m),",",".") & ","
							Next
							
							query="UPDATE configSessions SET stringCProducts='" & PCstring & "',stringCValues='" & VCstring & "',stringCCategories='" & CCstring & "' WHERE idconfigSession=" & pcCartArray(t,16)
							set rstemp=connTemp.execute(query)
							set rstemp=nothing

						end if 'Have BTO Additional Charges
						
					END IF
				
					pcCartArray(t,3)=BTODefaultPrice+AllItemPrices
					pcCartArray(t,30)=Round(ItemsDiscounts+0.001,2)
					pcCartArray(t,31)=AllCItemPrices
					pcCartArray(t,17)=pcCartArray(t,3)
					
					'====================
					' get product discount per quantity
					'====================					


					query="SELECT discountPerUnit,discountPerWUnit,percentage FROM discountsPerQuantity WHERE idProduct=" &pSIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
					set rstemp=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closeDb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if

					pOrigPrice=pcCartArray(t,3)

					if pcQDiscountType<>"1" then
						pOrigPrice=pOrigPrice-(pcCartArray(t,30)/pTotalQuantity)
					else
						pOrigPrice=pOrigPrice
					end if
					
					pcCartArray(t,15)=0

					if not rstemp.eof then
					 	' there are quantity discounts defined for that quantity 
					 	pDiscountPerUnit = rstemp("discountPerUnit")
					 	pDiscountPerWUnit = rstemp("discountPerWUnit")
					 	pPercentage = rstemp("percentage")

					 	if session("customerType")<>1 then
					 		if pPercentage = "0" then 
					 			pcCartArray(t,3)  = pcCartArray(t,3) - pDiscountPerUnit
								pcCartArray(t,15) = pcCartArray(t,15) + (pDiscountPerUnit * pTotalQuantity)
							else
								pcCartArray(t,3) = pcCartArray(t,3) - ((pDiscountPerUnit/100) * pOrigPrice)
								pcCartArray(t,15) = pcCartArray(t,15) + ((pDiscountPerUnit/100) * (pOrigPrice * pTotalQuantity))
							end if
						else
							if pPercentage = "0" then 
								pcCartArray(t,3)  = pcCartArray(t,3) - pDiscountPerWUnit
								pcCartArray(t,15) = pcCartArray(t,15) + (pDiscountPerWUnit * pTotalQuantity)
							else
								pcCartArray(t,3) = pcCartArray(t,3) - ((pDiscountPerWUnit/100) * pOrigPrice)
								pcCartArray(t,15) = pcCartArray(t,15) + ((pDiscountPerWUnit/100) * (pOrigPrice * pTotalQuantity))
							end if
						end if
					end if
					set rstemp=nothing
				
					
				END IF 'not serviceSpec
			end if 'pSIdProduct<>""
			
			'SM-S
			if UCase(scDB)="SQL" then
				query="SELECT Products.pcSC_ID,pcSales_BackUp.pcSales_TargetPrice FROM pcSales_BackUp INNER JOIN Products ON pcSales_BackUp.pcSC_ID=Products.pcSC_ID WHERE Products.idProduct=" & pSIdProduct
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
						pcCartArray(t,39)=tmpSCID '//Sale ID
					else
						pcCartArray(t,39)=0
					end if
				else
					pcCartArray(t,39)=0
				end if
				set rsQ=nothing
			else
				pcCartArray(t,39)=0
		end if
		'SM-E
		
		end if 'if product not remove
		session("pcCartSession")=pcCartArray
	next 'Loop around
	%>
	<!--#include file="inc-UpdPrdQtyDiscounts.asp"-->
	<!--#include file="inc-ReCalCrossSell.asp"-->
	<%
	session("pcCartSession")=pcCartArray
	
	'Calculate Product Promotions - START
	%>
	<!--#include file="inc_CalPromotions.asp"-->
	<%
	'Calculate Product Promotions - END
	
	'After reCalculation, check to see if wholesale amount is qualified
	'dim pCartTotalWeight, pCartQuantity, howMuch

	pCartTotalWeight=Cdbl(0)
	pCartQuantity=Cint(0)
	howMuch=CDbl(0)

	' calculate total price of the order, total weight and product total quantities
	pSubtotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
	pCartTotalWeight=Cdbl(calculateCartWeight(pcCartArray, ppcCartIndex))
	pCartQuantity=int(calculateCartQuantity(pcCartArray, ppcCartIndex))

	If session("customerType")=1 AND ppcCartIndex>0 Then
		if (calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase) and (Session("SFStrRedirectUrl")="") then  
			'response.redirect "msgb.asp?message="&Server.URLEncode(dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scWholesaleMinPurchase)&dictLanguage.Item(Session("language")&"_techErr_3")&"<BR><BR><a href=""viewCart.asp"">"&dictLanguage.Item(Session("language")&"_titles_3")&"</a> : <a href=""default.asp"">"&dictLanguage.Item(Session("language")&"_titles_5")&"</a></b>")
		end if
	End If
end If
'check for wholesale customer recalculation%>