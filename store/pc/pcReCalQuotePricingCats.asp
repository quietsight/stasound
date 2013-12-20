<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

Dim pPrice,dblpcCC_Price,tmpItemPrice,tmpItemWPrice

SUB updPricingCats(pcv_chkIdQuote,pIdProduct,pcv_chkIDConfig)
Dim query,rstemp,m,BTODefaultPrice

					'Check if this customer is logged in with a customer category
					if session("customerCategory")<>0 then
						query="SELECT idCC_Price, pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
						if NOT rstemp.eof then
							idCC_Price=rstemp("idCC_Price")
							dblpcCC_Price=rstemp("pcCC_Price")
							dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
							'if dblpcCC_Price>0 then
								strcustomerCategory="YES"
							'else
							'	strcustomerCategory="NO"
							'end if
						else
							strcustomerCategory="NO"
						end if
						set rstemp=nothing
					end if

					'====================
					' get original price
					'====================
					query="SELECT price, bToBPrice FROM products WHERE idProduct=" &pIdProduct
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
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
					pTotalQuantity=1
					IF pcv_chkIDConfig<>"0" and pcv_chkIDConfig<>"" then
						query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice,pcconf_Quantity FROM configWishlistSessions WHERE idconfigWishlistSession=" & pcv_chkIDConfig
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
							pTotalQuantity=rstemp("pcconf_Quantity")
						end if
						
						set rstemp=nothing
						
						if Pstring<>"" and uCase(Pstring)<>"NA" then
						
							ArrProduct=split(Pstring,",")
							ArrValue=Split(Vstring, ",")
							ArrCategory=Split(Cstring, ",")
							ArrQuantity=Split(Qstring,",")
							ArrPrice=split(Pricestring,",")
							
							For m=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								query="SELECT price, Wprice FROM configSpec_products WHERE specProduct="&pIdProduct&" AND configProduct=" &ArrProduct(m)
								set rstemp=server.CreateObject("ADODB.RecordSet")
								set rstemp=conntemp.execute(query)
					
								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
	 
								tmpItemWPrice_A=rstemp("Wprice")
								tmpItemPrice_A=rstemp("price")
								
								set rstemp=nothing
								
								tmpItemPrice=CheckPrdPrices(pIdProduct,ArrProduct(m),tmpItemPrice_A,tmpItemWPrice_A,0)
								tmpItemWPrice=CheckPrdPrices(pIdProduct,ArrProduct(m),tmpItemPrice_A,tmpItemWPrice_A,1)
								
								if Session("customerType")=1 then
									dblItemPrice=tmpItemWPrice
								else
									dblItemPrice=tmpItemPrice
								end if
								
								ArrPrice(m)=dblItemPrice
								
								tmpDefaultItemPrice=0
								
								if cdbl(ArrValue(m))<>0 then
									tmpDefaultItemPrice=GetDefaultPrice(pIdProduct,ArrCategory(m))*GetDefaultQty(pIdProduct,ArrCategory(m))
									tmpDef=ChkMultiSelectDef(pIdProduct,ArrCategory(m),ArrProduct(m))
									if tmpDef=2 then
										BTODefaultPrice=BTODefaultPrice+tmpDefaultItemPrice
										ArrValue(m)=cdbl(ArrPrice(m))-cdbl(tmpDefaultItemPrice)
									else
										if tmpDef=1 then
											tmpDefaultItemPrice=0
											ArrValue(m)=cdbl(ArrPrice(m))-cdbl(tmpDefaultItemPrice)
										end if
									end if
								end if
								
								AllItemPrices=AllItemPrices+cdbl(ArrPrice(m))-cdbl(tmpDefaultItemPrice)+(cdbl(ArrQuantity(m))-1)*cdbl(ArrPrice(m))
								
								'====================
								' get items discounts
								'====================
								query="SELECT quantityFrom,quantityUntil,discountperUnit,percentage,discountperWUnit FROM discountsPerQuantity WHERE IDProduct=" & ArrProduct(m)
								set rstemp=connTemp.execute(query)

								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closedb()
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
								Pstring=Pstring & ArrProduct(m) & ","
								Vstring=Vstring & ArrValue(m) & ","
								Cstring=Cstring & ArrCategory(m) & ","
								Qstring=Qstring & ArrQuantity(m) & ","
								Pricestring=Pricestring & ArrPrice(m) & ","
							Next
							
						end if 'Have BTO Items
						
						'Check BTO Addtional Charges
						query="SELECT stringCProducts,stringCValues,stringCCategories FROM configWishlistSessions WHERE idconfigWishlistSession=" & pcv_chkIDConfig
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
					
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closedb()
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
								query="SELECT price, Wprice FROM configSpec_Charges WHERE specProduct="&pIdProduct&" AND configProduct=" &ArrProduct(m)
								set rstemp=server.CreateObject("ADODB.RecordSet")
								set rstemp=conntemp.execute(query)
					
								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
	 
								tmpItemWPrice_A=rstemp("Wprice")
								tmpItemPrice_A=rstemp("price")
								
								set rstemp=nothing
								
								tmpItemPrice=CheckPrdPrices(pIdProduct,ArrProduct(m),tmpItemPrice_A,tmpItemWPrice_A,0)
								tmpItemWPrice=CheckPrdPrices(pIdProduct,ArrProduct(m),tmpItemPrice_A,tmpItemWPrice_A,1)
								
								if Session("customerType")=1 then
									dblItemPrice=tmpItemWPrice
								else
									dblItemPrice=tmpItemPrice
								end if
								
								ArrValue(m)=dblItemPrice
								
								tmpDefaultItemPrice=0
								
								if cdbl(ArrValue(m))<>0 then
									'tmpDefaultItemPrice=GetCDefaultPrice(pIdProduct,ArrCategory(m))
									ArrValue(m)=dblItemPrice '-tmpDefaultItemPrice
								end if
								
								AllCItemPrices=AllCItemPrices+ArrValue(m)
								
							Next
							
							PCstring=""
							VCstring=""
							CCstring=""
							For m=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								PCstring=PCstring & ArrProduct(m) & ","
								VCstring=VCstring & ArrValue(m) & ","
								CCstring=CCstring & ArrCategory(m) & ","
							Next
						end if 'Have BTO Additional Charges
						
					END IF
					
					pPrice=BTODefaultPrice+AllItemPrices
					pOrigPrice=pPrice

					'====================
					' get product discount per quantity
					'====================					


					query="SELECT discountPerUnit,discountPerWUnit,percentage FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
					set rstemp=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if


					if pcQDiscountType<>"1" then
						pOrigPrice=pOrigPrice-(ItemsDiscounts/pTotalQuantity)
					else
						pOrigPrice=pOrigPrice
					end if
					
					pQtyDisc=0

					if not rstemp.eof then
					 	' there are quantity discounts defined for that quantity 
					 	pDiscountPerUnit = rstemp("discountPerUnit")
					 	pDiscountPerWUnit = rstemp("discountPerWUnit")
					 	pPercentage = rstemp("percentage")

					 	if session("customerType")<>1 then
					 		if pPercentage = "0" then 
					 			pPrice  = pPrice - pDiscountPerUnit
								pQtyDisc = pQtyDisc + (pDiscountPerUnit * pTotalQuantity)
							else
								pPrice = pPrice - ((pDiscountPerUnit/100) * pOrigPrice)
								pQtyDisc = pQtyDisc + ((pDiscountPerUnit/100) * (pOrigPrice * pTotalQuantity))
							end if
						else
							if pPercentage = "0" then 
								pPrice  = pPrice - pDiscountPerWUnit
								pQtyDisc = pQtyDisc + (pDiscountPerWUnit * pTotalQuantity)
							else
								pPrice = pPrice - ((pDiscountPerWUnit/100) * pOrigPrice)
								pQtyDisc = pQtyDisc + ((pDiscountPerWUnit/100) * (pOrigPrice * pTotalQuantity))
							end if
						end if
					end if
					set rstemp=nothing
					
					pQtyDisc=round(pQtyDisc,2)
					ItemsDiscounts=round(ItemsDiscounts,2)
					pPrice=pOrigPrice*pTotalQuantity-pQtyDisc+AllCItemPrices
					ItemsDiscounts=ItemsDiscounts*(-1)
					
					query="UPDATE configWishlistSessions SET stringProducts='" & Pstring & "',stringValues='" & Vstring & "',stringCategories='" & Cstring & "',stringQuantity='" & Qstring & "',stringPrice='" & Pricestring & "',fPrice=" & pPrice & ",dPrice=" & ItemsDiscounts & ",pcconf_QDiscount=" & pQtyDisc & ",stringCProducts='" & PCstring & "',stringCValues='" & VCstring & "',stringCCategories='" & CCstring & "' WHERE idconfigWishlistSession=" & pcv_chkIDConfig
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
	
END SUB %>