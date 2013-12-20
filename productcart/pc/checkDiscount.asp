<% err.clear
pDiscountError=""
pFreeShip=""
piddiscount=""
pDiscountCode1=replace(pDiscountCode,"'","''")
query="SELECT iddiscount, onetime, expDate, idProduct, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, DiscountDesc, priceToDiscount, percentageToDiscount, pcDisc_StartDate, pcRetailFlag, pcWholesaleFlag, pcDisc_PerToFlatCartTotal, pcDisc_PerToFlatDiscount,pcDisc_IncExcPrd,pcDisc_IncExcCat,pcDisc_IncExcCust,pcDisc_IncExcCPrice FROM discounts WHERE discountcode='" &pDiscountCode1& "' AND active=-1"
set rsDisObj=Server.CreateObject("ADODB.RecordSet")
set rsDisObj=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rsDisObj=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rsDisObj.eof then
	pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4") 
else
	piddiscount=rsDisObj("iddiscount")
	ponetime=rsDisObj("onetime")
	pexpDate=rsDisObj("expDate")
	tmpidProduct=rsDisObj("idProduct")
	pquantityFrom=rsDisObj("quantityFrom")
	pquantityUntil=rsDisObj("quantityUntil")
	pweightFrom=rsDisObj("weightFrom")
	pweightUntil=rsDisObj("weightUntil")
	ppriceFrom=rsDisObj("priceFrom")
	ppriceUntil=rsDisObj("priceUntil")
	pDiscountDesc=rsDisObj("DiscountDesc")
	ppriceToDiscount=rsDisObj("priceToDiscount")
	ppercentageToDiscount=rsDisObj("percentageToDiscount")
	pStartDate=rsDisObj("pcDisc_StartDate")
	pcv_retail = rsDisObj("pcRetailFlag")
	pcv_wholeSale = rsDisObj("pcWholeSaleFlag")
	pcv_PerToFlatCartTotal = rsDisObj("pcDisc_PerToFlatCartTotal")
	pcv_PerToFlatDiscount = rsDisObj("pcDisc_PerToFlatDiscount")
	pcIncExcPrd=rsDisObj("pcDisc_IncExcPrd")
	pcIncExcCat=rsDisObj("pcDisc_IncExcCat")
	pcIncExcCust=rsDisObj("pcDisc_IncExcCust")
	pcIncExcCPrice=rsDisObj("pcDisc_IncExcCPrice")
	
	set rsDisObj=nothing
	
	'check to see if discount has been used for one use only for this customer specified
	If ponetime=true Then
		'check customer's id in database with iddiscount
		query="SELECT * FROM used_discounts WHERE idcustomer="&session("IDCustomer")&" AND iddiscount=" &piddiscount
		set rsUsedObj=Server.CreateObject("ADODB.RecordSet")
		set rsUsedObj=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsUsedObj=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if NOT rsUsedObj.eof then
			'discount has been used already by the customer
			pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
		end if
		set rsUsedObj=nothing
	Else
		'check to see if discount code is expired
		If pexpDate<>"" then
			expDate=pexpDate
			If datediff("d", Now(), expDate) <= 0 Then
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
			Else
				' check if the discount has defined the product   
				pVerPrdCode=-1     
				if isNull(tmpidProduct) or tmpidProduct=0 then
					' discount is across the board
				else
					' find out if the product is in the cart
					if findProduct(pcCartArray, ppcCartIndex, tmpidProduct)=0 then
						pVerPrdCode=0
					end if
				end if   
			end if
		end if
		
		'check to see if discount has start date
		If pStartDate<>"" then
			StartDate=pStartDate
			If datediff("d", Now(), StartDate) > 0 Then
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_43")
			End If
		end if
	end if

'************ FILTERS

	pcv_Filters=0
	pcv_FResults=0
	pcv_IDDiscount=piddiscount
	
	IF pcv_IDDiscount<>"" THEN	
								'Filter by Products
								query="select pcFPro_IDProduct from PcDFProds where pcFPro_IDDiscount=" & pcv_IDDiscount
								set rsF=server.CreateObject("ADODB.RecordSet")
	
								set rsF=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsF=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								if not rsF.eof then
									pcv_Filters=pcv_Filters+1
									tmpIDArr=rsF.getRows()
									set rsF=nothing
									intIDCount=ubound(tmpIDArr,2)
											tmpgotit=0
											for ik=0 to intIDCount
												if clng(QIDProduct)=clng(tmpIDArr(0,ik)) then
													tmpgotit=1
													exit for
												end if
											next
											if (pcIncExcPrd="0") AND (tmpgotit=1) then
												pcv_FResults=1
											else
												if (pcIncExcPrd="1") AND (tmpgotit=0) then
													pcv_FResults=1
												end if
											end if
								end if
								set rsF=nothing
								'End of Filter by Products

								'Filter by Categories
								if pcv_Filters=0 then
									query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount
									set rsF=server.CreateObject("ADODB.RecordSet")
									set rsF=connTemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
										set rsF=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if						
									if not rsF.eof then
										pcv_Filters=pcv_Filters+1
												query="select idcategory from categories_products where idproduct=" & QIDProduct
												set rsF=server.CreateObject("ADODB.RecordSet")
												set rsF=connTemp.execute(query)
												if err.number<>0 then
													call LogErrorToDatabase()
													set rsF=nothing
													call closedb()
													response.redirect "techErr.asp?err="&pcStrCustRefID
												end if												
												
												if not rsF.eof then
                                                	tmpCatArr=rsF.getRows()
                                                    set rsF=nothing
                                                    intCatCount=ubound(tmpCatArr,2)
                                                    tmpgotit=0
													'Check assigned categories
                                                    For ik=0 to intCatCount
														pcv_IDCat=tmpCatArr(o,ik)
														query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount & " and pcFCat_IDCategory=" & pcv_IDCat
														set rstempF=server.CreateObject("ADODB.RecordSet")
														set rstempF=connTemp.execute(query)
														if err.number<>0 then
															call LogErrorToDatabase()
															set rstempF=nothing
															call closedb()
															response.redirect "techErr.asp?err="&pcStrCustRefID
														end if
														if not rstempF.eof then
															set rstempF=nothing
                                                        	tmpgotit=1
                                                            exit for
														end if
                                                        set rstempF=nothing
                                                        'Check parent-categories
                                                        if (tmpgotit=0) AND (pcv_IDCat<>"1") then
                                                        	pcv_ParentIDCat=pcv_IDCat
															do while (tmpgotit=0) and (pcv_ParentIDCat<>"1")
																query="select idParentCategory from categories where idcategory=" & pcv_ParentIDCat

																set rstempF=server.CreateObject("ADODB.RecordSet")
																set rstempF=connTemp.execute(query)
																if err.number<>0 then
																	call LogErrorToDatabase()
																	set rstempF=nothing
																	call closedb()
																	response.redirect "techErr.asp?err="&pcStrCustRefID
																end if														
																if not rstempF.eof then
																	pcv_ParentIDCat=rstempF("idParentCategory")
																	if pcv_ParentIDCat<>"1" then
																		query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount & " and pcFCat_IDCategory=" & pcv_ParentIDCat & " and pcFCat_SubCats=1;"
																		set rstempF=server.CreateObject("ADODB.RecordSet")
																		set rstempF=connTemp.execute(query)
																		if err.number<>0 then
																			call LogErrorToDatabase()
																			set rstempF=nothing
																			call closedb()
																			response.redirect "techErr.asp?err="&pcStrCustRefID
																		end if
																		if not rstempF.eof then
																			tmpgotit=1
																		end if
																		set rstempF=nothing
																	end if
																end if
                                                                set rstempF=nothing
															loop
                                                        end if
														set rstempF=nothing
                                                        if tmpgotit=1 then
                                                            exit for
														end if
													Next
													if (pcIncExcCat="0") AND (tmpgotit=1) then
														pcv_FResults=1
													else
														if (pcIncExcCat="1") AND (tmpgotit=0) then
															pcv_FResults=1
														end if
													end if
												end if
                                                set rsF=nothing
									end if
								end if
								'End of Filter by Categories
								
								'Filter by Customers
								pcv_CustFilter=0
								query="select pcFCust_IDCustomer from PcDFCusts where pcFCust_IDDiscount=" & pcv_IDDiscount
								set rsF=server.CreateObject("ADODB.RecordSet")
								set rsF=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsF=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if								
								if not rsF.eof then
									pcv_Filters=pcv_Filters+1
									pcv_CustFilter=1
								end if
								set rsF=nothing
								
								if pcv_CustFilter=1 then
		
								query="select pcFCust_IDCustomer from PcDFCusts where pcFCust_IDDiscount=" & pcv_IDDiscount & " and pcFCust_IDCustomer=" & session("IDCustomer")
								set rsF=server.CreateObject("ADODB.RecordSet")
								set rsF=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsF=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if							
								if not rsF.eof then
									if (pcIncExcCust="0") then
										pcv_FResults=pcv_FResults+1
									end if
								else
									if (pcIncExcCust="1") then
										pcv_FResults=pcv_FResults+1
									end if
								end if
								set rsF=nothing
								
								end if
								'End of Filter by Customers
								
								
								'Filter by Customer Categories
								pcv_CustCatFilter=0
								
								query="select pcFCPCat_IDCategory from pcDFCustPriceCats where pcFCPCat_IDDiscount=" & pcv_IDDiscount
								set rsF=server.CreateObject("ADODB.RecordSet")
								set rsF=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsF=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if								
								if not rsF.eof then
									pcv_Filters=pcv_Filters+1
									pcv_CustCatFilter=1
								end if
								set rsF=nothing
								

								if pcv_CustCatFilter=1 then
		
								query="select pcDFCustPriceCats.pcFCPCat_IDCategory from pcDFCustPriceCats, Customers where pcDFCustPriceCats.pcFCPCat_IDDiscount=" & pcv_IDDiscount & " and pcDFCustPriceCats.pcFCPCat_IDCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
								set rsF=server.CreateObject("ADODB.RecordSet")
								set rsF=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsF=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if							
								if not rsF.eof then
									if (pcIncExcCPrice="0") then
										pcv_FResults=pcv_FResults+1
									end if
								else
									if (pcIncExcCPrice="1") then
										pcv_FResults=pcv_FResults+1
									end if
								end if
								set rsF=nothing
								
								end if
								'End of Filter by Customer Categories
								
								' Check to see if discount is filtered by reatil or wholesale.
		                         if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
							       pcv_Filters=pcv_Filters+1
								   if pcv_wholeSale = "1" and session("customertype") = 1 then
								   	pcv_FResults=pcv_FResults+1	
								   end if 
								   if pcv_retail = "1" and 	session("customertype") <> 1 Then
								    pcv_FResults=pcv_FResults+1
								   end if    
							     end if 
		                        ' end retail wholesale check
	
		if pcv_Filters<>pcv_FResults then
			pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_38")
		end if
		set rsF=nothing
	END IF

'****************** END OF FILTERS

	If pDiscountError="" Then
		if 1>=Int(pquantityFrom) and 1<=Int(pquantityUntil) and Int(pweight)>=Int(pweightFrom) and Int(pweight)<=Int(pweightUntil) and Cdbl(pSubTotal)>=Cdbl(ppriceFrom) and Cdbl(pSubTotal)<=Cdbl(ppriceUntil) then
		if pPriceToDiscount>0 or ppercentageToDiscount>0 then
			pDiscountDesc=pDiscountDesc
			pPriceToDiscount=cdbl(ppriceToDiscount)
			ppercentageToDiscount=ppercentageToDiscount
		else
			pFreeShip="FREE SHIPPING if you choose a shipping service that is compatible with this discount code"
		end if
		else
			pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5") 
		end if
	End If
end if
%>