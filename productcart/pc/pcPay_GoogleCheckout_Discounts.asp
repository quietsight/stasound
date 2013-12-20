<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<% 
'//////////////////////////////////////////////////////////////
' START: DISCOUNTS
'//////////////////////////////////////////////////////////////
'// Set errors to zero
pDiscountError=0
'// Discount Price is zero
discountTotal=ccur(0)
'// Set all Variables
pcv_IDDiscount=""
pcv_IDDiscount1=""
pcv_OneTime=""
expDate=""
dcIdProduct=""
dcQuantityFrom=""
dcQuantityUntil=""
dcWeightFrom=""
dcWeightUntil=""
dcPriceFrom=""
dcPriceUntil=""
pDiscountDesc=""
pPriceToDiscount=ccur(0)
ppercentageToDiscount=""
intPcSeparate=""
intPcAuto=""

pSubTotal=Session("pSubTotal")
pCartQuantity=Session("pCartQuantity")
pCartTotalWeight=Session("pCartTotalWeight")

'*************************************************************
' START: CHECK FOR DISCOUNT
'*************************************************************
query="SELECT iddiscount, onetime,expDate, idProduct, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, DiscountDesc, priceToDiscount, percentageToDiscount, pcSeparate, pcDisc_Auto, pcDisc_StartDate, pcRetailFlag, pcWholesaleFlag, pcDisc_PerToFlatCartTotal, pcDisc_PerToFlatDiscount,pcDisc_IncExcPrd,pcDisc_IncExcCat,pcDisc_IncExcCust,pcDisc_IncExcCPrice FROM discounts WHERE discountcode='" &pDiscountCode& "' AND active=-1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)	
if rs.eof then
	'// There are no discounts
	pcv_strDiscountErrorMsg="This is not a valid Code."
	pDiscountError=1
	GC="1"
else
	'// Set all the Discount Data
	pcv_IDDiscount=rs("iddiscount")
	pcv_IDDiscount1=pcv_IDDiscount
	pcv_OneTime=rs("onetime")
	expDate=rs("expDate")
	dcIdProduct=rs("idProduct")
	dcQuantityFrom=rs("quantityFrom")
	dcQuantityUntil=rs("quantityUntil")
	dcWeightFrom=rs("weightFrom")
	dcWeightUntil=rs("weightUntil")
	dcPriceFrom=rs("priceFrom")
	dcPriceUntil=rs("priceUntil")
	pDiscountDesc=rs("DiscountDesc")
	pPriceToDiscount=ccur(rs("priceToDiscount"))
	ppercentageToDiscount=rs("percentageToDiscount")
	intPcSeparate=rs("pcSeparate")
	intPcAuto=rs("pcDisc_Auto")
	pcv_StartDate=rs("pcDisc_StartDate")
	pcv_retail = rs("pcRetailFlag")
	pcv_wholeSale = rs("pcWholeSaleFlag")
	pcv_PerToFlatCartTotal = rs("pcDisc_PerToFlatCartTotal")
	pcv_PerToFlatDiscount = rs("pcDisc_PerToFlatDiscount")
	pcIncExcPrd=rs("pcDisc_IncExcPrd")
	pcIncExcCat=rs("pcDisc_IncExcCat")
	pcIncExcCust=rs("pcDisc_IncExcCust")
	pcIncExcCPrice=rs("pcDisc_IncExcCPrice")
	if intPcSeparate="" OR IsNull(intPcSeparate) then
		pcv_HaveSeparateCode=0
	else
		pcv_HaveSeparateCode=intPcSeparate
	end if
end if
'*************************************************************
' END: CHECK FOR DISCOUNT
'*************************************************************




Dim UsedDiscountCodes, intArryCnt
'*************************************************************
' START: THERE ARE DISCOUNTS
'*************************************************************
IF pDiscountError=0 THEN	

	'// Set filter variables 
	UsedDiscountCodes=Session("UsedDiscountCodes")
	intArryCnt=0
	CatCount=1
	CatFound=0
	intCodeCnt=0

	pcv_Filters=0
	pcv_FResults=0
	pcv_ProTotal=0

	'// Set the Discount Code to a temporary variable
	pTempDiscCode=pDiscountCode


	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start: Perform Checks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If pDiscountError=0 Then

		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: CHECK FOR MULTI USE CODES (Order)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if intPcSeparate="0" AND Session("SFUsedDiscountCodes")="YES" AND Session("SF"&elemCode)="" then
			'// This discount code may not be used with other discounts
			pcv_strDiscountErrorMsg = "This discount code may not be used with other discounts. Click ""Change Order"" above to use this code."
			pDiscountError=1
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: CHECK FOR MULTI USE CODES (Order)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: CHECK FOR VALID
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Check to see if discount code has expired
		If expDate<>"" then
			If datediff("d", Now(), expDate) <= 0 Then
				pcv_strDiscountErrorMsg = dictLanguage.Item(Session("language")&"_orderverify_21")
				pDiscountError=1 
			end if
		end if
		
		'// Check to see if discount has start date
		If pcv_startDate<>"" then
			StartDate=pcv_startDate
			If datediff("d", Now(), StartDate) > 0 Then
				pcv_strDiscountErrorMsg = dictLanguage.Item(Session("language")&"_orderverify_43")
				pDiscountError=1 
			End If
		end if
		
		
		If pDiscountError=0 Then
			if (Int(pCartQuantity)>=Int(dcQuantityFrom)) AND (Int(pCartQuantity)<=Int(dcQuantityUntil)) AND (Int(pCartTotalWeight)>=Int(dcWeightFrom)) AND (Int(pCartTotalWeight)<=Int(dcWeightUntil)) AND (ccur(pSubTotal)>=ccur(dcPriceFrom)) AND (ccur(pSubTotal)<=ccur(dcPriceUntil)) then
			else
				'// The discount code or gift certificate that you have entered is no longer valid
				pcv_strDiscountErrorMsg = dictLanguage.Item(Session("language")&"_orderverify_5")
				pDiscountError=1 
			end if
		End If
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: CHECK FOR VALID
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	End If
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// End: Perform Checks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// START: Filter Discount Codes
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	IF pcv_IDDiscount<>"" AND pDiscountError=0 THEN
	
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Filter By Product
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="select pcFPro_IDProduct from PcDFProds where pcFPro_IDDiscount=" & pcv_IDDiscount1
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)		
		if not rs.eof then
			do while not rs.eof
				pcProductList=session("pcCartSession")
				pcv_tmpIDPro=rs("pcFPro_IDProduct")
				for f=1 to session("pcCartIndex")
					if pcProductList(f,10)=0 then
						if (pcProductList(f,32)<>"") then
							pcv_tmpPPrd=split(pcProductList(f,32),"$$")
							pcv_tmpID=pcv_tmpPPrd(ubound(pcv_tmpPPrd))
						else
							pcv_tmpID=pcProductList(f,0)
						end if
						if (pcIncExcPrd="0") AND (clng(pcv_tmpID)=clng(pcv_tmpIDPro)) then
							itemIndex = f-1
							pcProductList(f,2) = session("pcUnitPrice"&itemIndex)						
							pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
						
						elseif (pcIncExcPrd="1") AND (clng(pcv_tmpID)<>clng(pcv_tmpIDPro)) then
							itemIndex = f-1
							pcProductList(f,2) = session("pcUnitPrice"&itemIndex)						
							pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
						
						else
							pcv_Filters=pcv_Filters+1
							pcv_strDiscountErrorMsg="Product(s) not eligible. "						
						end if
					end if
				next
				rs.MoveNext
			loop
		end if
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Filter By Product
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Filter by Categories
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if pcv_ProTotal=0 then
			query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)							
			if not rs.eof then					
				CategoryFilterMatch=0		
				pcProductList=session("pcCartSession")
				for f=1 to session("pcCartIndex")
					if pcProductList(f,10)=0 then						
						query="select idcategory from categories_products where idproduct=" & pcProductList(f,0)
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)											
						pcv_cattest=0						
						do while not rs.eof							
							IF pcv_cattest=0 THEN
								pcv_IDCat=rs("IDCategory")
								query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1 & " and pcFCat_IDCategory=" & pcv_IDCat
								set rstemp=server.CreateObject("ADODB.RecordSet")
								set rstemp=connTemp.execute(query)
								
								if not rstemp.eof then
									set rstemp=nothing
									if pcv_cattest=0 then
										itemIndex = f-1
										pcProductList(f,2) = session("pcUnitPrice"&itemIndex)															
										pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))											
										CategoryFilterMatch=1										
										pcv_cattest=1
									end if
								else
									set rstemp=nothing
									CatCount=0
									CatFound=0
									do while (CatCount<4) and (CatFound=0) and (pcv_IDCat<>"1")
										query="select idParentCategory from categories where idcategory=" & pcv_IDCat
										set rstemp=server.CreateObject("ADODB.RecordSet")
										set rstemp=connTemp.execute(query)												
										if not rstemp.eof then
											pcv_IDCat=rstemp("idParentCategory")
											set rstemp=nothing
										
											CatCount=CatCount+1
											if pcv_IDCat<>"1" then
												query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1 & " and pcFCat_IDCategory=" & pcv_IDCat & " and pcFCat_SubCats=1;"
												set rstemp=server.CreateObject("ADODB.RecordSet")
												set rstemp=connTemp.execute(query)
												if not rstemp.eof then
													if pcv_cattest=0 then														
														if (pcIncExcCat="0") then
															itemIndex = f-1
															pcProductList(f,2) = session("pcUnitPrice"&itemIndex)						
															pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))															
															CategoryFilterMatch=1
															CatFound=1														
															pcv_cattest=1

														end if

													else

														if (pcIncExcCat="1") then
															itemIndex = f-1
															pcProductList(f,2) = session("pcUnitPrice"&itemIndex)						
															pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))															
															CategoryFilterMatch=1
															CatFound=1														
															pcv_cattest=1
														end if
													end if
												end if
												set rstemp=nothing
											end if
										else
											CatCount=4
										end if
									loop
								end if
							END IF
							rs.MoveNext
						loop
						set rs=nothing
					end if
				next
				'// If no Categories Found Set the message.
				if CategoryFilterMatch=0 AND pcv_ProTotal=0 then
					pcv_Filters=pcv_Filters+1
					pcv_strDiscountErrorMsg="This category is not eligible. "	
				end if
			end if
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Filter by Categories
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Filter by Customers
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="select pcFCust_IDCustomer from PcDFCusts where pcFCust_IDDiscount=" & pcv_IDDiscount1
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)							
		if not rs.eof then
			pcv_srtNewIDCustomer=rs("pcFCust_IDCustomer")
			'/////////////////////////////////////////////////////////////////////////////////////////////				
			'// We can not Filter by Customer with Google Checkout, we don't who the customer is yet.
			'// By Default we will accept all customer for the coupon.
			'// To over-ride this behavior you can change the value of "pcv_intApplyCouponToAll" to zero
			'// For example: pcv_intApplyCouponToAll = 0
			'/////////////////////////////////////////////////////////////////////////////////////////////	
			pcv_intApplyCouponToAll = 1					
			
			'if pcv_srtNewIDCustomer=pcv_srtIDCustomer then '// We will never have a match with Google.
			if pcv_intApplyCouponToAll = 1 then				
				pcv_ProTotal=pcv_ProTotal '// Use the current SubTotal, as it could be Product Filtered.
			else
				pcv_Filters=pcv_Filters+1
				pcv_strDiscountErrorMsg="This coupon can not be used with Google Checkout. "
				pcv_ProTotal=0
			end if
			
		end if
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Filter by Customers
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		'// Clear the Filter Errors if we have a positive Discount Total
		if pcv_ProTotal > 0 then
			pcv_Filters=pcv_FResults '// This is bypass the an error message.
		end if

		if pcv_Filters<>pcv_FResults then
			'// The requirements for the coupon that you entered don't match your order details. 
			'// Please review this coupon's requirements and edit your order accordingly.
			pcv_strDiscountErrorMsg=pcv_strDiscountErrorMsg&"Please review the coupon's requirements."
			pDiscountError=1
		end if
		
	END IF	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// END: Filter Discount Codes
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	




	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// START: Calculate Discount Codes
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	if pDiscountError=0 then
		
		pTempPriceToDiscount=pPriceToDiscount
		pTempPercentageToDiscount=pPercentageToDiscount
		pTempIdDiscount=pcv_IDDiscount
		'// Calculate discount. Note: percentage does not affect shipment and payment prices
		if pTempPriceToDiscount>0 or pTempPercentageToDiscount>0 then
			if pcv_ProTotal=0 then
				pcv_ProTotal=pSubTotal
			end if
			tempPercentageToDiscount=(pTempPercentageToDiscount*(pcv_ProTotal)/100)				
			pcv_ProTotal=0
			tempPercentageToDiscount=RoundTo(tempPercentageToDiscount,.01)
			tempDiscountAmount=pTempPriceToDiscount + tempPercentageToDiscount			
			discountTotal=discountTotal + tempDiscountAmount
			pCheckSubtotal=pSubtotal-discountTotal
			if pCheckSubTotal<0 then
				tempDiscountAmount=tempDiscountAmount+pChecksubTotal
			end if
			if intArryCnt=0 then
				discountAmount=tempDiscountAmount
				intArryCnt=intArryCnt+1
			else
				discountAmount=discountAmount&","&tempDiscountAmount
				intArryCnt=intArryCnt+1
			end if		
		else
			pcIntIdShipService=Session(shippingMethod & "_id")
			if pcIntIdShipService<>"" then
				query="select pcFShip_IDShipOpt from pcDFShip where pcFShip_IDDiscount=" & pTempIdDiscount & " and pcFShip_IDShipOpt=" & pcIntIdShipService
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)								
				if not rs.eof then					
					discountAmount=Session(shippingMethod)
					Session("FreeShippingFlag")="1"				
				else
					pcv_strDiscountErrorMsg = "Free Shipping is not available on the selected shipping service using this coupon."
					pDiscountError=1
				end if
				set rs=nothing
			else
				pcv_strDiscountErrorMsg = "Free Shipping is not available on the selected shipping service using this coupon."
				pDiscountError=1
			end if
		end if
		
	else
		'// Set Defaults
	end if	
	if discountAmount="" then
		discountAmount=0
	end if
	discountAmount=money(discountAmount)
	discountAmount=pcf_CurrencyField(discountAmount)
	tempDiscountAmount=0	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// END: Calculate Discount Codes
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
END IF '// pcv_intDiscountsFlag=0
'*************************************************************
' END: THERE ARE DISCOUNTS
'*************************************************************




'*************************************************************
' START: TOTAL THE DISCOUNT
'*************************************************************

'// Check if there was a Gift Cert
if GC="1" then
	query="SELECT pcGCOrdered.pcGO_ExpDate, pcGCOrdered.pcGO_Amount, pcGCOrdered.pcGO_Status, products.Description FROM pcGCOrdered, products WHERE pcGCOrdered.pcGO_GcCode='"& pDiscountCode &"' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	
	IF rs.eof then
		pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4")	
		pcv_strDiscountErrorMsg = "We could not find this Gift Certficate Code."
		pDiscountError=1 '// dictLanguage.Item(Session("language")&"_orderverify_36")
	ELSE
		mTest=0
		pDiscountError=""
		pGCExpDate=rs("pcGO_ExpDate")
		pGCAmount=rs("pcGO_Amount")
		pGCStatus=rs("pcGO_Status")
		pDiscountDesc=rs("Description")
	
		if ccur(pGCAmount)<=0 then
			mTest=1
			'// pDiscountError=dictLanguage.Item(Session("language")&"_msg_3")
			pcv_strDiscountErrorMsg = "This Gift Certficate has an insufficient balance."
			pDiscountError=1 '//dictLanguage.Item(Session("language")&"_orderverify_36")
		end if
		
		if cint(pGCStatus)<>1 then
			mTest=1
			'// pDiscountError=dictLanguage.Item(Session("language")&"_msg_1")
			pcv_strDiscountErrorMsg = "This Gift Certficate Code is not active."
			pDiscountError=1 '//dictLanguage.Item(Session("language")&"_orderverify_36")
		end if
		
		if year(pGCExpDate)<>"1900" then
			if Date()>pGCExpDate then
				mTest=1
				'// pDiscountError=dictLanguage.Item(Session("language")&"_msg_2")
				pcv_strDiscountErrorMsg = "This Gift Certificate has expired."
				pDiscountError=1 '//dictLanguage.Item(Session("language")&"_orderverify_36")
			end if
		end if
	
		if mTest=0 then
			'// Have Available Amount
			GCAmount=pGCAmount
		end if
	END IF
	set rs=nothing
end if
		
				
'// Make sure the price is a valid number
if GCAmount="" then
	GCAmount="0.00"
end if
if NOT isNumeric(GCAmount) then 
	GCAmount="0.00"
end if

'// If there are no errors, then calculate the price
if pDiscountError=0 then
	IF GCAmount="0.00" then
		'// Have Coupon Amount
		pcv_strCodeType="coupon"
		discountTotal=discountAmount		
	ELSE
		pcv_strCodeType="gift-certificate"
		'// Have Gift Certificate Amount
		discountTotal=GCAmount
	END IF
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start: Check Used Discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Session("SFUsedDiscountCodes")="YES"
	Session("SF"&elemCode)="YES"
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// End: Check Used Discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

else
	'// Set Defaults
	if GC="" then
		pcv_strCodeType="coupon"
		discountTotal="0.00"
	else
		pcv_strCodeType="gift-certificate"
		discountTotal="0.00"
		pcv_strDiscountErrorMsg = "We could not locate this Gift Certificate or Coupon."
	end if
end if
'*************************************************************
' END: TOTAL THE DISCOUNT
'*************************************************************




'*************************************************************
' START: CUT OFFS
'*************************************************************
if pcv_strCodeType="gift-certificate" AND Session("SFLast")="coupon" then
	discountTotal="0.00"
	pcv_strDiscountErrorMsg = "Sorry, you can not use a Gift Certificate with a Coupon."
end if

if pcv_strCodeType="coupon" AND Session("SFLast")="gift-certificate" then
	discountTotal="0.00"
	pcv_strDiscountErrorMsg = "Sorry, you can not use a Coupon with a Gift Certificate."
end if
if Session("SFLast")="" then
	Session("SFLast")=pcv_strCodeType
end if
'*************************************************************
' END: CUT OFFS
'*************************************************************

'//////////////////////////////////////////////////////////////
' END: DISCOUNTS
'//////////////////////////////////////////////////////////////

%>


