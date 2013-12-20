<%		'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
err.number=0
Dim rstemp, pIdOrder, pOID, pnValid, pOrderStatus

pnValid=0

pOID=(int(session("GWOrderId"))-scpre)

if pOID = "" then
	pOID = 0
	pnValid=1
end if

if NOT validNum(pOID) then
	pnValid=1
end if

'// Order number not valid
if pnValid=1 then

'// Order number is valid
else
	
	'// Get order status and customer ID
	call openDb()
	
	query = "SELECT orders.idCustomer,orders.orderStatus FROM orders WHERE orders.idOrder =" & pOID
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
	end if 
		
	'// Get the customer ID if the session is empty
	if int(Session("idcustomer")) = 0 then
		Session("idcustomer") = rs("idCustomer")
	end if
	
	'// Start Order Details section
	pIdOrder=pOID
			
	query="SELECT orders.orderDate, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.customerType, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.comments, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_shippingPhone, orders.shippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.idOrder, orders.rmaCredit, orders.ordPackageNum, orders.ord_DeliveryDate, orders.ord_OrderName, orders.ord_VAT,orders.pcOrd_CatDiscounts, orders.paymentDetails, orders.gwAuthCode, orders.gwTransId, orders.paymentCode, orders.pcOrd_GWTotal FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
	
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
		response.redirect "msg.asp?message=35"     
	end if 
						
	Dim pidCustomer, porderDate, pfirstname, plastname,pcustomerCompany, pphone, paddress, pzip, pstate, pcity, pcountryCode, pcomments, pshippingAddress, pshippingState, pshippingCity, pshippingCountryCode, pshippingZip, paddress2, pshippingFullName, pshippingCompany, pshippingAddress2, pshippingPhone, pcustomerType
	
	pidCustomer=Session("idcustomer")
	porderDate=rs("orderDate")
	'porderDate=showdateFrmt(porderDate)
	pfirstname=rs("name")
	plastName=rs("lastName")
	pCustomerName=pfirstname& " " & plastName
	pcustomerCompany=rs("customerCompany")
	pphone=rs("phone")
	pcustomerType=rs("customerType")
	paddress=rs("address")
	pzip=rs("zip")
	pstate=rs("stateCode")
	if pstate="" then
		pstate=rs("state")
	end if
	pcity=rs("city")
	pcountryCode=rs("countryCode")
	pshippingAddress=rs("shippingAddress")
	pshippingState=rs("shippingStateCode")
	if pshippingState="" then
		pshippingState=rs("shippingState")
	end if
	pshippingCity=rs("shippingCity")
	pshippingCountryCode=rs("shippingCountryCode")
	pshippingZip=rs("shippingZip")
	pshippingPhone=rs("pcOrd_shippingPhone")
	pshippingFullName=rs("shippingFullName")
	paddress2=rs("address2")
	pshippingCompany=rs("shippingCompany")
	pshippingAddress2=rs("shippingAddress2")
	pidOrder=rs("idOrder")
	pord_VAT=rs("ord_VAT")
	pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
	if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
		pcv_CatDiscounts="0"
	end if
	pOrd_GWTotal=rs("pcOrd_GWTotal")
	
	query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.unitPrice, ProductsOrdered.xfdetails,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray  "
	'BTO ADDON-S
	If scBTO=1 then
		query=query&", ProductsOrdered.idconfigSession"
	End If
	'BTO ADDON-E
	query=query&", pcPO_GWOpt, pcPO_GWNote, pcPO_GWPrice, products.description, products.sku, orders.total, orders.paymentDetails, orders.taxamount, orders.shipmentDetails, orders.discountDetails, orders.orderstatus,orders.processDate, orders.shipdate, orders.shipvia, orders.trackingNum, orders.returnDate, orders.returnReason, orders.iRewardPoints, orders.iRewardValue, orders.iRewardPointsCustAccrued, orders.taxdetails, orders.dps, orders.pcOrd_CatDiscounts FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idCustomer=" &Session("idcustomer")& " AND orders.idOrder=" &pIdOrder
	
	set rsOrdObj=conntemp.execute(query)		
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsOrdObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if rsOrdObj.eof then
		set rsOrdObj=nothing
		call closeDb()
		response.redirect "msg.asp?message=35"
	end if 

	Dim pidProduct, pquantity, punitPrice, pxfdetails, pidconfigSession, pdescription
	Dim pSku, pcDPs, ptotal, ppaymentDetails,ptaxamount,pshipmentDetails, pdiscountDetails
	Dim pprocessDate, pshipdate, pshipvia, ptrackingNum, preturnDate, preturnReason, piRewardPoints, piRewardValue
	Dim piRewardPointsCustAccrued,ptaxdetails, pOpPrices, rsObjOptions, pRowPrice, count, rsConfigObj,stringProducts
	Dim stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory,i, s,OptPrice,xfdetails, xfarray, q				
	Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
	Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
	Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
	
	
	'**************************************************************
	' START: Line Item Details
	'**************************************************************
	count=0
	tmpFinalRowPriceTotal=0
	do while not rsOrdObj.eof
		tmpFinalRowPrice=0
		pidProduct=rsOrdObj("idProduct")
		pquantity=rsOrdObj("quantity")
		punitPrice=rsOrdObj("unitPrice")
		QDiscounts=rsOrdObj("QDiscounts")
		ItemsDiscounts=rsOrdObj("ItemsDiscounts")
		
		'BTO ADDON-S
		if scBTO=1 then
			pidconfigSession=rsOrdObj("idconfigSession")
			if pidconfigSession="" then
				pidconfigSession="0"
			end if
		End If
		'BTO ADDON-E

		pdescription=rsOrdObj("description")
		pSku=rsOrdObj("sku")
		ptotal=rsOrdObj("total")
		ppaymentDetails=trim(rsOrdObj("paymentDetails"))
		ptaxamount=rsOrdObj("taxamount")
		pshipmentDetails=rsOrdObj("shipmentDetails")
		pdiscountDetails=rsOrdObj("discountDetails")
		piRewardValue=rsOrdObj("iRewardValue")
		ptaxdetails=rsOrdObj("taxdetails")		
		pCatDiscounts=rsOrdObj("pcOrd_CatDiscounts")
		pIdConfigSession=trim(pidconfigSession)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Row Price
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'BTO ADDON-S
		TotalUnit=0
		If scBTO=1 then
			pIdConfigSession=trim(pidconfigSession)
			if pIdConfigSession<>"0" then 
				query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
				set rsConfigObj=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsConfigObj=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				stringProducts=rsConfigObj("stringProducts")
				stringValues=rsConfigObj("stringValues")
				stringCategories=rsConfigObj("stringCategories")
				stringQuantity=rsConfigObj("stringQuantity")
				stringPrice=rsConfigObj("stringPrice")
				ArrProduct=Split(stringProducts, ",")
				ArrValue=Split(stringValues, ",")
				ArrCategory=Split(stringCategories, ",")
				ArrQuantity=Split(stringQuantity, ",")
				ArrPrice=Split(stringPrice, ",")
				set rsConfigObj=nothing
				for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
					set rsConfigObj=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsConfigObj=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if NOT validNum(ArrQuantity(i)) then
						pIntQty=1
					else
						pIntQty=ArrQuantity(i)
					end if
					if (CDbl(ArrValue(i))<>0) or (((ArrQuantity(i)-1)*pQuantity>0) and (ArrPrice(i)>0)) then
						if (ArrQuantity(i)-1)>=0 then
							UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
						else
							UPrice=0
						end if
						TotalUnit=TotalUnit+((ArrValue(i)+UPrice)*pQuantity)
					end if
					set rsConfigObj=nothing
				next
			end if 
		End If 
		'BTO ADDON-E

		if TotalUnit>0 then
			punitPrice1=punitPrice
			if pIdConfigSession<>"0" then
				pRowPrice1=Cdbl(pquantity * punitPrice1) - QDiscounts - ItemsDiscounts
				punitPrice1=Round(pRowPrice1/pquantity,2)
			else
				pRowPrice1=Cdbl(pquantity * punitPrice1)
			end if
		else
			punitPrice1=punitPrice
			if pIdConfigSession<>"0" AND pIdConfigSession<>"" then
				pRowPrice1=Cdbl(pquantity * punitPrice1) - QDiscounts - ItemsDiscounts
			else
				pRowPrice1=Cdbl(pquantity * punitPrice1)
				punitPrice1=Round(pRowPrice1/pquantity,2)
			end if
		end if
	
		'// Final Row Price
		tmpFinalRowPrice = money(pRowPrice1)
		pcv_strUnitRowPrice = money(pRowPrice1/pquantity) '// Use the money function to synch with the Invoice.	
		tmpUnitRowPrice = pcv_strUnitRowPrice
		tmpUnitRowPrice = pcf_CurrencyField(tmpUnitRowPrice)
		'response.Write tmpUnitRowPrice
		'response.End()
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Row Price
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		If pshippingAddress="" then
			'// Grab shipping address from shipping...
			pshippingAddress=pAddress
			pshippingAddress2=pAddress2
			pshippingCity=pCity
			pshippingState=pState
			pshippingZip=pZip
			pshippingCountryCode=pCountryCode
		End if
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Add Line Item to Array
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
		if count=0 then
			count=count+1
		end if
		
		pdescription = replace(pdescription, """","")
		pdescription = replace(pdescription, "&quot;","")
		pdescription = replace(pdescription, ":","")
		pSKU = replace(pSKU, """","")
		pSKU = replace(pSKU, ":","")
		pSKU = replace(pSKU, "&quot;","")
		prdString=prdString&"|"&tmpUnitRowPrice&"::"&pquantity&"::"&pSKU&"::"&pdescription&".::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Add Line Item to Array
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		count=count+1
		tmpFinalRowPriceTotal=tmpFinalRowPriceTotal+tmpFinalRowPrice
		rsOrdObj.movenext  
	loop
	'**************************************************************
	' END: Line Item Details
	'**************************************************************
	


	'**************************************************************
	' START: Processing Charges
	'**************************************************************
	dim payment, PaymentType, PayCharge
	payment = split(ppaymentDetails,"||")
	err.clear
	PaymentType=payment(0)
	If payment(1)="" then
		if err.number<>0 then
			PayCharge=0
		end if
		PayCharge=0
	else
		PayCharge=payment(1)
	end If
	err.number=0				
	if PayCharge>0 then
		pcv_strFinalPayCharge = money(PayCharge)  '// Use the money function to synch with the Invoice.
		tmpFinalPayCharge = pcv_strFinalPayCharge
		tmpFinalPayCharge = pcf_CurrencyField(tmpFinalPayCharge)			
		count=count+1
	Else
		pcv_strFinalPayCharge = money(0)			
	end if
	'**************************************************************
	' START: Processing Charges
	'**************************************************************
	
	
	'response.Write pcf_CurrencyField(pcv_strFinalPayCharge)
	if pcv_strFinalPayCharge<>0 then
		prdString=prdString&"|"&pcv_strFinalPayCharge&"::1::PC::Processing Charges.::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if

	'**************************************************************
	' START: Category Discounts
	'**************************************************************
	Dim pcv_strFinalCatDiscounts
	if pCatDiscounts>0 then
		pcv_strFinalCatDiscounts = money(pCatDiscounts) '// Use the money function to synch with the Invoice.
		tmpFinalCatDiscounts = pcv_strFinalCatDiscounts
		tmpFinalCatDiscounts = pcf_CurrencyField(tmpFinalCatDiscounts)
		tmpFinalCatDiscounts = "-" & tmpFinalCatDiscounts
		'response.Write(tmpFinalCatDiscounts)
		'response.End()
		
		count=count+1
	Else
		pcv_strFinalCatDiscounts = money(0)			
	end if
	'**************************************************************
	' END: Category Discounts
	'**************************************************************
	
	
	'response.Write -(pcf_CurrencyField(pcv_strFinalCatDiscounts))
	if pcv_strFinalCatDiscounts<>0 then
		prdString=prdString&"|-"&pcv_strFinalCatDiscounts&"::1::CD::Category Discounts.::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if

'Discounts - discountDetails
if instr(pdiscountDetails,",") then
	pcvDiscountDetailsArry=split(pdiscountDetails,",")
	pcvGoTo = ubound(pcvDiscountDetailsArry)
else
	pcvGoTo = 0
	pcvDiscountDetails=pdiscountDetails
end if
pcv_strFinalMDiscountTotal=0
for pdisc=0 to pcvGoTo
	if pcvGoTo<>0 then
		pcvDiscountDetails=pcvDiscountDetailsArry(pdisc)
	end if
	if instr(pcvDiscountDetails,"- ||") then 
		pMDiscounts= split(pcvDiscountDetails,"- ||")
		pMDiscountAmt=trim(pMDiscounts(1))
		if NOT isNumeric(pMDiscountAmt) then
			pMDiscountAmt=0
		end if
	Else
		pMDiscountAmt=0
	end if

	err.number=0				
	if pMDiscountAmt>0 then
		pcv_strFinalMDiscount = money(pMDiscountAmt)  '// Use the money function to synch with the Invoice.
		tmpFinalMDiscountCharge = pcv_strFinalMDiscount
		tmpFinalMDiscountCharge = pcf_CurrencyField(tmpFinalMDiscountCharge)			
		count=count+1
	Else
		tmpFinalMDiscountCharge = money(0)			
	end if
	if tmpFinalMDiscountCharge<>0 then
		prdString=prdString&"|-"&tmpFinalMDiscountCharge&"::1::MD"&pdisc&"::Discount.::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if
	pcv_strFinalMDiscountTotal=pcv_strFinalMDiscountTotal+pcv_strFinalMDiscount
next

'Reward Points Value - iRewardValue
	'response.Write -(pcf_CurrencyField(piRewardValue))
	if piRewardValue<>0 then
		prdString=prdString&"|-"&piRewardValue&"::1::RP::Points Discounts.::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if

'Gift Wrapping Charges - pcOrd_GWTotal
	'response.Write -(pcf_CurrencyField(pOrd_GWTotal))
	if pOrd_GWTotal<>0 then
		prdString=prdString&"|"&pOrd_GWTotal&"::1::GW::Gift Wrapping.::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if

	'**************************************************************
	' START: Shipping Charges
	'**************************************************************
	dim shipping, varShip, Shipper, Service, Postage, serviceHandlingFee		
	shipping=split(pshipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
			varShip="0"
			response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
		else
			Shipper=shipping(0)
			Service=shipping(1)
			Postage=trim(shipping(2))
			if ubound(shipping)=>3 then
				serviceHandlingFee=trim(shipping(3))
				if NOT isNumeric(serviceHandlingFee) then
					serviceHandlingFee=0
				end if
			else
				serviceHandlingFee=0
			end if
		end if
	else
		varShip="0"
	end if 						
	if varShip<>"0" then
		pcv_strFinalShipCharge = Postage
	end if
	
	if pcv_strFinalShipCharge>0 then
		pcv_strFinalShipCharge = money(pcv_strFinalShipCharge) '// Use the money function to synch with the Invoice.		
	 else
		pcv_strFinalShipCharge = money(0)			
	end if 
	'**************************************************************
	' END: Shipping Charges
	'**************************************************************

	'response.Write pcf_CurrencyField(pcv_strFinalShipCharge)
	Service = replace(Service, """","")
	Service = replace(Service, ":","")
	if pcv_strFinalShipCharge<>0 then
		prdString=prdString&"|"&pcv_strFinalShipCharge&"::1::SH::"&Service&"::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if 

	'**************************************************************
	' START: Service Handling Charges
	'**************************************************************
	if serviceHandlingFee>0 then
		pcv_strFinalServiceCharge = money(serviceHandlingFee) '// Use the money function to synch with the Invoice.		
	 else
		pcv_strFinalServiceCharge = money(0)			
	end if 
	'**************************************************************
	' END: Service Handling Charges
	'**************************************************************

	'response.Write pcf_CurrencyField(pcv_strFinalServiceCharge)
	if pcv_strFinalServiceCharge<>0 then
		prdString=prdString&"|"&pcv_strFinalServiceCharge&"::1::HC::Handling::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if

	'**************************************************************
	' START: Tax Charges
	'**************************************************************
	' If the store is using VAT and VAT is> 0, don't show any taxes here, but show VAT after the total
	pcv_strFinalTax = 0
	if pord_VAT>0 then
	else
		if isNull(ptaxDetails) or trim(ptaxDetails)="" then
			pcv_strFinalTax = money(ptaxAmount)
		else
			dim taxArray, taxDesc
			taxArray=split(ptaxDetails,",")
			for i=0 to (ubound(taxArray)-1)
				taxDesc=split(taxArray(i),"|") 
				pcv_strFinalTax = cCur(pcv_strFinalTax) + cCur(money(taxDesc(1)))
			next
		end if 
	end if
	' If the store is using VAT and VAT> 0, show it here
	if pord_VAT>0 then
		pcv_strFinalTax = cCur(pcv_strFinalTax) + cCur(money(pord_VAT))
	end if
	
	if pcv_strFinalTax>0 then
		pcv_strFinalTax = money(pcv_strFinalTax) '// Use the money function to synch with the Invoice.			
	 else
		pcv_strFinalTax = money(0)			
	end if 
	'**************************************************************
	' START: Tax Charges
	'**************************************************************

	'response.Write pcf_CurrencyField(pcv_strFinalTax)
	if pcv_strFinalTax<>0 then
		prdString=prdString&"|"&pcv_strFinalTax&"::1::TX::Taxes::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if

	'**************************************************************
	' START: Order Total
	'**************************************************************
	pcv_strFinalTotal = money(ptotal) '// Use the money function to synch with the Invoice.
	'**************************************************************
	' END: Order Total
	'**************************************************************

	
	'// Everything is currently in the "Money Display" formatting so that it's identical to the Order Confirmation.
	'// We need to convert the "Money Display" formatting into something that can be used in a calculation.
	pcv_strFinalTotal=cCur(pcf_CurrencyField(pcv_strFinalTotal))
	pcv_strFinalShipCharge=cCur(pcf_CurrencyField(pcv_strFinalShipCharge))
	pcv_strFinalServiceCharge=cCur(pcf_CurrencyField(pcv_strFinalServiceCharge))
	pcv_strFinalTax=cCur(pcf_CurrencyField(pcv_strFinalTax))
	pcv_strFinalPayCharge=cCur(pcf_CurrencyField(pcv_strFinalPayCharge))
	
	'// Perform the "ItemTotal" Calculation
	ItemTotal = pcv_strFinalTotal - (pcv_strFinalShipCharge+pcv_strFinalServiceCharge+pcv_strFinalTax+pcv_strFinalPayCharge-pcv_strFinalMDiscountTotal)
	pcv_ItemizedTotal=(tmpFinalRowPriceTotal+pcv_strFinalPayCharge-pcv_strFinalCatDiscounts-pcv_strFinalMDiscountTotal-piRewardValue+pOrd_GWTotal+pcv_strFinalShipCharge+pcv_strFinalServiceCharge+pcv_strFinalTax)
	if pcv_ItemizedTotal>pcv_strFinalTotal then
		'There must be a discount in bundles or otherwise that isn't saved in the database
		pcv_ExtraDiscount=pcv_ItemizedTotal-pcv_strFinalTotal
		prdString=prdString&"|-"&pcv_ExtraDiscount&"::1::OD::Other Discounts::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if
	if pcv_ItemizedTotal<pcv_strFinalTotal then
		'There must be an extra charge that isn't saved in the database
		pcv_ExtraCharges=pcv_strFinalTotal-pcv_ItemizedTotal
		prdString=prdString&"|"&pcv_ExtraCharges&"::1::OC::Other Charges::"
		if IsTestmode="1" then
			prdString=prdString&"{TEST}{TESTD}"
		end if
	end if
end if '// End if order number is valid

'// Close Db connections
call closeDB()
'// Format For Field
Public Function pcf_CurrencyField(moneyAMT)	
if scDecSign = "," then
	moneyAMT=replace(moneyAMT,".","")
	moneyAMT=replace(moneyAMT,",",".")		
else
	moneyAMT=replace(moneyAMT,",","")
end if
pcf_CurrencyField=moneyAMT
End Function
%>