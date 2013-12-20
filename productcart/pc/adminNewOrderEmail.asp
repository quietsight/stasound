<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions

'create body of email
storeAdminEmail=Cstr("")

storeAdminEmail=""
storeAdminEmail=storeAdminEmail & vbCrLf & dictLanguage.Item(Session("language")&"_adminMail_1") & (scpre + int(pIdOrder)) & vbCrLf &vbCrLf 

' Order summary starts here ...
storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_2") & vbCrlf
storeAdminEmail=storeAdminEmail & "===================" & vbCrlf
storeAdminEmail=storeAdminEmail & pName & " " & pLName & vbCrLf

If Trim(pCustomerCompany) <> "" Then
	storeAdminEmail=storeAdminEmail & pCustomerCompany & vbCrLf
End If
storeAdminEmail=storeAdminEmail & paddress & vbCrLf
if paddress2 <> "" then
	storeAdminEmail=storeAdminEmail & paddress2 & vbCrLf
end if
storeAdminEmail=storeAdminEmail & pCity & ", "
storeAdminEmail=storeAdminEmail & pStateCode & " "
storeAdminEmail=storeAdminEmail & pzip & vbCrLf
storeAdminEmail=storeAdminEmail & pCountryCode & vbCrLf
storeAdminEmail=storeAdminEmail & pPhone & vbCrLf
storeAdminEmail=storeAdminEmail & pEmail & vbCrLf

'Start Special Customer Fields
session("sf_nc_custfields")=""
session("pcSFCustFieldsExist")=""
query="SELECT pcCField_ID,pcCField_Name,pcCField_FieldType,pcCField_Value,pcCField_Length,pcCField_Maximum,pcCField_Required,pcCField_PricingCategories,pcCField_ShowOnReg,pcCField_ShowOnCheckout,'',pcCField_Description FROM pcCustomerFields;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if not rs.eof then
	session("pcSFCustFieldsExist")="YES"
	session("sf_nc_custfields")=rs.GetRows()
end if
set rs=nothing

if session("pcSFCustFieldsExist")="YES" AND Session("idCustomer")<>0 then
	pcArr=session("sf_nc_custfields")
	For k=0 to ubound(pcArr,2)
		pcArr(10,k)=""
		query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & Session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.eof then
			pcArr(10,k)=rs("pcCFV_Value")
		end if
		set rs=nothing
		if trim(pcArr(10,k))<>"" then
			storeAdminEmail=storeAdminEmail & pcArr(1,k) & ": " & pcArr(10,k) & VBCRLF
		end if
	Next
	session("sf_nc_custfields")=pcArr
end if
'End of Special Customer Fields

if Session("pcSFCRecvNews")="1" then
	storeAdminEmail=storeAdminEmail & "Signed up for the store newsletter" & vbCrLf
end if

if (Session("pcSFIDrefer")<>"") and (Session("pcSFIDrefer")<>"0") then
	query="select [name] from Referrer where IDRefer=" & Session("pcSFIDrefer")
	set rstempObj=connTemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rstempObj=nothing
		'//close any connections
		call closedb()
		'//redirect to error page
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if not rstempObj.eof then
		pRefer=dictLanguage.Item(Session("language")&"_adminMail_3") & rstempObj("name") & VBCRLF
	end if
	set rstempObj=nothing
	storeAdminEmail=storeAdminEmail & pRefer & vbCrLf & vbCrLf
else
	storeAdminEmail=storeAdminEmail & vbCrLf
end if

dim rsShipTypeObj
query="select ordShipType from Orders where idOrder=" & pIDOrder
set rsShipTypeObj=server.CreateObject("ADODB.RecordSet")
set rsShipTypeObj=connTemp.execute(query)

pOrdShipType=rsShipTypeObj("ordShipType")
set rsShipTypeObj=nothing

if pOrdShipType=0 then
	pTempDisShipType="Residential"
else
	pTempDisShipType="Commercial" 
end if

storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_4") & vbCrlf
storeAdminEmail=storeAdminEmail & "====================" & vbCrlf
If Trim(pshippingAddress) <> "" Then
	if pShippingFullName<>"" then
		storeAdminEmail=storeAdminEmail & pShippingFullName& vbCrLf
	end if
	if pShippingCompany<>"" then
		storeAdminEmail=storeAdminEmail & pShippingCompany& vbCrLf
	end if
	storeAdminEmail=storeAdminEmail & pshippingAddress & vbCrLf
	if pshippingAddress2<>"" then
		storeAdminEmail=storeAdminEmail & pshippingAddress2 & vbCrLf
	end if
	storeAdminEmail=storeAdminEmail & pshippingCity & ", "
	storeAdminEmail=storeAdminEmail & pshippingStateCode & " "
	storeAdminEmail=storeAdminEmail & pshippingZip & vbCrLf
	storeAdminEmail=storeAdminEmail & pshippingCountryCode & vbCrLf
	storeAdminEmail=storeAdminEmail & trim(pshippingPhone) & vbCrLf
Else
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_5") & vbCrLf
End if 
storeAdminEmail=storeAdminEmail & vbCrLf
storeAdminEmail=storeAdminEmail & "Shipping Type: " & pTempDisShipType & vbCrLf
storeAdminEmail=storeAdminEmail & vbCrLf

'shipping details
shipping=split(pshipmentDetails,",")
if ubound(shipping)>1 then
	if NOT isNumeric(shipping(2)) then
		storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_6") & pShipmentDetails & vbCrLf
		Service=""
		Postage=0
	else
		Shipper=shipping(0)
		Service=shipping(1)
		Postage=shipping(2)
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
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_6") & pShipmentDetails & vbCrLf
	Service=""
	Postage=0
end if

If DFShow="1" Then
	storeAdminEmail=storeAdminEmail & DFLabel & " " & pord_DeliveryDate & vbCrlf
End If

if Service<>"" then
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_7") & Service & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_35") & pOrdPackageNum & vbCrLf
end if
storeAdminEmail=storeAdminEmail & vbCrLf

'offline payment details
paymentdetails=split(trim(pPaymentDetails),"||")
if instr(PaymentType,"FREE") AND len(PaymentType)<6 then
	paymentCharge=0
else
	storeAdminEmail = storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_8") & paymentdetails(0) & vbCrLf
	
	'// Custom Field "paymnta_c.asp"
	if Session("pcSFpAccNum2") <> "" then
		storeAdminEmail = storeAdminEmail & trim(paymentdetails(0)) & ": " & Session("pcSFpAccNum2") & vbCrLf		
	end if
	Session("pcSFpAccNum2")=""
	
	'// Customer Field "paymnta_customcard.asp"
	if Session("pcSFSpecialFields") <> "" then
		storeAdminEmail = storeAdminEmail & Session("pcSFSpecialFields")
	end if
	Session("pcSFSpecialFields")=""
	
	if ubound(paymentdetails)>0 then
		paymentCharge=trim(paymentdetails(1))
		If NOT isNumeric(paymentCharge) then
			paymentCharge=0
		End if
	else
		paymentCharge=0
	end if
	storeAdminEmail = storeAdminEmail & vbCrLf
end if

'GGG Add-on start

query="select pcOrd_IDEvent,pcOrd_GWTotal from Orders where idOrder=" & pIDOrder
set rs19=connTemp.execute(query)

gIDEvent=rs19("pcOrd_IDEvent")
if gIDEvent<>"" then
else
gIDEvent="0"
end if

pGWTotal=rs19("pcOrd_GWTotal")
if pGWTotal<>"" then
else
pGWTotal="0"
end if

if gIDEvent<>"0" then
	query="select pcEvents.pcEv_name,pcEvents.pcEv_Date,customers.name,customers.lastname from pcEvents,Customers where Customers.idcustomer=pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
	set rs1=connTemp.execute(query)
	
	geName=rs1("pcEv_name")
	geDate=rs1("pcEv_Date")
	if year(geDate)="1900" then
		geDate=""
	end if
	if gedate<>"" then
		if scDateFrmt="DD/MM/YY" then
			gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
		else
			gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
		end if
	end if
	gReg=rs1("name") & " " & rs1("lastname")
	
	storeAdminEmail=storeAdminEmail & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_36") & geName & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_37") & geDate & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_38") & gReg & vbCrLf
	storeAdminEmail=storeAdminEmail & vbCrLf
end if
'GGG Add-on end

'discount details
'Check if more then one discount code was utilized
if instr(pdiscountDetails,",") then
	DiscountDetailsArry=split(pdiscountDetails,",")
	intArryCnt=ubound(DiscountDetailsArry)
else
	intArryCnt=0
end if
pTotalDiscountAmount=0
for k=0 to intArryCnt
	if intArryCnt=0 then
		pTempDiscountDetails=pdiscountDetails
	else
		pTempDiscountDetails=DiscountDetailsArry(k)
	end if
	if instr(pTempDiscountDetails,"- ||") then 
		discounts= split(pTempDiscountDetails,"- ||")
		pdiscountDesc=discounts(0)
		pdiscountAmt=trim(discounts(1))
		pIsNumeric=1
		if NOT isNumeric(pdiscountAmt) then
			pdiscountAmt=0
			pIsNumeric=0
		end if
		if (pdiscountAmt>0 OR pdiscountAmt=0) AND pIsNumeric=1 then
			storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_9") & pdiscountDesc & vbCrLf
		end if
	Else
		pdiscountAmt=0
	end if
	pTotalDiscountAmount=pTotalDiscountAmount+pdiscountAmt
Next

'GGG Add-on start
query="select pcOrd_GCDetails,pcOrd_GCAmount from Orders where idOrder=" & pIDOrder
set rs19=connTemp.execute(query)

GCDetails=rs19("pcOrd_GCDetails")
GCAmountTotal=rs19("pcOrd_GCAmount")
if GCAmountTotal="" OR IsNull(GCAmountTotal) then
	GCAmountTotal=0
end if

if GCDetails<>"" then
GCArr=split(GCDetails,"|g|")
intGCCount=ubound(GCArr)
For y=0 to intGCCount
if GCArr(y)<>"" then
GCInfo=split(GCArr(y),"|s|")
query="select Products.Description from pcGCOrdered,Products where pcGCOrdered.pcGO_GcCode like '" & GCInfo(0) & "' and products.idproduct=pcGCOrdered.pcGO_idproduct"
set rs19=connTemp.execute(query)

	pGCName=rs19("Description")
	pdiscountAmt=cdbl(GCInfo(2))
	if pdiscountAmt>0 then
		storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_9a") & pGCName & " (" & GCInfo(0) & ")" & vbCrLf
	end if
end if
Next
end if
'GGG Add-on end

'Reward points used or accrued on this order?
iDollarValue=0

If (int(piRewardValue) > 0) OR (int(piRewardPointsCustAccrued) > 0) then
	storeAdminEmail=storeAdminEmail & vbCrLf
	'Did we use points or accrue points?
	If int(piRewardPointsCustAccrued) > 0 Then 'Accrued
		iDollarValue=int(piRewardPointsCustAccrued) * (RewardsPercent / 100)
		storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_10") & int(piRewardPointsCustAccrued) & " " & RewardsLabel & dictLanguage.Item(Session("language")&"_adminMail_11") & scCurSign&money(iDollarValue) & vbCrLf
	End If
	If int(piRewardValue) > 0 Then
		storeAdminEmail=storeAdminEmail & scCurSign&money(piRewardValue) & dictLanguage.Item(Session("language")&"_adminMail_12") & RewardsLabel & "." & vbCrLf
	End If			
	storeAdminEmail=storeAdminEmail & vbCrLf
End If 

'order details
storeAdminEmail=storeAdminEmail & vbCrLf

storeAdminEmail=storeAdminEmail & FixedField(20, "L", dictLanguage.Item(Session("language")&"_adminMail_13"))
storeAdminEmail=storeAdminEmail & FixedField(40, "R", dictLanguage.Item(Session("language")&"_adminMail_14"))
storeAdminEmail=storeAdminEmail & FixedField(10, "R", dictLanguage.Item(Session("language")&"_adminMail_15"))
storeAdminEmail=storeAdminEmail & vbCrLf

storeAdminEmail=storeAdminEmail & FixedField(50, "R", "==================================================")
storeAdminEmail=storeAdminEmail & FixedField(10, "R", "==========")
storeAdminEmail=storeAdminEmail & FixedField(10, "R", "==========")
storeAdminEmail=storeAdminEmail & vbCrLf
iSubtotal=0

query="SELECT products.idproduct,products.sku, products.description, ProductsOrdered.pcSC_ID, quantity, unitPrice, xfdetails"
'BTO ADDON-S
if scBTO=1 then
	query=query&" ,idconfigSession"
end if
'BTO ADDON-E
query=query&", ProductsOrdered.QDiscounts, ProductsOrdered.ItemsDiscounts, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.pcPO_GWOpt, ProductsOrdered.pcPO_GWNote, ProductsOrdered.pcPO_GWPrice, ProductsOrdered.pcPrdOrd_BundledDisc FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& pIdOrder

set rsOrderDetails=conntemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rsOrderDetails=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
	
Do While Not rsOrderDetails.EOF
	pidProduct=rsOrderDetails("idproduct")
	psku=rsOrderDetails("sku")
	pdescription=rsOrderDetails("description")
	pdescription=ClearHTMLTags2(pdescription,0)
	pcSCID=rsOrderDetails("pcSC_ID")
	if IsNull(pcSCID) OR len(pcSCID)=0 then
		pcSCID=0
	end if
	pqty=rsOrderDetails("quantity")
	pPrice=rsOrderDetails("unitPrice")
	xfdetails=replace(rsOrderDetails("xfdetails"),"&lt;BR&gt;",vbcrlf)
	xfdetails=replace(xfdetails,"<BR>",vbcrlf)
	if scBTO=1 then
		pIdConfigSession=rsOrderDetails("idconfigSession")
	end if
	QDiscounts=rsOrderDetails("QDiscounts")
	ItemsDiscounts=rsOrderDetails("ItemsDiscounts")	
	
	'// Product Options Arrays
	pcv_strSelectedOptions = rsOrderDetails("pcPrdOrd_SelectedOptions") ' Column 11
	pcv_strOptionsPriceArray = rsOrderDetails("pcPrdOrd_OptionsPriceArray") ' Column 25
	pcv_strOptionsArray = rsOrderDetails("pcPrdOrd_OptionsArray") ' Column 4
	
	'GGG Add-on start	
	pGWOpt=rsOrderDetails("pcPO_GWOpt")
	if pGWOpt<>"" then
	else
	pGWOpt="0"
	end if 
	pGWText=rsOrderDetails("pcPO_GWNote")
	pGWPrice=rsOrderDetails("pcPO_GWPrice")
	if pGWPrice<>"" then
	else
	pGWPrice="0"
	end if
	'GGG Add-on end
	pcPrdOrd_BundledDisc=rsOrderDetails("pcPrdOrd_BundledDisc")

	pExtendedPrice=pPrice*pqty
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_16")&pqty & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_17")&psku & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_18")&pdescription & vbCrLf
	storeAdminEmail=storeAdminEmail & "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" & vbcrlf
	'BTO ADDON-S
	TotalUnit=0
	if scBTO=1 then
		if pIdConfigSession<>"0" then
			query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsConfigObj=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsConfigObj=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
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
			for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
				query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pidProduct 
				set rsQ=server.CreateObject("ADODB.RecordSet") 
				set rsQ=conntemp.execute(query)										
				btDisplayQF=rsQ("displayQF")
				set rsQ=nothing
											
				query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
				set rsQ=connTemp.execute(query)
				tmpMinQty=1
				if not rsQ.eof then
					tmpMinQty=rsQ("pcprod_minimumqty")
					if IsNull(tmpMinQty) or tmpMinQty="" then
						tmpMinQty=1
					else
						if tmpMinQty="0" then
							tmpMinQty=1
						end if
					end if
				end if
				set rsQ=nothing
				tmpDefault=0
				query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					tmpDefault=rsQ("cdefault")
					if IsNull(tmpDefault) or tmpDefault="" then
						tmpDefault=0
					else
						if tmpDefault<>"0" then
						 	tmpDefault=1
						end if
					end if
				end if
				set rsQ=nothing
				
				query="SELECT products.sku, categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
				set rsConfigObj=conntemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rsConfigObj=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				pcv_strBtoItemSku = rsConfigObj("sku")
				pcv_strBtoItemSku=ClearHTMLTags2(pcv_strBtoItemSku,0)
				pcv_strBtoItemName = rsConfigObj("description")
				pcv_strBtoItemName=ClearHTMLTags2(pcv_strBtoItemName,0)
				pcv_strBtoItemCat=rsConfigObj("categoryDesc")
				pcv_strBtoItemCat=ClearHTMLTags2(pcv_strBtoItemCat,0)
				storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
				dispStr = ""
				dispStr = pcv_strBtoItemCat &": "& pcv_strBtoItemName
				dispStr = dispStr & " - SKU: " & pcv_strBtoItemSku
				if btDisplayQF=True then
					if clng(ArrQuantity(i))>1 then
						dispStr = dispStr & " - QTY: " & ArrQuantity(i)
					end if
				end if
				dispStr = replace(dispStr,"&quot;", chr(34))
				tStr = dispStr
				wrapPos=50
				if len(dispStr) > 50 then
					tStr = WrapString(50, dispStr)
				end if
				storeAdminEmail=storeAdminEmail & FixedField(50, "L", tStr)

				if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
					if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
						if tmpDefault=1 then
							UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
						else
							UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
						end if
					else
						UPrice=0
					end if
					TotalUnit=TotalUnit+cdbl((ArrValue(i)+UPrice)*pQty)
					storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money((ArrValue(i)+UPrice)*pQty))
				else
					if tmpDefault=1 then
						storeAdminEmail=storeAdminEmail & FixedField(10, "R", dictLanguage.Item(Session("language")&"_defaultnotice_1"))
					end if
				end if
				storeAdminEmail=storeAdminEmail & vbCrLf
				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 50
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(50, dispStr)
					storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
					storeAdminEmail=storeAdminEmail  & FixedField(50, "L", tStr)
					storeAdminEmail=storeAdminEmail  & vbCrLf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					storeAdminEmail=storeAdminEmail  & FixedField(10, "L", "")
					storeAdminEmail=storeAdminEmail  & FixedField(50, "L", dispStr)
					storeAdminEmail=storeAdminEmail  & vbCrLf
				end if
				set rsConfigObj=nothing
			next
		end if
	end if
	'BTO ADDON-E
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Add first 50 characters of options on a separate line
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
		pcv_strSelectedOptions = ""
	end if
	
	If len(pcv_strSelectedOptions)>0 Then
			'// Add the header "OPTIONS"		
			storeAdminEmail=storeAdminEmail & FixedField(10, "L", "OPTIONS")
			
			'#####################
			' START LOOP
			'#####################				
			'// Generate Our Local Arrays from our Stored Arrays  
			
			' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
			pcArray_strSelectedOptions = ""					
			pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))
			
			' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
			pcArray_strOptionsPrice = ""
			pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))
			
			' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
			pcArray_strOptions = ""
			pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))
			
			' Get Our Loop Size
			pcv_intOptionLoopSize = 0
			pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)

			' Display Our Options
			For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
				dispStr = ""
							
				'//There isnt a header after the first one so we indent
				if pcv_intOptionLoopCounter >0 then
					storeAdminEmail=storeAdminEmail & FixedField(10, "L", " ")
				end if
			
				dispStr = pcArray_strOptions(pcv_intOptionLoopCounter)
				dispStr = replace(dispStr,"&quot;", chr(34))
				tStr = dispStr
				wrapPos=50
				if len(dispStr) > 50 then
					tStr = WrapString(50, dispStr)
				end if
				storeAdminEmail=storeAdminEmail & FixedField(50, "L", tStr)

				tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
													
				if tempPrice="" or tempPrice=0 then
					storeAdminEmail=storeAdminEmail & FixedField(10, "R", " ")
					storeAdminEmail=storeAdminEmail & vbCrLf
				else 
					storeAdminEmail=storeAdminEmail & FixedField(10, "R", "")
					storeAdminEmail=storeAdminEmail & vbCrLf
				end if
				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 50
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(50, dispStr)
					storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
					storeAdminEmail=storeAdminEmail  & FixedField(50, "L", tStr)
					storeAdminEmail=storeAdminEmail  & vbCrLf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					storeAdminEmail=storeAdminEmail  & FixedField(10, "L", "")
					storeAdminEmail=storeAdminEmail  & FixedField(50, "L", dispStr)
					storeAdminEmail=storeAdminEmail  & vbCrLf
				end if
				
			Next
			'#####################
			' END LOOP
			'#####################	
			
			storeAdminEmail=storeAdminEmail & vbCrLf
			
	End If
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Add first 50 characters of options on a separate line
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	
	If len(xfdetails)>3 then
		xfarray=split(xfdetails,"|")
		for q=lbound(xfarray) to ubound(xfarray)
			storeAdminEmail=storeAdminEmail & xfarray(q)
			storeAdminEmail=storeAdminEmail & vbCrLf & vbcrlf
		next
		storeAdminEmail=storeAdminEmail & vbCrLf
	End If
	
	pPrice1=pPrice
	pExtendedPrice1=pExtendedPrice
	
	if TotalUnit>0 then
		pExtendedPrice1=pExtendedPrice1-TotalUnit
		pPrice1=Round(pExtendedPrice1/pqty,2)
	end if	

	tmpText1=""
	if pcSCID>"0" then
		tmpText1=tmpText1 & FixedField(50, "L", dictLanguage.Item(Session("language")&"_adminMail_19S"))
	else
		tmpText1=tmpText1 & FixedField(50, "L", dictLanguage.Item(Session("language")&"_adminMail_19"))
	end if
	if money(pPrice1)=money(pExtendedPrice1) then
		tmpText1=tmpText1 & FixedField(10, "R","")
	else
		tmpText1=tmpText1 & FixedField(10, "R", scCurSign & money(pPrice1))
	end if
	tmpText1=tmpText1 & FixedField(10, "R", scCurSign & money(pExtendedPrice1)) & vbCrLf	
	storeAdminEmail=replace(storeAdminEmail,"BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" & vbcrlf,tmpText1)
	if pcPrdOrd_BundledDisc>0 then
		storeAdminEmail=storeAdminEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_custOrdInvoice_36"))
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", " ")
		storeAdminEmail=storeAdminEmail & FixedField(10, "R","-" & scCurSign & money(pcPrdOrd_BundledDisc))  & vbCrLf
	end if
	storeAdminEmail=storeAdminEmail & vbCrLf

	'BTO ADDON-S
	Charges=0
	if scBTO=1 then
		if pIdConfigSession<>"0" then
		if (ItemsDiscounts<>"") and (ItemsDiscounts<>"0") then
		storeAdminEmail=storeAdminEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_adminMail_31"))
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", " ")
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", "-" & scCurSign & money(ItemsDiscounts))  & vbCrLf
		end if
		
		'BTO Additional Charges
		'Add customizations if there are any
		if pIdConfigSession<>"0" then
			query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsConfigObj=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsConfigObj=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
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
				Charges=Charges+Cdbl(ArrCValue(i))
			next
			
			storeAdminEmail=storeAdminEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_adminMail_34"))
			storeAdminEmail=storeAdminEmail & FixedField(10, "R", " ")
			storeAdminEmail=storeAdminEmail & FixedField(10, "R", " ") & vbCrLf
			
			for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
				query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
				set rsConfigObj=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsConfigObj=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
				dispStr = rsConfigObj("categoryDesc")&": "&rsConfigObj("description")
				dispStr = replace(dispStr,"&quot;", chr(34))
				tStr = dispStr
				wrapPos=50
				if len(dispStr) > 50 then
					tStr = WrapString(50, dispStr)
				end if
				storeAdminEmail=storeAdminEmail & FixedField(50, "L", tStr)

				if ArrCValue(i)<>0 then
					storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCursign & money(ArrCValue(i)))
				end if
				storeAdminEmail=storeAdminEmail & vbCrLf

				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 50
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(50, dispStr)
					storeAdminEmail=storeAdminEmail & FixedField(10, "L", "")
					storeAdminEmail=storeAdminEmail  & FixedField(50, "L", tStr)
					storeAdminEmail=storeAdminEmail  & vbCrLf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					storeAdminEmail=storeAdminEmail  & FixedField(10, "L", "")
					storeAdminEmail=storeAdminEmail  & FixedField(50, "L", dispStr)
					storeAdminEmail=storeAdminEmail  & vbCrLf
				end if
				set rsConfigObj=nothing
			next
			end if
		end if
			'BTO Additional Charges
			iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(ItemsDiscounts)+cdbl(Charges)-cdbl(pcPrdOrd_BundledDisc)
		else
			iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(pcPrdOrd_BundledDisc)
	end if
	else
		iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(pcPrdOrd_BundledDisc)
	end if
	'======================================
		if (QDiscounts<>"") and (QDiscounts<>"0") then
		storeAdminEmail=storeAdminEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_adminMail_32"))
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", " ")
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", "-" & scCurSign & money(QDiscounts)) & vbCrLf
		end if
	iSubTotal=iSubtotal-cdbl(QDiscounts)

	cdblCmprTmp1=(pPrice*pqty)
	cdblCmprTmp2=(pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges)

	if cdblCmprTmp2<>cdblCmprTmp1 then
		storeAdminEmail=storeAdminEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_adminMail_33"))
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", " ")
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money((pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges))) & vbCrLf
	end if
	
	'GGG Add-on start
	if pGWOpt<>"0" then
	query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
	set rsG=connTemp.execute(query)
	if not rsG.eof then
	pGWOptName=rsG("pcGW_OptName")
	storeAdminEmail=storeAdminEmail & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_39") & pGWOptName & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_40") & scCurSign & money(pGWPrice) & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_41") & pGWText & vbCrLf
	end if
	end if
	'GGG Add-on end
		
	storeAdminEmail=storeAdminEmail & vbCrLf& vbCrLf
		
	rsOrderDetails.MoveNext
loop

' Break then start totals ...
storeAdminEmail=storeAdminEmail & vbCrLf
storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_20"))
storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(iSubTotal))
storeAdminEmail=storeAdminEmail & vbCrLf

' processing fees ...
if paymentCharge<>0 then
	storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_21"))
	storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(paymentCharge))
	storeAdminEmail=storeAdminEmail & vbCrLf
end if

'DiscountCode/Rewards Pts., when applicable...
ptotalDiscounts=pTotalDiscountAmount+piRewardValue+pcOrd_CatDiscounts+GCAmountTotal
if ptotalDiscounts>0 then
	if piRewardValue>0 then
		storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_22")&RewardsLabel&dictLanguage.Item(Session("language")&"_adminMail_23"))
	else
		storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_24"))
	end if
	storeAdminEmail=storeAdminEmail & FixedField(11, "R", "(-"&scCurSign & money(ptotalDiscounts)&")")
	storeAdminEmail=storeAdminEmail & vbCrLf
End If

'GGG Add-on start
If pGWTotal<>"0" Then
	storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_39"))
	storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(pGWTotal))
	storeAdminEmail=storeAdminEmail & vbCrLf
End If
'GGG Add-on end

' Shipping, when applicable ...
If Postage<>0 Then
	storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_26"))
	storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(Postage))
	storeAdminEmail=storeAdminEmail & vbCrLf
End If

' Shipping & handling fees, when applicable ...
If serviceHandlingFee>0 Then
	storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_27"))
	storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(serviceHandlingFee))
	storeAdminEmail=storeAdminEmail & vbCrLf
End If

' Sales tax, when applicable ...
if pord_VAT>0 then
	If ptaxAmount>"0" Then
		storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_orderverify_35"))
		storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(pord_VAT))
		storeAdminEmail=storeAdminEmail & vbCrLf
	End If
else
	if isNull(ptaxDetails) OR trim(ptaxDetails)="" then 
		If ptaxAmount>"0" Then
			storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_25"))
			storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(ptaxAmount))
			storeAdminEmail=storeAdminEmail & vbCrLf
		End If
	else 
		taxArray=split(ptaxDetails,",")
		tempTaxAmount=0
		for i=0 to (ubound(taxArray)-1)
			taxDesc=split(taxArray(i),"|")
			storeAdminEmail=storeAdminEmail & FixedField(60, "R", taxDesc(0)&":")
			storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(taxDesc(1)))
			storeAdminEmail=storeAdminEmail & vbCrLf
		next 
	end if
end if

' Grand total ...
	storeAdminEmail=storeAdminEmail & FixedField(60, "R", "===========")
	storeAdminEmail=storeAdminEmail & FixedField(10, "R", "===========")
	storeAdminEmail=storeAdminEmail & vbCrLf
	storeAdminEmail=storeAdminEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_adminMail_28"))
	storeAdminEmail=storeAdminEmail & FixedField(10, "R", scCurSign & money(ptotal))
	storeAdminEmail=storeAdminEmail & vbCrLf

'Check for comments by customer
If pcomments<>"" then
	storeAdminEmail=storeAdminEmail & vbCrLf
	storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_sendMail_80") & pcomments
	storeAdminEmail=storeAdminEmail & vbCrLf
End If

'Check for affiliate
If idAffiliate<>1 AND idAffiliate<>"" then
	storeAdminEmail=storeAdminEmail & vbCrLf & dictLanguage.Item(Session("language")&"_adminMail_29")& idAffiliate
	storeAdminEmail=storeAdminEmail & vbCrLf & dictLanguage.Item(Session("language")&"_adminMail_30") & scCurSign & money(affiliatePay)
End If

'Downloadable Product Lincense(s)
storeAdminEmail=storeAdminEmail & vbCrLf

IF DPOrder="1" AND pOrderStatus="3" then
	query="select IdProduct from DPRequests WHERE IdOrder=" & qry_ID & ";"
	pidorder=qry_ID
	set rs11=connTemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rs11=nothing
		'//close any connections
		call closedb()
		'//redirect to error page
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	do while not rs11.eof
		pIdProduct=rs11("idProduct")
		query="select * from Products,DProducts where products.idproduct=" & pIdProduct & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
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
		if not rs.eof then
			pProductName=rs("Description")
			pURLExpire=rs("URLExpire")
			pExpireDays=rs("ExpireDays")	
			pLicense=rs("License")
			pLL1=rs("LicenseLabel1")
			pLL2=rs("LicenseLabel2")
			pLL3=rs("LicenseLabel3")
			pLL4=rs("LicenseLabel4")
			pLL5=rs("LicenseLabel5")
	
			query="select RequestSTR from DPRequests where idproduct=" & pIdProduct & " and idorder=" & pidorder & " and idcustomer=" & pidcustomer
			set rs19=connTemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rs19=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			pdownloadStr=rs19("RequestSTR")

			SPath1=Request.ServerVariables("PATH_INFO")
			mycount1=0
			do while mycount1<2
				if mid(SPath1,len(SPath1),1)="/" then
					mycount1=mycount1+1
				end if
				if mycount1<2 then
					SPath1=mid(SPath1,1,len(SPath1)-1)
				end if
			loop
			SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1

			if Right(SPathInfo,1)="/" then
				pdownloadStr=SPathInfo & "pc/pcdownload.asp?id=" & pdownloadStr					
			else
				pdownloadStr=SPathInfo & "/pc/pcdownload.asp?id=" & pdownloadStr
			end if

			storeAdminEmail=storeAdminEmail & "======================================================================" & vbcrlf & vbcrlf
	
			storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_sendMail_28") & pProductName & vbcrlf & vbcrlf
			storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_sendMail_29")
			if (pURLExpire<>"") and (pURLExpire="1") then
				if date()-(CDate(pprocessDate)+pExpireDays)<0 then
					storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_sendMail_30") & (CDate(pprocessDate)+pExpireDays)-date() & dictLanguage.Item(Session("language")&"_sendMail_31") & vbcrlf & vbcrlf
				else
					if date()-(CDate(pprocessDate)+pExpireDays)=0 then
						storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_sendMail_32") & vbcrlf & vbcrlf
					else
						storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_sendMail_33") & vbcrlf & vbcrlf
					end if
				end if
			else
				storeAdminEmail=storeAdminEmail & ":" & vbcrlf & vbcrlf
			end if
			storeAdminEmail=storeAdminEmail & pdownloadStr & vbcrlf &vbcrlf

			if (pLicense<>"") and (pLicense="1") then
				query="select * from DPLicenses where idproduct=" & pIdProduct & " and idorder=" & pidorder
				set rs19=connTemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rs19=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				TempLicStr=""
				do while not rs19.eof
					TempLic=""
					Lic1=rs19("Lic1")
					if trim(Lic1)<>"" then
						TempLic=TempLic & pLL1 & ": " & Lic1 & vbcrlf
					end if
					Lic2=rs19("Lic2")
					if trim(Lic2)<>"" then
						TempLic=TempLic & pLL2 & ": " & Lic2 & vbcrlf
					end if
					Lic3=rs19("Lic3")
					if trim(Lic3)<>"" then
						TempLic=TempLic & pLL3 & ": " & Lic3 & vbcrlf
					end if
					Lic4=rs19("Lic4")
					if trim(Lic4)<>"" then
						TempLic=TempLic & pLL4 & ": " & Lic4 & vbcrlf
					end if
					Lic5=rs19("Lic5")
					if trim(Lic5)<>"" then
						TempLic=TempLic & pLL5 & ": " & Lic5 & vbcrlf
					end if
					if TempLic<>"" then
						TempLic=TempLic & vbcrlf
						TempLicStr=TempLicStr & TempLic
					end if
				rs19.movenext
				loop
				if TempLicStr<>"" then
					TempLicStr=dictLanguage.Item(Session("language")&"_sendMail_34") & vbcrlf & vbcrlf & TempLicStr
					storeAdminEmail=storeAdminEmail & TempLicStr & vbcrlf
				end if
			end if

		end if
	rs11.MoveNext
	loop
	storeAdminEmail=storeAdminEmail & "======================================================================" & vbcrlf & vbcrlf
end if

'Start SDBA
'Back-Ordered products Area%>
<!--#include file="inc_BackOrderEmail.asp"-->
<%
storeAdminEmail=storeAdminEmail & pcv_BackOrderStr
'End SDBA
	
storeAdminEmail=storeAdminEmail & vbCrLf & vbCrLf
storeAdminEmail=replace(storeAdminEmail,"''",chr(39))
%>