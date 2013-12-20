<%
'///////////////////////////////
'// SHOW VAT ID
'// Change "0" to "-1" to show
'///////////////////////////////
pcv_strShowVatId=0 ' -1
'///////////////////////////////

'///////////////////////////////
'// SHOW INTL. ID
'// Change "0" to "1" to show
'///////////////////////////////
pcv_strShowSSN=0 ' -1
'///////////////////////////////


'===========================
'send email
'===========================
'GGG Add-on start

query="SELECT pcOrd_OrderKey, pcOrd_IDEvent, pcOrd_GWTotal, pcOrd_GCDetails, pcOrd_GCAmount FROM orders WHERE idOrder=" & qry_ID
set rs19=server.CreateObject("ADODB.RecordSet")
set rs19=conntemp.execute(query)
if NOT rs19.eof then
	pcOrderKey=rs19("pcOrd_OrderKey")
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
	GCDetails=rs19("pcOrd_GCDetails")
	GCAmountTotal=rs19("pcOrd_GCAmount")
	if GCAmountTotal="" OR IsNull(GCAmountTotal) then
		GCAmountTotal=0
	end if
end if
set rs19=nothing

geHideAddress=0

if gIDEvent<>"0" then

	query="select pcEvents.pcEv_Notify, pcEvents.pcEv_name, pcEvents.pcEv_Date, pcEvents.pcEv_HideAddress, customers.name, customers.lastname, customers.email from pcEvents, Customers where Customers.idcustomer=pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
	set rs1=server.CreateObject("ADODB.RecordSet")
	set rs1=connTemp.execute(query)

	geNotify=rs1("pcEv_Notify")
	if geNotify<>"" then
	else
	geNotify="0"
	end if
	
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
	
	geHideAddress=rs1("pcEv_HideAddress")
	if geHideAddress<>"" then
	else
	geHideAddress="0"
	end if
	
	gReg=rs1("name") & " " & rs1("lastname")
	gRegemail=rs1("email")
	
	set rs1 = nothing
	
end if
'GGG Add-on end

' compile emails
customerEmail=Cstr("")
' Build body of message ...

customerEmail=""
'Customized message from store owner, entered on the Email Settings page

pCustomerFullName = pName&" "&pLName

If (scConfirmEmail<>"" and pcv_CustomerReceived=0) or (scReceivedEmail<>"" and pcv_CustomerReceived=1) Then
	todaydate=showDateFrmt(now())
	if pcv_CustomerReceived=1 then
		personalmessage=replace(scReceivedEmail,"<br>", vbCrlf)
	else
		personalmessage=replace(scConfirmEmail,"<br>", vbCrlf)
	end if
	personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
	personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
	personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
	personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",pCustomerFullName)
	personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(qry_ID)))
	personalmessage=replace(personalmessage,"<ORDER_DATE>",pcv_OrderDate)
	personalmessage=replace(personalmessage,"''",chr(39))
	personalmessage=replace(personalmessage,"//","/")
	personalmessage=replace(personalmessage,"http:/","http://")
	personalmessage=replace(personalmessage,"https:/","https://")
	customerEmail=customerEmail & vbCrLf & personalmessage & vbCrLf
End If

if pcv_AdmComments<>"" then
	customerEmail=customerEmail & vbCrLf & replace(pcv_AdmComments,"''","'") & vbCrLf
end if

If pcOrderKey<>"" then
	customerEmail=customerEmail & "----------------------------------------------------------------------------------------------" & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_storeEmail_30") & pcOrderKey & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_storeEmail_31") &  vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_storeEmail_32") & vbCrLf
	customerEmail=customerEmail & "----------------------------------------------------------------------------------------------" & vbCrLf
End If

' START - Order summary starts here ...
IF pcv_CustomerReceived=0 THEN 'Order Confirmation E-mail
		
customerEmail=customerEmail & vbCrLf & dictLanguage.Item(Session("language")&"_sendMail_2") & vbCrlf
customerEmail=customerEmail & "===================" & vbCrlf
customerEmail=customerEmail & pCustomerFullName & vbCrLf

if pVATID<>"" AND pcv_strShowVatId=-1 Then
	customerEmail=customerEmail & pVATID & vbCrLf
end if

If pSSN<>"" AND pcv_strShowSSN=-1 Then
	customerEmail=customerEmail & pSSN & vbCrLf
End If

If Trim(pCustomerCompany) <> "" Then
	customerEmail=customerEmail & pCustomerCompany & vbCrLf
End If
		
customerEmail=customerEmail & paddress & vbCrLf
if pAddress2<>"" then
	customerEmail=customerEmail & pAddress2 & vbCrLf	
end if
customerEmail=customerEmail & pCity & ", "
if pState = "" then
	customerEmail=customerEmail & pStateCode & " "
	else
	customerEmail=customerEmail & pState & " "
end if
customerEmail=customerEmail & pzip & vbCrLf
customerEmail=customerEmail & pCountryCode & vbCrLf
customerEmail=customerEmail & pPhone &vbCrLf
customerEmail=customerEmail & pEmail  & vbCrLf & vbCrLf 

if geHideAddress=0 then 
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_3") & vbCrlf
	customerEmail=customerEmail & "====================" & vbCrlf
	If Trim(pshippingAddress) <> "" Then
		if pShippingFullName<>"" then
			customerEmail=customerEmail & pShippingFullName& vbCrLf
		end if
		if trim(pshippingCompany)<>"" then
			customerEmail=customerEmail & pshippingCompany & vbCrLf
		end if
		customerEmail=customerEmail & pshippingAddress & vbCrLf
		if trim(pshippingAddress2)<>"" then
			customerEmail=customerEmail & pshippingAddress2 & vbCrLf
		end if
		customerEmail=customerEmail & pshippingCity & ", "
		if pshippingState = "" then
			customerEmail=customerEmail & pshippingStateCode & " "
			else
			customerEmail=customerEmail & pshippingState & " "
		end if
		customerEmail=customerEmail & pshippingZip & vbCrLf
		customerEmail=customerEmail & pshippingCountryCode & vbCrLf
		customerEmail=customerEmail & pshippingPhone & vbCrLf
	Else
		customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_4") & vbCrLf
	End if 
End if
	
customerEmail=customerEmail & vbCrLf

'get shipping details...
shipping=split(pshipmentDetails,",")
if ubound(shipping)>1 then
	if NOT isNumeric(trim(shipping(2))) then
		customerEmail=customerEmail & ship_dictLanguage.Item(Session("language")&"_noShip_a") & vbCrLf
		pShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_a")
		pShipmentPriceToAdd="0"
		Postage=0
		serviceHandlingFee=0
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
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_5") & Service & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_41") & pOrdPackageNum & vbCrLf
else
	customerEmail=customerEmail & ship_dictLanguage.Item(Session("language")&"_noShip_a") & vbCrLf
	pShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	pShipmentPriceToAdd="0"
	Postage=0
	serviceHandlingFee=0
end if  

' If the store is collecting the delivery date
' and it's not empty, then show it.
If DFShow = "1" and pord_DeliveryDate <> "//" Then
	customerEmail=customerEmail & vbCrlf & DFLabel & " " & pord_DeliveryDate & vbCrlf
End If

'offline payment details
paymentdetails=split(trim(pPaymentDetails),"||")
if ubound(paymentdetails)>0 then
	paymentCharge=trim(paymentdetails(1))
	If NOT isNumeric(paymentCharge) then
		paymentCharge=0
	End if
else
	paymentCharge=0
end if

'GGG Add-on start
if gIDEvent<>"0" then

	customerEmail=customerEmail & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_55") & geName & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_56") & geDate & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_57") & gReg & vbCrLf
	customerEmail=customerEmail & vbCrLf

end if
'GGG Add-on end
	
'get discount details...
if instr(pdiscountDetails,",") then
	DiscountDetailsArry=split(pdiscountDetails,",")
	intArryCnt=ubound(DiscountDetailsArry)
else
	intArryCnt=0
end if
dim discounts, discountType 

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
		pdiscountAmt=discounts(1)
		pIsNumeric=1
		if NOT isNumeric(pdiscountAmt) then
			pdiscountAmt=0
			pIsNumeric=0
		end if
		if (pdiscountAmt>0 OR pdiscountAmt=0) AND pIsNumeric=1 then
			customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_6") & pdiscountDesc & vbCrLf
		end if
	Else
		pdiscountAmt=0
	end if
	pTotalDiscountAmount=pTotalDiscountAmount+pdiscountAmt
Next
 
If RewardsActive=1 And ( (piRewardPointsCustAccrued > 0) Or (piRewardPoints > 0)) Then 
	customerEmail=customerEmail & vbCrLf
	'Did we use points or accrue points?
	If piRewardPointsCustAccrued > 0 AND piRewardPoints=0 Then 'Accrued
		iDollarValue=piRewardPointsCustAccrued * (RewardsPercent / 100)
		customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_7") & piRewardPointsCustAccrued & " " & RewardsLabel & dictLanguage.Item(Session("language")&"_sendMail_8") & scCurSign &money(iDollarValue) & vbCrLf
	End If
	If piRewardPoints > 0 Then
		customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_9") & money(piRewardValue) & dictLanguage.Item(Session("language")&"_sendMail_10") & RewardsLabel & "!" & vbCrLf
	End If			
	customerEmail=customerEmail & vbCrLf
End If 

' Begin order details ...

'GGG Add-on start
'Add bookmarks
customerEmail=customerEmail & "AAAAAAAAAA"
'GGG Add-on end

customerEmail=customerEmail & vbCrLf

' Column headings ...
customerEmail=customerEmail & FixedField(20, "L", dictLanguage.Item(Session("language")&"_sendMail_11"))
customerEmail=customerEmail & FixedField(40, "R", dictLanguage.Item(Session("language")&"_sendMail_12"))
customerEmail=customerEmail & FixedField(10, "R", dictLanguage.Item(Session("language")&"_sendMail_13"))
customerEmail=customerEmail & vbCrLf

'Column Dividers
customerEmail=customerEmail & FixedField(50, "R", "==================================================")
customerEmail=customerEmail & FixedField(10, "R", "==========")
customerEmail=customerEmail & FixedField(10, "R", "==========")
customerEmail=customerEmail & vbCrLf
iSubtotal=0
	
query="SELECT products.sku, products.idproduct, products.description, products.price, ProductsOrdered.pcSC_ID, quantity, unitPrice, xfdetails"
'BTO ADDON-S
if scBTO=1 then
	query=query&", idconfigSession"
end if
'BTO ADDON-E
query=query&",ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts,ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.pcPO_GWOpt,ProductsOrdered.pcPO_GWNote,ProductsOrdered.pcPO_GWPrice FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& qry_ID
set rsOrderDetails=server.CreateObject("ADODB.RecordSet")
set rsOrderDetails=conntemp.execute(query)

Do While Not rsOrderDetails.EOF
	psku=rsOrderDetails("sku")
	pidproduct=rsOrderDetails("idproduct")
	pdescription=rsOrderDetails("description")
	pdescription=ClearHTMLTags2(pdescription,0)
	pIndPrice=rsOrderDetails("price")
	
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

	pExtendedPrice=pPrice*pqty
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_14")&pqty & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_15")&psku & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_16")&pdescription & vbCrLf
	customerEmail=customerEmail & "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" & vbcrlf

	'BTO ADDON-S
	TotalUnit=0
	if scBTO=1 then
		'Add customizations if there are any
		if pIdConfigSession<>"0" then
			query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsConfigObj=conntemp.execute(query)
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
				
				query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
				set rsConfigObj=conntemp.execute(query)
				pcv_strBtoItemName = rsConfigObj("description")
				pcv_strBtoItemName=ClearHTMLTags2(pcv_strBtoItemName,0)
				pcv_strBtoItemCat=rsConfigObj("categoryDesc")
				pcv_strBtoItemCat=ClearHTMLTags2(pcv_strBtoItemCat,0)
				query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i) 
				set rsObj1=conntemp.execute(query)
				customerEmail=customerEmail & FixedField(10, "L", "")
				dispStr = ""
				dispStr = pcv_strBtoItemCat &": "& pcv_strBtoItemName
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
				customerEmail=customerEmail & FixedField(50, "L", tStr)
				
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
					customerEmail=customerEmail & FixedField(10, "R", scCurSign & money((ArrValue(i)+UPrice)*pQty))
				else
					if tmpDefault=1 then
						customerEmail=customerEmail & FixedField(10, "R", dictLanguage.Item(Session("language")&"_defaultnotice_1"))
					end if
				end if
				customerEmail=customerEmail & vbCrLf
				
				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 50
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(50, dispStr)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(50, "L", tStr)
					customerEmail=customerEmail & vbCrLf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(50, "L", dispStr)
					customerEmail=customerEmail & vbCrLf
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
			customerEmail=customerEmail & FixedField(10, "L", "OPTIONS")

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
					customerEmail=customerEmail & FixedField(10, "L", " ")
				end if
			
				dispStr = pcArray_strOptions(pcv_intOptionLoopCounter)
				dispStr = replace(dispStr,"&quot;", chr(34))
				tStr = dispStr
				wrapPos=50
				if len(dispStr) > 50 then
					tStr = WrapString(50, dispStr)
				end if
				customerEmail=customerEmail & FixedField(50, "L", tStr)
								
				tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
													
				if tempPrice="" or tempPrice=0 then
					customerEmail=customerEmail & FixedField(10, "R", " ")
					customerEmail=customerEmail & vbCrLf
				else 
					customerEmail=customerEmail & FixedField(10, "R", "")
					customerEmail=customerEmail & vbCrLf
				end if
				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 50
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(50, dispStr)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(50, "L", tStr)
					customerEmail=customerEmail & vbCrLf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(50, "L", dispStr)
					customerEmail=customerEmail & vbCrLf
				end if
			Next
			'#####################
			' END LOOP
			'#####################

			customerEmail=customerEmail & vbCrLf

	End If
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Add first 50 characters of options on a separate line
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	'show xtra options
	If len(xfdetails)>3 then
		xfarray=split(xfdetails,"|")
		for q=lbound(xfarray) to ubound(xfarray)
			customerEmail=customerEmail & replace((xfarray(q)),"&quot;","""")
			customerEmail=customerEmail & vbCrLf & vbcrlf
		next
		customerEmail=customerEmail & vbCrLf
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
	customerEmail=replace(customerEmail,"BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB" & vbcrlf,tmpText1)
	customerEmail=customerEmail & vbCrLf
		
	'BTO ADDON-S
	Charges=0
	if scBTO=1 then
		if pIdConfigSession<>"0" then
		if (ItemsDiscounts<>"") and (ItemsDiscounts<>"0") then
		customerEmail=customerEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_sendMail_37"))
		customerEmail=customerEmail & FixedField(10, "R", " ")
		customerEmail=customerEmail & FixedField(10, "R", "-" & scCurSign & money(ItemsDiscounts))  & vbCrLf
		end if
		
		'BTO ADDON-S
	if scBTO=1 then
		'Add customizations if there are any
		if pIdConfigSession<>"0" then
			query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsConfigObj=conntemp.execute(query)
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
			
			customerEmail=customerEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_sendMail_40"))
			customerEmail=customerEmail & FixedField(10, "R", " ")
			if Charges<>0 then
				customerEmail=customerEmail & FixedField(10, "R", " ") & vbCrLf
			else
				customerEmail=customerEmail & FixedField(10, "R", " ") & vbCrLf
			end if			

			for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
				query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
				set rsConfigObj=conntemp.execute(query)
				dispStr =""
				customerEmail=customerEmail & FixedField(10, "L", "")
				dispStr = rsConfigObj("categoryDesc")&": "&rsConfigObj("description")
				dispStr = replace(dispStr,"&quot;", chr(34))
				tStr = dispStr
				wrapPos=50
				if len(dispStr) > 50 then
					tStr = WrapString(50, dispStr)
				end if
				customerEmail=customerEmail & FixedField(50, "L", tStr)
				
				if ArrCValue(i)<>0 then
					customerEmail=customerEmail & FixedField(10, "R", scCursign & money(ArrCValue(i)))
				end if
				customerEmail=customerEmail & vbCrLf
				
				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 50
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(50, dispStr)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(50, "L", tStr)
					customerEmail=customerEmail & vbCrLf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					customerEmail=customerEmail & FixedField(10, "L", "")
					customerEmail=customerEmail & FixedField(50, "L", dispStr)
					customerEmail=customerEmail & vbCrLf
				end if

				set rsConfigObj=nothing
			next
			end if
		end if
	end if 
	'BTO ADDON-E
		iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges)
		
		else
		iSubTotal=iSubtotal + (pPrice*pqty)
		end if
	else	
	iSubTotal=iSubtotal + (pPrice*pqty)
	end if
		
	'======================================
		if (QDiscounts<>"") and (QDiscounts<>"0") then
		customerEmail=customerEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_sendMail_38"))
		customerEmail=customerEmail & FixedField(10, "R", " ")
		customerEmail=customerEmail & FixedField(10, "R", "-" & scCurSign & money(QDiscounts)) & vbCrLf
		end if
	iSubTotal=iSubtotal-cdbl(QDiscounts)

	cdblCmprTmp1=(pPrice*pqty)
	cdblCmprTmp2=(pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges)

	if cdblCmprTmp2<>cdblCmprTmp1 then
		customerEmail=customerEmail & FixedField(50, "L", dictLanguage.Item(Session("language")&"_sendMail_39"))
		customerEmail=customerEmail & FixedField(10, "R", " ")
		customerEmail=customerEmail & FixedField(10, "R", scCurSign & money((pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges))) & vbCrLf
	end if
	
	'GGG Add-on start
	if pGWOpt<>"0" then
	query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
	set rsG=connTemp.execute(query)
	if not rsG.eof then
	pGWOptName=rsG("pcGW_OptName")
	customerEmail=customerEmail & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_63") & pGWOptName & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_64") & scCurSign & money(pGWPrice) & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_65") & pGWText & vbCrLf
	end if
	end if
	'GGG Add-on end
		
	customerEmail=customerEmail & vbCrLf& vbCrLf

rsOrderDetails.MoveNext
loop
set rsOrderDetails=nothing

' Break then start totals ...
customerEmail=customerEmail & vbCrLf
customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_19"))
customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(iSubTotal))
customerEmail=customerEmail & vbCrLf

'GGG Add-on start
'Add bookmarks
customerEmail=customerEmail & "AAAAAAAAAA"
'GGG Add-on end

' processing charges, when applicable ...
If paymentcharge<>0 Then
	customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_20"))
	customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(paymentcharge))
	customerEmail=customerEmail & vbCrLf
End If
	
'DiscountCode, when applicable...Category Discounts
ptotalDiscounts=pTotalDiscountAmount+piRewardValue+pcOrd_CatDiscounts+GCAmountTotal
if ptotalDiscounts>0 then
	if piRewardValue>0 then
		customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_21") &RewardsLabel&dictLanguage.Item(Session("language")&"_sendMail_22"))
	else
		customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_23"))
	end if
	customerEmail=customerEmail & FixedField(11, "R", "(-"&scCurSign & money(ptotalDiscounts)&")")
	customerEmail=customerEmail & vbCrLf
End If

'GGG Add-on start
If pGWTotal<>"0" Then
	customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_63"))
	customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(pGWTotal))
	customerEmail=customerEmail & vbCrLf
End If
'GGG Add-on end
	
' Shipping, when applicable ...
If Postage>0 Then
	customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_25"))
	customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(Postage))
	customerEmail=customerEmail & vbCrLf
End If
	
'Handling, when applicable ...
If serviceHandlingFee>0 Then
	customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_26"))
	customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(serviceHandlingFee))
	customerEmail=customerEmail & vbCrLf
End If

' Sales tax, when applicable ...
if pord_VAT>0 then
	If ptaxAmount>"0" Then
		customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_orderverify_35"))
		customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(pord_VAT))
		customerEmail=customerEmail & vbCrLf
	End If
else
	if isNull(ptaxDetails) OR trim(ptaxDetails)="" then 
		If ptaxAmount>"0" Then
			customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_24"))
			customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(ptaxAmount))
			customerEmail=customerEmail & vbCrLf
		End If
	else 
		taxArray=split(ptaxDetails,",")
		for i=0 to (ubound(taxArray)-1)
			taxDesc=split(taxArray(i),"|")
			customerEmail=customerEmail & FixedField(60, "R", taxDesc(0)&":")
			customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(taxDesc(1)))
			customerEmail=customerEmail & vbCrLf
		next 
	end if
end if

customerEmail=customerEmail & FixedField(60, "R", "===========")
customerEmail=customerEmail & FixedField(10, "R", "===========")
customerEmail=customerEmail & vbCrLf
customerEmail=customerEmail & FixedField(60, "R", dictLanguage.Item(Session("language")&"_sendMail_27"))
customerEmail=customerEmail & FixedField(10, "R", scCurSign & money(ptotal))
customerEmail=customerEmail & vbCrLf

' Check for comments by customer
If pcomments<>"" then
	customerEmail=customerEmail & vbCrLf
	customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_80") & pcomments
	customerEmail=customerEmail & vbCrLf
End If

customerEmail=customerEmail & vbCrLf

'GGG Add-on start

IF (GCDetails<>"") then
CustomerEmail=customerEmail & "======================================================================" & vbcrlf & vbcrlf
	
CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_46") & vbcrlf & vbcrlf

GCArr=split(GCDetails,"|g|")
intGCCount=ubound(GCArr)
For y=0 to intGCCount
if GCArr(y)<>"" then
	GCInfo=split(GCArr(y),"|s|")
	pGiftCode=GCInfo(0)
	pGiftUsed=GCInfo(2)

	query="select products.IDProduct,products.Description from pcGCOrdered,Products where products.idproduct=pcGCOrdered.pcGO_idproduct and pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if not rs.eof then
		pIdproduct=rs("idproduct")
		pName=rs("Description")
		pCode=pGiftCode
		CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_47") & pName & vbcrlf

		query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
		set rs19=server.CreateObject("ADODB.RecordSet")
		set rs19=connTemp.execute(query)					
		if not rs19.eof then
			CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_48") & rs19("pcGO_GcCode") & vbcrlf
			CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_49") & scCurSign & money(pGiftUsed) & vbcrlf & vbcrlf
			pGCAmount=rs19("pcGO_Amount")
			if cdbl(pGCAmount)<=0 then
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_50") & vbcrlf
			else
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_51") & scCurSign & money(pGCAmount) & vbcrlf
				pExpDate=rs19("pcGO_ExpDate")
				if year(pExpDate)="1900" then
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_52") & vbcrlf 
				else
					if scDateFrmt="DD/MM/YY" then
						pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
					else
						pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
					end if
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_53") & pExpDate & vbcrlf
				end if
				pGCStatus=rs19("pcGO_Status")
				if pGCStatus="1" then
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_54") & dictLanguage.Item(Session("language")&"_sendMail_54a") & vbcrlf
				else
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_54") & dictLanguage.Item(Session("language")&"_sendMail_54b") & vbcrlf
				end if
			end if
			CustomerEmail=customerEmail & vbcrlf & vbcrlf
		end if '// if not rs19.eof then
		set rs19=nothing
		
	end if '// if not rs.eof then
	set rs=nothing
end if
Next
	CustomerEmail=customerEmail & "======================================================================" & vbcrlf & vbcrlf
END IF

'GGG Add-on end

IF DPOrder="1" then
	
	query="select IdProduct from DPRequests WHERE IdOrder=" & qry_ID
	pidorder=qry_ID
	set rs11=server.CreateObject("ADODB.RecordSet")
	set rs11=connTemp.execute(query)
	do while not rs11.eof
		pIdproduct=rs11("idproduct")
		
		query="SELECT * from Products,DProducts WHERE products.idproduct=" & pIdproduct & " AND DProducts.idproduct=Products.idproduct AND products.downloadable=1"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)	
		pcv_strShowSection = 0
		if not rs.eof then			
			pName=rs("Description")
			pURLExpire=rs("URLExpire")
			pExpireDays=rs("ExpireDays")	
			pLicense=rs("License")
			pLL1=rs("LicenseLabel1")
			pLL2=rs("LicenseLabel2")
			pLL3=rs("LicenseLabel3")
			pLL4=rs("LicenseLabel4")
			pLL5=rs("LicenseLabel5")
			pAddtoMail=rs("AddtoMail")
			pcv_strShowSection = -1
		end if
		set rs=nothing
		
		If pcv_strShowSection = -1 Then
		
			query="select RequestSTR from DPRequests where idproduct=" & pIdproduct & " and idorder=" & pidorder & " and idcustomer=" & pidcustomer
			set rs19=server.CreateObject("ADODB.RecordSet")
			set rs19=connTemp.execute(query)
			pdownloadStr=rs19("RequestSTR")
			set rs19=nothing
			
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

			CustomerEmail=customerEmail & "======================================================================" & vbcrlf & vbcrlf
	
			CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_28") & pName & vbcrlf & vbcrlf
			CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_29")
			if (pURLExpire<>"") and (pURLExpire="1") then
				if date()-(CDate(pprocessDate)+pExpireDays)<0 then
					CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_30") & (CDate(pprocessDate)+pExpireDays)-date() & dictLanguage.Item(Session("language")&"_sendMail_31") & vbcrlf & vbcrlf
				else
					if date()-(CDate(pprocessDate)+pExpireDays)=0 then
						CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_32") & vbcrlf & vbcrlf
					else
						CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_33") & vbcrlf & vbcrlf
					end if
				end if
			else
				CustomerEmail=CustomerEmail & ":" & vbcrlf & vbcrlf
			end if
			CustomerEmail=CustomerEmail & pdownloadStr & vbcrlf & vbcrlf
			CustomerEmail=CustomerEmail & dictLanguage.Item(Session("language")&"_DownloadURLNote_1") & vbcrlf & vbcrlf

			if (pLicense<>"") and (pLicense="1") then
				
				query="SELECT * FROM DPLicenses WHERE idproduct=" & rs11("idproduct") & " AND idorder=" & pidorder
				set rs19=server.CreateObject("ADODB.RecordSet")
				set rs19=connTemp.execute(query)
				TempLicStr=""
				do while not rs19.eof
					TempLic=""
					Lic1=rs19("Lic1")
					if Lic1<>"" then
						TempLic=TempLic & pLL1 & ": " & Lic1 & vbcrlf
					end if
					Lic2=rs19("Lic2")
					if Lic2<>"" then
						TempLic=TempLic & pLL2 & ": " & Lic2 & vbcrlf
					end if
					Lic3=rs19("Lic3")
					if Lic3<>"" then
						TempLic=TempLic & pLL3 & ": " & Lic3 & vbcrlf
					end if
					Lic4=rs19("Lic4")
					if Lic4<>"" then
						TempLic=TempLic & pLL4 & ": " & Lic4 & vbcrlf
					end if
					Lic5=rs19("Lic5")
					if Lic5<>"" then
						TempLic=TempLic & pLL5 & ": " & Lic5 & vbcrlf
					end if
					if TempLic<>"" then
						TempLic=TempLic & vbcrlf
						TempLicStr=TempLicStr & TempLic
					end if
				
					rs19.movenext
				loop
				set rs19=nothing
				
				if TempLicStr<>"" then
					TempLicStr=dictLanguage.Item(Session("language")&"_sendMail_34") & vbcrlf & vbcrlf & TempLicStr
					CustomerEmail=customerEmail & TempLicStr & vbcrlf
				end if
				
			end if

			if pAddtoMail<>"" then
				CustomerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_35") & vbcrlf & vbcrlf & pAddtoMail & vbcrlf & vbcrlf
			end if
			
		End If '// If pcv_strShowSection = -1 Then
	
		rs11.MoveNext
	loop
	set rs11=nothing
	
	CustomerEmail=customerEmail & "======================================================================" & vbcrlf & vbcrlf
end if

'GGG Add-on start
IF pGCs="1" then
	
	query="SELECT idproduct FROM ProductsOrdered WHERE idOrder="& qry_ID
	pidorder=qry_ID
	set rs11=server.CreateObject("ADODB.RecordSet")
	set rs11=connTemp.execute(query)
	do while not rs11.eof
		pIdproduct=rs11("idproduct")
		
		query="select products.Description,pcGCOrdered.pcGO_GcCode from Products,pcGCOrdered where products.idproduct=" & pIdproduct & " and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)	
		pcv_strShowSection = 0
		if not rs.eof then			
			pName=rs("Description")
			pCode=rs("pcGO_GcCode")
			pcv_strShowSection = -1
		end if
		set rs=nothing
		
		If pcv_strShowSection = -1 Then
			
			customerEmail=customerEmail & "======================================================================" & vbcrlf
			customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_42") & vbcrlf
			customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_43") & pName & vbcrlf & vbcrlf
			
				'// START - Gift Certificate Recipient information
				query="select pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg from Orders WHERE idOrder="& qry_ID
				set rs20=Server.CreateObject("ADODB.Recordset")
				set rs20=connTemp.execute(query)
				if not rs20.eof then
					pcvGcRecipientName=rs20("pcOrd_GcReName")
					if pcvGcRecipientName<>"" then
						customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_NotifyRe_3") & pcvGcRecipientName & vbcrlf
					end if
					pcvGcRecipientEmail=rs20("pcOrd_GcReEmail")
					if pcvGcRecipientEmail<>"" then
						customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_NotifyRe_4") & pcvGcRecipientEmail & vbcrlf
					end if
					customerEmail=customerEmail & vbcrlf
				end if
				set rs20=nothing
				'// END - Gift Certificate Recipient information

			
				query="SELECT pcGO_GcCode,pcGO_ExpDate FROM pcGCOrdered WHERE pcGO_idproduct=" & pIdproduct & " AND pcGO_idorder=" & qry_ID
				set rs19=server.CreateObject("ADODB.RecordSet")
				set rs19=connTemp.execute(query)				
				do while not rs19.eof
					customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_44") & rs19("pcGO_GcCode") & vbcrlf
					pExpDate=rs19("pcGO_ExpDate")
					 
					if year(pExpDate)="1900" then
						customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_45b") & vbcrlf & vbcrlf
					else
						if scDateFrmt="DD/MM/YY" then
							pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
						else
							pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
						end if
						customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_45") & pExpDate & vbcrlf & vbcrlf
					end if
					rs19.movenext
				loop
				set rs19 = nothing
				
				customerEmail=customerEmail & vbcrlf

		End If '// If pcv_strShowSection = -1 Then
		
		rs11.MoveNext
	loop
	set rs11 = nothing
	
	CustomerEmail=customerEmail & "======================================================================" & vbcrlf & vbcrlf
END IF
'GGG Add-on end

'Start SDBA
'Back-Ordered products Area
%>

<!--#include file="../pc/inc_BackOrderEmail.asp"-->

<%
customerEmail=customerEmail & pcv_BackOrderStr

'// Create a link to receive customer confirmation about separate shipments
if (scAllowSeparate="1") and (pcv_BackOrderStr<>"") and (pcv_CustomerReceived=0) and ((request.querystring("Submit4")<>"") or (pcv_SubmitType=3)) then
	strPath=Request.ServerVariables("PATH_INFO")
	iCnt=0
	do while iCnt<1
		if mid(strPath,len(strPath),1)="/" then
			iCnt=iCnt+1
		end if
		if iCnt<1 then
			strPath=mid(strPath,1,len(strPath)-1)
		end if
	loop
	
	strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
	
	strPathInfo=replace(strPathInfo,"/" & scAdminFolderName,"")
				
	if Right(strPathInfo,1)="/" then
	else
		strPathInfo=strPathInfo & "/"
	end if
	
	'///////////////////////////////////////////////////
	'// START: DO
	'///////////////////////////////////////////////////
	DO
		Tn1=""
		For dd=1 to 100
			Randomize
			myC=Fix(3*Rnd)
			Select Case myC
				Case 0: 
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
				Case 1: 
					Randomize
					Tn1=Tn1 & Cstr(Fix(10*Rnd))
				Case 2: 
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+97)		
			End Select		
		Next

		ReqExist=0
	
		query="SELECT IDOrder FROM Orders WHERE pcOrd_CustRequestStr='" & Tn1 & "'" 
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		if not rstemp.eof then
			ReqExist=1
		end if
		set rstemp=nothing
		
	LOOP UNTIL ReqExist=0
	'///////////////////////////////////////////////////
	'// END: DO
	'///////////////////////////////////////////////////
	
	query="Update Orders Set pcOrd_CustRequestStr='" & Tn1 & "' WHERE idorder=" & qry_ID
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
	
end if
	
if (scAllowSeparate="1") and (pcv_BackOrderStr<>"") and (pcv_CustomerReceived=0) then
	query="SELECT pcOrd_CustRequestStr FROM Orders WHERE idorder=" & qry_ID 
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	Tn1=rstemp("pcOrd_CustRequestStr")
	set rstemp=nothing
	
	strPathInfo=strPathInfo & "pc/sds_AllowSeparateShip.asp?req=" & Tn1
	
	'Add request link to the Customer Confirmation E-mail
	if pcv_CustomerReceived=0 then
	customerEmail=customerEmail & ship_dictLanguage.Item(Session("language")&"_custconfirm_msg_1") & vbcrlf
	customerEmail=customerEmail & strPathInfo & vbcrlf & vbcrlf
	end if
end if
'End SDBA

customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_sendMail_36") & scCompanyName & "." & vbCrLf & vbCrLf
customerEmail=replace(CustomerEmail,"''",chr(39))

'GGG Add-on start
'Del bookmarks
tempMail=split(customerEmail,"AAAAAAAAAA")
customerEmail=replace(customerEmail,"AAAAAAAAAA","")
'GGG Add-on end

'GGG Add-on start

if gIDEvent<>"0" then
	RegEmail=""
	RegEmail=dictLanguage.Item(Session("language")&"_sendMail_58") & gReg & "," & vbCrLf & vbcrlf
	RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_59") & vbcrlf
	RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_55") & geName & vbCrLf
	RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_56") & geDate & vbCrLf
	RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_60") & pCustomerFullName & vbcrlf
	RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_61") & vbcrlf
	RegEmail=RegEmail & dictLanguage.Item(Session("language")&"_sendMail_62") & scpre+int(qry_ID) & vbcrlf
	RegEmail=RegEmail & tempMail(1) & vbcrlf
	RegEmail=RegEmail & scCompanyName & vbCrLf & vbCrLf
	if geNotify="1" then
	call sendmail (scCompanyName, scEmail, gRegemail, "Someone purchased some gifts off your Gift Registry", replace(RegEmail, "&quot;", chr(34)))
	end if
end if
'GGG Add-on end



'GGG Add-on start
ReciEmail=""

IF pGCs="1" then
	
	query="select idproduct from ProductsOrdered WHERE idOrder="& qry_ID
	pidorder=qry_ID
	set rs11=Server.CreateObject("ADODB.Recordset")
	set rs11=connTemp.execute(query)
	do while not rs11.eof
		pIdproduct=rs11("idproduct")
		
		query="select products.Description,pcGCOrdered.pcGO_GcCode,pcGc.pcGc_EOnly from Products,pcGc,pcGCOrdered where products.idproduct=" & pIdproduct & " and pcGC.pcGc_IDProduct=products.idproduct and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& qry_ID
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		pcv_strRunSection = 0
		if not rs.eof then			
			pName=rs("Description")
			pCode=rs("pcGO_GcCode")
			pEOnly=rs("pcGc_EOnly")
			pcv_strRunSection = -1
		end if
		set rs = nothing
	
		If pcv_strRunSection = -1 Then
			
				query="select pcGO_Amount,pcGO_GcCode,pcGO_ExpDate from pcGCOrdered where pcGO_idproduct=" & pIdproduct & " and pcGO_idorder=" & qry_ID
				set rs19=Server.CreateObject("ADODB.Recordset")
				set rs19=connTemp.execute(query)				
				do while not rs19.eof
					
					pAmount=rs19("pcGO_Amount")
					if pAmount<>"" then
					else
						pAmount="0"
					end if					
					ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_68") & scCurSign & money(pAmount) & vbcrlf					
					ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_69") & rs19("pcGO_GcCode") & vbcrlf
					pExpDate=rs19("pcGO_ExpDate")
					 
					if year(pExpDate)="1900" then
						ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_45b") & vbcrlf
					else
						if scDateFrmt="DD/MM/YY" then
							pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
						else
							pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
						end if
						ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_70") & pExpDate & vbcrlf
					end if
					
					if pEOnly="1" then
						ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_71") & vbcrlf
					end if
					ReciEmail=ReciEmail & vbcrlf
					
					rs19.movenext
				loop
				set rs19=nothing
				
		End If '// If pcv_strRunSection = -1 Then
		
		rs11.MoveNext
	loop
	set rs11=nothing
	
	query="select pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg from Orders WHERE idOrder="& qry_ID
	set rs11=connTemp.execute(query)	
	GcReName=rs11("pcOrd_GcReName")
	GcReEmail=rs11("pcOrd_GcReEmail")
	GcReMsg=rs11("pcOrd_GcReMsg")
	pCustomerFullNamePlusEmail=pCustomerFullName & " (" & pEmail & ")"
	
	set rs11 = nothing	
	if GcReEmail<>"" then
		if GcReName<>"" then
		else
			GcReName=GcReEmail
		end if
		ReciEmail1=replace(dictLanguage.Item(Session("language")&"_sendMail_66"),"<recipient name>",GcReName)
		ReciEmail2=replace(dictLanguage.Item(Session("language")&"_sendMail_67"),"<customer name>",pCustomerFullNamePlusEmail)
		if GcReMsg<>"" then
			ReciEmail3=replace(dictLanguage.Item(Session("language")&"_sendMail_72"),"<customer name>",pCustomerFullNamePlusEmail) & vbcrlf & GcReMsg & vbcrlf
		else
		ReciEmail3=""
		end if
		ReciEmail=ReciEmail1 & vbcrlf & vbcrlf & ReciEmail2 & vbcrlf & vbcrlf & ReciEmail & ReciEmail3
		ReciEmail=ReciEmail & vbcrlf & scCompanyName & vbCrLf & scStoreURL & vbcrlf & vbCrLf
		call sendmail (scCompanyName, scEmail, GcReEmail,pCustomerFullName & dictLanguage.Item(Session("language")&"_sendMail_73"), replace(ReciEmail, "&quot;", chr(34)))
	end if

END IF
'GGG Add-on end

END IF
' END - Order summary
%>