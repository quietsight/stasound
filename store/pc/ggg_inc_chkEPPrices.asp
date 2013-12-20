<%
Function updPrices(pcv_chkIDProduct,pcv_chkIDConfig,pIDOptionArray,gmode)

'Get Product Default Price
query="Select price,btoBPrice from products where IDProduct=" & pcv_chkIDProduct
set rsA=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsA=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
	
ProductWPrice=rsA("btoBPrice")

If (Session("customerType")=1) and (ProductWPrice>0) Then
	pPrice=ProductWPrice
else
	pPrice=rsA("price")
end if
	
pfPrice=pPrice

set rsA=nothing
	
'************** BTO Configuration ****************
	
IF (pcv_chkIDConfig<>"0") and (pcv_chkIDConfig<>"") THEN

query="SELECT * FROM configSessions where idProduct=" & pcv_chkIDProduct &" and idconfigSession=" & pcv_chkIDConfig
Set rsA=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsA=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcv_Midproduct=rsA("idproduct")

stringProducts=rsA("stringProducts")
stringValues=rsA("stringValues")
stringCategories=rsA("stringCategories")
stringQuantity=rsA("stringQuantity")
stringPrice=rsA("stringPrice")
stringCProducts=rsA("stringCProducts")
stringCValues=rsA("stringCValues")
stringCCategories=rsA("stringCCategories")

set rsA=nothing

ArrProduct=Split(stringProducts, ",")
ArrValue=Split(stringValues, ",")
ArrCategory=Split(stringCategories, ",")
ArrQuantity=Split(stringQuantity, ",")
ArrPrice=Split(stringPrice, ",")
ArrCProduct=Split(stringCProducts, ",")
ArrCValue=Split(stringCValues, ",")
ArrCCategory=Split(stringCCategories, ",")

myUpdateQuote=false

sProducts=""
sValues=""
sCategories=""
sQuantity=""
sPrice=""

RmvItems=0

'Check BTO Items
IF ArrProduct(0)<>"na" then
for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
	DefaultPrice=0
	query="select price,Wprice from configSpec_products where configProductCategory=" & ArrCategory(i) & " and cdefault=1 and specProduct=" & pcv_Midproduct
	set rsB=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsB=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rsB.eof then
		ItemWPrice=cdbl(rsB("Wprice"))
		If (Session("customerType")=1) and (ItemWPrice<>0) Then
			DefaultPrice=ItemWPrice
		else
			DefaultPrice=cdbl(rsB("price"))
		end if
	else
		myUpdateQuote=True
	end if
	
	set rsB=nothing
	
	NewPrice=0
	query="select price,Wprice,multiSelect from configSpec_products where configProductCategory=" & ArrCategory(i) & " and configProduct=" & ArrProduct(i) & " and specProduct=" & pcv_Midproduct
	set rsB=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rsB.eof then
		NewPrice=cdbl(rsB("Wprice"))
		If (Session("customerType")=1) and (NewPrice<>0) Then
			NewPrice=cdbl(rsB("Wprice"))
		else
			NewPrice=cdbl(rsB("price"))
		end if
		sProducts=sProducts & ArrProduct(i) & ","
		sCategories=sCategories & ArrCategory(i) & ","
		sQuantity=sQuantity & ArrQuantity(i) & ","
		sPrice=sPrice & NewPrice & ","
		NPrice=0
		multiSelect=rsB("multiSelect")
	
		if (cdbl(NewPrice)-cdbl(DefaultPrice)<>cdbl(ArrValue(i))) and (multiSelect<>"") and (multiSelect<>-1) then
			NPrice=cdbl(NewPrice)-cdbl(DefaultPrice)
			sValues=sValues & NPrice & ","
			myUpdateQuote=True
		else
			if (cdbl(NewPrice)<>cdbl(ArrValue(i))) and (multiSelect=-1) then
				NPrice=cdbl(NewPrice)
				sValues=sValues & NPrice & ","
				myUpdateQuote=True
			else
				NPrice=cdbl(ArrValue(i))
				sValues=sValues & NPrice & ","
			end if
		end if
	else	
		myUpdateQuote=True
	end if
	set rsB=nothing
next

query="select configProductCategory,requiredCategory,configProduct,Wprice,price,multiSelect from configSpec_products where specProduct=" & pcv_Midproduct & " and cdefault=1"
set rsB=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsB=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
do while not rsB.eof
	conCAT=rsB("configProductCategory")
	RCAT=rsB("requiredCategory")
	if (instr(sCategories,conCAT & ",")=0) and (RCAT=True) then
		myUpdateQuote=True
		sProducts=sProducts & rsB("configProduct") & ","
		sCategories=sCategories & rsB("configProductCategory") & ","
		if rsb("multiSelect")=-1 then
			ItemWPrice=cdbl(rsB("Wprice"))
			If (Session("customerType")=1) and (ItemWPrice>0) Then
				sValues=sValues & ItemWPrice & ","
			else
				sValues=sValues & rsB("price") & ","
			end if
		else	
			sValues=sValues & "0,"
		end if
		sQuantity=sQuantity & "1,"
		ItemWPrice=cdbl(rsB("Wprice"))
		If (Session("customerType")=1) and (ItemWPrice>0) Then
			sPrice=sPrice & ItemWPrice & ","
		else
			sPrice=sPrice & rsB("price") & ","
		end if
	end if
	rsB.movenext
loop
set rsB=nothing
END IF

if sProducts="" then
sProducts="na"
sValues="na"
sCategories="na"
sQuantity="na"
sPrice="na"
end if
'End Check BTO Items

sCProducts=""
sCValues=""
sCCategories=""

'Check Additional Charges
IF ArrCProduct(0)<>"na" then
for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
	DefaultPrice=0
	query="select price,Wprice from configSpec_Charges where configProductCategory=" & ArrCCategory(i) & " and cdefault=1 and specProduct=" & pcv_Midproduct
	set rsB=conntemp.execute(query)
		if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rsB.eof then
		ItemWPrice=cdbl(rsB("Wprice"))
		If (Session("customerType")=1) and (ItemWPrice<>0) Then
			DefaultPrice=ItemWPrice
		else
			DefaultPrice=cdbl(rsB("price"))
		end if
	else
		myUpdateQuote=True
	end if
	set rsB=nothing
	
	NewPrice=0
	query="select price,Wprice,multiSelect from configSpec_Charges where configProductCategory=" & ArrCCategory(i) & " and configProduct=" & ArrCProduct(i) & " and specProduct=" & pcv_Midproduct
	set rsB=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsB=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rsB.eof then
		ItemWPrice=cdbl(rsB("Wprice"))
		If (Session("customerType")=1) and (ItemWPrice<>0) Then
			NewPrice=ItemWPrice
		else
			NewPrice=cdbl(rsB("price"))
		end if
		sCProducts=sCProducts & ArrCProduct(i) & ","
		sCCategories=sCCategories & ArrCCategory(i) & ","
		NPrice=0
		multiSelect=rsB("multiSelect")
		if (cdbl(NewPrice)<>cdbl(ArrCValue(i))) then
			sCValues=sCValues & NewPrice & ","
			myUpdateQuote=True
		else
			NPrice=cdbl(ArrCValue(i))
			sCValues=sCValues & NPrice & ","
		end if
	else	
		myUpdateQuote=True
	end if
	set rsB=nothing	
next

query="select configProductCategory,requiredCategory,configProduct,Wprice,price,multiSelect from configSpec_Charges where specProduct=" & pcv_Midproduct & " and cdefault=1"
set rsB=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
do while not rsB.eof
	conCAT=rsB("configProductCategory")
	RCAT=rsB("requiredCategory")
	if (instr(sCCategories,conCAT & ",")=0) and (RCAT=True) then
		myUpdateQuote=True
		sCProducts=sCProducts & rsB("configProduct") & ","
		sCCategories=sCCategories & rsB("configProductCategory") & ","
		If (Session("customerType")=1) and (rsB("Wprice")<>0) Then
			sCValues=sCValues & rsB("Wprice") & ","
		else
			sCValues=sCValues & rsB("price") & ","
		end if
	end if
	rsB.movenext
loop
set rsB=nothing
END IF

if sCProducts="" then
sCProducts="na"
sCValues="na"
sCCategories="na"
end if

'End check Additional Charges

	query="UPDATE configSessions SET stringProducts='"&sProducts&"',stringValues='"&sValues&"',stringCategories='"&sCategories&"',stringQuantity='" & sQuantity & "',stringPrice='" & sPrice & "',stringCProducts='"&sCProducts&"',stringCValues='"&sCValues&"',stringCCategories='"&sCCategories&"' WHERE idconfigSession="& pcv_chkIDConfig
	set rsA=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsA=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rsA=nothing
	
	'Recalculate Products Total Price
	
	
	query="select configProductCategory,requiredCategory,configProduct,Wprice,price,multiSelect from configSpec_products where specProduct=" & pcv_Midproduct & " and cdefault=1"
	set rsB=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsB=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	do while not rsB.eof
		conCAT=rsB("configProductCategory")
		RCAT=rsB("requiredCategory")
		if (instr(sCategories,conCAT & ",")=0) or (rsB("multiSelect")=-1) then
			ItemWPrice=cdbl(rsB("Wprice"))
			If (Session("customerType")=1) and (ItemWPrice>0) Then
				RmvItems=RmvItems+cdbl(ItemWPrice)
			else
				RmvItems=RmvItems+cdbl(rsB("price"))
			end if
		end if
		rsB.movenext
	loop
	set rsB=nothing
	
	'Recalculate BTO Default Price
	query="SELECT categories.idCategory, categories.categoryDesc, products.idProduct, products.description, configSpec_products.configProductCategory, configSpec_products.price, configSpec_products.Wprice, products.weight FROM (configSpec_products INNER JOIN products ON configSpec_products.configProduct = products.idProduct) INNER JOIN categories ON configSpec_products.configProductCategory = categories.idCategory WHERE (((configSpec_products.specProduct)="&pcv_chkIDProduct&") AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort,products.description;"
	set rsA=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsA=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if NOT rsA.eof then 
		iAddDefaultPrice=Cdbl(0)
		do until rsA.eof
			ItemWPrice=rsA("Wprice")
			If (Session("customerType")=1) and (ItemWPrice>0) Then
				iAddDefaultPrice=Cdbl(iAddDefaultPrice+rsA("Wprice"))
			else
				iAddDefaultPrice=Cdbl(iAddDefaultPrice+rsA("price"))
			End if
		rsA.moveNext
		loop
		set rsA=nothing
		pPrice=Cdbl(pPrice+iAddDefaultPrice)
	end if
	set rsA=nothing
'End Recalculate Default Price

	
	pfPrice=pPrice-RmvItems

	'Recalculate BTO Item Prices
	query="SELECT idProduct, dtCreated, stringProducts, stringValues, stringCategories,stringQuantity,stringPrice, stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pcv_chkIDConfig
	set rsA=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsA=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	pIdProduct=rsA("idProduct")
	pdtCreated=rsA("dtCreated")
	pstringProducts = rsA("stringProducts")
	pstringValues = rsA("stringValues")
	pstringCategories = rsA("stringCategories")
	pstringQuantity = rsA("stringQuantity")
	pstringPrice = rsA("stringPrice")	
	pstringCProducts = rsA("stringCProducts")
	pstringCValues = rsA("stringCValues")
	pstringCCategories = rsA("stringCCategories")
	
	set rsA=nothing

	ArrProduct = Split(pstringProducts, ",")
	ArrValue = Split(pstringValues, ",")
	ArrCategory = Split(pstringCategories, ",")
	ArrQuantity = Split(pstringQuantity, ",")
	ArrPrice = Split(pstringPrice, ",")
	ArrCProduct = Split(pstringCProducts, ",")
	ArrCValue = Split(pstringCValues, ",")
	ArrCCategory = Split(pstringCCategories, ",")
	
	pQty=1
	
	pfPrice=pfPrice*pQty
	
	customizations=0
	for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
		if (ArrQuantity(i)-1)>=0 then
			UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
		else
			UPrice=0
		end if
		UPrice= UPrice + ArrValue(i)
		customizations=customizations+UPrice
	next
	pfPrice=pfPrice+(pQty*customizations)
	
	itemsDiscounts=0
	for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
		query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i)
		set rsA=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsA=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		TempDiscount=0
		do while not rsA.eof
			QFrom=rsA("quantityFrom")
			QTo=rsA("quantityUntil")
			DUnit=rsA("discountperUnit")
			QPercent=rsA("percentage")
			DWUnit=rsA("discountperWUnit")
			if (DWUnit=0) and (DUnit>0) then
				DWUnit=DUnit
			end if
			

			TempD1=0
			if (clng(ArrQuantity(i)*pQty)>=clng(QFrom)) and (clng(ArrQuantity(i)*pQty)<=clng(QTo)) then
				if QPercent="-1" then
					if session("customerType")=1 then
						TempD1=ArrQuantity(i)*pQty*ArrPrice(i)*0.01*DWUnit
					else
						TempD1=ArrQuantity(i)*pQty*ArrPrice(i)*0.01*DUnit
					end if
				else
					if session("customerType")=1 then
						TempD1=ArrQuantity(i)*pQty*DWUnit
					else
						TempD1=ArrQuantity(i)*pQty*DUnit
					end if
				end if
			end if
			TempDiscount=TempDiscount+TempD1
			rsA.movenext
		loop
		set rsA=nothing
		itemsDiscounts=ItemsDiscounts+TempDiscount
	next			

'pfPrice=pfPrice-ItemsDiscounts
ItemsDiscounts=ItemsDiscounts*(-1)

Charges=0

IF gmode=0 THEN	
	for i = lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
		UPrice=ArrCValue(i)
		Charges=Charges+UPrice
	next

	pfPrice=pfPrice+Charges
	
END IF

END IF
'************** End of BTO Configuration ****************

'******* Products Options ************************
IF gmode=0 THEN
	
Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc, pcv_strSelectedOptions
Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	IF len(pIDOptionArray)>0 AND pIDOptionArray<>"NULL" THEN 
	
		pcv_strSelectedOptions = pIDOptionArray
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
	
	END IF
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
	
	pfPrice = pfPrice+cdbl(pcv_strOptionsPriceTotal)

END IF
'******* End of Products Options ************************

updPrices=pfPrice
	
END Function %>
