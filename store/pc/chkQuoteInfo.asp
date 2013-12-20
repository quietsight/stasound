<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.


SUB updQuoteInfo(pcv_chkIdQuote,pcv_chkIDProduct,pcv_chkIDConfig)

query="SELECT * FROM configWishlistSessions where idProduct=" & pcv_chkIDProduct &" and idconfigWishlistSession=" & pcv_chkIDConfig
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
	tmpconfigProduct=0
	pcv_minqty=1
	query="select configProduct,price,Wprice from configSpec_products where configProductCategory=" & ArrCategory(i) & " and cdefault<>0 and specProduct=" & pcv_Midproduct
	set rsB=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsB=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	if not rsB.eof then
	tmpconfigProduct=rsB("configProduct")
	ItemWPrice=cdbl(rsB("Wprice"))
	If (Session("customerType")=1) and (ItemWPrice<>0) Then
		DefaultPrice=ItemWPrice
	else
		DefaultPrice=cdbl(rsB("price"))
	end if
	else
	myUpdateQuote=True
	end if
	
	if tmpconfigProduct<>0 then
		query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & tmpconfigProduct & ";"
		set rsB=conntemp.execute(query)
		if not rsB.eof then
			pcv_minqty=rsB("pcprod_minimumqty")
			if IsNull(pcv_minqty) or pcv_minqty="" then
				pcv_minqty=1
			end if
			if pcv_minqty="0" then
				pcv_minqty=1
			end if
		end if
		set rsB=nothing
	end if
	
	DefaultPrice=DefaultPrice*pcv_minqty
	
	NewPrice=0
	query="select price,Wprice,multiSelect from configSpec_products where configProductCategory=" & ArrCategory(i) & " and configProduct=" & ArrProduct(i) & " and specProduct=" & pcv_Midproduct
	set rsB=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsB=nothing
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
	if clng(ArrProduct(i))=clng(tmpconfigProduct) then
		NewPrice=NewPrice*pcv_minqty
	end if
	if (cdbl(NewPrice)-cdbl(DefaultPrice)<>cdbl(ArrValue(i))) and (multiSelect<>"") and (multiSelect<>-1) then
		NPrice=cdbl(NewPrice)-cdbl(DefaultPrice)
		sValues=sValues & NPrice & ","
		myUpdateQuote=True
	else
	if (cdbl(NewPrice)<>cdbl(ArrValue(i))) and (multiSelect=-1) and (clng(ArrProduct(i))<>clng(tmpconfigProduct)) then
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
	

next

	query="select configProductCategory,requiredCategory,configProduct,Wprice,price,multiSelect from configSpec_products where specProduct=" & pcv_Midproduct & " and cdefault<>0"
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
		query="SELECT categories.idcategory,products.idproduct FROM categories,products,categories_products WHERE categories.idCategory=" & conCAT & " and products.idproduct=" & rsB("configProduct") & " and products.removed=0 and products.active=-1 and categories_products.idCategory=" & conCAT & " and categories_products.idproduct=" & rsB("configProduct") 
		set rsC=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsC=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		IF NOT rsC.eof THEN
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
		END IF
	end if
	rsB.movenext
	loop
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
	query="select price,Wprice from configSpec_Charges where configProductCategory=" & ArrCCategory(i) & " and cdefault<>0 and specProduct=" & pcv_Midproduct
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
	if (cdbl(NewPrice)<>cdbl(ArrCValue(i))) and (multiSelect<>"") and (multiSelect<>-1) then
		sCValues=sCValues & NewPrice & ","
		myUpdateQuote=True
	else
		NPrice=cdbl(ArrCValue(i))
		sCValues=sCValues & NPrice & ","
	end if
	else	
		myUpdateQuote=True
	end if
	

next

	query="select configProductCategory,requiredCategory,configProduct,Wprice,price,multiSelect from configSpec_Charges where specProduct=" & pcv_Midproduct & " and cdefault<>0"
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
	if (instr(sCCategories,conCAT & ",")=0) and (RCAT=True) then
	query="SELECT categories.idcategory,products.idproduct FROM categories,products,categories_products WHERE categories.idCategory=" & conCAT & " and products.idproduct=" & rsB("configProduct") & " and products.removed=0 and products.active=-1 and categories_products.idCategory=" & conCAT & " and categories_products.idproduct=" & rsB("configProduct") 
	set rsC=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsC=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	IF NOT rsC.eof THEN
	myUpdateQuote=True
	sCProducts=sCProducts & rsB("configProduct") & ","
	sCCategories=sCCategories & rsB("configProductCategory") & ","
	If (Session("customerType")=1) and (rsB("Wprice")<>0) Then
	sCValues=sCValues & rsB("Wprice") & ","
	else
	sCValues=sCValues & rsB("price") & ","
	end if
	END IF
	end if
	rsB.movenext
	loop

END IF

if sCProducts="" then
sCProducts="na"
sCValues="na"
sCCategories="na"
end if

'End check Additional Charges

'Update Quote Infor
IF myUpdateQuote=True then
	query="UPDATE configWishlistSessions SET stringProducts='"&sProducts&"',stringValues='"&sValues&"',stringCategories='"&sCategories&"',stringQuantity='" & sQuantity & "',stringPrice='" & sPrice & "',stringCProducts='"&sCProducts&"',stringCValues='"&sCValues&"',stringCCategories='"&sCCategories&"' WHERE idconfigWishlistSession="& pcv_chkIDConfig
	set rsA=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsA=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if	
	'Recalculate Quote Total Price

	query="select configSpec_products.configProductCategory,configSpec_products.requiredCategory,configSpec_products.configProduct,configSpec_products.Wprice,configSpec_products.price,configSpec_products.multiSelect,products.pcprod_minimumqty from configSpec_products,products where configSpec_products.configProduct=products.idproduct AND specProduct=" & pcv_Midproduct & " and cdefault<>0"
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
	pcv_minqty=rsB("pcprod_minimumqty")
	if IsNull(pcv_minqty) or pcv_minqty="" then
		pcv_minqty=1
	end if
	if pcv_minqty="0" then
		pcv_minqty=1
	end if
	if (instr(sCategories,conCAT & ",")=0) and (RCAT=True) then
	else
	if (instr(sCategories,conCAT & ",")=0) or ((rsB("multiSelect")=-1) AND (instr(sProducts,rsB("configProduct") & ",")=0))then
		ItemWPrice=cdbl(rsB("Wprice"))
		If (Session("customerType")=1) and (ItemWPrice>0) Then
			RmvItems=RmvItems+cdbl(ItemWPrice)*pcv_minqty
		else
			RmvItems=RmvItems+cdbl(rsB("price"))*pcv_minqty
		end if
	end if
	end if
	rsB.movenext
	loop
	
	'Recalculate BTO Default Price
	query="Select pcProd_BTODefaultPrice,pcProd_BTODefaultWPrice from products where IDProduct=" & pcv_chkIDProduct
	set rsA=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsA=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	ProductWPrice=rsA("pcProd_BTODefaultWPrice")
	
	If (Session("customerType")=1) and (ProductWPrice>0) Then
		pPrice=ProductWPrice
	else
		pPrice=rsA("pcProd_BTODefaultPrice")
	end if
	set rsA=nothing
	
'End Recalculate Default Price
	
	pfPrice=pPrice-RmvItems
	pDefaultPrice=pfPrice
	
	'Recalculate BTO Item Prices
	query="SELECT idProduct, dtCreated, xfdetails, fPrice, dPrice, stringProducts, stringValues, stringCategories,stringQuantity,stringPrice, stringCProducts, stringCValues, stringCCategories,pcconf_Quantity FROM configWishlistSessions WHERE idconfigWishlistSession=" & pcv_chkIDConfig
	set rsA=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsA=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	pIdProduct=rsA("idProduct")
	pdtCreated=rsA("dtCreated")
	pxfdetails=rsA("xfdetails") 
	pstringProducts = rsA("stringProducts")
	pstringValues = rsA("stringValues")
	pstringCategories = rsA("stringCategories")
	pstringQuantity = rsA("stringQuantity")
	pstringPrice = rsA("stringPrice")	
	pstringCProducts = rsA("stringCProducts")
	pstringCValues = rsA("stringCValues")
	pstringCCategories = rsA("stringCCategories")
	pQuantity=rsA("pcconf_Quantity")
	if (pQuantity<>"") then
	else
	pQuantity="1"
	end if
	ArrProduct = Split(pstringProducts, ",")
	ArrValue = Split(pstringValues, ",")
	ArrCategory = Split(pstringCategories, ",")
	ArrQuantity = Split(pstringQuantity, ",")
	ArrPrice = Split(pstringPrice, ",")
	ArrCProduct = Split(pstringCProducts, ",")
	ArrCValue = Split(pstringCValues, ",")
	ArrCCategory = Split(pstringCCategories, ",")
	
	pfPrice=pfPrice*pQuantity
	
	customizations=0
	for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
		pcv_minqty=1
		query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pIdProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			query="SELECT products.pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcv_minqty=rsQ("pcprod_minimumqty")
				if IsNull(pcv_minqty) or pcv_minqty="" then
					pcv_minqty=1
				end if
				if pcv_minqty="0" then
					pcv_minqty=1
				end if
			else
				pcv_minqty=1
			end if
			set rsQ=nothing
		end if
		set rsQ=nothing
		if (ArrQuantity(i)-pcv_minqty)>=0 then
			UPrice=(ArrQuantity(i)-pcv_minqty)*ArrPrice(i)
		else
			UPrice=0
		end if
		UPrice= UPrice + ArrValue(i)
		customizations=customizations+UPrice
	next
	pfPrice=pfPrice+(pQuantity*customizations)
	
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
				if (clng(ArrQuantity(i)*pQuantity)>=clng(QFrom)) and (clng(ArrQuantity(i)*pQuantity)<=clng(QTo)) then
				if QPercent="-1" then
				if session("customerType")=1 then
				TempD1=ArrQuantity(i)*pQuantity*ArrPrice(i)*0.01*DWUnit
				else
				TempD1=ArrQuantity(i)*pQuantity*ArrPrice(i)*0.01*DUnit
				end if
				else
				if session("customerType")=1 then
				TempD1=ArrQuantity(i)*pQuantity*DWUnit
				else
				TempD1=ArrQuantity(i)*pQuantity*DUnit
				end if
				end if
				end if
				TempDiscount=TempDiscount+TempD1
				rsA.movenext
		loop
	itemsDiscounts=ItemsDiscounts+TempDiscount
	next			

if pcQDiscountType<>"1" then
pfPrice1=pfPrice-(itemsDiscounts/pQuantity)
else
pfPrice1=pDefaultPrice
end if

pfPrice=pfPrice-ItemsDiscounts
ItemsDiscounts=ItemsDiscounts*(-1)

	Charges=0
	
	for i = lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
		UPrice=ArrCValue(i)
		Charges=Charges+UPrice
	next

pfPrice=pfPrice+Charges

'*********** Check Quantity Discounts **************
query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &pcv_chkIDProduct& " AND quantityFrom<=" &pQuantity& " AND quantityUntil>=" &pQuantity
set rsDisc=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsDisc=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
QDiscounts=0

if not rsDisc.eof then

 	pDiscountPerUnit = rsDisc("discountPerUnit")
 	pDiscountPerWUnit = rsDisc("discountPerWUnit")
 	pPercentage = rsDisc("percentage")

 	if session("customerType")<>1 then
 		if pPercentage = "0" then 
			QDiscounts = pDiscountPerUnit * pQuantity
		else
			QDiscounts = (pDiscountPerUnit/100) * pfPrice1
		end if
	else
		if pPercentage = "0" then 
			QDiscounts = pDiscountPerWUnit * pQuantity
		else
			QDiscounts = (pDiscountPerWUnit/100) * pfPrice1
		end if
	end if
end if

pfPrice=pfPrice-QDiscounts

'*********** End of Check Quantity Discounts *******

query="UPDATE configWishlistSessions SET fPrice="&pfPrice&",dPrice=" & ItemsDiscounts & ",pcconf_QDiscount=" & QDiscounts & " WHERE idconfigWishlistSession=" & pcv_chkIDConfig
set rsA=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsA=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

END IF
'End Update Quote Infor
	
END SUB %>