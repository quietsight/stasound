<%
' findProduct
function findProduct(pcCartArray, indexCart, pIdProduct)
 
 dim f 
 findProduct=Cint(0) 
 if indexCart>0 then
  for f=1 to indexCart
    if pcCartArray(f,10)=0 and int(pcCartArray(f,0))=int(pIdProduct) then
       findProduct=-1
    end if
  next 
 end if  
 set f=nothing   
end function

'check for oversized product
' findProduct
function oversizecheck(pcCartArray, indexCart)
	dim f 
	if indexCart>0 then
		oversizecheck=""
		for f=1 to indexCart
			If pcCartArray(f,20)="-1" Then
				'response.write ""
			Else
				if pcCartArray(f,10)=0 and pcCartArray(f,23)<>"NO" then
					OSArray=split(pcCartArray(f,23),"||")
					if ubound(OSArray)>3 then
						for i=1 to pcCartArray(f,2)
							oversizecheck=oversizecheck&"1|||"&pcCartArray(f,23)&"||"&pcCartArray(f,6)&"||"&pcCartArray(f,3)&","
						next
					end if
				end if
				if pcCartArray(f,16)<>"" then
					'//Get config info
					query="SELECT stringProducts, stringPrice, stringQuantity FROM configSessions WHERE idconfigSession="&pcCartArray(f,16)&";"
					set rsBTOChkObj=server.CreateObject("ADODB.RecordSet")
					set rsBTOChkObj=conntemp.execute(query)
					pcv_itemString = rsBTOChkObj(0)
					pcv_itemPrice = rsBTOChkObj(1)
					pcv_itemQty = rsBTOChkObj(2)
					pcv_itemStringArry = split(pcv_itemString,",")
					pcv_itemPriceArry = split(pcv_itemPrice,",")
					pcv_itemQtyArry = split(pcv_itemQty,",")
					
					for iOSChkCnt=lbound(pcv_itemStringArry) to ubound(pcv_itemStringArry)-1
						pcv_tempPrdChk = pcv_itemStringArry(iOSChkCnt)
						query="SELECT weight, oversizeSpec FROM products WHERE idProduct="&pcv_tempPrdChk&";"
						set rsOSChkObj=server.CreateObject("ADODB.RecordSet")
						set rsOSChkObj=conntemp.execute(query)
						pcv_OSChkWeight=rsOSChkObj(0)
						pcv_OSChkSpec=rsOSChkObj(1)
						if pcv_OSChkSpec<>"NO" then
							for i=1 to pcCartArray(f,2)
								for iICnt=1 to pcv_itemQtyArry(iOSChkCnt)
									'//oversized, get array
									oversizecheck=oversizecheck&"2|||"&pcv_OSChkSpec&"||"&pcv_OSChkWeight&"||"&pcv_itemPriceArry(iOSChkCnt)&","
								next
							next
						end if
					next
				end if
			End If
		next
	end if  
	set f=nothing   
end function

'If Editing an existing order
function eoversizecheck(pcOSCheckOrderNumber)
	dim f 
	query="SELECT ProductsOrdered.quantity, products.OverSizeSpec, products.weight, ProductsOrdered.idOrder FROM ProductsOrdered INNER JOIN products ON ProductsOrdered.idProduct = products.idProduct WHERE (((ProductsOrdered.idOrder)="&pcOSCheckOrderNumber&"));"
	set rsOSC=server.CreateObject("ADODB.RecordSet")
	set rsOSC=connTemp.execute(query)
	eoversizecheck=""
	do until rsOSC.eof
		pcOS_Quantity=rsOSC("quantity")
		pcOS_OverSizeSpec=rsOSC("OverSizeSpec")
		pcOS_weight=rsOSC("weight")
		if pcOS_OverSizeSpec<>"NO" then
			OSArray=split(pcOS_OverSizeSpec,"||")
			if ubound(OSArray)>3 then
				for i=1 to pcOS_Quantity
					eoversizecheck=eoversizecheck&"1|||"&pcOS_OverSizeSpec&"||"&pcOS_weight&","
				next
			end if
		end if
		rsOSC.MoveNext
	loop
	set rsOSC=nothing
	set f=nothing   
end function

' count cart Rows
function countCartRows(pcCartArray, indexCart)
 
 dim cont, f
 
 cont=Cint(0)
 if indexCart>0 then
  for f=1 to indexCart
    if pcCartArray(f,10)=0 then
     cont=cont+1
    end if
  next
 else
  cont=0
 end if
 
 countCartRows=cont
 set f=nothing 
 set cont=nothing
 
end function

' Cart Amount
function calculateCartTotal(pcCartArray, indexCart)
	dim f, total
	'SB S
	Dim subInstArr
	'SB E
	total=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then  
			'SB S
			If Not (len(pcCartArray(f,38))>0) Then pcCartArray(f,38)=0
		    if pcCartArray(f,38) > 0 then 
				subInstArr = split(getSubInstallVals(pcCartArray(f,38)),",")
			else
			    subInstArr = split("0,0,0,0",",") 
			end if 
			'SB E
			if pcCartArray(f,16)<>"" then
				'SB S
			 	if subInstArr(2) = "1" Then 
			   		'// Trial price no discounts 
					total = total + (pcCartArray(f,2) * cdbl(subInstArr(3)))
			 	else
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
				end if  
				'SB E
			else
				'SB S
		    	if subInstArr(2) = "1" Then 
			   		'// Trial price no discounts 
					total = total + (pcCartArray(f,2)* cdbl(subInstArr(3)))
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3))  
				end if  
			end if
				'SB E
			end if
			if (pcCartArray(f,27)>"0") AND (pcCartArray(f,28)>"0") then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if
		end if
	next
	calculateCartTotal=ccur(total)
	set f=nothing
	set total=nothing 
end function


' Cart Ship Amount
function calculateShipCartTotal(pcCartArray, indexCart)
	dim f, total
	total=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND pcCartArray(f,20)=0 then   
			if pcCartArray(f,16)<>"" then
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3))  
				end if  
			end if  
		end if
	next
	calculateShipCartTotal=total  
	set f=nothing
	set total=nothing 
end function

function calculateCategoryDiscounts(pcCartArray, indexCart)

	Dim TmpProList(100,5)
	for f=1 to indexCart
		TmpProList(f,0)=pcCartArray(f,0)
		TmpProList(f,1)=pcCartArray(f,10)
		TmpProList(f,3)=pcCartArray(f,2)
		TmpProList(f,4)=0
		if pcCartArray(f,10)=0 then
			'Get RowPrice
			pRowPrice=0
			pRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))
			pRowPrice=pRowPrice + ccur(pcCartArray(f,2) * pcCartArray(f,5))	
			if trim(pcCartArray(f,30))<>"" AND trim(pcCartArray(f,30))>"0" then
				pRowPrice=pRowPrice-pcCartArray(f,30)
			end if
			if trim(pcCartArray(f,31))<>"" AND trim(pcCartArray(f,31))>"0" then
				pRowPrice=pRowPrice+ccur(pcCartArray(f,31))
			end if
			if trim(pcCartArray(f,15))<>"" AND trim(pcCartArray(f,15))>"0" then
				pRowPrice=pRowPrice-ccur(pcCartArray(f,15))
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then 
				pRowPrice = ( ccur(pRowPrice) + ccur(TmpProList(cint(pcCartArray(f,27)),2)) ) - ( ( ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28) ) ) * pcCartArray(f,2) )
			end if
			TmpProList(f,2)=pRowPrice
		end if
	next
			
	' ------------------------------------------------------
	' START - Calculate category-based quantity discounts
	' ------------------------------------------------------
	CatDiscTotal=0
	
	query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
	set rsCDObj=server.CreateObject("ADODB.RecordSet")
	set rsCDObj=conntemp.execute(query)

	Do While not rsCDObj.eof
		CatSubQty=0
		CatSubTotal=0
		CatSubDiscount=0
	
		For f=1 to indexCart
			if (TmpProList(f,1)=0) and (TmpProList(f,4)=0) then 
				query="select idproduct from categories_products where idcategory=" & rsCDObj("IDCat") & " and idproduct=" & TmpProList(f,0)
				set rsCDObjtemp=server.CreateObject("ADODB.RecordSet")
				set rsCDObjtemp=connTemp.execute(query)
				
				if not rsCDObjtemp.eof then
					CatSubQty=CatSubQty+TmpProList(f,3)
					CatSubTotal=CatSubTotal+TmpProList(f,2)
					TmpProList(f,4)=1
				end if
				set rsCDObjtemp=nothing
			end if
		Next

		if CatSubQty>0 then
			query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rsCDObj("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
			set rsCDObjtemp=server.CreateObject("ADODB.RecordSet")
			set rsCDObjtemp=conntemp.execute(query)
	
			if not rsCDObjtemp.eof then
				' there are quantity discounts defined for that quantity 
				pDiscountPerUnit=rsCDObjtemp("pcCD_discountPerUnit")
				pDiscountPerWUnit=rsCDObjtemp("pcCD_discountPerWUnit")
				pPercentage=rsCDObjtemp("pcCD_percentage")
				pbaseproductonly=rsCDObjtemp("pcCD_baseproductonly")
				set rsCDObjtemp=nothing
				
				if session("customerType")<>1 then  'customer is a normal user
					if pPercentage="0" then 
						CatSubDiscount=pDiscountPerUnit*CatSubQty
					else
						CatSubDiscount=(pDiscountPerUnit/100) * CatSubTotal
					end if
				else  'customer is a wholesale customer
					if pPercentage="0" then 
						CatSubDiscount=pDiscountPerWUnit*CatSubQty
					else
						CatSubDiscount=(pDiscountPerWUnit/100) * CatSubTotal
					end if
				end if
			end if
		end if

		CatDiscTotal=CatDiscTotal+CatSubDiscount
		rsCDObj.MoveNext
		loop
		set rsCDObj=nothing				
		'// Round the Category Discount to two decimals
		if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
			CatDiscTotal = Round(CatDiscTotal,2)
		end if
		' ------------------------------------------------------
		' END - Calculate category-based quantity discounts
		' ------------------------------------------------------
	calculateCategoryDiscounts=CatDiscTotal
end function



Function CheckTaxEpt(pcv_IDPro,pcv_StateCode)
	Dim rsD,pcArrayD,intCountD,pcv_mc
	
	query="SELECT pcTEpt_StateCode FROM pcTaxEpt WHERE pcTEpt_StateCode='" & pcv_StateCode & "' AND ((pcTEpt_ProductList like '" & pcv_IDPro & ",%') OR (pcTEpt_ProductList like '%," & pcv_IDPro & ",%'))"
	set rsD=connTemp.execute(query)
	
	if not rsD.eof then
		set rsD=nothing
		CheckTaxEpt=1
		Exit Function
	end if
	set rsD=nothing
	
	query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcv_IDPro & ";"
	set rsD=connTemp.execute(query)
	
	if not rsD.eof then
		pcArrayD=rsD.getRows()
		intCountD=ubound(pcArrayD,2)
		set rsD=nothing
		
		For pcv_mc=0 to intCountD
			query="SELECT pcTEpt_StateCode FROM pcTaxEpt WHERE pcTEpt_StateCode='" & pcv_StateCode & "' AND ((pcTEpt_CategoryList like '" & pcArrayD(0,pcv_mc) & ",%') OR (pcTEpt_CategoryList like '%," & pcArrayD(0,pcv_mc) & ",%'))"
			set rsD=connTemp.execute(query)
			if not rsD.eof then
				set rsD=nothing
				CheckTaxEpt=1
				Exit Function
			end if
		Next
	end if
set rsD=nothing

End Function

Function CheckTaxEptZone(pcv_IDPro,pcv_ZoneRateID)
	Dim query, rsD,pcArrayD,intCountD,pcv_mc
	CheckTaxEptZone=0
	
	query="SELECT pcTaxEpt.pcTaxZoneRate_ID FROM pcTaxEpt WHERE pcTaxEpt.pcTaxZoneRate_ID=" & pcv_ZoneRateID & " AND ((pcTEpt_ProductList like '" & pcv_IDPro & ",%') OR (pcTEpt_ProductList like '%," & pcv_IDPro & ",%'))"
	set rsD=connTemp.execute(query)
	
	if not rsD.eof then
		set rsD=nothing
		CheckTaxEptZone=1
		Exit Function
	end if
	set rsD=nothing
	
	query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcv_IDPro & ";"
	set rsD=connTemp.execute(query)
	if not rsD.eof then
		pcArrayD=rsD.getRows()
		intCountD=ubound(pcArrayD,2)
		set rsD=nothing
		
		For pcv_mc=0 to intCountD
			query="SELECT pcTaxZoneRate_ID FROM pcTaxEpt WHERE pcTaxZoneRate_ID=" & pcv_ZoneRateID & " AND ((pcTEpt_CategoryList like '" & pcArrayD(0,pcv_mc) & ",%') OR (pcTEpt_CategoryList like '%," & pcArrayD(0,pcv_mc) & ",%'))"
			set rsD=connTemp.execute(query)
			if not rsD.eof then
				set rsD=nothing
				CheckTaxEptZone=1
				Exit Function
			end if
		Next
	end if
set rsD=nothing

End Function

function NontaxableItems(tmpIDConfigSession,parentQty)
	Dim rs,query,i,tmpResult
	Dim stringProducts,stringValues,stringCategories,Qstring,Pstring,ArrProduct,ArrValue,ArrCategory,ArrQuantity,ArrPrice
	Dim TempDiscount,TempD1,QFrom,QTo,DUnit,QPercent,DWUnit

	tmpResult=0

	IF tmpIDConfigSession<>"" AND IsNumeric(tmpIDConfigSession) THEN
		query="SELECT stringProducts, stringValues, stringCategories,stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & tmpIDConfigSession
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
							
		if not rs.eof then
			stringProducts=rs("stringProducts")
			stringValues=rs("stringValues")
			stringCategories=rs("stringCategories")
			Qstring=rs("stringQuantity")
			Pstring=rs("stringPrice")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			ArrQuantity=Split(Qstring,",")
			ArrPrice=split(Pstring,",")
			set rs=nothing
						
			if ArrProduct(0)="na" then
			else
				For i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					if ArrProduct(i)<>"" then
						query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i) & " AND notax<>0;"
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpResult=tmpResult+ArrQuantity(i)*ArrPrice(i)*parentQty
							
							query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i)
							set rs=connTemp.execute(query)
		 
							TempDiscount=0
							do while not rs.eof
								QFrom=rs("quantityFrom")
								QTo=rs("quantityUntil")
								DUnit=rs("discountperUnit")
								QPercent=rs("percentage")
								DWUnit=rs("discountperWUnit")
								if (DWUnit=0) and (DUnit>0) then
									DWUnit=DUnit
								end if
								

								TempD1=0
								if (clng(ArrQuantity(i)*parentQty)>=clng(QFrom)) and (clng(ArrQuantity(i)*parentQty)<=clng(QTo)) then
									if QPercent="-1" then
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DUnit
										end if
									else
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*DUnit
										end if
									end if
								end if
								TempDiscount=TempDiscount+TempD1
								rs.movenext
							loop
							set rs=nothing
							tmpResult=tmpResult-TempDiscount
							
						end if
						set rs=nothing
					end if
				Next
			end if
		end if
		set rs=nothing
	END IF

	NontaxableItems=tmpResult

End function

function NontaxableZoneItems(tmpIDConfigSession,parentQty,zoneRateID)

	Dim rs,query,i,tmpResult
	Dim stringProducts,stringValues,stringCategories,Qstring,Pstring,ArrProduct,ArrValue,ArrCategory,ArrQuantity,ArrPrice
	Dim TempDiscount,TempD1,QFrom,QTo,DUnit,QPercent,DWUnit

	tmpResult=0

	IF tmpIDConfigSession<>"" AND IsNumeric(tmpIDConfigSession) THEN
		query="SELECT stringProducts, stringValues, stringCategories,stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & tmpIDConfigSession
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
							
		if not rs.eof then
			stringProducts=rs("stringProducts")
			stringValues=rs("stringValues")
			stringCategories=rs("stringCategories")
			Qstring=rs("stringQuantity")
			Pstring=rs("stringPrice")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			ArrQuantity=Split(Qstring,",")
			ArrPrice=split(Pstring,",")
			set rs=nothing
						
			if ArrProduct(0)="na" then
			else
				For i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					if ArrProduct(i)<>"" then
						ztest=CheckTaxEptZone(ArrProduct(i),zoneRateID)
						
						if ztest=1 then
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i)
						else
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i) & " AND notax<>0;"
						end if
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpResult=tmpResult+ArrQuantity(i)*ArrPrice(i)*parentQty
							
							query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i)
							set rs=connTemp.execute(query)
		 
							TempDiscount=0
							do while not rs.eof
								QFrom=rs("quantityFrom")
								QTo=rs("quantityUntil")
								DUnit=rs("discountperUnit")
								QPercent=rs("percentage")
								DWUnit=rs("discountperWUnit")
								if (DWUnit=0) and (DUnit>0) then
									DWUnit=DUnit
								end if
								

								TempD1=0
								if (clng(ArrQuantity(i)*parentQty)>=clng(QFrom)) and (clng(ArrQuantity(i)*parentQty)<=clng(QTo)) then
									if QPercent="-1" then
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*ArrPrice(i)*0.01*DUnit
										end if
									else
										if session("customerType")=1 then
											TempD1=ArrQuantity(i)*parentQty*DWUnit
										else
											TempD1=ArrQuantity(i)*parentQty*DUnit
										end if
									end if
								end if
								TempDiscount=TempDiscount+TempD1
								rs.movenext
							loop
							set rs=nothing
							tmpResult=tmpResult-TempDiscount
							
						end if
						set rs=nothing
					end if
				Next
			end if
		end if
		set rs=nothing
	END IF


	IF tmpIDConfigSession<>"" AND IsNumeric(tmpIDConfigSession) THEN
		query="SELECT stringQuantity, stringPrice, stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & tmpIDConfigSession
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
							
		if not rs.eof then
			stringProducts=rs("stringCProducts")
			stringValues=rs("stringCValues")
			stringCategories=rs("stringCCategories")
			Qstring=rs("stringQuantity")
			Pstring=rs("stringPrice")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			ArrQuantity=Split(Qstring,",")
			ArrPrice=split(Pstring,",")
			set rs=nothing
						
			if ArrProduct(0)="na" then
			else
				For i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					if ArrProduct(i)<>"" then
						ztest=CheckTaxEptZone(ArrProduct(i),zoneRateID)
						
						if ztest=1 then
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i)
						else
							query="SELECT idproduct FROM Products WHERE idproduct=" & ArrProduct(i) & " AND notax<>0;"
						end if
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpResult=tmpResult+1*ArrValue(i)
						end if
						set rs=nothing
					end if
				Next
			end if
		end if
		set rs=nothing
	END IF

	NontaxableZoneItems=tmpResult

End function

' Cart Taxable Amount
function calculateTaxableTotal(pcCartArray, indexCart)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	'SB S
	Dim subInstArr
	'SB E
	
	total=0

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			'SB S
			If Not (len(pcCartArray(f,38))>0) Then pcCartArray(f,38)=0
			if pcCartArray(f,38) > 0 then 
				subInstArr = split(getSubInstallVals(pcCartArray(f,38)),",")
			else
				subInstArr = split("0,0,0,0",",") 
			end if
			'SB E
			if pcCartArray(f,16)<>"" then
				'SB S
				if subInstArr(2) = "1" Then 
					total = total + (pcCartArray(f,2) * cdbl(subInstArr(3))) - NontaxableItems(pcCartArray(f,16),pcCartArray(f,2))
				else
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31) - NontaxableItems(pcCartArray(f,16),pcCartArray(f,2))
				end if
				'SB E
			else
				'SB S
				If subInstArr(2) = "1" Then
					total = total + (pcCartArray(f,2) * cdbl(subInstArr(3)))
				Else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if
				End If
				'SB E
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if

		end if
	next
	
	calculateTaxableTotal=total
	set f=nothing
	set total=nothing 
end function

'SB S
function calculateTaxableTotal_SB(pcCartArray, indexCart)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31) - NontaxableItems(pcCartArray(f,16),pcCartArray(f,2))
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if

		end if
	next
	
	calculateTaxableTotal_SB=total
	set f=nothing
	set total=nothing 
end function
'SB E

' Cart Taxable Amount
function calculateTaxableZoneTotal(pcCartArray, indexCart, zoneRateID)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		mtest=0
		
		mtest=CheckTaxEptZone(pcCartArray(f,0),zoneRateID)

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31) - NontaxableZoneItems(pcCartArray(f,16),pcCartArray(f,2),zoneRateID)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total=total + ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if

		end if
	next
	
	calculateTaxableZoneTotal=total

	set f=nothing
	set total=nothing 
end function


function checkTaxExempt(pcCartArray, indexCart, TaxZoneRateID)
	Dim f, checkVar, pcv_StateCode, pcv_Country
	
	strCheckVar=""

	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND (pcv_StateCode<>"") and (ucase(pcv_Country)="CA") then
			strCheckVar=strCheckVar&CheckTaxEptZone(pcCartArray(f,0),TaxZoneRateID)&","
		end if
	next
	
	if instr(strCheckVar,"1") then
		if instr(strCheckVar,"0") then
			checkVar=0
		else
			checkVar=1
		end if
	else
		checkVar=0
	end if
	
	checkTaxExempt=checkVar
end function

' Cart Weight
function calculateCartWeight(pcCartArray, indexCart)
	dim f, totalWeight
	totalWeight=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then   
			totalWeight=totalWeight + (pcCartArray(f,6)*pcCartArray(f,2))
		end if
	next  
	if cdbl(totalWeight)>0 AND cdbl(totalWeight)<1 then
		totalWeight=1
	end if
	totalWeight=round(totalWeight,0)
	calculateCartWeight=totalWeight  
	set f=nothing
	set totalWeight=nothing
end function

' Cart Weight
function calculateShipWeight(pcCartArray, indexCart)
	dim f, totalWeight
	totalWeight=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND pcCartArray(f,20)=0 then
			totalWeight=totalWeight + (pcCartArray(f,6)*pcCartArray(f,2))
		end if
	next  
	if cdbl(totalWeight)>0 AND cdbl(totalWeight)<1 then
		totalWeight=1
	end if
	totalWeight=round(totalWeight,0) 
	calculateShipWeight=totalWeight  
	set f=nothing
	set totalWeight=nothing
end function

' Cart Surcharge
function calculateTotalProductSurcharge(pcCartArray, indexCart)
	dim f, totalSurcharge
	dim fQty, fSurcharge1, fSurcharge2
	totalSurcharge=0
	'Create a new temporary array to group like ProductIDs regardless of option
	Dim SCArray(100,5)
	G = Cint(0)
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then
			if G>0 then
				var_update = 0
				for h = 0 to G
					if SCArray(h,0)=pcCartArray(f,0) then 
						SCArray(h,1)=Cint(SCArray(h,1))+Cint(pcCartArray(f,2))
						var_update = 1
					end if
				next
				if var_update = 0 then
					SCArray(G,0)=pcCartArray(f,0)
					SCArray(G,1)=pcCartArray(f,2)
					SCArray(G,2)=pcCartArray(f,36)
					SCArray(G,3)=pcCartArray(f,37)
				end if
			else
				SCArray(G,0)=pcCartArray(f,0)
				SCArray(G,1)=pcCartArray(f,2)
				SCArray(G,2)=pcCartArray(f,36)
				SCArray(G,3)=pcCartArray(f,37)
			end if
			G = G + 1
		end if
	Next
	
	for t=0 to G
		fQty = SCArray(t,1) 'quantity ordered of product
		fSurcharge1 = SCArray(t,2) 'Initial Surcharge
		fSurcharge2 = SCArray(t,3) 'Additional Surcharge
		totalSurcharge=ccur(totalSurcharge) + ccur(SCArray(t,2))
		if fQty > 1 AND fSurcharge2>0 then
			totalSurcharge=ccur(totalSurcharge) + (ccur(SCArray(t,3))*(CLng(fQty)-1))
		end if
	next 
	calculateTotalProductSurcharge=totalSurcharge  
	set f=nothing
	set totalSurcharge=nothing
end function

' Cart Product Quantity
function calculateCartQuantity(pcCartArray, indexCart)
	dim f, totalQuantity
	totalQuantity=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 then   
			totalQuantity=totalQuantity + pcCartArray(f,2)
		end if
	next  
	calculateCartQuantity=totalQuantity  
	set f=nothing
	set totalQuantity=nothing
end function
'RP ADDON-S

' Cart Points Rewarded Quantity
function calculateCartRewards(pcCartArray, indexCart)
	dim f, totalQuantity
	totalQuantity = 0
	for f=1 to indexCart
		if pcCartArray(f,10) = 0 then  
			if (pcCartArray(f,22)="") OR IsNull(pcCartArray(f,22)) then
				pcCartArray(f,22)=0
			end if
			totalQuantity = totalQuantity + (pcCartArray(f,2) * cdbl(pcCartArray(f,22)))
			'BTO Additional Charges Reward Points
			if (pcCartArray(f,29)<>"") and (pcCartArray(f,29)<>"0") then
				totalQuantity = totalQuantity + pcCartArray(f,29)
			end if
		end if
	next   
	calculateCartRewards = totalQuantity  
	set f			= nothing
	set totalQuantity	= nothing
end function
'RP ADDON-E

' Cart Product Quantity
function calculateCartShipQuantity(pcCartArray, indexCart)
	dim f, totalShipQuantity
	totalShipQuantity=0
	for f=1 to indexCart
		if pcCartArray(f,10)=0 AND pcCartArray(f,20)=0 then   
			totalShipQuantity=totalShipQuantity + pcCartArray(f,2)
		end if
	next  
	calculateCartShipQuantity=totalShipQuantity  
	set f=nothing
	set totalShipQuantity=nothing
end function


' Check session lost
function checkSessionLost(pcCartArray, pcCartIndex) 
	if pcCartIndex="" then
		' session is lost, initialize all variables
		Session.Timeout=25
		Session("idCustomer")=Cint(0)
		Session("language")=Cstr("english")
		Session("pcCartIndex")=Cint(0)
		ReDim pcCartArray(50, 18)
		Session("pcCartSession")=pcCartArray      
		checkSessionLost=1
	else
		checkSessionLost=0
	end if
end function


'////////////////////////////////////////////////////////////////////////////////////////
'// START: VAT CALCULATIONS
'////////////////////////////////////////////////////////////////////////////////////////

'// Removes VAT from UnitPrice when purchased outside the European Union.
Function pcf_VAT(Price, ProductID)
	Dim notax
	If pcv_IsEUMemberState=1 OR pcv_IsEUMemberState=0 Then
		notax="-1"
		If not validNum(ProductID) Then 
			ProductID = 0
		End If		
		If ProductID<>0 Then '// Product			
			query="SELECT products.notax FROM products WHERE idProduct=" & ProductID & " AND configOnly=0 AND removed=0 " 
			set rsVAT=server.CreateObject("ADODB.RecordSet")
			set rsVAT=connTemp.execute(query)
			If NOT rsVAT.eof Then
				notax=rsVAT("notax")
			End If
			set rsVAT=nothing
			If ptaxVAT="1" AND notax <> "-1" Then
				if pcv_IsEUMemberState=0 then
					pcf_VAT=pcf_RemoveVAT(Price, ProductID)
					Exit Function
				else
					pcf_VAT=Price
					Exit Function
				end if
			Else
				pcf_VAT=Price
				Exit Function
			End If
		Else '// VAT Item - No Product ID		
			If ptaxVAT="1" Then
				if pcv_IsEUMemberState=0 then
					pcf_VAT=pcf_RemoveVAT(Price, ProductID)
					Exit Function
				else
					pcf_VAT=Price
					Exit Function
				end if
			Else
				pcf_VAT=Price
				Exit Function
			End If
		End If
	Else '// If pcv_IsEUMemberState=1 OR pcv_IsEUMemberState=0 Then
		pcf_VAT=Price
		Exit Function
	End If
End Function

'// Removes VAT from Price
Function pcf_RemoveVAT(Price,ProductID)
	pcf_RemoveVAT=Price/(1+pcf_VATRate(ptaxVATRate_Code, ProductID)/100)
End Function

'// Determines the correct VAT from a Product ID
Function pcf_VATRate(StateCode,ProductID)
	If StateCode="0" OR StateCode="" Then
		pcf_VATRate=ptaxVATrate '// Default Rate
		Exit Function
	Else
		if not validNum(ProductID) then ProductID = 0
		if ProductID="" OR ProductID=0 OR isNULL(ProductID)=True then
			pcf_VATRate=ptaxVATrate '// Default Rate
			Exit Function
		else
			query="SELECT pcProductsVATRates.pcVATRate_ID, pcVATRates.pcVATRate_Rate, pcVATRates.pcVATRate_ID "
			query=query&"FROM pcProductsVATRates, pcVATRates "
			query=query&"WHERE pcProductsVATRates.pcVATRate_ID=pcVATRates.pcVATRate_ID AND pcProductsVATRates.idProduct="&ProductID	
			Set rsVAT=Server.CreateObject("ADODB.Recordset")  
			set rsVAT=connTemp.execute(query)
			if not rsVAT.eof then
				pcf_VATRate=rsVAT("pcVATRate_Rate") '// Category Rate
			else
				pcf_VATRate=ptaxVATrate '// Default Rate
			end if
			set rsVAT=nothing			
			Exit Function
		end if		
	End If
End Function

'// Determines if a country is apart of the European Union
' 1 = YES
' 0 = NO
' 2 = Not Applicable
Function pcf_IsEUMemberState(CountryCode)
	pcf_IsEUMemberState=2
	If (CountryCode<>"" AND isNULL(CountryCode)=False) Then
		if (UCASE(CountryCode)=UCASE(scCompanyCountry)) AND ptaxVAT="1" then
			pcf_IsEUMemberState=1
			Exit Function
		end if	
	End If
	If ptaxVAT="1" AND (CountryCode<>"" AND isNULL(CountryCode)=False) Then		
		query="SELECT pcVATCountries.pcVATCountry_Code From pcVATCountries WHERE pcVATCountries.pcVATCountry_Code='"&UCASE(CountryCode)&"';"
		set rsVAT=Server.CreateObject("ADODB.Recordset")
		set rsVAT=connTemp.execute(query)
		if not rsVAT.eof then
			pcf_IsEUMemberState=1
		else
			pcf_IsEUMemberState=0
		end if
		set rsVAT=nothing
	End If
End Function


' Cart VAT Total and Remove VAT
function calculateNoVATTotal(pcCartArray, indexCart, TotalStandardDiscount, TotalCategoryDiscount)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0
	grandtotal=0
	
	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		total=0
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3)) 
				end if	 
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total = ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if
		end if

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// This Line Item represents what % of the Total Discount?  							
		Proportional_total = RoundTo((total/ApplicableDisountTotal),.01)
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Line Item after discount
		ApplicableDisount_total = (TotalStandardDiscount * Proportional_total)
		total = (total - ApplicableDisount_total)	
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		if pcv_strrApplicableCategories<>"" then
			pcArray_ApplicableCategories = split(pcv_strrApplicableCategories, ",")
			for y=0 to ubound(pcArray_ApplicableCategories)-1 '// For Each Category Discount Available
				pcArray_ApplicableCategory = split(pcArray_ApplicableCategories(y), chr(124))
				tmpApplicableCategoryID = pcArray_ApplicableCategory(1)
				tmpCategorySubTotal = pcArray_ApplicableCategory(0)
			
				ApplicableCategoryItem=False
				'response.Write("Category: " & tmpApplicableCategoryID & "<br /><br />")
				if pcv_strApplicableProducts <> "" then
					pcArray_ApplicableProducts = split(pcv_strApplicableProducts, ",")
					for x=0 to ubound(pcArray_ApplicableProducts)-1 '// Loop through all Products						
						pcArray_ApplicableProduct = split(pcArray_ApplicableProducts(x), chr(124))
						tmpProductID = pcArray_ApplicableProduct(0)
						tmpCategoryID = pcArray_ApplicableProduct(1)		
						'response.Write("Product: " & tmpProductID & " - " & pcCartArray(f,0) & "<br />")			
						if (tmpProductID = pcCartArray(f,0)) AND (tmpCategoryID = tmpApplicableCategoryID) then '// This Product is Applicable to this Category
							ApplicableCategoryItem=True
							'response.Write("Hit: " & tmpCategorySubTotal & "<br />")
						end if
					next
				end if  '// if pcv_strApplicableProducts <> "" then
				
				If ApplicableCategoryItem=True AND tmpCategorySubTotal>0 Then
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// This Line Item represents what % of the Total Category?  							
				ProportionalCat_total = RoundTo((total/tmpCategorySubTotal),.01)
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// Line Item after Category
				ApplicableCatDisount_total = (TotalCategoryDiscount * ProportionalCat_total)
				total = (total - ApplicableCatDisount_total)	
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				End If
				
			next
		end if '// if pcv_strrApplicableCategories<>"" then	



		total = pcf_RemoveVAT(total, pcCartArray(f,0))
		grandtotal = grandtotal + total
	next
	
	calculateNoVATTotal=grandtotal
	set f=nothing
	set total=nothing 
end function


' Cart VAT Total
function calculateVATTotal(pcCartArray, indexCart, TotalStandardDiscount, TotalCategoryDiscount)
	Dim f, total,pcv_StateCode,pcv_Country,mtest
	
	total=0
	grandtotal=0
	
	If trim(pcStrShippingStateCode)<> "" then
		pcv_StateCode=trim(pcStrShippingStateCode)
	else
		If trim(pcStrBillingStateCode)<> "" then
			pcv_StateCode=trim(pcStrBillingStateCode)
		End if
	end if

	if trim(pcStrShippingCountryCode)<>"" then
		pcv_Country=trim(pcStrShippingCountryCode)
	else
		if trim(pcStrBillingCountryCode)<>"" then
			pcv_Country=trim(pcStrBillingCountryCode)
		end if
	end if
	
	for f=1 to indexCart
		total=0
		mtest=0
		if (pcv_StateCode<>"") and (ucase(pcv_Country)="US") then
			mtest=CheckTaxEpt(pcCartArray(f,0),pcv_StateCode)
		end if

		if (pcCartArray(f,10)=0) AND (pcCartArray(f,19)=0) AND (mtest=0) then
			if pcCartArray(f,16)<>"" then
				total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - (pcCartArray(f,15)+pcCartArray(f,30)) +pcCartArray(f,31)
			else
				if pcCartArray(f,15)<>"0" then
					total=total + pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,17)) - pcCartArray(f,15)
				else
					total = pcCartArray(f,2)*(pcCartArray(f,5)+pcCartArray(f,3))
				end if 	 
			end if
			if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
				total = ( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) 
			end if
		end if
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// This Line Item represents what % of the Total Discount?  							
		Proportional_total = RoundTo((total/ApplicableDisountTotal),.01)
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Discount Distribution %
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Line Item after discount
		ApplicableDisount_total = (TotalStandardDiscount * Proportional_total)
		total = (total - ApplicableDisount_total)	
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Distribute Discounts based off % above
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			



		if pcv_strrApplicableCategories<>"" then
			pcArray_ApplicableCategories = split(pcv_strrApplicableCategories, ",")
			for y=0 to ubound(pcArray_ApplicableCategories)-1 '// For Each Category Discount Available
				pcArray_ApplicableCategory = split(pcArray_ApplicableCategories(y), chr(124))
				tmpApplicableCategoryID = pcArray_ApplicableCategory(1)
				tmpCategorySubTotal = pcArray_ApplicableCategory(0)
			
				ApplicableCategoryItem=False
				'response.Write("Category: " & tmpApplicableCategoryID & "<br /><br />")
				if pcv_strApplicableProducts <> "" then
					pcArray_ApplicableProducts = split(pcv_strApplicableProducts, ",")
					for x=0 to ubound(pcArray_ApplicableProducts)-1 '// Loop through all Products						
						pcArray_ApplicableProduct = split(pcArray_ApplicableProducts(x), chr(124))
						tmpProductID = pcArray_ApplicableProduct(0)
						tmpCategoryID = pcArray_ApplicableProduct(1)		
						'response.Write("Product: " & tmpProductID & " - " & pcCartArray(f,0) & "<br />")			
						if (tmpProductID = pcCartArray(f,0)) AND (tmpCategoryID = tmpApplicableCategoryID) then '// This Product is Applicable to this Category
							ApplicableCategoryItem=True
							'response.Write("Hit: " & tmpCategorySubTotal & "<br />")
						end if
					next
				end if  '// if pcv_strApplicableProducts <> "" then
				
				If ApplicableCategoryItem=True AND tmpCategorySubTotal>0 Then
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// This Line Item represents what % of the Total Category?  							
				ProportionalCat_total = RoundTo((total/tmpCategorySubTotal),.01)
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Category Distribution %
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// Line Item after Category
				ApplicableCatDisount_total = (TotalCategoryDiscount * ProportionalCat_total)
				total = (total - ApplicableCatDisount_total)	
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Distribute Category based off % above
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				End If
				
			next
		end if '// if pcv_strrApplicableCategories<>"" then			



		grandtotal = grandtotal + total
	next

	calculateVATTotal=grandtotal
	set f=nothing
	set total=nothing 
end function
'////////////////////////////////////////////////////////////////////////////////////////
'// END: VAT CALCULATIONS
'////////////////////////////////////////////////////////////////////////////////////////




'////////////////////////////////////////////////////////////////////////////////////////
'// START: (Re)check stock levels of current cart
'////////////////////////////////////////////////////////////////////////////////////////

function checkCartStockLevels(pcCartArray, indexCart, aryBads)

	Dim intCCSLindex, intCCSLcounter, intCCSLfound
	Dim strCCSLSQL, strCCSLWarn
	Dim aryCCSLitems, aryCCSLids
	Dim objCCSLrs

    ' Initialize the 'bads' array (an array tied to the index of pcCartArray, indicating which lines are bad)
	ReDim aryBads(indexCart)

    ' If cart configuration allows purchasing out-of-stock items, we're all done here
	If scOutofstockpurchase=0 Then Exit Function

	ReDim aryCCSLitems(1,0)

    intCCSLcounter = -1
	strCCSLwarn = ""

	for intCCSLIndex=1 to indexCart

	    intCCSLfound = -1								' Init to -1 indicating 'not found'

		if pcCartArray(intCCSLindex,10)=0 then			' If this product has not been deleted from cart, then
		      intCCSLcounter = intCCSLcounter + 1
			  ReDim preserve aryCCSLitems(1, intCCSLcounter)
		      aryCCSLitems(0,intCCSLcounter) = pcCartArray(intCCSLindex, 0)
		      aryCCSLitems(1,intCCSLcounter) = pcCartArray(intCCSLindex, 2)		   
		End If
		
	Next
	
	' If there were no viable items found in the cart array (e.g. all were deleted?) then exit now
	If intCCSLcounter = -1 Then Exit Function

	' Unspool/Serialize the idProducts in the array
	ReDim aryCCSLids(intCCSLcounter)
	For intCCSLindex=0 to intCCSLcounter
	   aryCCSLids(intCCSLIndex) = aryCCSLitems(0, intCCSLindex)
	Next
	strCCSLSQL = "SELECT idproduct, stock, Description, noStock, pcProd_BackOrder FROM products WHERE idproduct in (" & Join(aryCCSLids, ",") & ") and noStock=0 and pcProd_BackOrder=0"
	Set objCCSLrs = connTemp.execute(strCCSLSQL)
	If objCCSLrs.eof Then
	   objCCSLrs.close
	   Set objCCSLrs = Nothing
	   Exit Function
	Else
	   aryCCSLrecs = objCCSLrs.getrows()
	End If
	objCCSLrs.close
	Set objCCSLrs = Nothing

	ReDim aryBads(UBound(aryCCSLitems,2))
	For intCCSLindex=0 to UBound(aryCCSLitems,2)

       aryBads(intCCSLindex) = 0							' Flag this line as OK (for now)

	   For intCCSLindex2 = 0 To UBound(aryCCSLrecs, 2)

	      If CLng(aryCCSLitems(0, intCCSLindex)) = CLng(aryCCSLrecs(0, intCCSLindex2)) Then
		     If CLng(aryCCSLitems(1, intCCSLindex)) > CLng(aryCCSLrecs(1, intCCSLindex2)) Then ' Is overstock!
			    strCCSLwarn = strCCSLwarn & ("<li>" & aryCCSLrecs(2, intCCSLindex2) & " (we currently have " & aryCCSLrecs(1, intCCSLindex2) & " in stock)</li>")

			    aryBads(intCCSLindex) = -1				' Flag this line as insufficient stock level

			 End If
		     Exit For
		     
	      End If

	   next
	   
	Next

	If Len(strCCSLwarn)>0 Then checkCartStockLevels = dictLanguage.Item(Session("language")&"__alert_14") & "<ul>" & strCCSLwarn & "</ul>"

end function

'////////////////////////////////////////////////////////////////////////////////////////
'// START: get customer ID from order ID
'////////////////////////////////////////////////////////////////////////////////////////
function getCustIDfromOrder(idOrder)
	query="SELECT idcustomer FROM orders WHERE idOrder="&idOrder
	set rsCFO=server.CreateObject("ADODB.RecordSet")
	set rsCFO=connTemp.execute(query)
	if rsCFO.eof then
		getCustIDfromOrder=0
	else
		getCustIDfromOrder=rsCFO("idcustomer")
	end if
	set rsCFO=nothing  
end function
'////////////////////////////////////////////////////////////////////////////////////////
'// END: get customer ID from order ID
'////////////////////////////////////////////////////////////////////////////////////////

%>