<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "instPrdCharge.asp"
' This page is handles BTO Products
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/bto_language.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<%
Response.Buffer = True

Dim query, conntemp, rstemp, ptotalQuantity

'GGG Add-on start

'Check Shopping Cart if It is used for a Gift Registry
if Session("Cust_BuyGift")<>"" then
  response.redirect "msg.asp?message=100"      
end if

'GGG Add-on end

if session("idcustomer")="" then
	session("idCustomer")	= Cint(0)
	session("idAffiliate")	= Cint(1)     
	session("language") = Cstr("english")
	session("pcCartIndex") = Cint(0)
	dim pcCartArray(100,45)
	session("pcCartSession") = pcCartArray    
end if

pcCartArray	= Session("pcCartSession")
ppcCartIndex = getUserInput(request.Form("pcCartIndex"),0)
' check for errors
if err.number>0 then
	response.redirect "msg.asp?message=47"
end if
fpc=getUserInput(request("ConfigCartIndex"),0)
pTotalQuantity = Cint(0)

' check for bound quantity in cart
if countCartRows(pcCartArray, ppcCartIndex) = scQtyLimit then
  response.redirect "msg.asp?message=39"         
end if

' get data from viewPrd form
piBTOQuote_rec=getUserInput(request("iBTOQuote_rec.x"),0)
if piBTOQuote_rec="" then
	piBTOQuote_rec=getUserInput(request("iBTOQuote_rec"),0)
end if
pidconf=getUserInput(request("idconf"),10)
pBTOQuote = getUserInput(request("iBTOQuote.x"),0)
if pBTOQuote="" then
	pBTOQuote = getUserInput(request("iBTOQuote"),0)
end if
pIdProduct = getUserInput(request("idproduct"),10)
pQuantity = pcCartArray(ppcCartIndex,2)
'how many products are part of this configured product
pJcnt = getUserInput(request.Form("jCnt"),4)
'how many categories are multiselect
CB_CatCnt = getUserInput(request.Form("CB_CatCnt"),10)
	if not validNum(CB_CatCnt) then
		call closeDB()
		response.redirect("msg.asp?message=207")
	end if
'price of configured product
pGrTotal = getUserInput(request.Form("GrandTotal"),10)
pGrTotal=replace(pGrTotal,scCurSign,"")
if scDecSign="." then
	pGrTotal=replace(pGrTotal,",","")
else
	pGrTotal=replace(pGrTotal,".","")
	pGrTotal=replace(pGrTotal,",",".")
end if
pGrTotal1=pGrTotal
pfPrice=pGrTotal

pcv_QDisc=session("QDisc" & pIDProduct)
if pcv_QDisc="" then
pcv_QDisc="0"
end if

pGrandTotal2=getUserInput(request.Form("GrandTotal2"),10)
pGrandTotal2=replace(pGrandTotal2,scCurSign,"")
if scDecSign="." then
	pGrandTotal2=replace(pGrandTotal2,",","")
else
	pGrandTotal2=replace(pGrandTotal2,".","")
	pGrandTotal2=replace(pGrandTotal2,",",".")
end if

' if cannot get quantity get quantity 1 (from listing)
if pQuantity="" then
	pQuantity=1
end if

if Int(pQuantity)>Int(scAddLimit) then
   response.redirect "msg.asp?message=40"          
end if

call opendb()

' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function

'insert product configuration into configSessions

'create strings
Dim Pstring, Vstring, Cstring, tempVar, tempCatarray, tempString, strArray
Pstring = ""
Vstring = ""
Cstring = ""
Cweight = 0
FirstCnt = getUserInput(request.form("FirstCnt"),0)
If FirstCnt<>"" then
	for i = 1 to FirstCnt
		tempVar = getUserInput(request.form("CAT"&i),0)
		MS=getUserInput(request.form("MS"&i),0)
		If MS="" then
		tempCatarray = split(tempVar,"G")
		tempString = getUserInput(request.form(tempVar),0)
		strArray = split(tempString, "_")
		If strArray(0)<>0 then
			Cstring = Cstring & tempCatarray(1) & ","
			Pstring = Pstring & strArray(0) & ","
			Vstring = Vstring & strArray(1) & ","
			if cdbl(strArray(2))>0 then
			Cweight = Cweight + cdbl(strArray(2))
			else
			query="SELECT pcprod_QtyToPound FROM Products WHERE idproduct=" & strArray(0)
			set rsW=connTemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsW=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			itemWeight=0
			if not rsW.eof then
			item_QtyToPound=rsW("pcprod_QtyToPound")
			if item_QtyToPound>0 then
				itemWeight=(16/item_QtyToPound)
				if scShipFromWeightUnit="KGS" then
					itemWeight=(1000/item_QtyToPound)
				end if
			end if
			end if
			Cweight = Cweight + itemWeight
			set rsW=nothing
			end if
		End If
		End If
	next
end if
'continue strings with multiselect items
Dim c, p, tempCatVar, tempMSPrd
if CB_CatCnt<>"" then
	for c=1 to CB_CatCnt
		tempCatVar=getUserInput(request.form("CB_CatID"&c),0)
		tempPrdCntVar=getUserInput(request.form("PrdCnt"&tempCatVar),0)
		for p=1 to tempPrdCntVar
			tempMSPrd = getUserInput(request.form("Cat"&tempCatVar&"_"&"Prd"&p),0)
			tempString = getUserInput(request.form("CAG"&tempCatVar&tempMSPrd),0)
			if tempString<>"" then
				strArray = split(tempString, "_")
				If strArray(0)<>0 then
					Cstring = Cstring & tempCatVar & ","
					Pstring = Pstring & strArray(0) & ","
					Vstring = Vstring & strArray(1) & ","
					if cdbl(strArray(2))>0 then
						Cweight = Cweight + cdbl(strArray(2))
					else
						query="SELECT pcprod_QtyToPound FROM Products WHERE idproduct=" & strArray(0)
						set rsW=connTemp.execute(query)
						if err.number<>0 then
							'//Logs error to the database
							call LogErrorToDatabase()
							'//clear any objects
							set rsW=nothing
							'//close any connections
							call closedb()
							'//redirect to error page
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						itemWeight=0
						if not rsW.eof then
							item_QtyToPound=rsW("pcprod_QtyToPound")
							if item_QtyToPound>0 then
								itemWeight=(16/item_QtyToPound)
								if scShipFromWeightUnit="KGS" then
									itemWeight=(1000/item_QtyToPound)
								end if
							end if
						end if
						Cweight = Cweight + itemWeight
						set rsW=nothing
					end if
				End If
			End IF
		next
	next
end if

if Pstring<>"" then
	query="select * from configSpec_Charges where specproduct=" & pIDproduct & " order by catSort, prdSort"
	set rs4=connTemp.execute(query)
	if err.number<>0 then
		'//Logs error to the database
		call LogErrorToDatabase()
		'//clear any objects
		set rs4=nothing
		'//close any connections
		call closedb()
		'//redirect to error page
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	A=split(Cstring,",")
	B=split(PString,",")
	C=split(Vstring,",")
	Count=0
	Dim A1(100),B1(100),C1(100)

	Do while not rs4.eof

		For k=lbound(B) to ubound(B)
			if B(k)<>"" then
				if (clng(B(k))=clng(rs4("configProduct"))) and (clng(A(k))=clng(rs4("configProductCategory"))) then
					A1(Count)=A(k)
					B1(Count)=B(k)
					C1(Count)=C(k)
					Count=Count+1
				end if
			end if	
		Next

	rs4.MoveNext
	Loop
	Cstring=""
	Pstring=""
	Vstring=""
	For k=0 to Count-1
		Cstring=Cstring & A1(k) & ","
		Pstring=Pstring & B1(k) & ","
		Vstring=Vstring & C1(k) & ","
	Next
end if

if Pstring="" then
Pstring="na"
Vstring="na"
Cstring="na"
end if

discountcodetemp=getUserInput(request("discountcode"),0)
If trim(discountcodetemp)<>"" then
	discountcodetemp1=replace(discountcodetemp,"'","''")
	query="SELECT onetime,iddiscount FROM discounts WHERE discountcode='"&discountcodetemp1&"' AND active=-1"
	set rs=Server.CreateObject("ADODB.Recordset")
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
	if rs.eof then
		call closeDb()
		response.redirect "PrdAddCharges.asp?idproduct="&pIdProduct&"&pcCartIndex=" &fpc&"&msg="&Server.Urlencode("The discount code you entered is not valid.")
	end if
	If rs("onetime")=true Then
		'check customer's id in database with iddiscount
		query="SELECT * FROM used_discounts WHERE iddiscount="&rs("iddiscount")&" AND idcustomer="&session("idCustomer")
	 	set rsDisObj=conntemp.execute(query)
		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rsDisObj=nothing
			'//close any connections
			call closedb()
			'//redirect to error page
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		if rsDisObj.eof then
		else
			call closeDb()
			response.redirect "PrdAddCharges.asp?idproduct="&pIdProduct&"&pcCartIndex=" &fpc&"&msg="&Server.Urlencode("The discount code you entered is no longer valid.")
		end if
	End if
End If

' get discount certificate data
pDiscountError=Cstr("")
pDiscountCode=discountcodetemp

if pDiscountCode="" then
	noCode="1"
	pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_3") 
else
	pDiscountCode1=replace(pDiscountCode,"'","''")
	query="SELECT iddiscount, onetime, expDate, idProduct, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, DiscountDesc, priceToDiscount, percentageToDiscount, pcDisc_StartDate FROM discounts WHERE discountcode='" &pDiscountCode1& "' AND active=-1"

 	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4") 
	else 
		piddiscount=rstemp("iddiscount")
		ponetime=rstemp("onetime")
		pexpDate=rstemp("expDate")
		tmpidProduct=rstemp("idProduct")
		pquantityFrom=rstemp("quantityFrom")
		pquantityUntil=rstemp("quantityUntil")
		pweightFrom=rstemp("weightFrom")
		pweightUntil=rstemp("weightUntil")
		ppriceFrom=rstemp("priceFrom")
		ppriceUntil=rstemp("priceUntil")
		pDiscountDesc=rstemp("DiscountDesc")
		ppriceToDiscount=rstemp("priceToDiscount")
		ppercentageToDiscount=rstemp("percentageToDiscount")
		pStartDate=rstemp("pcDisc_StartDate")
		set rstemp=nothing
	 
		'check to see if discount has been used for one use only for this customer specified
		If ponetime=true Then
			'check customer's id in database with iddiscount
			query="SELECT * FROM used_discounts WHERE iddiscount=" &piddiscount& " AND idcustomer="&session("idCustomer")
			set rsDisObj=conntemp.execute(query)
		
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsDisObj=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		
			if NOT rsDisObj.eof then
				'discount has been used already by the customer
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
			end if
			set rsDisObj=nothing
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
		If pDiscountError="" Then
			if Int(pCartQuantity)>=Int(pquantityFrom) and Int(pCartQuantity)<=Int(pquantityUntil) and Int(pCartTotalWeight)>=Int(pweightFrom) and Int(pCartTotalWeight)<=Int(pweightUntil) and Cdbl(pSubTotal)>=Cdbl(ppriceFrom) and Cdbl(pSubTotal)<=Cdbl(ppriceUntil) then
				pcv_DiscountDesc=pDiscountDesc
			pcv_PriceToDiscount=cdbl(pPriceToDiscount)
				pcv_percentageToDiscount=ppercentageToDiscount
			else
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5") 
			end if
		End If
	end if
end if

discountTotal=Cdbl(0)
' discounts
if pDiscountError="" then    
	discountTotal=Cdbl(0)
	' calculate discount. Note: percentage dont affects shipment and payment prices
	if pcv_PriceToDiscount>0 or pcv_percentageToDiscount>0 then 
		discountTotal=pcv_PriceToDiscount + (pcv_percentageToDiscount*(pGrTotal1)/100)
	end if
end if

SPstring=session("SPstring")
SVstring=session("SVstring")
SCstring=session("SCstring")
xString=session("SxString")
pGrTotal1=session("SpGrTotal1")*pQuantity
PDiscounts=session("SdiscountTotal")
Qstring=session("SQstring")
Pricestring=session("SPricestring")
pConfigKey=session("pConfigKey")

discountTotal=discountTotal + cdbl(PDiscounts)
ArrProduct=split(SPstring,",")
ArrQuantity=split(Qstring,",")
ArrPrice=split(Pricestring,",")
itemsDiscounts=0
for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
 query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(i)
 set rs99=connTemp.execute(query)
		if err.number<>0 then
			'//Logs error to the database
			call LogErrorToDatabase()
			'//clear any objects
			set rs99=nothing
			'//close any connections
			call closedb()
			'//redirect to error page
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
 
 TempDiscount=0
 do while not rs99.eof
 				QFrom=rs99("quantityFrom")
				QTo=rs99("quantityUntil")
				DUnit=rs99("discountperUnit")
				QPercent=rs99("percentage")
				DWUnit=rs99("discountperWUnit")
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
	rs99.movenext
	loop
	itemsDiscounts=ItemsDiscounts+TempDiscount
	next
	
if ItemsDiscounts>0 then
pGrTotal1=cdbl(pGrTotal1)-cdbl(ItemsDiscounts)
end if



pGrTotal1=cdbl(pGrTotal1)+Cdbl(getUserInput(request.form("CHGTotal"),0))-pcv_QDisc

pIdConfigSession=pcCartArray(ppcCartIndex,16)
pidconf=pIdConfigSession
'If this is a quote, add to quote sessions and then to wishlist and then redirect to wishlist page
Dim pTodayDate
pTodayDate=Date()
if SQL_Format="1" then
	pTodayDate=Day(pTodayDate)&"/"&Month(pTodayDate)&"/"&Year(pTodayDate)
else
	pTodayDate=Month(pTodayDate)&"/"&Day(pTodayDate)&"/"&Year(pTodayDate)
end if

If pBTOQuote<>"" or piBTOQuote_rec<>"" then
		if piBTOQuote_rec<>"" then
			query="UPDATE configWishlistSessions SET stringProducts='"&SPstring&"',stringValues='"&SVstring&"',stringCategories='"&SCstring&"',xfdetails='"&xString&"',fPrice="&pGrTotal1&",dPrice=" & discountTotal & ",stringCProducts='" & Pstring & "',stringCValues='" & Vstring & "',stringCCategories='" & Cstring & "',pcconf_Quantity=" & pQuantity & ",pcconf_QDiscount=" & pcv_QDisc & " WHERE idconfigWishlistSession="&pidConf
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsConf=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pIdConfigWishlistSession=pidConf
		else
			err.clear
			if scDB="Access" then
				query="INSERT INTO configWishlistSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,xfdetails,dtCreated,fPrice,dPrice,stringCProducts,stringCValues,stringCCategories,stringQuantity,stringPrice,pcconf_Quantity,pcconf_QDiscount) VALUES ("&pConfigKey &","&pIdProduct&",'"&SPstring&"','"&SVstring&"','"&SCstring&"','"&xString&"',#"&pTodayDate&"#,"&pGrTotal1&"," & discountTotal & ",'" & Pstring & "','" & Vstring & "','" & Cstring & "','" & Qstring & "','" & Pricestring & "'," & pQuantity & "," & pcv_QDisc & ")"
			else
				query="INSERT INTO configWishlistSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,xfdetails,dtCreated,fPrice,dPrice,stringCProducts,stringCValues,stringCCategories,stringQuantity,stringPrice,pcconf_Quantity,pcconf_QDiscount) VALUES ("&pConfigKey &","&pIdProduct&",'"&SPstring&"','"&SVstring&"','"&SCstring&"','"&xString&"','"&pTodayDate&"',"&pGrTotal1&"," & discountTotal & ",'" & Pstring & "','" & Vstring & "','" & Cstring & "','" & Qstring & "','" & Pricestring & "'," & pQuantity & "," & pcv_QDisc & ")"
			end if
			set rsConf=Server.CreateObject("ADODB.Recordset")
			set rsConf=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsConf=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		end if
		if scDB="Access" then
			query="SELECT configWishlistSessions.idconfigWishlistSession FROM configWishlistSessions WHERE (((configWishlistSessions.configKey)="&pConfigKey&") AND ((configWishlistSessions.dtCreated)=#"&pTodayDate&"#));"
		else
			query="SELECT idconfigWishlistSession FROM configWishlistSessions WHERE configKey="&pConfigKey&" AND dtCreated='"&pTodayDate&"';"
		end if
		set rsConf=Server.CreateObject("ADODB.Recordset")
		set rsConf=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsConf=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		Dim pIdConfigWishlistSession
		pIdConfigWishlistSession = rsConf("idconfigWishlistSession")
		
	set rsConf=nothing
	response.redirect "Custwl.asp?discountcode=" & discountcodetemp & "&idConf="&pIdConfigWishlistSession&"&idproduct="&pIdProduct&"&redirecturl=" & Server.URLEncode("Custwl.asp?idConf="&pIdConfigWishlistSession&"&discountcode=" & discountcodetemp &"&idproduct="&pIdProduct) 
	response.end
'else
else
	if Pstring="" then
	else
		
		query="Update configSessions set stringCProducts='" & Pstring & "',stringCValues='" & Vstring & "',stringCCategories='" & Cstring & "' WHERE IdConfigSession=" & pIdConfigSession
		set rsConf=Server.CreateObject("ADODB.Recordset")
		set rsConf=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsConf=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
	end if
'end if
end if


' insert new basket line, is not in the cart
pTotalQuantity = pQuantity


' deleted mark
pcCartArray(ppcCartIndex,10) = 0 
pWeight=pcCartArray(ppcCartIndex,6)
pcCartArray(ppcCartIndex,6) = pWeight + Cweight	 
pcCartArray(ppcCartIndex,31) = Cdbl(getUserInput(request.form("CHGTotal"),0))

	'// Start Reward Points
	If RewardsActive <> 0 Then
		'Calculate BTO Configured Product RP

		pcv_BTORP=0
		if Ucase(PString)<>"NA" then
			ArrProduct=split(Pstring,",")
			For i=lbound(ArrProduct) to ubound(ArrProduct)
				if trim(ArrProduct(i))<>"" then
					query="SELECT iRewardPoints FROM Products WHERE idproduct=" & ArrProduct(i)
					set rstemp=connTemp.execute(query)
					if err.number<>0 then
						'//Logs error to the database
						call LogErrorToDatabase()
						'//clear any objects
						set rstemp=nothing
						'//close any connections
						call closedb()
						'//redirect to error page
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if not rstemp.eof then
						pcv_BTORP=pcv_BTORP+clng(rstemp("iRewardPoints"))
					end if
					set rstemp=nothing
				end if
			Next
		end if
		
		pcCartArray(ppcCartIndex,29) = pcv_BTORP
	End If
	'// Start Reward Points

session("pcCartSession") = pcCartArray

NeedReCalculate=1%>
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="pcReCalPricesLogin.asp"-->
<%
'Calculate Product Promotions - START
%>
<!--#include file="inc_CalPromotions.asp"-->
<%
'Calculate Product Promotions - END

call closeDB()
call clearLanguage()

'Clear custom input data session
session("SFxfield1_" & pIdProduct)=""
session("SFxfield2_" & pIdProduct)=""
session("SFxfield3_" & pIdProduct)=""

' redirect to cart view
Session("pcSessionID")=Session.SessionID '// browser test session
response.redirect "viewCart.asp?cs=1" '// cs = Check Session. Initializes the session check script.
%>