<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "AddFQuoteToCart.asp"
'
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<%
Response.Buffer = True

Dim query, conntemp, rstemp, pcCartArray

function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function


Dim ItemsDiscounts
ItemsDiscounts=0
Dim Charges
Charges=0
Dim pDefaultPrice
pDefaultPrice=0

pcCartArray=Session("pcCartSession")
ppcCartIndex=Session("pcCartIndex")
call openDb()

pidconfigWishlistSession=getUserInput(request("idconf"),0)

if pidconfigWishlistSession="" or pidconfigWishlistSession="0" then
	redirect "Custquotesview.asp"
end if

query="SELECT idCustomer,idProduct,IdQuote,DiscountCode FROM WishList WHERE idconfigWishlistSession=" & pidconfigWishlistSession & " AND QSubmit=3 AND idcustomer=" & session("idCustomer") & ";"
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "Custquotesview.asp"
else
	pidCustomer=rs("idCustomer")
	pidProduct=rs("idProduct")
	pidQuote=rs("idQuote")
	pDiscountCode=rs("DiscountCode")
end if
set rs=nothing

query="SELECT idProduct, fPrice, stringProducts, stringValues, stringCategories, stringCProducts,stringCValues,stringCCategories,stringQuantity,stringPrice,pcconf_Quantity, dPrice, pcconf_QDiscount, xfdetails FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
	
	pIdProduct=rs("idProduct")
	pfPrice=rs("fPrice")
	pstringProducts = rs("stringProducts")
	pstringValues = rs("stringValues")
	pstringCategories = rs("stringCategories")
	pstringCProducts = rs("stringCProducts")
	pstringCValues = rs("stringCValues")
	pstringCCategories = rs("stringCCategories")
	stringQuantity = rs("stringQuantity")
	stringPrice = rs("stringPrice")
	pQuantity=rs("pcconf_Quantity")
	if (pQuantity<>"") then
	else
		pQuantity="1"
	end if
	ItemsDiscounts=rs("dPrice")
	if IsNull(ItemsDiscounts) or ItemsDiscounts="" then
		ItemsDiscounts=0
	end if
	if ItemsDiscounts<0 then
		ItemsDiscounts=-1*ItemsDiscounts
	end if
	QDisCounts=rs("pcconf_QDiscount")
	if IsNull(QDisCounts) or QDisCounts="" then
		QDisCounts=0
	end if
	pxfdetails=rs("xfdetails")
	BTOCharges=0
	if pstringCValues<>"na" then
		ArrBTO=split(pstringCValues,",")
		For i=lbound(ArrBTO) to ubound(ArrBTO)
			if trim(ArrBTO(i)<>"") then
				BTOCharges=BTOCharges+cdbl(ArrBTO(i))
			end if
		Next
	end if
	punitPrice=Round((pfPrice+ItemsDiscounts-QDisCounts-BTOCharges)/pQuantity,2)
	pOrigPrice=Round((pfPrice+ItemsDiscounts+QDisCounts-BTOCharges)/pQuantity,2)
	
set rs=nothing
	
query="SELECT products.price,products.btoBprice,products.description,products.sku, products.weight,products.pcprod_QtyToPound,products.emailText, products.deliveringTime, products.pcSupplier_ID, products.cost, products.stock, products.notax, products.noshipping, products.iRewardPoints, products.pcProd_Surcharge1, products.pcProd_Surcharge2 FROM Products WHERE idProduct=" & trim(pidProduct)
set rstemp=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

	pcv_price=rstemp("price")
	if customertype=1 then
		if rs("btoBPrice")>"0" then
			pcv_price=rstemp("btoBPrice")
		end if
	end if

	pDescription = rstemp("description")
	pWeight	= rstemp("weight")
	pcv_QtyToPound=rstemp("pcprod_QtyToPound")
	if pcv_QtyToPound>0 then
		pWeight=(16/pcv_QtyToPound)
		if scShipFromWeightUnit="KGS" then
			pWeight=(1000/pcv_QtyToPound)
		end if
	end if
	pSku = rstemp("sku")
	pEmailText = rstemp("emailText")
	pDeliveringTime	= rstemp("deliveringTime")
	pIdSupplier	= rstemp("pcSupplier_ID")
	pCost = rstemp("cost")
	pStock = rstemp("stock")
	pnotax = rstemp("notax")
	pnoshipping = rstemp("noshipping")
	iRewardPoints=rstemp("iRewardPoints")
	pcv_Surcharge1 = rstemp("pcProd_Surcharge1")
	pcv_Surcharge2 = rstemp("pcProd_Surcharge2")
	set rstemp=nothing


'-- Create new Product Config Session
' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function

pConfigKey=trim(randomNumber(9999)&randomNumber(9999))

session("pConfigKey")=pConfigKey

Dim pTodayDate
pTodayDate=Date()
if SQL_Format="1" then
	pTodayDate=Day(pTodayDate)&"/"&Month(pTodayDate)&"/"&Year(pTodayDate)
else
	pTodayDate=Month(pTodayDate)&"/"&Day(pTodayDate)&"/"&Year(pTodayDate)
end if

if scDB="Access" then
	query="INSERT INTO configSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,dtCreated,stringQuantity,stringPrice,stringCProducts,stringCValues,stringCCategories) VALUES ("&pConfigKey &","&pIdProduct&",'"&pstringProducts&"','"&pstringValues&"','"&pstringCategories&"',#"&pTodayDate&"#,'" & stringQuantity & "','" & stringPrice & "','" & pstringCProducts & "','" & pstringCValues & "','" & pstringCCategories & "')"
else
	query="INSERT INTO configSessions (configKey,idproduct,stringProducts,stringValues,stringCategories,dtCreated,stringQuantity,stringPrice,stringCProducts,stringCValues,stringCCategories) VALUES ("&pConfigKey &","&pIdProduct&",'"&pstringProducts&"','"&pstringValues&"','"&pstringCategories&"','"&pTodayDate&"','" & stringQuantity & "','" & stringPrice & "','" & pstringCProducts & "','" & pstringCValues & "','" & pstringCCategories & "')"
end if
set rs=conntemp.execute(query)

query="select idConfigSession from configSessions order by idConfigSession desc"
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
			
pIdConfigSession=rs("IdConfigSession")

set rs=nothing

'-- END Create new Product Config Session
'---------------------		

'---------------------------
'Calculate BTO Items Weights
IF pIdConfigSession<>"" then
	query="SELECT stringProducts,stringCProducts,stringQuantity FROM configSessions WHERE idconfigSession=" & pIdConfigSession
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
		
	stringProducts=rs("stringProducts")
	stringCProducts=rs("stringCProducts")
	ArrProduct=Split(stringProducts, ",")
	ArrCProduct=Split(stringCProducts, ",")
	Qstring=rs("stringQuantity")
	ArrQuantity=Split(Qstring,",")
		
	CWeight=0

	if ArrProduct(0)<>"na" then
		for j=lbound(ArrProduct) to (UBound(ArrProduct)-1)
			query="SELECT weight,pcprod_QtyToPound FROM products WHERE IDProduct=" & ArrProduct(j)
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
		
			if not rs.eof then
				ItemWeight=rs("weight")
				pcv_QtyToPound=rs("pcprod_QtyToPound")
				if pcv_QtyToPound>0 then
					ItemWeight=cdbl(16/pcv_QtyToPound)
					if scShipFromWeightUnit="KGS" then
						ItemWeight=cdbl(1000/pcv_QtyToPound)
					end if
				end if

				CWeight=CWeight+cdbl(ItemWeight*clng(ArrQuantity(j)))
			end if
		next
	end if

		if ArrCProduct(0)<>"na" then
			for j=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
				query="SELECT weight,pcprod_QtyToPound FROM products WHERE IDProduct=" & ArrCProduct(j)
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
				
				if not rs.eof then
		
					ItemWeight=rs("weight")
					
					pcv_QtyToPound=rs("pcprod_QtyToPound")
					if pcv_QtyToPound>0 then
						ItemWeight=(16/pcv_QtyToPound)
						if scShipFromWeightUnit="KGS" then
							ItemWeight=(1000/pcv_QtyToPound)
						end if
					end if
		
					CWeight=CWeight+cdbl(ItemWeight)
		
				end if
			next
		end if	
	
	END IF 'Have ID Config Session
	'--- END Calculate BTO Items Weights	
	'---------------------------

	pcv_FQuotes=session("sf_FQuotes")
	pcv_AddQuote=0
	if pcv_FQuotes="" then	
		ppcCartIndex = ppcCartIndex + 1
		session("pcCartIndex")	= ppcCartIndex
		session("sf_FQuotes") = pcv_FQuotes & "****" & pIdProduct & "****"
		pcv_AddQuote=1
	else
		if Instr(pcv_FQuotes,"****" & pIdProduct & "****")=0 then
			ppcCartIndex = ppcCartIndex + 1
			session("pcCartIndex")	= ppcCartIndex
			session("sf_FQuotes") = pcv_FQuotes & "****" & pIdProduct & "****"
			pcv_AddQuote=1
		end if
	end if

IF pcv_AddQuote=1 THEN
	pcCartArray(ppcCartIndex,0) = pIdProduct 
	pcCartArray(ppcCartIndex,1) = pDescription
	pcCartArray(ppcCartIndex,2) = pQuantity
	
	' add price or BtoB price depending on customer type
	pcCartArray(ppcCartIndex,3) = punitPrice

	pcCartArray(ppcCartIndex,8)="" '// not in use anymore
	
	
	pcCartArray(ppcCartIndex,4)= "" '// store array of product "option groups: options"

	pcCartArray(ppcCartIndex,5) = 0 '// Total Cost of all Options
	pcCartArray(ppcCartIndex,23)= pOverSizeSpec
	pcCartArray(ppcCartIndex,25)= "" '// Array of Individual Options Prices
	pcCartArray(ppcCartIndex,26)= "" '// Array of Options Prices 
	'pcCartArray(ppcCartIndex,27)="" '// Not in use anymore - VERIFY FOR Crosssell
	'pcCartArray(ppcCartIndex,28)="" '// Not in use anymore - VERIFY FOR Crosssell	 	 	 	 
	pcCartArray(ppcCartIndex,6) = pWeight + Cweight
	pcCartArray(ppcCartIndex,7) = pSku
	pcCartArray(ppcCartIndex,9) = pDeliveringTime  
	
	' deleted mark
	pcCartArray(ppcCartIndex,10) = 0 
	
	pcCartArray(ppcCartIndex,11) = "" '// Array of Individual Selected Options Id Numbers
	pcCartArray(ppcCartIndex,12) = "" '// Not in use anymore
	pcCartArray(ppcCartIndex,13) = pIdSupplier
	pcCartArray(ppcCartIndex,14) = pCost
	pcCartArray(ppcCartIndex,16) = pIdConfigSession
	pcCartArray(ppcCartIndex,19) = pnotax
	pcCartArray(ppcCartIndex,20) = pnoshipping
	pcCartArray(ppcCartIndex,36) = pcv_Surcharge1
	pcCartArray(ppcCartIndex,37) = pcv_Surcharge2
	
	'// Fix the XFields for display
	if trim(pxfdetails)<>"" then
	
		xfieldsarray=split(pxfdetails,"||")
		dispStr = ""			
			
		for i=lbound(xfieldsarray)to (UBound(xfieldsarray)-1)
			xfields=split(xfieldsarray(i),"|")
			query="SELECT xfield FROM xfields WHERE idxfield="&xfields(0)
			set rs=Server.CreateObject("ADODB.Recordset")
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
			
			xfielddesc=rs("xfield")
			set rs=nothing
			dispStr = dispStr & xfielddesc&": "&xfields(1) & "|"									
		next
		pxfdetails = dispStr
	end if

	pcCartArray(ppcCartIndex,21)=  replace(pxfdetails,"|","<br>")

	pTotalQuantity = pQuantity
	
	'-----------------------------
	'ReCalculate BTO Items Discounts

	pcCartArray(ppcCartIndex,30)=cdbl(ItemsDiscounts)
	
	'End ReCulculate BTO Items Discounts		
	'------------------------------------
	
	'------------------------------------
	'BTO Additional Charges

	pcCartArray(ppcCartIndex,31) = Cdbl(BTOCharges) 
	
	'End BTO Additional Charges
	'------------------------------------

	
	if Session("customerType")=1 OR session("customerCategory")<>0 then
		pcCartArray(ppcCartIndex,18)=1
	else
		pcCartArray(ppcCartIndex,18)=0
	end if

	pcCartArray(ppcCartIndex,17) = pOrigPrice
	pcCartArray(ppcCartIndex,15) = QDiscounts

	'// Start Reward Points
	If RewardsActive <> 0 Then
		pcCartArray(ppcCartIndex,22) = Clng(iRewardPoints)
	End If
	'// End Reward Points
	
	'SM-S
	if UCase(scDB)="SQL" then
		query="SELECT Products.pcSC_ID,pcSales_BackUp.pcSales_TargetPrice FROM pcSales_BackUp INNER JOIN Products ON pcSales_BackUp.pcSC_ID=Products.pcSC_ID WHERE Products.idProduct=" & pIdProduct
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			tmpSCID=rsQ("pcSC_ID")
			If IsNull(tmpSCID) then
				tmpSCID=0
			End If
			tmpTarget=rsQ("pcSales_TargetPrice")
			if IsNull(tmpTarget) then
				tmpTarget=0
			end if
			if ((clng(tmpTarget)=0) AND (session("customerCategory")=0) AND (session("customerType")<>"1")) OR ((clng(tmpTarget)=-1) AND (session("customerCategory")=0) AND (session("customerType")="1")) OR ((clng(tmpTarget)=clng(session("customerCategory"))) AND (clng(tmpTarget)>0)) then
				pcCartArray(ppcCartIndex,39)=tmpSCID '//Sale ID
			else
				pcCartArray(ppcCartIndex,39)=0
			end if
		else
			pcCartArray(ppcCartIndex,39)=0
		end if
		set rsQ=nothing
	else
		pcCartArray(ppcCartIndex,39)=0
	end if
	'SM-E
	
END IF

session("pcCartSession") = pcCartArray

call closedb()

response.redirect "viewcart.asp"
%>

