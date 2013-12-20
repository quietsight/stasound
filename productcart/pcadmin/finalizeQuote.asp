<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<% 
dim query, conntemp, rs
call opendb()

msg=""

query="SELECT wishlist.idCustomer, wishlist.idProduct, wishlist.idconfigWishlistSession, wishlist.QDate,wishlist.Qsubmit FROM wishlist INNER JOIN configWishlistSessions ON wishlist.idconfigWishlistSession = configWishlistSessions.idconfigWishlistSession WHERE (((wishlist.QSubmit)>0) AND ((wishlist.IDQuote)="&request("idquote")&"));"

set rs=server.CreateObject("ADODB.RecordSet")							
set rs=conntemp.execute(query)
					
if err.number <> 0 then
	set rs=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in mysavequote: "&err.description) 
end if
	
idcustomer=rs("idcustomer")
idproduct=rs("idproduct")
idconf=rs("idconfigWishlistSession")
qdate=rs("qdate")
pcv_qsubmit=rs("qsubmit")
if IsNull(pcv_qsubmit) or pcv_qsubmit="" then
	pcv_qsubmit=0
end if
set rs=nothing
idquote=request("idquote")
pidconfigWishlistSession=idconf
query="select email,customerType from customers where idcustomer=" & idcustomer
set rs=conntemp.execute(query)
if not rs.eof then
	custEmail=rs("email")
	customertype=cint(rs("customerType"))
else
	customertype=0
end if
set rs=nothing
	
	query="SELECT idProduct, fPrice, stringProducts, stringValues, stringCategories, stringCProducts,stringCValues,stringCCategories,stringQuantity,stringPrice,pcconf_Quantity, dPrice, pcconf_QDiscount FROM configWishlistSessions WHERE idconfigWishlistSession=" & pidconfigWishlistSession
	set rs=conntemp.execute(query)
	
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
	QDisCounts=rs("pcconf_QDiscount")
	if IsNull(QDisCounts) or QDisCounts="" then
		QDisCounts=0
	end if
	set rs=nothing
	
	query="SELECT price,btoBprice FROM Products WHERE idProduct=" & trim(pidProduct)
	set rs=conntemp.execute(query)
	pcv_price=rs("price")
	if customertype=1 then
		if rs("btoBPrice")>"0" then
			pcv_price=rs("btoBPrice")
		end if
	end if

	ArrProduct = Split(pstringProducts, ",")
	ArrValue = Split(pstringValues, ",")
	ArrCategory = Split(pstringCategories, ",")
	ArrQuantity = Split(stringQuantity, ",")
	ArrPrice = Split(stringPrice, ",")
	ArrCProduct = Split(pstringCProducts, ",")
	ArrCValue = Split(pstringCValues, ",")
	ArrCCategory = Split(pstringCCategories, ",")
	
	new_ArrProduct = ""
	new_ArrValue = ""
	new_ArrCategory = ""
	new_ArrQuantity = ""
	new_ArrPrice = ""
	if pstringProducts<>"na" then
	For i=lbound(ArrProduct) to ubound(ArrProduct)
		if trim(ArrProduct(i))<>"" then
			new_ArrProduct = new_ArrProduct & ArrProduct(i) & ","
			new_ArrValue = new_ArrValue & cdbl(ArrPrice(i)) & ","
			new_ArrCategory = new_ArrCategory & ArrCategory(i) & ","
			new_ArrQuantity = new_ArrQuantity & ArrQuantity(i) & ","
			new_ArrPrice = new_ArrPrice & ArrPrice(i) & ","
			pcv_price=pcv_price+cdbl(ArrQuantity(i)*ArrPrice(i))
		end if
	Next
	end if
	
	if new_ArrProduct="" then
		new_ArrProduct = "na"
		new_ArrValue = "na"
		new_ArrCategory = "na"
		new_ArrQuantity = "na"
		new_ArrPrice = "na"
	end if

	new_ArrCProduct = ""
	new_ArrCValue = ""
	new_ArrCCategory = ""
	pcv_price1=0
	if pstringCProducts<>"na" then
	For i=lbound(ArrCProduct) to ubound(ArrCProduct)
		if trim(ArrCProduct(i))<>"" then
			new_ArrCProduct = new_ArrCProduct & ArrCProduct(i) & ","
			new_ArrCValue = new_ArrCValue & cdbl(ArrCValue(i)) & ","
			new_ArrCCategory = new_ArrCCategory & ArrCCategory(i) & ","
			pcv_price1=pcv_price1+cdbl(ArrCValue(i))
		end if
	Next
	end if
	if new_ArrCProduct="" then
		new_ArrCProduct = "na"
		new_ArrCValue = "na"
		new_ArrCCategory = "na"
	end if
	
	if ItemsDiscounts="" then
		ItemsDiscounts=0
	end if
	
	if cdbl(ItemsDiscounts)<0 then
		ItemsDiscounts=-1*ItemsDiscounts
	end if
	
	if QDisCounts="" then
		QDisCounts=0
	end if
	
	pcv_price=(pcv_price*pQuantity)+pcv_price1-ItemsDiscounts-QDisCounts
	
	query="UPDATE configWishlistSessions SET fPrice=" & pcv_price & ", stringProducts='" & new_ArrProduct & "', stringValues='" & new_ArrValue & "', stringCategories='" & new_ArrCategory & "', stringCProducts='" & new_ArrCProduct & "',stringCValues='" & new_ArrCValue & "',stringCCategories='" & new_ArrCCategory & "',stringQuantity='" & new_ArrQuantity & "',stringPrice='" & new_ArrPrice & "',pcconf_Quantity=" & pQuantity & ", dPrice=" & -1*ItemsDiscounts & ",pcconf_QDiscount=" & QDisCounts & " WHERE idconfigWishlistSession=" & pidconfigWishlistSession
	set rs=connTemp.execute(query)
	set rs=nothing
	
	query="SELECT QSubmit FROM wishlist WHERE IDQuote=" & idquote
	set rs=connTemp.execute(query)
	if rs("QSubmit")>="1" then
		query="UPDATE wishlist SET QSubmit=3 WHERE IDQuote=" & idquote
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	set rs=nothing
	call closedb()
	
	pcv_strSubject=bto_dictLanguage.Item(Session("language")&"_quotenotice_1")
	pcv_NoticeEmail=bto_dictLanguage.Item(Session("language")&"_quotenotice_2")
	pcv_NoticeEmail=pcv_NoticeEmail & idquote & bto_dictLanguage.Item(Session("language")&"_quotenotice_3") & vbcrlf
	
	dim tempURL
	if scPcFolder<>"" then
		tempURL=scStoreURL&"/"&scPcFolder&"/pc/Custquotesview.asp"
	else
		tempURL=scStoreURL&"/pc/Custquotesview.asp"
	end if
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/Custquotesview.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
	
	pcv_NoticeEmail=pcv_NoticeEmail & tempURL & vbcrlf & vbcrlf & scCompanyName & vbCrLf & scStoreURL & vbcrlf & vbCrLf
	
	call sendmail (scCompanyName, scEmail, custEmail, pcv_strSubject, pcv_NoticeEmail)
	
	msg="The quote was updated successfully!"
	response.redirect "srcQuotesa.asp?s=1&msg="&msg

%>