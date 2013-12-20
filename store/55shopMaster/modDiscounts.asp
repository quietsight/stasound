<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Modify Discount by Code" %>
<% Section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<% Dim rstemp, connTemp, query, iddiscount, editAction

iddiscount=request("iddiscount")
editAction=request("editAction")

dim intRequestSubmit
intRequestSubmit=0

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check for Unique Discount Code
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// Discount Description
discountdesc=Request("discountdesc")
if instr(discountdesc,"'")>0 then
	response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&msg=The description cannot include an apostrophe.")
	response.End()
end if
discountdesc=replace(discountdesc,"&quot;","""")
discountdesc=replace(discountdesc,"'","''")
discountdesc=replace(discountdesc,",","")
if Request("Submit1")<>"" then
	if discountdesc="" then
		discountdesc="No Description"
	end if
end if

discountdesc2=replace(Request("discountdesc2"),"&quot;","""")
discountdesc2=replace(discountdesc2,"'","''")
discountdesc2=replace(discountdesc2,",","")
if Request("Submit1")<>"" then
	if discountdesc2="" then
		discountdesc2="No Description"
	end if
end if

If (discountdesc<>discountdesc2) AND (Request("Submit1")<>"") Then
	call openDb()
	query="SELECT discountdesc FROM discounts WHERE discountdesc='"& discountdesc &"' "
	set rsValidateDC=server.CreateObject("ADODB.RecordSet")
	set rsValidateDC=conntemp.execute(query)
	if not rsValidateDC.eof then
		set rsValidateDC=nothing
		call closeDb()
		response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&editAction=1&msg=A discount already exists with that description.")
	end if
	set rsValidateDC=nothing
	call closeDb()
End If

'// Discount Code
discountcode=replace(Request("discountcode"),"'","")
discountcode2=replace(Request("discountcode2"),"'","")
discountcode=replace(Request("discountcode"),",","")
discountcode2=replace(Request("discountcode2"),",","")

If (discountcode<>discountcode2) AND (Request("Submit1")<>"") Then
	call openDb()
	query="SELECT discountcode FROM discounts WHERE discountcode='"& discountcode &"' "
	set rsValidateDC=server.CreateObject("ADODB.RecordSet")
	set rsValidateDC=conntemp.execute(query)
	if not rsValidateDC.eof then
		set rsValidateDC=nothing
		call closeDb()
		response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&editAction=1&msg=A discount already exists with that code.")
	end if
	set rsValidateDC=nothing
	call closeDb()
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check for Unique Discount Code
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Filter by Product
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request("submit2")<>"" then
	intRequestSubmit=1
	Count1=request("Count1")
	if Count1="" then
	Count1="0"
	end if
	
	For i=1 to Count1
		if request("Pro" & i)="1" then
		IDPro=request("IDPro" & i)
		call opendb()
		query="DELETE FROM pcDFProds WHERE pcFPro_IDDiscount=" & iddiscount & " AND pcFPro_IDProduct=" & IDPro
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		call closedb()
		end if
	Next
	
	response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&editAction=1")
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Filter by Product
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Filter by Category
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if (request("submit3")<>"") OR (request("submit3A")<>"") then
	intRequestSubmit=1
	Count2=request("Count2")
	if Count2="" then
	Count2="0"
	end if
	
	if (request("submit3")<>"") then
		For i=1 to Count2
			if request("CAT" & i)="1" then
			IDCat=request("IDCat" & i)
			call opendb()
			query="DELETE FROM pcDFCats WHERE pcFCat_IDDiscount=" & iddiscount & " AND pcFCat_IDCategory=" & IDCat
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
			call closedb()
			end if
		Next
	else
		For i=1 to Count2
			if request("CAT" & i)="1" then
			IDCat=request("IDCat" & i)
			IncSub=request("IncSub" & i)
			if IncSub="" then
				IncSub=0
			end if
			call opendb()
			query="UPDATE pcDFCats SET pcFCat_SubCats=" & IncSub & " WHERE pcFCat_IDDiscount=" & iddiscount & " AND pcFCat_IDCategory=" & IDCat
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
			call closedb()
			end if
		Next
	end if
	
	response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&editAction=1")

end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Filter by Category
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Filter by Customer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request("submit4")<>"" then
	intRequestSubmit=1
	Count3=request("Count3")
	if Count3="" then
		Count3="0"
	end if

	For i=1 to Count3
		if request("Cust" & i)="1" then
		IDCust=request("IDCust" & i)
		call opendb()
		query="DELETE FROM pcDFCusts WHERE pcFCust_IDDiscount=" & iddiscount & " AND pcFCust_IDCustomer=" & IDCust
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		call closedb()
		end if
	Next
	
	response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&editAction=1")
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Filter by Customer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Filter by Customer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request("submit8")<>"" then

	intRequestSubmit=1
	Count4=request("Count4")
	if Count4="" then
		Count4="0"
	end if

	For i=1 to Count4
		if request("CustCat" & i)="1" then
		IDCustCat=request("IDCustCat" & i)
		call opendb()
		query="DELETE FROM pcDFCustPriceCats WHERE pcFCPCat_IDDiscount=" & iddiscount & " AND pcFCPCat_IDCategory=" & IDCustCat
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		call closedb()
		end if
	Next
	
	response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&editAction=1")
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Filter by Customer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Security
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if Session("Admin_DC_Status")="ok" then
	Session("Admin_DC_Status")=""
	'// response.redirect "ModDiscounts.asp" & "?" & Session("Addmin_DC_Query")
else
	if Request("GoURL")<>"" then
	Session("Admin_DC_Status")="ok"
	Session("Addmin_DC_Query")=pcv_Query
	response.redirect Request("GoURL")
	end if
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Security
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


msg=""


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Update Shippng Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if (Request("Submit1")<>"") or (intRequestSubmit=1) then
	discountType=Request("discountType")
	if Request("Submit1")<>"" then
		call opendb()
		query="DELETE FROM pcDFShip WHERE pcFShip_IDDiscount=" & iddiscount
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		call closedb()
	end if
	ShipCount=-1
	If discountType="1" then
		pricetodiscount=replacecomma(Request("pricetodiscount"))
		if pricetodiscount="" or not isNumeric(pricetodiscount) then
			msg="Please enter a valid discount amount"
		end if
		percentagetodiscount="0"
		pcDisc_PerToFlatDiscount=0
		pcDisc_PerToFlatCartTotal=0
	Else

	If discountType="2" then
		percentagetodiscount=Request("percentagetodiscount")
		if not validNum(percentagetodiscount) then
			msg="Please enter a valid discount amount"
		end if
		pricetodiscount="0"
		pcDisc_PerToFlatCartTotal=replacecomma(request("pcDisc_PerToFlatCartTotal"))
		if pcDisc_PerToFlatCartTotal="" then
			pcDisc_PerToFlatCartTotal=0
		end if
		if pcDisc_PerToFlatCartTotal=0 then
			pcDisc_PerToFlatDiscount=0
		else
			pcDisc_PerToFlatDiscount=replacecomma(request("pcDisc_PerToFlatDiscount"))
			if pcDisc_PerToFlatDiscount="" then
				pcDisc_PerToFlatDiscount=0
				pcDisc_PerToFlatCartTotal=0
			end if
		end if
	Else
		SHIPS=split(request("pcv_intShippingDiscount"),", ")
		if (Request("Submit1")<>"") then
		for i=lbound(SHIPS) to ubound(SHIPS)
			if SHIPS(i)<>"" then
				call opendb()
				query="INSERT INTO pcDFShip (pcFShip_IDDiscount, pcFShip_IDShipOpt) VALUES (" & iddiscount & "," &SHIPS(i)& ")"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
				call closedb()
			end if
		next
		pricetodiscount="0"
		percentagetodiscount="0"
		pcDisc_PerToFlatDiscount=0
		pcDisc_PerToFlatCartTotal=0
		end if
		ShipCount=ubound(SHIPS)
	End if
	end if

	discountcode=replace(Request("discountcode"),"'","")	

	expDate=Request("expDate")
	if msg="" AND trim(expDate) <> "" then
		if not isDate(expDate) then
			msg="Please enter a valid expiration date"
		end if
		'// Reverse "International" Date Format for comparison and db entry 
		if scDateFrmt="DD/MM/YY" then
			pcArray_FixDate=split(expDate,"/")
			expDate=pcArray_FixDate(1) & "/" & pcArray_FixDate(0) & "/" & pcArray_FixDate(2)
		end if	
	end if
		
	startDate=Request("startDate")
	if msg="" AND trim(startDate) <> "" then
		if not isDate(startDate) then
			msg="Please enter a valid start date"
		end if
		'// Reverse "International" Date Format for comparison and db entry 
		if scDateFrmt="DD/MM/YY" then
			pcArray_FixDate=split(startDate,"/")
			startDate=pcArray_FixDate(1) & "/" & pcArray_FixDate(0) & "/" & pcArray_FixDate(2)
		end if	
	end if
	
	'Check that exp date is after start date, if there is one
	if msg="" AND trim(startDate) <> "" AND trim(expDate) <> "" then
		if datediff("d", startDate, expDate)<1 then
			msg="Your Start Date must be at least one day before your Expiration Date."
		end if
	end if
		
	if msg="" AND discountcode="" then
		msg="You must supply a discount code. Discount was not updated."
	End if
else
	ShipCount=-1
	if request("pcv_intShippingDiscount")<>"" then
		SHIPS=split(request("pcv_intShippingDiscount"),", ")
		ShipCount=ubound(SHIPS)
	end if
		discountType=Request("discountType")
		pricetodiscount=Request("pricetodiscount")
		percentagetodiscount=Request("percentagetodiscount")
		discountcode=replace(Request("discountcode"),"'","")
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Update Shippng Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Set From Data
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
discountdesc=replace(Request("discountdesc"),"&quot;","""")
discountdesc=replace(discountdesc,"'","''")
if Request("Submit1")<>"" then
	if discountdesc="" then
		discountdesc="No Description"
	end if
end if

active=Request("active")
if active="" then
	active=0
end if

archive=Request("archive")
if archive="2" then
	active="2"
end if

onetime=Request("onetime")
if onetime="" then
	onetime=0
end if

quantityfrom=Request("quantityfrom")
if quantityfrom="" then
	quantityfrom=0
end if

quantityUntil=Request("quantityuntil")
if quantityuntil="" then
	quantityuntil=9999
end if

weightfrom=Request("weightfrom")
if weightfrom="" then
	weightfrom=0
end if

weightfromSub=Request("weightfromSub")
if weightfromSub="" then
	weightfromSub=0
end if

weightUntil=Request("weightuntil")
weightUntilSub=Request("weightuntilSub")
if weightuntil="" AND weightUntilSub="" then
	weightuntil="99"
	weightUntilSub="0"
end if
if (weightUntil="") AND (weightUntilSub<>"") then
	weightuntil="99"
end if
if (weightUntil<>"") AND (weightUntilSub="") then
	weightUntilSub="0"
end if

pricefrom=replacecomma(Request("pricefrom"))
if pricefrom="" then
	pricefrom=0
end if

priceUntil=replacecomma(Request("priceuntil"))
if priceuntil="" then
	priceuntil=9999
end if

pcSeparate=Request("pcSeparate")
if pcSeparate="" then
	pcSeparate=0
end if

idproduct=Request("idproduct")
if idproduct="" then
	idproduct=0
end if

pcAuto=Request("pcAuto")
if pcAuto="" then
	pcAuto=0
end if

pcIncExcPrd=Request("IncExcPrd")
if pcIncExcPrd="" then
	pcIncExcPrd=0
end if

pcIncExcCat=Request("IncExcCat")
if pcIncExcCat="" then
	pcIncExcCat=0
end if

pcIncExcCust=Request("IncExcCust")
if pcIncExcCust="" then
	pcIncExcCust=0
end if

pcIncExcCPrice=Request("IncExcCPrice")
if pcIncExcCPrice="" then
	pcIncExcCPrice=0
end if

pcRetail=Request("Retail")
if pcRetail="" then
	pcRetail="0"
end if

pcWholesale=Request("Wholesale")
if pcWholesale="" then
	pcWholesale="0"
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Set From Data
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


sMode=Request("Submit1")


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Update
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if (sMode="Update") and (msg="") then

	if scShipFromWeightUnit="KGS" then
		weightfrom=(int(weightfrom)*1000)+int(weightfromSub)
		WeightUntil=(WeightUntil*1000)+WeightUntilSub
	else
		weightfrom=(weightfrom*16)+weightfromSub
		WeightUntil=(WeightUntil*16)+WeightUntilSub
	end if

	If expDate<>"" Then
		if SQL_Format="1" then
			expDate=(day(expDate)&"/"&month(expDate)&"/"&year(expDate))
		else
			expDate=(month(expDate)&"/"&day(expDate)&"/"&year(expDate))
		end if
	End If
	
	If startDate<>"" Then
		if SQL_Format="1" then
			startDate=(day(startDate)&"/"&month(startDate)&"/"&year(startDate))
		else
			startDate=(month(startDate)&"/"&day(startDate)&"/"&year(startDate))
		end if
	End If
	
	call openDb()
	
	if scDB="SQL" then
		strDtDelim="'"
	else
		strDtDelim="#"
	end if
	query="UPDATE discounts SET pricetodiscount="&pricetodiscount&",percentagetodiscount="&percentagetodiscount&",discountcode='"&discountcode&"',discountdesc='"&discountdesc&"', active="&active
	If expDate<>"" Then
		query=query & ",expDate="&strDtDelim&expDate&strDtDelim
	Else
		query=query & ",expDate=NULL"
	End If
	If startDate<>"" Then
		query=query & ",pcDisc_StartDate="&strDtDelim&startDate&strDtDelim
	Else
		query=query & ",pcDisc_StartDate=NULL"
	End If
	query=query & ", onetime="&onetime&", quantityfrom="&quantityfrom&", quantityuntil="&quantityuntil&", weightfrom="&weightfrom&", weightuntil="& weightuntil&", pricefrom="&pricefrom&", priceuntil="&priceuntil&", idProduct="&idProduct&", pcSeparate="&pcSeparate&", pcDisc_Auto="&pcAuto&", pcRetailFlag ="&pcRetail&", pcWholesaleFlag="&pcWholeSale&", pcDisc_PerToFlatCartTotal="&pcDisc_PerToFlatCartTotal&", pcDisc_PerToFlatDiscount="&pcDisc_PerToFlatDiscount&",pcDisc_IncExcPrd=" & pcIncExcPrd & ",pcDisc_IncExcCat=" & pcIncExcCat & ",pcDisc_IncExcCust=" & pcIncExcCust & ",pcDisc_IncExcCPrice=" & pcIncExcCPrice & " WHERE iddiscount="&iddiscount
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	call closedb()
	
	response.redirect("modDiscounts.asp?mode=Edit&iddiscount="&iddiscount&"&editAction=1")
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Update
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


iMode=Request.QueryString("mode")


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Edit Mode
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if iMode="Edit" then
	
	call openDb()
	
	query="SELECT discountdesc, pricetodiscount, percentagetodiscount, discountcode, active, expDate, onetime, quantityfrom, quantityuntil, weightfrom, weightuntil, pricefrom, priceuntil, pcSeparate, pcDisc_Auto, pcDisc_StartDate ,pcRetailFlag, pcWholesaleFlag, pcDisc_PerToFlatCartTotal, pcDisc_PerToFlatDiscount,pcDisc_IncExcPrd,pcDisc_IncExcCat,pcDisc_IncExcCust,pcDisc_IncExcCPrice FROM discounts WHERE iddiscount="& iddiscount
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	discountdesc=rs("discountdesc")
	if discountdesc<>"" AND isNULL(discountdesc)=False then		
		discountdesc=replace(discountdesc,"""","&quot;")
	end if
	pricetodiscount=rs("pricetodiscount")
	if pricetodiscount<>"" then
	else
		pricetodiscount="0"
	end if
	if pricetodiscount>"0" then
		discounttype="1"
	end if

	percentagetodiscount=rs("percentagetodiscount")
	if percentagetodiscount<>"" then
	else
		percentagetodiscount="0"
	end if
	if percentagetodiscount>"0" then
		discounttype="2"
	end if

	discountcode=rs("discountcode")
	active=rs("active")
	expDate=rs("expDate")
	if not isNull(expDate) and trim(expDate)<>"" then
		expDate=ShowDateFrmt(expDate)
	end if
	if trim(expDate) =  "//" then
		expDate=""
	end if
	
	startDate=rs("pcDisc_StartDate")
	if not isNull(startDate) and trim(startDate)<>"" then
		startDate=ShowDateFrmt(startDate)
	end if
	if trim(startDate) =  "//" then
		startDate=""
	end if
	
	onetime=rs("onetime")
	quantityfrom=rs("quantityfrom")
	if quantityfrom="" then
		quantityfrom="0"
	end if
	quantityUntil=rs("quantityuntil")
	if quantityuntil="" then
		quantityuntil="9999"
	end if
	weightfrom=rs("weightfrom")
	if weightfrom="" then
		weightfrom="0"
	end if
	weightUntil=rs("weightuntil")
	if weightuntil="" then
		weightuntil="9999"
	end if
	if weightfrom>"0" then
	if scShipFromWeightUnit="KGS" then
		weightfrom1=weightfrom
		weightfrom=fix(int(weightfrom)/1000)
		weightfromSub=int(weightfrom1)-int(weightfrom*1000)
	else
		weightfrom1=weightfrom
		weightfrom=fix(int(weightfrom)/16)
		weightfromSub=int(weightfrom1)-int(weightfrom*16)
	end if
	end if
	if (weightUntil<>"9999") and (weightUntil>"0") then
	if scShipFromWeightUnit="KGS" then
		WeightUntil1=WeightUntil
		WeightUntil=fix(int(WeightUntil)/1000)
		WeightUntilSub=int(WeightUntil1)-int(WeightUntil*1000)
	else
		WeightUntil1=WeightUntil
		WeightUntil=fix(int(WeightUntil)/16)
		WeightUntilSub=int(WeightUntil1)-int(WeightUntil*16)
	end if
	end if
	pricefrom=rs("pricefrom")
	if pricefrom="" then
		pricefrom="0"
	end if
	priceUntil=rs("priceuntil")
	if priceuntil="" then
		priceuntil="9999"
	end if
	pcSeparate=rs("pcSeparate")
	pcAuto=rs("pcDisc_Auto")
	pcRetail = rs("pcRetailFlag")
	pcWholeSale = rs("pcWholeSaleFlag")
	pcDisc_PerToFlatCartTotal = rs("pcDisc_PerToFlatCartTotal")
	pcDisc_PerToFlatDiscount = rs("pcDisc_PerToFlatDiscount")
	pcIncExcPrd=rs("pcDisc_IncExcPrd")
	pcIncExcCat=rs("pcDisc_IncExcCat")
	pcIncExcCust=rs("pcDisc_IncExcCust")
	pcIncExcCPrice=rs("pcDisc_IncExcCPrice")

	idproduct="0"
	set rs=nothing

	query="SELECT pcFShip_IDShipOpt FROM pcDFShip WHERE pcFShip_IDDiscount="& iddiscount
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	Dim SHIPS(300)
	For Count=0 to 299
		SHIPS(Count)=""
	next
	if not rs.eof then
		discounttype="3"
		ShipCount=-1
		do while not rs.eof
			ShipCount=ShipCount+1
			SHIPS(ShipCount)=rs("pcFShip_IDShipOpt")
			rs.MoveNext
		loop
	end if
	set rs=nothing
	call closedb()
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Edit Mode
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Delete Mode
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if iMode="Del" then
	iddiscount=request.querystring("iddiscount")

	call openDb()
	Set rs=Server.CreateObject("ADODB.Recordset")
	query="DELETE FROM discounts WHERE iddiscount="& iddiscount
	rs.Open query, connTemp, adOpenStatic, adLockReadOnly
	
	query="DELETE FROM pcDFProds WHERE pcFPro_IDDiscount="& iddiscount
	rs.Open query, connTemp, adOpenStatic, adLockReadOnly
	
	query="DELETE FROM pcDFCats WHERE pcFCat_IDDiscount="& iddiscount
	rs.Open query, connTemp, adOpenStatic, adLockReadOnly
	
	query="DELETE FROM pcDFCusts WHERE pcFCust_IDDiscount="& iddiscount
	rs.Open query, connTemp, adOpenStatic, adLockReadOnly
	
	query="DELETE FROM pcDFShip WHERE pcFShip_IDDiscount="& iddiscount
	rs.Open query, connTemp, adOpenStatic, adLockReadOnly
	
		
	query="DELETE FROM pcDFCustPriceCats WHERE pcFCPCat_IDDiscount="& iddiscount
	rs.Open query, connTemp, adOpenStatic, adLockReadOnly

		
	connTemp.Close
	Set rs=Nothing
	Set connTemp=Nothing
	
	response.redirect("AdminDiscounts.asp")
End if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Delete Mode
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: On Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

call opendb()

pcv_Filter=0
query="SELECT pcFPro_IDProduct FROM pcDFProds WHERE pcFPro_IDDiscount=" & iddiscount
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if not rs.eof then

	query="DELETE FROM pcDFCats WHERE pcFCat_IDDiscount=" & iddiscount
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_Filter=1
	
else

	query="SELECT pcFCat_IDCategory FROM pcDFCats WHERE pcFCat_IDDiscount=" & iddiscount
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if not rs.eof then
		query="DELETE FROM pcDFProds WHERE pcFPro_IDDiscount=" & iddiscount
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcv_Filter=2
	end if
	
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ENd: On Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<script>
function Form1_Validator(theForm)
{
	if (theForm.discountdesc.value.indexOf("'")>=0)
	{
		alert("You cannot use apostrophes in the discount description.");
		theForm.discountdesc.focus();
		return (false);
	}
	if (theForm.discountcode.value.indexOf("'")>=0)
	{
		alert("You cannot use apostrophes in the discount code.");
		theForm.discountcode.focus();
		return (false);
	}
	if (theForm.discountdesc.value.indexOf(",")>=0)
	{
		alert("You cannot use commas in the discount description.");
		theForm.discountdesc.focus();
		return (false);
	}
	if (theForm.discountcode.value.indexOf(",")>=0)
	{
		alert("You cannot use commas in the discount code.");
		theForm.discountcode.focus();
		return (false);
	}
	if (theForm.discountcode.value=="")
	{
		alert("Please enter a discount code.");
		theForm.discountcode.focus();
		return (false);
	}
	if (theForm.discountdesc.value=="")
	{
		alert("Please enter a discount description.");
		theForm.discountdesc.focus();
		return (false);
	}
	if (theForm.clicksav.value=="1")
	{
		if ((theForm.discount1.value=="") || (theForm.discount1.value=="0"))
	{
		alert("Please select a discount type.");
			theForm.pricetodiscount.focus();
			return (false);
	}
	if ((theForm.discount1.value=="3") && ((theForm.pcv_intShippingDiscount.value=="0") || (theForm.pcv_intShippingDiscount.value=="")))
	{
			alert("Please select a shipping service.");
		theForm.pcv_intShippingDiscount.focus();
			return( false);
	}
	}
	return (true);
}
</script>

<form method="post" name="hForm" action="modDiscounts.asp?<%=pcv_Query%>" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type=hidden value="<%=discounttype%>" name="discount1">          
<table class="pcCPcontent">
		<%
	  	 if editAction="1" and msg="" then 
	  		msg="Discount code updated successfully! <br />Return to the <a href=AdminDiscounts.asp>Discounts by Code summary</a> page."
			msgType=1
		 end if	
		%>
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>      
      <tr>
				<th colspan="3">General Information</th>
			</tr>   
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr> 
		<tr>
				<td nowrap width="20%">Discount Description:</td>
				<td colspan="2" width="80%">
					<input name="discountdesc" id="discountdesc" size="30" value="<%=discountdesc%>">&nbsp;Shown to the customer during checkout and on order receipts.
					<input name="discountdesc2" type="hidden" size="40" value="<%=discountdesc%>">
            </td>
		</tr>
		<tr>
			<td>Discount Code:</td>
				<td colspan="2">
					<input name="discountcode" id"discountcode" size="30" value="<%=discountcode%>">&nbsp;Used by customers to apply the discount to an order
					<input name="discountcode2" type="hidden" value="<%=discountcode%>">
            </td>
		</tr>
		<tr>
				<td colspan="3"><hr></td>
		</tr>
		<tr>
				<td colspan="3">Type of Discount:</td>
		</tr>
		<tr>
				<td colspan="3"> 
					<table width="100%" border="0" cellspacing="0" cellpadding="2">
		<tr>
							<td width="5%" align="right"><input type="radio" name="discountType" value="1" onClick="hForm.discount1.value='1';" <%if discountType=1 then%>checked<%end if%> class="clearBorder"></td>
							<td width="25%">Price Discount</td>
							<td width="70%"><%=scCurSign%>&nbsp;<input name="pricetodiscount" size="16" value="<%=money(pricetodiscount)%>"></td>
					</tr>
			
					<tr>
					<td align="right">
					<input type="radio" name="discountType" value="2" onClick="hForm.discount1.value='2';" <%if discountType=2 then%>checked<%end if%> class="clearBorder"></td>
					<td>Percent Discount</td>
			        <td>%&nbsp;<input name="percentagetodiscount" size="16" value="<%=percentagetodiscount%>"></td>
					</tr>
					<tr>
			          <td>&nbsp;</td>
					  <td colspan="2">
				      	<div class="pcCPnotes">Use the settings below to set values that will allow you to change the &quot;Percent Discount&quot; to a &quot;Price Discount&quot; fee based on the cart total. Setting the values to zero (0) will ignore this feature and always use the percentage set above.</div>
				        <div style="padding: 5px; margin-bottom: 15px;">
				            If the total ordered is over: <%=scCurSign%>&nbsp;
				            <input type="text" name="pcDisc_PerToFlatCartTotal" value="<%=money(pcDisc_PerToFlatCartTotal)%>" size="16">
				            ... switch to the following flat discount: <%=scCurSign%>&nbsp;
				            <input type="text" name="pcDisc_PerToFlatDiscount" value="<%=money(pcDisc_PerToFlatDiscount)%>" size="16">
				        </div>
			        </td>
					  </tr>
					<tr>
						<td align="right" valign="top">
							<input type="radio" name="discountType" value="3" onClick="hForm.discount1.value='3';" <%if discountType=3 then%>checked<%end if%> class="clearBorder"></td>
						<td valign="top">Free Shipping <div class="pcCPnotes">To select multiple shipping services use the CTRL key.</div>
						</td>
						<td>
							<select name="pcv_intShippingDiscount" size="5" multiple>
							<%
							query="SELECT idshipservice, servicePriority, serviceDescription FROM shipService WHERE serviceActive=-1 ORDER BY servicePriority;"
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							do until rs.eof%>
							<option value="<%=rs("idshipservice")%>" <%
									for i=0 to ShipCount
			
			
									if SHIPS(i)<>"" then
									if clng(SHIPS(i))=clng(rs("idshipservice")) then%>selected<%end if
									end if
									next
									%>><%=rs("serviceDescription")%></option>
							<%
							rs.MoveNext
							loop
							set rs=nothing
							%>
							</select>
						</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>  
			<tr><th colspan="3">Status and Expiration</th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td>Active:</td>
				<td colspan="2">Yes <input type="checkbox" name="active" value="-1" <%if active="-1" then%>checked<%end if%><%if active="2" then%>disabled<%end if%> class="clearBorder"></td>
			</tr>
			<tr>
				<td>Archive:</td>
				<td colspan="2">Yes <input type="checkbox" name="archive" value="2" <%if active="2" then%>checked<%end if%> onclick="if (this.checked) {document.hForm.active.disabled=true; document.hForm.active.checked=false} else {document.hForm.active.disabled=false}" class="clearBorder"></td>
			</tr>
			<tr>
				<td>Start date:</td>
				<td colspan="2">
					<input type="text" name="startDate" value="<%=startDate%>">
					<span class="pcCPnotes">Format: <%=lcase(scDateFrmt)%></span>
				</td>
			</tr>
			<tr>
				<td>Expiration date:</td>
				<td colspan="2">
					<input type="text" name="expDate" value="<%=expDate%>">
					<span class="pcCPnotes">Format: <%=lcase(scDateFrmt)%></span>
				</td>
			</tr>
			<tr>
				<td>One Time Only:</td>
				<td colspan="2">Yes 
				<input type="checkbox" name="onetime" value="-1" <%if onetime="-1" then%>checked<%end if%> class="clearBorder">&nbsp;Check this option for discounts that a customer can only use once</td>
			</tr>
			<tr>
            	<td valign="top">Automatically Apply?</td>
				<td colspan="2">
				  <input name="pcAuto" type="radio" value="1" class="clearBorder" <%if pcAuto="1" then%>checked<%end if%>>Yes&nbsp; 
                  <input name="pcAuto" type="radio" value="0" class="clearBorder" <%if pcAuto<>"1" then%>checked<%end if%>>No&nbsp;|&nbsp;
				  The discount is automatically applied to the order, when applicable, and shown on the order verification page.
				  </td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<th colspan="3">Parameters that Restrict Applicability</th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr> 
			<tr>
				<td>Order Quantity:</td>
				<td nowrap>From: 
				<input name="quantityFrom" size="6" value="<%=quantityFrom%>">
                </td>
				<td nowrap>To: 
				<input name="quantityUntil" size="6" value="<%=quantityUntil%>">
                </td>
			</tr>
			<%
				if scShipFromWeightUnit="KGS" then
					u1="kg"
					u2="g"
				else
					u1="lb"
					u2="oz"
				end if
			%>
			<tr> 
				<td>Order Weight:</td>
				<td nowrap>From: 
				<input name="WeightFrom" size="2" value="<%=WeightFrom%>">
				&nbsp;<%=u1%>&nbsp;
			
				<input name="WeightFromSub" size="2" value="<%=WeightFromSub%>">
				&nbsp;<%=u2%>
				</td>
				<td nowrap>To: 
				<input name="WeightUntil" size="2" value="<%=WeightUntil%>">
				&nbsp;<%=u1%>&nbsp;
			
				<input name="WeightUntilSub" size="2" value="<%=WeightUntilSub%>">
				&nbsp;
				<%=u2%>
                </td>
			</tr>
			<tr>
				<td>Order Amount:</td>
				<td nowrap>From: <%=scCurSign%> 
				<input name="priceFrom" size="6" value="<%=money(priceFrom)%>">
                </td>
				<td nowrap>To: <%=scCurSign%> 
				<input name="priceUntil" size="6" value="<%=money(priceUntil)%>">
				</td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>
			<tr>
            	<td>Use with others?</td>
				<td colspan="2">
					<input name="pcSeparate" type="radio" value="1" class="clearBorder" <%if pcSeparate="1" then%>checked<%end if%>>Yes&nbsp; 
					<input name="pcSeparate" type="radio" value="0" class="clearBorder" <%if pcSeparate<>"1" then%>checked<%end if%>>No&nbsp;|&nbsp;If "yes" this discount code can be used with other discount codes.</td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="3"><h2>Filter by Product(s)</h2>
                If no products are selected, the discount applies to orders that contain any products.
                </td>
			</tr>
            <tr>
				<td colspan="3"><input type="radio" name="IncExcPrd" value="0" class="clearBorder" <%if pcIncExcPrd<>"1" then%>checked<%end if%>> Include selected products&nbsp;&nbsp;<input type="radio" name="IncExcPrd" value="1" class="clearBorder" <%if pcIncExcPrd="1" then%>checked<%end if%>> Exclude selected products </td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">
						<%
						query="SELECT products.idproduct,products.description FROM products,pcDFProds WHERE products.idproduct=pcDFProds.pcFPro_IDProduct and pcDFProds.pcFPro_IDDiscount=" & iddiscount &" order by products.description"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
						Count1=0
						do while not rs.eof
							Count1=Count1+1
							pIDProduct=rs("IDProduct")
							pName=rs("description")
							%>
							<tr>
							<td><%=pName%></td><td>
							<input type="checkbox" name="Pro<%=Count1%>" value="1" class="clearBorder">
							<input type=hidden name="IDPro<%=Count1%>" value="<%=pIDProduct%>">
							</td>
							</tr>
							<%rs.MoveNext
						loop%>
							<tr>
							<td colspan="2">
							<%if Count1>0 then%>
							<a href="javascript:checkAllPrd();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllPrd();">Uncheck All</a>
							<script language="JavaScript">
							<!--
							function checkAllPrd() {
							for (var j = 1; j <= <%=count1%>; j++) {
							box = eval("document.hForm.Pro" + j); 
							if (box.checked == false) box.checked = true;
								 }
							}
							
							function uncheckAllPrd() {
							for (var j = 1; j <= <%=count1%>; j++) {
							box = eval("document.hForm.Pro" + j); 
							if (box.checked == true) box.checked = false;
								 }
							}
									
							//-->
							</script>
							<%else%>
									No Items to display.
							<%end if%>
							</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
				<td colspan="3">
				<%if Count1>0 then%>
				<input type="hidden" name="Count1" value="<%=Count1%>">
				<input type="submit" name="submit2" value="Remove Selected Product(s)">
				&nbsp;
				<%end if%>
				<input type="submit" name="submit5" value="Add Products" onclick="document.hForm.GoURL.value='addprdsToDc.asp?idcode=<%=iddiscount%>';" <%if pcv_Filter=2 then%>disabled="disabled"<%end if%>>
				</td>
				</tr>
				<tr>
                    <td colspan="3" class="pcCPspacer"><img src="images/pc_admin.gif" width="85" height="19" alt="Or"></td>
				</tr>
				<tr>
				<td colspan="3"><h2>Filter by Category</h2>
                If no categories and no products are selected, the discount applies to orders that contain any products.
                </td>
                </tr>
                <tr>
				<td colspan="3"><input type="radio" name="IncExcCat" value="0" class="clearBorder" <%if pcIncExcCat<>"1" then%>checked<%end if%>> Include selected categories&nbsp;&nbsp;<input type="radio" name="IncExcCat" value="1" class="clearBorder" <%if pcIncExcCat="1" then%>checked<%end if%>> Exclude selected categories</td>
				</tr>
                <tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">
						<%
						query="SELECT categories.idcategory,categories.categoryDesc,pcDFCats.pcFCat_SubCats FROM categories,pcDFCats WHERE categories.idcategory=pcDFCats.pcFCat_IDCategory and pcDFCats.pcFCat_IDDiscount=" & iddiscount &" order by categories.categoryDesc"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
						Count2=0
						if not rs.eof then
						%>
						<tr><td></td><td>Include Subcategories</td><td></td></tr>
						<%
						end if
						do while not rs.eof
							Count2=Count2+1
							pIDCat=rs("IDCategory")
							pName=rs("categoryDesc")
							pSubCats=rs("pcFCat_SubCats")
							if pSubCats<>"" then
							else
							pSubCats="0"
							end if%>
							<tr>
							<td><%=pName%></td>
							<td>
							<input type="checkbox" name="IncSub<%=Count2%>" value="1" <%if pSubCats="1" then%>checked<%end if%> class="clearBorder">
							</td>
							<td>
							<input type="checkbox" name="CAT<%=Count2%>" value="1" class="clearBorder">
							<input type="hidden" name="IDCat<%=Count2%>" value="<%=pIDCAT%>" >
							</td></tr>
							<%rs.MoveNext
						loop
						
						set rs=nothing
						%>
								<tr>
								<td colspan="2">
								<%if Count2>0 then%>
								<a href="javascript:checkAllCat();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllCat();">Uncheck All</a>
								<script language="JavaScript">
								<!--
								function checkAllCat() {
								for (var j = 1; j <= <%=count2%>; j++) {
								box = eval("document.hForm.CAT" + j); 
								if (box.checked == false) box.checked = true;
									 }
								}
								
								function uncheckAllCat() {
								for (var j = 1; j <= <%=count2%>; j++) {
								box = eval("document.hForm.CAT" + j); 
								if (box.checked == true) box.checked = false;
									 }
								}
								
								//-->
								</script>
								<%else%>
								No categories to display.
								<%end if%>
								</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
					<td colspan="3">
					<%if Count2>0 then%>
					<input type="hidden" name="Count2" value="<%=Count2%>">
					<input type="submit" name="submit3A" value="Update Selected Categories">
					&nbsp;
					<input type="submit" name="submit3" value="Remove Selected Categories">
					&nbsp;
					<%end if%>
					<input type="submit" name="submit6" value="Add Categories" onclick="document.hForm.GoURL.value='addcatsToDc.asp?idcode=<%=iddiscount%>';" <%if pcv_Filter=1 then%>disabled="disabled"<%end if%>>
					</td>
					</tr>
					<tr>
				<td colspan="3" class="pcCPspacer"><div class="pcCPnotes" style="margin: 10px 0 10px 0;"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Note about this feature" hspace="5">You cannot use <em><strong>Filter by Product(s)</strong></em> and <em><strong>Filter by Category</strong></em> at the same time. <a href="http://wiki.earlyimpact.com/productcart/marketing-discounts_by_code#limiting_applicability" target="_blank">See the documentation</a> for details.</div></td>
			</tr>
					<tr>
				<td colspan="3"><h2>Filter by Customer(s)</h2>
                If no customers are selected, the discount can be used by anyone.
                </td>
                    </tr>
                    <tr>
					<td colspan="3"><input type="radio" name="IncExcCust" value="0" class="clearBorder" <%if pcIncExcCust<>"1" then%>checked<%end if%>> Include selected customers&nbsp;&nbsp;<input type="radio" name="IncExcCust" value="1" class="clearBorder" <%if pcIncExcCust="1" then%>checked<%end if%>> Exclude selected customers</td>
					</tr>
                    <tr>
					<td colspan="3">
						<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">	
							<%
							query="SELECT customers.idcustomer,customers.name,customers.lastname FROM customers,pcDFCusts WHERE customers.idcustomer=pcDFCusts.pcFCust_IDCustomer and pcDFCusts.pcFCust_IDDiscount=" & iddiscount &" order by customers.name"
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							Count3=0
							do while not rs.eof
								Count3=Count3+1
								pIDCust=rs("IDCustomer")
								pName=rs("name") & " " & rs("lastname")%>
								<tr>
								<td><%=pName%></td><td>
								<input type="checkbox" name="Cust<%=Count3%>" value="1" class="clearBorder">
								<input type="hidden" name="IDCust<%=Count3%>" value="<%=pIDCust%>" >
								</td></tr>
								<%rs.MoveNext
							loop
							set rs=nothing
							call closedb()
							%>
							<tr>
							<td colspan="2">
							<%if Count3>0 then%>
							<a href="javascript:checkAllCust();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllCust();">Uncheck All</a>
							<script language="JavaScript">
							<!--
							function checkAllCust() {
							for (var j = 1; j <= <%=count3%>; j++) {
							box = eval("document.hForm.Cust" + j); 
							if (box.checked == false) box.checked = true;
								 }
							}
							
							function uncheckAllCust() {
							for (var j = 1; j <= <%=count3%>; j++) {
							box = eval("document.hForm.Cust" + j); 
							if (box.checked == true) box.checked = false;
								 }
							}
							
							//-->
							</script>
							<%else%>
								No Customers to display.
							<%end if%>
                           </td>
						</tr>
							</table>
						</td>
					</tr>
					<tr>
					<td colspan="3">
					<%if Count3>0 then%>
					<input type="hidden" name="Count3" value="<%=Count3%>">
					<input type="submit" name="submit4" value="Remove Selected Customer(s)">
					&nbsp;
					<%end if%>
					<input type="submit" name="submit7" value="Add Customers" onclick="document.hForm.GoURL.value='addcustsToDc.asp?idcode=<%=iddiscount%>';">
					</td>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>
					<tr>
				<td colspan="3"><h2>Filter by Customer Pricing Category</h2>
				If no categories are selected, the discount can be used by anyone.</td>
			</tr>
            <tr>
				<td colspan="3"><input type="radio" name="IncExcCPrice" value="0" class="clearBorder" <%if pcIncExcCPrice<>1 then%>checked<%end if%>> Include selected customer categories&nbsp;&nbsp;<input type="radio" name="IncExcCPrice" value="1" class="clearBorder" <%if pcIncExcCPrice="1" then%>checked<%end if%>> Exclude selected customer categories</td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">
						<% pcArray=split(session("admin_DiscFCustPriceCATs"),",")
						Count4=0
						call opendb()
																	
							query="SELECT pcCustomerCategories.idcustomerCategory, pcCustomerCategories.pcCC_Name FROM pcCustomerCategories , pcDFCustPriceCats  Where pcCustomerCategories.idcustomerCategory = pcDFCustPriceCats.pcFCPCat_IDCategory And pcDFCustPriceCats.pcFCPCat_IDDiscount ="&iddiscount &" Order by pcCustomerCategories.pcCC_Name ;"
							
					 
							
							SET rs=Server.CreateObject("ADODB.RecordSet")
							SET rs=conntemp.execute(query)
							if NOT rs.eof then
							Do while not RS.eof			
							Count4 = Count4 + 1 					
							intIdcustomerCategory=rs("idcustomerCategory")
							strpcCC_Name=rs("pcCC_Name")%>								
								<tr>
									<td><%=strpcCC_Name%></td><td>
										<input type="checkbox" name="CustCat<%=Count4%>" value="1" class="clearBorder">
										<input type="hidden" name="IDCustCat<%=Count4%>" value="<%=intIdcustomerCategory%>">
									</td>
								</tr>
								<%
							
						 
						rs.movenext
						loop
						end if
						set rs = nothing							
						call closedb() %>
						<tr>
							<td colspan="2">
							<%if Count4>0 then%>
								<a href="javascript:checkAllCustCat();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllCustCat();">Uncheck All</a>
								<script language="JavaScript">
								<!--
								function checkAllCustCat() {
								for (var j = 1; j <= <%=count4%>; j++) {
								box = eval("document.hForm.CustCat" + j); 
								if (box.checked == false) box.checked = true;
								}
								}
								
								function uncheckAllCustCat() {
								for (var j = 1; j <= <%=count4%>; j++) {
								box = eval("document.hForm.CustCat" + j); 
								if (box.checked == true) box.checked = false;
								}
								}
								
								//-->
								</script>
							<%else%>
								No Customer Pricing Categories to display.
							<%end if%>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="3">
					<%if Count4>0 then%>
						<input type="hidden" name="Count4" value="<%=Count4%>">
						<input type="submit" name="submit8" value="Remove Selected Pricing Categories">
						&nbsp;
					<%end if%>
					<input type="submit" name="submit9" value="Add Pricing Catgeories" onclick="document.hForm.GoURL.value='addCustPriceCatsToDc.asp?idcode=<%=iddiscount%>';">
				</td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>  
			<tr>
			  <td colspan="3"><h2>Filter by Customer Type</h2>
If no Types are selected, the discount can be used by anyone (check all that apply):</td>
		    </tr>
			<tr>
			  <td >Retail Customers:</td>
		      <td ><input type="checkbox" name="Retail" value="1" class="clearBorder" <%if pcRetail ="1" then %> checked <% end if %> >&nbsp;</td>
		      <td >&nbsp;</td>
			</tr>
			<tr>
				<td>Wholesale Customers: </td>
			    <td><input type="checkbox" name="Wholesale" value="1" class="clearBorder"  <%if pcWholeSale ="1" then %> checked <% end if %>>&nbsp;</td>
			    <td>&nbsp;</td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>  
					<tr>
					<td colspan="3" align="center">
						<input type="submit" name="submit1" value="Update" class="submit2" onclick="hForm.clicksav.value='1';">
						<input type="hidden" name="clicksav" value="">
						&nbsp;
						<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			            &nbsp;
			            <input type="button" name="discountHome" value="View All" onClick="document.location.href='adminDiscounts.asp'">
						<input type="hidden" name="iddiscount" value="<%=iddiscount%>">
			            <input type="hidden" name="editAction" value="1">
						</td>
					</tr>
				</table>
<input type="hidden" name="GoURL" value="">
</form>
<!--#include file="AdminFooter.asp"-->