<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pageTitle="Export Ordered Products"%>
<% Section="products" %>
<%PmAdmin=2%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->  
<!--#include file="../includes/SQLFormat.txt"-->
<%
totalrecords=0
Dim connTemp
call opendb()

Function getEventName(tmpIDEP)
Dim rs1,query
	query="SELECT pcEvents.pcEv_Name FROM pcEvProducts,pcEvents WHERE pcEvProducts.pcEP_ID=" & tmpIDEP & " AND pcEvents.pcEv_IDEvent=pcEvProducts.pcEP_IDEvent;"
	set rs1=connTemp.execute(query)
	getEventName=rs1("pcEv_Name")
	set rs1=nothing
End Function

Function getGWOptName(tmpIDGWOpt)
Dim rs1,query
	query="SELECT pcGW_OptName FROM pcGWOptions WHERE pcGW_IDOpt=" & tmpIDGWOpt & ";"
	set rs1=connTemp.execute(query)
	getGWOptName=rs1("pcGW_OptName")
	set rs1=nothing
End Function

Function getBTOConf(pIdConfigSession)
Dim rs1,rs2,query,tmpResult
	tmpResult=""
	
	query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
	set rs1=server.CreateObject("ADODB.RecordSet")
	set rs1=connTemp.execute(query)

	stringProducts=rs1("stringProducts")
	stringValues=rs1("stringValues")
	stringCategories=rs1("stringCategories")
	stringQuantity=rs1("stringQuantity")
	stringPrice=rs1("stringPrice")
	ArrProduct=Split(stringProducts, ",")
	ArrValue=Split(stringValues, ",")
	ArrCategory=Split(stringCategories, ",")
	ArrQuantity=Split(stringQuantity, ",")
	ArrPrice=Split(stringPrice, ",")
	tmpResult="Customizations:<br>"
	for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
		query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
		set rs2=server.CreateObject("ADODB.RecordSet")
		set rs2=connTemp.execute(query)
		pcategoryDesc=rs2("categoryDesc")
		pdescription=rs2("description")
		psku=rs2("sku")
		if NOT isNumeric(ArrQuantity(i)) then
			pIntQty=1
		else
			pIntQty=ArrQuantity(i)
		end if
		tmpResult=tmpResult & pcategoryDesc & ": " & psku & " - " & pdescription
		if pIntQty>1 then
			tmpResult=tmpResult & " - QTY: " & ArrQuantity(i)
		end if
		tmpResult=tmpResult & "<br>"
		set rs2=nothing
	next
	set rs1=nothing
	
	query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
	set rs1=server.CreateObject("ADODB.RecordSet")
	set rs1=connTemp.execute(query)
				
	stringCProducts=rs1("stringCProducts")
	stringCValues=rs1("stringCValues")
	stringCCategories=rs1("stringCCategories")
	ArrCProduct=Split(stringCProducts, ",")
	ArrCValue=Split(stringCValues, ",")
	ArrCCategory=Split(stringCCategories, ",")
	if ArrCProduct(0)<>"na" then
		tmpResult=tmpResult & "Additional Charges:<br>"
		for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
			query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
			set rs2=server.CreateObject("ADODB.RecordSet")
			set rs2=connTemp.execute(query)
			pcategoryDesc=rs2("categoryDesc")
			pdescription=rs2("description")
			psku=rs2("sku")
			tmpResult=tmpResult & pcategoryDesc & ": " & psku & " - " & pdescription & "<br>"
			set rs2=nothing
		next
	end if
	set rs1=nothing

	getBTOConf=tmpResult
End Function

pcv_OrderID=request("pcv_OrderID")
pcv_PrdSKU=request("pcv_PrdSKU")
pcv_PrdName=request("pcv_PrdName")
pcv_UnitPrice=request("pcv_UnitPrice")
pcv_Units=request("pcv_Units")
pcv_WholesalePrice=request("pcv_WholesalePrice")
pcv_TotalPrice=request("pcv_TotalPrice")
pcv_BTOConf=request("pcv_BTOConf")
pcv_POptions=request("pcv_POptions")
pcv_QDiscounts=request("pcv_QDiscounts")
pcv_IDiscounts=request("pcv_IDiscounts")
pcv_EventName=request("pcv_EventName")
pcv_GWOption=request("pcv_GWOption")
pcv_GWPrice=request("pcv_GWPrice")
pcv_PackageID=request("pcv_PackageID")
pcv_PCost=request("pcv_PCost")
pcv_Margin=request("pcv_Margin")


Dim strTDateVar, strTDateVar2, DateVar, DateVar2
strTDateVar=Request("FromDate")
DateVar=strTDateVar
if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
end if
strTDateVar2=Request("ToDate")
DateVar2=strTDateVar2
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	if err.number<>0 then
		DateVar=Request("FromDate")
		DateVar2=Request("ToDate")
	end if
end if
err.clear

if (DateVar<>"") and IsDate(DateVar) then
    if SQL_Format = "1" then DateVar = day(DateVar) & "/" & month(DateVar) & "/" & year(DateVar)
	if scDB="Access" then
		query1=" AND orders.orderDate >=#" & DateVar & "# "
	else
		query1=" AND orders.orderDate >='" & DateVar & "' "
	end if
else
	query1=""
end if
if (DateVar2<>"") and IsDate(DateVar2) then
    if SQL_Format = "1" then DateVar2 = day(DateVar2) & "/" & month(DateVar2) & "/" & year(DateVar2)
	if scDB="Access" then
		query2=" AND orders.orderDate <=#" & DateVar2 & "# "
	else
		query2=" AND orders.orderDate <='" & DateVar2 & "' "
	end if
else
	query2=""	
end if

query="SELECT ProductsOrdered.IDOrder,Products.SKU,Products.Description,ProductsOrdered.unitPrice,ProductsOrdered.quantity,Products.btoBPrice,ProductsOrdered.idconfigSession,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts,ProductsOrdered.pcPackageInfo_ID,ProductsOrdered.pcPrdOrd_SelectedOptions,ProductsOrdered.pcPrdOrd_OptionsPriceArray,ProductsOrdered.pcPrdOrd_OptionsArray,ProductsOrdered.pcPO_EPID,ProductsOrdered.pcPO_GWOpt,ProductsOrdered.pcPO_GWPrice,Products.Cost FROM Products INNER JOIN (ProductsOrdered INNER JOIN Orders ON ProductsOrdered.IDOrder=Orders.IDOrder) ON Products.idproduct=ProductsOrdered.idproduct WHERE ((orders.orderstatus>2 AND orders.orderstatus<5) OR (orders.orderstatus>6 AND orders.orderstatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & query1 & query2 & " ORDER BY ProductsOrdered.IDOrder,ProductsOrdered.IDProduct;"
Set rstemp=Server.CreateObject("ADODB.Recordset")

set rstemp=connTemp.execute(query)

IF rstemp.eof THEN
%>
<!--#include file="AdminHeader.asp"-->
		<div class="pcCPmessage">
			Your search did not return any results. <a href="exportData.asp">Back</a>.
		</div>
<!--#include file="AdminFooter.asp"-->
<%
ELSE

Dim strCol, Count, HTMLResult
set StringBuilderObj = new StringBuilder
CSVResult=""
Count = 0

Function GenFileName()
	dim fname
	fname="File-"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime)) & "-"
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

StringBuilderObj.append "<table>" & vbcrlf
StringBuilderObj.append "<tr>" & vbcrlf
if pcv_OrderID="1" then
StringBuilderObj.append "<td><b>Order ID</b></td>" & vbcrlf
end if
if pcv_PrdSKU="1" then
StringBuilderObj.append "<td><b>Product SKU</b></td>" & vbcrlf
end if
if pcv_PrdName="1" then
StringBuilderObj.append "<td><b>Product Name</b></td>" & vbcrlf
end if
if pcv_UnitPrice="1" then
StringBuilderObj.append "<td><b>Unit Price</b></td>" & vbcrlf
end if
if pcv_Units="1" then
StringBuilderObj.append "<td><b>Units</b></td>" & vbcrlf
end if
if pcv_WholesalePrice="1" then
StringBuilderObj.append "<td><b>Wholesale Price</b></td>" & vbcrlf
end if
if pcv_BTOConf="1" then
StringBuilderObj.append "<td><b>BTO Configuration</b></td>" & vbcrlf
end if
if pcv_POptions="1" then
StringBuilderObj.append "<td><b>Product Options</b></td>" & vbcrlf
end if
if pcv_QDiscounts="1" then
StringBuilderObj.append "<td><b>Quantity Discounts</b></td>" & vbcrlf
end if
if pcv_IDiscounts="1" then
StringBuilderObj.append "<td><b>Items Discounts</b></td>" & vbcrlf
end if
if pcv_EventName="1" then
StringBuilderObj.append "<td><b>Event Name</b></td>" & vbcrlf
end if
if pcv_GWOption="1" then
StringBuilderObj.append "<td><b>Gift Wrapping</b></td>" & vbcrlf
end if
if pcv_GWPrice="1" then
StringBuilderObj.append "<td><b>Gift Wrapping Price</b></td>" & vbcrlf
end if
if pcv_PackageID="1" then
StringBuilderObj.append "<td><b>Package ID</b></td>" & vbcrlf
end if
if pcv_PCost="1" then
StringBuilderObj.append "<td><b>Product Cost</b></td>" & vbcrlf
end if
if pcv_Margin="1" then
StringBuilderObj.append "<td><b>Margins</b></td>" & vbcrlf
end if
if pcv_TotalPrice="1" then
StringBuilderObj.append "<td><b>Total Price</b></td>" & vbcrlf
end if

StringBuilderObj.append "</tr>" & vbcrlf

do while not rsTemp.eof
				
	count=count + 1
	tmp_IDOrder=Clng(scpre+rstemp("IDOrder"))
	tmp_PrdSKU=rstemp("sku")
	tmp_PrdName=ClearHTMLTags2(rstemp("description"),0)
	tmp_UnitPrice=rstemp("unitPrice")
	tmp_quantity=rstemp("quantity")
	tmp_WPrice=rstemp("btoBPrice")
	tmp_idconf=rstemp("idconfigSession")
	tmp_QDisc=rstemp("QDiscounts")
	tmp_IDisc=rstemp("ItemsDiscounts")
	tmp_IDPackage=rstemp("pcPackageInfo_ID")
	tmp_SelectOpts=rstemp("pcPrdOrd_SelectedOptions")
	tmp_OptPrices=rstemp("pcPrdOrd_OptionsPriceArray")
	tmp_OptArr=rstemp("pcPrdOrd_OptionsArray")
	tmp_IDEP=rstemp("pcPO_EPID")
	tmp_IDGWOpt=rstemp("pcPO_GWOpt")
	tmp_GWPrice=rstemp("pcPO_GWPrice")
	tmp_PrdCost=rstemp("Cost")
	
	if IsNull(tmp_IDPackage) or tmp_IDPackage="" then
		tmp_IDPackage="0"
	end if
	if tmp_IDPackage="0" then
		tmp_IDPackage="N/A"
	end if
	
	if IsNull(tmp_GWPrice) or tmp_GWPrice="" then
		tmp_GWPrice="0"
	end if
	if tmp_GWPrice="0" then
		tmp_GWPrice="N/A"
	else
		tmp_GWPrice=scCurSign & money(tmp_GWPrice)
	end if
	
	if IsNull(tmp_OptArr) or tmp_OptArr="" then
		tmp_OptList="N/A"
	else
		tmp_OptList=replace(tmp_OptArr,"|","<br>")
	end if
	
	if IsNull(tmp_idconf) or tmp_idconf="" then
		tmp_idconf="0"
	end if
	if tmp_idconf="0" then
		tmp_BTOConf="N/A"
	else
		tmp_BTOConf=getBTOConf(tmp_idconf)
	end if
	
	if IsNull(tmp_IDEP) or tmp_IDEP="" then
		tmp_IDEP="0"
	end if
	if tmp_IDEP="0" then
		tmp_EventName="N/A"
	else
		tmp_EventName=getEventName(tmp_IDEP)
	end if
	
	if IsNull(tmp_IDGWOpt) or tmp_IDGWOpt="" then
		tmp_IDGWOpt="0"
	end if
	if tmp_IDGWOpt="0" then
		tmp_GWOptName="N/A"
	else
		tmp_GWOptName=getGWOptName(tmp_IDGWOpt)
	end if
	
	if IsNull(tmp_QDisc) or tmp_QDisc="" then
		tmp_QDisc="0"
	end if
	if tmp_QDisc="0" then
		tmp_QDisc="N/A"
	else
		tmp_QDisc=scCurSign & money(tmp_QDisc)
	end if
	
	if IsNull(tmp_IDisc) or tmp_IDisc="" then
		tmp_IDisc="0"
	end if
	if tmp_IDisc="0" then
		tmp_IDisc="N/A"
	else
		tmp_IDisc=scCurSign & money(tmp_IDisc)
	end if
	
	if IsNull(tmp_PrdCost) or tmp_PrdCost="" then
		tmp_PrdCost="0"
	end if
	if tmp_PrdCost="0" then
		tmp_PrdCost="N/A"
	else
		tmp_PrdCost=scCurSign & money(tmp_PrdCost)
		tmp_PrdCost1=tmp_PrdCost
	end if
	
	if IsNull(tmp_WPrice) or tmp_WPrice="" then
		tmp_WPrice="0"
	end if
	if tmp_WPrice="0" then
		tmp_WPrice="N/A"
	else
		tmp_WPrice=scCurSign & money(tmp_WPrice)
	end if
	tmp_TotalPrice=scCurSign & money(tmp_UnitPrice*tmp_quantity)
	tmp_UnitPrice=scCurSign & money(tmp_UnitPrice)
	
	if tmp_PrdCost<>"N/A" then
		tmp_Margin=scCurSign & money((tmp_UnitPrice*tmp_quantity)-(tmp_PrdCost1*tmp_quantity))
	else
		tmp_Margin="N/A"
	end if
	
	StringBuilderObj.append "<tr>" & vbcrlf
	if pcv_OrderID="1" then
	StringBuilderObj.append "<td>" & tmp_IDOrder & "</td>" & vbcrlf
	end if
	if pcv_PrdSKU="1" then
	StringBuilderObj.append "<td>" & tmp_PrdSKU & "</td>" & vbcrlf
	end if
	if pcv_PrdName="1" then
	StringBuilderObj.append "<td>" & tmp_PrdName & "</td>" & vbcrlf
	end if
	if pcv_UnitPrice="1" then
	StringBuilderObj.append "<td>" & tmp_UnitPrice & "</td>" & vbcrlf
	end if
	if pcv_Units="1" then
	StringBuilderObj.append "<td>" & tmp_quantity & "</td>" & vbcrlf
	end if
	if pcv_WholesalePrice="1" then
	StringBuilderObj.append "<td>" & tmp_WPrice & "</td>" & vbcrlf
	end if
	if pcv_BTOConf="1" then
	StringBuilderObj.append "<td>" & tmp_BTOConf & "</td>" & vbcrlf
	end if
	if pcv_POptions="1" then
	StringBuilderObj.append "<td>" & tmp_OptList & "</td>" & vbcrlf
	end if
	if pcv_QDiscounts="1" then
	StringBuilderObj.append "<td>" & tmp_QDisc & "</td>" & vbcrlf
	end if
	if pcv_IDiscounts="1" then
	StringBuilderObj.append "<td>" & tmp_IDisc & "</td>" & vbcrlf
	end if
	if pcv_EventName="1" then
	StringBuilderObj.append "<td>" & tmp_EventName & "</td>" & vbcrlf
	end if
	if pcv_GWOption="1" then
	StringBuilderObj.append "<td>" & tmp_GWOptName & "</td>" & vbcrlf
	end if
	if pcv_GWPrice="1" then
	StringBuilderObj.append "<td>" & tmp_GWPrice & "</td>" & vbcrlf
	end if
	if pcv_PackageID="1" then
	StringBuilderObj.append "<td>" & tmp_IDPackage & "</td>" & vbcrlf
	end if
	if pcv_PCost="1" then
	StringBuilderObj.append "<td>" & tmp_PrdCost & "</td>" & vbcrlf
	end if
	if pcv_Margin="1" then
	StringBuilderObj.append "<td>" & tmp_Margin & "</td>" & vbcrlf
	end if
	if pcv_TotalPrice="1" then
	StringBuilderObj.append "<td>" & tmp_TotalPrice & "</td>" & vbcrlf
	end if
	StringBuilderObj.append "</tr>" & vbcrlf
	rsTemp.MoveNext
loop
StringBuilderObj.append "</table>" & vbcrlf
set rstemp=nothing
call closedb()
Response.ContentType = "application/vnd.ms-excel"
%>
<%=StringBuilderObj.toString()%>
<%
set StringBuilderObj = nothing
END IF
%>