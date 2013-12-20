<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Global Product Changes" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<%pageTitle="Global Product Changes" %>
<% 
if request("nav")="1" then
	section="services"
else
	section="products"
end if%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<% 
'on error resume next

Dim query, conntemp, rstemp, rstemp4, rstemp5
Dim TmpCatList
TmpCatList=""

Sub pcv_GetSubCats(tmpIDParent)
Dim query,rs,i,intCount,pcArr

	query="SELECT idcategory FROM categories WHERE idParentCategory=" & tmpIDParent & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
		set rs=nothing
		For i=0 to intCount
			TmpCatList=TmpCatList & "," & pcArr(0,i)
			call pcv_GetSubCats(pcArr(0,i))
		Next
	end if
	set rs=nothing
End sub

call openDB()

if request("action")="update" then
	CP=request("CP1")
	
	'Start SDBA
	tmpquery=""
	if CP="9" then
		if request("pcIDDropshipper")<>"0" then
			tmpquery=",pcDropShippersSuppliers "
		end if
	end if
	'End SDBA
	
   if CP="2" then
	if request("idcategory")<>"0" then
	query="SELECT DISTINCT products.idProduct,products.price,products.listprice,products.btoBprice FROM products, categories_products" & tmpquery & " WHERE products.removed=0 "
	else
	query="SELECT DISTINCT products.idProduct,products.price,products.listprice,products.btoBprice FROM products" & tmpquery & " WHERE products.removed=0 "
	end if	
	else
	if CP="1" then
	query="SELECT DISTINCT products.idProduct,products.price,products.listprice,products.btoBprice FROM products" & tmpquery & " WHERE products.removed=0 "
	else
	query="SELECT DISTINCT products.idProduct,products.price,products.listprice,products.btoBprice FROM products" & tmpquery & " WHERE products.removed=0 "
	end if
	end if
	
	TempStr1=""
	
	if CP="1" then
		TempStr1="All products in your store."	
	end if
	
	if CP="2" then
	if request("idcategory")<>"0" then
		idcategory=request("idcategory")
		if request("incSubCats")<>"1" then
			query=query & " AND categories_products.idCategory=" &idcategory & " AND products.idProduct=categories_products.idProduct"
		else
			TmpCatList=request("idcategory")
			call pcv_GetSubCats(TmpCatList)
			query=query & " AND categories_products.idCategory IN (" & TmpCatList & ") AND products.idProduct=categories_products.idProduct"
		end if
		query1="SELECT categoryDesc FROM categories WHERE idcategory=" & idcategory
		set rs=connTemp.execute(query1)
		if not rs.eof then
			pcv_CatDesc=rs("categoryDesc")
		end if
		set rs=nothing	
		TempStr1="All products in the following category: " & pcv_CatDesc
		if request("incSubCats")="1" then
			TempStr1=TempStr1 & "<br>Include sub-categories"
		end if
	else
		TempStr1="All products in your store."	
	end if
	end if
	
	if CP="3" then
		query=query & "AND products.sku like '%"&replace(request("sku"),"'","''")&"%'"
		TempStr1="All products whose part number (SKU) contains: " & request("sku")
	end if
	
	if CP="4" then
		query=query & "AND ((products.description like '%"&replace(request("nd"),"'","''")&"%') OR (products.details like '%" &replace(request("nd"),"'","''")& "%'))"
		TempStr1="All products whose name or description contains: " & request("nd")
	end if	
	
	if CP="5" then
		if request("hpType")="2" then
		myStr11=" list price "
		query=query & "AND products.listprice>="&replacecomma(request("hprice"))
		else
		myStr11=" online price "
		query=query & "AND products.price>="&replacecomma(request("hprice"))
		end if
		TempStr1="All products whose" & myStr11 & "is higher than: " & scCurSign & request("hprice")
	end if	

	if CP="6" then
		if request("lpType")="2" then
		myStr11=" list price "
		query=query & "AND products.listprice<="&replacecomma(request("lprice"))
		else
		myStr11=" online price "
		query=query & "AND products.price<="&replacecomma(request("lprice"))
		end if
		TempStr1="All products whose" & myStr11 & "is lower than: " & scCurSign & request("lprice")
	end if
	
	if CP="7" then
		if request("IDBrand")>"0" then
			query=query & "AND products.IDBrand="&request("IDBrand")
		end if
		TempStr1="All products whose product brand is: "
		if request("IDBrand")="0" then
		TempStr1=TempStr1 & "All brands"
		else
		query1="Select * from Brands where IDBrand=" & request("IDBrand")
		set rstemp1=connTemp.execute(query1)
		TempStr1=TempStr1 & rstemp1("BrandName")
		set rstemp1=nothing
		end if
	end if
	
	'Start SDBA
	if CP="8" then
		if request("pcIDSupplier")>"0" then
			query=query & "AND products.pcSupplier_ID="&request("pcIDSupplier")
		end if
		TempStr1="All products whose supplier is: "
		if request("pcIDSupplier")="0" then
		TempStr1=TempStr1 & "All suppliers"
		else
		query1="Select pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers where pcSupplier_ID=" & request("pcIDSupplier")
		set rstemp1=connTemp.execute(query1)
		TempStr1=TempStr1 & rstemp1("pcSupplier_Company")
		if rstemp1("pcSupplier_FirstName") & rstemp1("pcSupplier_LastName")<>"" then
			TempStr1=TempStr1 & " (" & rstemp1("pcSupplier_FirstName") & " " & rstemp1("pcSupplier_LastName") & ")"
		end if
		set rstemp1=nothing
		end if
	end if
	if CP="9" then
		if request("pcIDDropshipper")<>"0" then
			pcArr=split(request("pcIDDropshipper"),"_")
			query=query & "AND products.pcDropShipper_ID=" & pcArr(0)
			query=query & "AND pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & pcArr(1) & " AND products.pcDropShipper_ID=" & pcArr(0)
		end if
		TempStr1="All products whose drop-shipper is: "
		if request("pcIDSupplier")="0" then
		TempStr1=TempStr1 & "All drop-shippers"
		else
		if pcArr(1)="1" then
		query1="Select pcSupplier_Company As A,pcSupplier_FirstName As B,pcSupplier_LastName As C from pcSuppliers where pcSupplier_ID=" & pcArr(0) & " AND pcSupplier_IsDropShipper=1;"
		else
		query1="Select pcDropShipper_Company As A,pcDropShipper_FirstName As B,pcDropShipper_LastName As C from pcDropShippers where pcDropShipper_ID=" & pcArr(0) & ";"
		end if
		set rstemp1=connTemp.execute(query1)
		TempStr1=TempStr1 & rstemp1("A")
		if rstemp1("B") & rstemp1("C")<>"" then
			TempStr1=TempStr1 & " (" & rstemp1("B") & " " & rstemp1("C") & ")"
		end if
		set rstemp1=nothing
		end if
	end if
	'End SDBA
		
	if CP="10" then
		if request("pcv_instock")>"0" then
			query=query & "AND products.stock>0"
			TempStr1="All products are in stock"
		else
			query=query & "AND products.stock<=0"
			TempStr1="All products are out of stock"
		end if
	end if
	
	if CP="11" then
		TempStr1="All products are "
		pcv_tmp1=0
		' Standard = yes | BTO = no | BTO Item = no -> STANDARD ONLY
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")="") then
			query=query & " AND products.serviceSpec=0 AND products.configOnly=0"
			pcv_tmp1=1
		end if
		' Standard = no | BTO = yes | BTO Item = no -> BTO ONLY
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")="") then
			query=query & " AND products.serviceSpec<>0"
			pcv_tmp1=2
		end if
		' Standard = no | BTO = no | BTO Item = yes -> BTO ITEM ONLY
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")<>"") then
			query=query & " AND products.configOnly<>0"
			pcv_tmp1=3
		end if
		' Standard = yes | BTO = yes | BTO Item = no -> STANDARD and BTO
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")="") then
			query=query & " AND products.configOnly=0"
			pcv_tmp1=4
		end if
		' Standard = yes | BTO = no | BTO Item = yes -> STANDARD and BTO ITEM
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")<>"") then
			query=query & " AND products.serviceSpec=0"
			pcv_tmp1=5
		end if
		' Standard = no | BTO = yes | BTO Item = yes -> BTO and BTO ITEM
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")<>"") then
			query=query & " AND (products.serviceSpec<>0 OR products.configOnly<>0)"
			pcv_tmp1=6
		end if
		' Standard = no | BTO = no | BTO Item = no
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")="") then
			pcv_tmp1=7
		end if
		' Standard = yes | BTO = yes | BTO Item = yes
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")<>"") then
			pcv_tmp1=8
		end if
		if scBTO<>1 then
			TempStr1=TempStr1 & "standard products "
		else
			Select Case pcv_tmp1
				Case 1: TempStr1=TempStr1 & "standard products "
				Case 2: TempStr1=TempStr1 & "BTO products "
				Case 3: TempStr1=TempStr1 & "BTO items "
				Case 4: TempStr1=TempStr1 & "standard products or BTO products "
				Case 5: TempStr1=TempStr1 & "standard products or BTO items "
				Case 6: TempStr1=TempStr1 & "BTO products or BTO items "
				Case 8: TempStr1=TempStr1 & "standard products or BTO products or BTO items "
			End Select
		end if
		if (request("pcv_prdtype4")<>"") then
			if (request("pcv_prdtype5")<>"") then
				query=query & " AND ((products.Downloadable=1)"
			else
				query=query & " AND products.Downloadable=1"
			end if
			if pcv_tmp1=7 then
				TempStr1=TempStr1 & "downloadable products "
			else
				TempStr1=TempStr1 & "and they are downloadable products "
			end if
			pcv_tmp1=9
		end if
		if (request("pcv_prdtype5")<>"") then
			if pcv_tmp1=9 then
				query=query & " OR (products.pcprod_GC=1))"
			else
				query=query & " AND products.pcprod_GC=1"
			end if
			if pcv_tmp1=7 then
				TempStr1=TempStr1 & "gift certificates "
			else
				if pcv_tmp1=9 then
					TempStr1=TempStr1 & "or gift certificates "
				else
					TempStr1=TempStr1 & "and they are gift certificates "
				end if
			end if
		end if		
	end if
	
	if CP="13" then
		query=query & "AND products.idproduct NOT IN (SELECT DISTINCT categories_products.idproduct FROM categories_products)"
		TempStr1="All Products that are not assigned to any category"
	end if
	
	If scDB="SQL" Then
		'Prevent Products in a Current Sales from being modified
		smSalesQuery = " AND products.idProduct NOT IN (select  pcSales_BackUp.idProduct FROM pcSales_BackUp WHERE pcSales_BackUp.idProduct = products.IdProduct)"
		query = query & smSalesQuery
	End If
	
	set rstemp=connTemp.execute(query)
	
	count=0

	DO WHILE not rstemp.eof
		count=count+1
		rstemp.MoveNext
	LOOP
	
	UP=request("UP1")
	TempStr2=""
	
	if UP="1" then
		priceSelect=request("priceSelect")
		Select Case priceSelect
		Case "1": TempStr2="Change Online Price"
		Case "2": TempStr2="Change List Price"
		Case "3": TempStr2="Change Wholesale Price"
		Case Else:
			if instr(priceSelect,"CC_")>0 then
				tmp_Arr=split(priceSelect,"CC_")
				tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & tmp_Arr(1)
				set rstemp4=connTemp.execute(tmpquery)
				TempStr2="Change Price in Pricing Category: '" & rstemp4("pcCC_Name") & "'"
				set rstemp4=nothing
			end if
		End Select
		
		Select Case request("cpriceType")
		Case "1": TempStr2=TempStr2 & " by: " & request("cprice") & "%"
		Case "2": TempStr2=TempStr2 & " by: " & scCurSign & request("cprice")
		End Select
		if request("cpriceRound")="1" then
			TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest integer."
		else
			if request("cpriceRound")="2" then
				TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest hundredth."
			end if
		end if
	end if
	
	if UP="2" then
		priceSelect1=request("priceSelect1")
		priceSelect2=request("priceSelect2")
		Select Case priceSelect1
		Case "1": TempStr2="Make the online price " & request("wprice") & "% "
		Case "2": TempStr2="Make the list price " & request("wprice") & "% "
		Case "3": TempStr2="Make the wholesale price " & request("wprice") & "% "
		Case Else:
			if instr(priceSelect1,"CC_")>0 then
				tmp_Arr=split(priceSelect1,"CC_")
				tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & tmp_Arr(1)
				set rstemp4=connTemp.execute(tmpquery)
				TempStr2="Change Price in Pricing Category: '" & rstemp4("pcCC_Name") & "' " & request("wprice") & "% "
				set rstemp4=nothing
			end if
		End Select
		Select Case priceSelect2
		Case "1": TempStr2=TempStr2 & "of the online price."
		Case "2": TempStr2=TempStr2 & "of the list price."
		Case "3": TempStr2=TempStr2 & "of the wholesale price."
		'Start SDBA
		Case "4": TempStr2=TempStr2 & "of the product cost."
		'End SDBA
		Case Else:
			if instr(priceSelect2,"CC_")>0 then
				tmp_Arr=split(priceSelect2,"CC_")
				tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & tmp_Arr(1)
				set rstemp4=connTemp.execute(tmpquery)
				TempStr2=TempStr2 & "of the price in pricing category: '" & rstemp4("pcCC_Name") & "'"
				set rstemp4=nothing
			end if
		End Select
		if request("cpriceRound1")="1" then
			TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest integer."
		else
			if request("cpriceRound1")="2" then
				TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest hundredth."
			end if
		end if
	end if
	
	if UP="3" then
		COption=request("coption")
		Select Case COption
		Case "1":
			TempStr2="The 'Show Savings' option will be assigned to the selected products."
		Case "2":
			TempStr2="The 'Special' option will be assigned to the selected products."
		Case "3":
			TempStr2="The 'No Taxable' option will be assigned to the selected products."
		Case "4":
			TempStr2="The 'Not for Sale' option will be assigned to the selected products."
		Case "5":
			TempStr2="The 'Free or No Shipping' option will be assigned to the selected products."	
		Case "6":
			TempStr2="The the selected products will be made active."				
		Case "7":
			TempStr2="The 'Disregard Stock' option will be assigned to the selected products."			
		Case "8":
			TempStr2="The 'Display No Shipping Text' option will be assigned to the selected products."
		'Start SDBA
		Case "9":
			TempStr2="The 'Back-Ordering' option will be assigned to the selected products."
		Case "10":
			TempStr2="The 'Low Inventory Notification' option will be assigned to the selected products."
		Case "11":
			TempStr2="The 'Drop-Shipped' option will be assigned to the selected products."
		'End SDBA
		Case "12":
			TempStr2="The 'Oversized' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "Height: " & request("pcv_height") & " - Width: " & request("pcv_width") & " - Length: " & request("pcv_length") & " (inches)"
		Case "13":
			TempStr2="The 'Downloadable Product' option will be assigned to the selected products."
		Case "14":
			TempStr2="The 'Make download URL expire' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Downloadable Products only</i>"
		Case "15":
			TempStr2="The 'Deliver license with order confirmation' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Downloadable Products only</i>"
		Case "16":
			TempStr2="The 'Gift Certificate' option will be assigned to the selected products."
		Case "17":
			TempStr2="The 'Does not expire' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Gift Certificates only</i>"
		Case "18":
			TempStr2="The 'Expires on the Date' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Gift Certificates only</i>"
		Case "19":
			TempStr2="The 'Expires N days after purchase' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Gift Certificates only</i>"
		Case "20":
			TempStr2="The 'Electronic Only' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Gift Certificates only</i>"
		Case "21":
			TempStr2="The 'Use default generator' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Gift Certificates only</i>"
		Case "22":
			TempStr2="The 'Use custom generator' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Gift Certificates only</i>"
		Case "23":
			TempStr2="The 'Hide BTO Price' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "24":
			TempStr2="The 'Hide Default Configuration' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "25":
			TempStr2="The 'Disallow purchasing - Show Prices' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "26":
			TempStr2="The 'Disallow purchasing - Hide Prices' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "27":
			TempStr2="The 'Skip Product Details Page' option will be assigned to the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "28":
			TempStr2="The 'Featured Product' option will be assigned to the selected products."
		Case "29":
			TempStr2="The option 'Hide SKU on Product Details Page' will be assigned to the selected products."
		Case "30":
			TempStr2="The 'Image Magnifier' will be enabled for the selected products."
		End Select
	end if
	if UP="4" then
		ROption=request("roption")
		Select Case ROption
		Case "1":
			TempStr2="The 'Show Savings' option will be removed from the selected products."
		Case "2":
			TempStr2="The 'Special' option will be removed from the selected products."
		Case "3":
			TempStr2="The 'No Taxable' option will be removed from the selected products."
		Case "4":
			TempStr2="The 'Not for Sale' option will be removed from the selected products."
		Case "5":
			TempStr2="The 'Free or No Shipping' option will be removed from the selected products."	
		Case "6":
			TempStr2="The selected products will be made inactive."				
		Case "7":
			TempStr2="The 'Disregard Stock' option will be removed from the selected products."	
		Case "8":
			TempStr2="The 'Display No Shipping Text' option will be removed from the selected products."
		'Start SDBA
		Case "9":
			TempStr2="The 'Back-Ordering' option will be removed from the selected products."
		Case "10":
			TempStr2="The 'Low Inventory Notification' option will be removed from the selected products."
		Case "11":
			TempStr2="The 'Drop-Shipped' option will be removed from the selected products."
		'End SDBA
		Case "12":
			TempStr2="The 'Oversized' option will be removed from the selected products."
		Case "13":
			TempStr2="The 'Downloadable Product' option will be removed from the selected products."
		Case "14":
			TempStr2="The 'Make download URL expire' option will be removed from the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Downloadable Products only</i>"
		Case "15":
			TempStr2="The 'Deliver license with order confirmation' option will be removed from the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Downloadable Products only</i>"
		Case "16":
			TempStr2="The 'Gift Certificate' option will be removed from the selected products."
		Case "17":
			TempStr2="The 'Electronic Only' option will be removed from the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects Gift Certificates only</i>"
		Case "18":
			TempStr2="The 'Hide BTO Price' option will be removed from the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "19":
			TempStr2="The 'Hide Default Configuration' option will be removed from the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "20":
			TempStr2="The 'Disallow purchasing' options will be removed from the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "21":
			TempStr2="The 'Skip Product Details Page' option will be removed from the selected products.<br>"
			TempStr2=TempStr2 & "<i>* Note: This option affects BTO Products only</i>"
		Case "22":
			TempStr2="The 'Featured Product' option will be removed from the selected products."
		Case "23":
			TempStr2="The option 'Hyde SKU on Product Details Page' will be removed from the selected products."
		Case "24":
			TempStr2="The 'Image Magnifier' will be disabled for the selected products."
		End Select
	end if
	if UP="5" then
		PToption=request("ptoption")
		Select Case PTOption
		Case "1":
			TempStr2="The type of selected products will be changed to 'Standard Product'."
		Case "2":
			TempStr2="The type of selected products will be changed to 'BTO Product'."
		Case "3":
			TempStr2="The type of selected products will be changed to 'BTO Only Item'."
		End Select
	end if
	if UP="6" then
		weight=request("weight")
		weight_oz=request("weight_oz")
		weight_kg=request("weight_kg")
		weight_g=request("weight_g")
		weight_units=request("weight_units")				
		TempStr2="Set the weight to: "
		If scShipFromWeightUnit="KGS" then
			TempStr2=TempStr2 & weight_kg & " kg " & weight_g & " g "
			if weight_units>"0" then
				TempStr2=TempStr2 & "<br>Units to make 1 kg: " & weight_units 
			end if
		Else
			TempStr2=TempStr2 & weight & " lbs. " & weight_oz & " ozs. "
			if weight_units>"0" then
				TempStr2=TempStr2 & "<br>Units to make 1 lb: " & weight_units 
			end if
		End if
	end if
	'Start SDBA	
	if UP="7" then
		numoptions=request("numoptions")
		numvalue=request("numvalue")
		Select Case numoptions
			Case "1":
				pcv_tmp1="Stock Level"
				pcv_tmp2=" Units"
			Case "2":
				pcv_tmp1="Cost"
				pcv_tmp2=" " & scCurSign
			Case "3":
				pcv_tmp1="Reorder Level"
				pcv_tmp2=" Units"
			Case "4":
				pcv_tmp1="Ship within N Days"
				pcv_tmp2=" Days"
			Case "5":
				pcv_tmp1=RewardsLabel
				pcv_tmp2=" points"
			Case "6":
				pcv_tmp1="URL will expire after N days"
				pcv_tmp2=" Days<br><i>* Note: It is the options for Downloadable Products only.</i>"
			Case "7":
				pcv_tmp1="Expires after N Days"
				pcv_tmp2=" Days<br><i>* Note: It is the options for Gift Certificates only.</i>"
		End Select
		TempStr2="Set the '" & pcv_tmp1 & "' to: " & numvalue & pcv_tmp2
	end if
	'End SDBA
	if UP="8" then
		minimumqty=request("minimumqty")
		qtyvalidate=request("qtyvalidate")
		TempStr2="Set the ""Minimum quantity customers can buy"" to: " & minimumqty & " Units"
		if qtyvalidate="1" then
			TempStr2=TempStr2 & "<br>Force purchase of multiples of minimum"
		end if
	end if
	'Start SDBA	
	if UP="9" then
		query1="Select pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers where pcSupplier_ID=" & request("pcToIDSupplier")
		set rstemp1=connTemp.execute(query1)
		TempStr2="Move filtered products to the supplier: " & rstemp1("pcSupplier_Company")
		if rstemp1("pcSupplier_FirstName") & rstemp1("pcSupplier_LastName")<>"" then
			TempStr2=TempStr2 & " (" & rstemp1("pcSupplier_FirstName") & " " & rstemp1("pcSupplier_LastName") & ")"
		end if
		set rstemp1=nothing
	end if
	if UP="10" then
		pcArr=split(request("pcToIDDropshipper"),"_")
		if pcArr(1)="1" then
			query1="Select pcSupplier_Company As A,pcSupplier_FirstName As B,pcSupplier_LastName As C from pcSuppliers where pcSupplier_ID=" & pcArr(0) & " AND pcSupplier_IsDropShipper=1;"
		else
			query1="Select pcDropShipper_Company As A,pcDropShipper_FirstName As B,pcDropShipper_LastName As C from pcDropShippers where pcDropShipper_ID=" & pcArr(0) & ";"
		end if
		set rstemp1=connTemp.execute(query1)
		TempStr2="Move filtered products to the drop-shipper: " & rstemp1("A")
		if rstemp1("B") & rstemp1("C")<>"" then
			TempStr2=TempStr2 & " (" & rstemp1("B") & " " & rstemp1("C") & ")"
		end if
		set rstemp1=nothing
	end if
	'End SDBA
	if UP="11" then
		stroptions=request("stroptions")
		strvalue=request("strvalue")
		pcv_tmp2=""
		Select Case stroptions
			Case "1":
				pcv_tmp1="Downloadable File Location"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "2":
				pcv_tmp1="Local license generator"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "3":
				pcv_tmp1="Remote license generator"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "4":
				pcv_tmp1="License Field (1)"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "5":
				pcv_tmp1="License Field (2)"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "6":
				pcv_tmp1="License Field (3)"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "7":
				pcv_tmp1="License Field (4)"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "8":
				pcv_tmp1="License Field (5)"
				pcv_tmp2="<br><i>* Note: This option affects Downloadable Products only.</i>"
			Case "9":
				pcv_tmp1="Custom generator file name"
				pcv_tmp2="<br><i>* Note: This option affects Gift Certificates only.</i>"
			Case "10":
				pcv_tmp1="Expiration Date"
				pcv_tmp2="<br><i>* Note: This option affects Gift Certificates only.</i>"
		End Select
		TempStr2="Set the '" & pcv_tmp1 & "' to: '" & strvalue & "'" & pcv_tmp2
	end if
	if UP="12" then
		query1="Select BrandName from Brands where idbrand=" & request("pcToIDBrand")
		set rstemp1=connTemp.execute(query1)
		TempStr2="Move filtered products to the brand: " & rstemp1("BrandName")
		set rstemp1=nothing
	end if
	if UP="13" then
		Select Case request("pcv_displayLayout")
			Case "": TempStr2="Set Page layout is 'Use Default'"
			Case "c" : TempStr2="Set Page layout is 'Two Columns-Image on Right'"
			Case "l": TempStr2="Set Page layout is 'Two Columns-Image on Left'"
			Case "o": TempStr2="Set Page layout is 'One-Column'"
		End Select
	end if
	if UP="14" then
		query1="SELECT categoryDesc FROM categories WHERE idcategory=" & request("ToIDCategory")
		set rstemp1=connTemp.execute(query1)
		TempStr2="Assign selected products to the following category: " & rstemp1("categoryDesc")
		set rstemp1=nothing
	end if
	
	if UP="15" then
		TempStr2="Google Shopping Settings: Set the '"
		Select Case request("goSett")
			Case "1": TempStr2=TempStr2 & "Google Product Category"
			Case "2": TempStr2=TempStr2 & "Google Shopping - Gender"
	        Case "3": TempStr2=TempStr2 & "Google Shopping - Age"
			Case "4": TempStr2=TempStr2 & "Google Shopping - Color"
			Case "5": TempStr2=TempStr2 & "Google Shopping - Size"
			Case "6": TempStr2=TempStr2 & "Google Shopping - Pattern"
			Case "7": TempStr2=TempStr2 & "Google Shopping - Material"
		End Select
		TempStr2=TempStr2 & "' to the value: '" & request("goValue") & "'"
	end if
%>
<form name="UpdateForm" action="globalChangesconfirm.asp?action=update" method="post" class="pcForms">
<table class="pcCPcontent">
		<input type="hidden" name="nav" value="<%=request("nav")%>">
		<input type="hidden" name="CP1" value="<%=request("CP1")%>">
		<input type="hidden" name="UP1" value="<%=request("UP1")%>">
		<input type="hidden" name="idcategory" value="<%=request("idcategory")%>">
		<input type="hidden" name="sku" value="<%=request("sku")%>">
		<input type="hidden" name="nd" value="<%=request("nd")%>">
		<input type="hidden" name="hpType" value="<%=request("hpType")%>">
		<input type="hidden" name="hprice" value="<%=request("hprice")%>">
		<input type="hidden" name="lpType" value="<%=request("lpType")%>">
		<input type="hidden" name="lprice" value="<%=request("lprice")%>">
		<input type="hidden" name="IDBrand" value="<%=request("IDBrand")%>">
		<input type="hidden" name="pcIDSupplier" value="<%=request("pcIDSupplier")%>">
		<input type="hidden" name="pcIDDropshipper" value="<%=request("pcIDDropshipper")%>">
		<input type="hidden" name="pcv_instock" value="<%=request("pcv_instock")%>">
		<input type="hidden" name="pcv_prdtype1" value="<%=request("pcv_prdtype1")%>">
		<input type="hidden" name="pcv_prdtype2" value="<%=request("pcv_prdtype2")%>">
		<input type="hidden" name="pcv_prdtype3" value="<%=request("pcv_prdtype3")%>">
		<input type="hidden" name="pcv_prdtype4" value="<%=request("pcv_prdtype4")%>">
		<input type="hidden" name="pcv_prdtype5" value="<%=request("pcv_prdtype5")%>">
		<input type="hidden" name="IDCustCAT" value="<%=request("IDCustCAT")%>">

		<input type="hidden" name="cprice" value="<%=request("cprice")%>">
		<input type="hidden" name="cpricetype" value="<%=request("cpricetype")%>">
		<input type="hidden" name="cpriceRound" value="<%=request("cpriceRound")%>">
		<input type="hidden" name="cpriceRound1" value="<%=request("cpriceRound1")%>">
		<input type="hidden" name="priceSelect" value="<%=request("priceSelect")%>">
		<input type="hidden" name="priceSelect1" value="<%=request("priceSelect1")%>">
		<input type="hidden" name="priceSelect2" value="<%=request("priceSelect2")%>">
		<input type="hidden" name="wprice" value="<%=request("wprice")%>">
		<input type="hidden" name="coption" value="<%=request("coption")%>">
		<input type="hidden" name="roption" value="<%=request("roption")%>">
		<input type="hidden" name="ptoption" value="<%=request("ptoption")%>">
		<input type="hidden" name="nfsmsg" value="<%=request("nfsmsg")%>">
		<input type="hidden" name="pcv_height" value="<%=request("pcv_height")%>">
		<input type="hidden" name="pcv_width" value="<%=request("pcv_width")%>">
		<input type="hidden" name="pcv_length" value="<%=request("pcv_length")%>">
		<input type="hidden" name="weight" value="<%=request("weight")%>">
		<input type="hidden" name="weight_oz" value="<%=request("weight_oz")%>">
		<input type="hidden" name="weight_kg" value="<%=request("weight_kg")%>">
		<input type="hidden" name="weight_g" value="<%=request("weight_g")%>">
		<input type="hidden" name="weight_units" value="<%=request("weight_units")%>">
		<input type="hidden" name="numoptions" value="<%=request("numoptions")%>">
		<input type="hidden" name="numvalue" value="<%=request("numvalue")%>">
		<input type="hidden" name="minimumqty" value="<%=request("minimumqty")%>">
		<input type="hidden" name="qtyvalidate" value="<%=request("qtyvalidate")%>">
		<input type="hidden" name="calSelect1" value="<%=request("calSelect1")%>">
		<input type="hidden" name="pcToIDSupplier" value="<%=request("pcToIDSupplier")%>">
		<input type="hidden" name="pcToIDDropshipper" value="<%=request("pcToIDDropshipper")%>">
		<input type="hidden" name="stroptions" value="<%=request("stroptions")%>">
		<input type="hidden" name="strvalue" value="<%=request("strvalue")%>">
		<input type="hidden" name="pcToIDBrand" value="<%=request("pcToIDBrand")%>">
		<input type="hidden" name="pcv_displayLayout" value="<%=request("pcv_displayLayout")%>">
		<input type="hidden" name="ToIDCategory" value="<%=request("ToIDCategory")%>">
		
		<input type="hidden" name="goSett" value="<%=request("goSett")%>">
		<%goValue=request("goValue")
		if goValue<>"" then
			goValue=replace(goValue,"""","&quot;")
			goValue=replace(goValue,"&","&amp;")
			goValue=replace(goValue,">","&gt;")
			goValue=replace(goValue,"<","&lt;")
		end if%>
		<input type="hidden" name="goValue" value="<%=goValue%>">
		
		<input type="hidden" name="incSubCats" value="<%=request("incSubCats")%>">
		<input type="hidden" name="TmpCatList" value="<%=TmpCatList%>">
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<th>Review and Confirm Your Global Changes</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
				<p>You are about to apply the following change to your product database:</p>
				<ol>
				<li>The change will affect <span class="pcCPnotes"><%=count%> products</span><br /><br /></li>
				<li>The following change will apply:<br /><br /><span class="pcCPnotes"><%=TempStr2%></span><br /><br /></li>
				<li>The change will be applied to:<br /><br /><span class="pcCPnotes"><%=TempStr1%></span></li>
				</ol>
			</td>
		</tr>
		<tr>
			<td align="center"><p><strong>Would you like to proceed?</strong></p></td>
		</tr>
		<tr>
			<td class="pcCPspacer"><hr></td>
		</tr>
		<tr>
			<td align="center">
				<input type="submit" name="submit" value="Yes, apply these changes" onClick="pcf_Open_globalChanges();" class="submit2">&nbsp;
				<input type="button" name="back" value="No, return to the previous page" onClick="javascript:history.back()">
				<%
                '// Loading Window
                '	>> Call Method with OpenHS();
                response.Write(pcf_ModalWindow("The changes that you requested are being applied. Please wait...", "globalChanges", 300))
                %>
			</td>
		</tr>
</table>                    
</form>	
<%ELSE%>
<script language="JavaScript">
<!--
	
function isDigit(s)
{
var test=""+s;
if(test=="+"||test=="-"||test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}
	
function isDigit1(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit1(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit1(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}	

function Form1_Validator(theForm)
{

	if (theForm.CP1.value == "0")
 	{
		    alert("Please select one of the available product selection options before proceeding.");
		    return (false);
	}
	
 	if (theForm.CP1.value == "3")
 	{
			if (theForm.sku.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.sku.focus();
		    return (false);
		    }
	}
	if (theForm.CP1.value == "4")
  	{
			if (theForm.nd.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.nd.focus();
		    return (false);
		    }
	}
	if (theForm.CP1.value == "5")
  	{
		  	if (theForm.hpType.value == "")
			{
		    alert("Please select a price type.");
		    theForm.hpType.focus();
		    return (false);
		    }
			if (theForm.hprice.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.hprice.focus();
		    return (false);
		    }
			if (allDigit(theForm.hprice.value) == false)
			{
		    alert("Please enter a right value for this field.");
		    theForm.hprice.focus();
		    return (false);
		    }		    
	}
	if (theForm.CP1.value == "6")
	{
			if (theForm.lpType.value == "")
			{
		    alert("Please select a price type.");
		    theForm.lpType.focus();
		    return (false);
		    }
			if (theForm.lprice.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.lprice.focus();
		    return (false);
		    }
			if (allDigit(theForm.lprice.value) == false)
			{
		    alert("Please enter a right value for this field.");
		    theForm.lprice.focus();
		    return (false);
		    }		    
	}
	
	<%query="Select BrandName from Brands order by BrandName asc"
	set rstemp4=connTemp.execute(query)
	if not rstemp4.eof then%>
	if (theForm.CP1.value == "7")
  	{
			if (theForm.IDBrand.value == "")
			{
		    alert("Please select a brand.");
		    theForm.IDBrand.focus();
		    return (false);
		    }
	}
	<%end if
	set rstemp4=nothing%>
	<%'Start SDBA
	'Get Suppliers List
	query="Select pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers order by pcSupplier_Company asc"
	set rs=connTemp.execute(query)
	if not rs.eof then%>
	if (theForm.CP1.value == "8")
  	{
			if (theForm.pcIDSupplier.value == "")
			{
		    alert("Please select a supplier.");
		    theForm.pcIDSupplier.focus();
		    return (false);
		    }
	}
	<%end if
	set rs=nothing
	'End SDBA%>
	
	<%'Start SDBA
	'Get Drop-Shippers List
	query="SELECT pcDropShipper_ID,pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName,0 FROM pcDropShippers UNION (SELECT pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName,1 FROM pcSuppliers WHERE pcSupplier_IsDropShipper=1) ORDER BY pcDropShipper_Company ASC"
	set rs=connTemp.execute(query)
	if not rs.eof then%>
	if (theForm.CP1.value == "9")
  	{
			if (theForm.pcIDDropshipper.value == "")
			{
		    alert("Please select a supplier.");
		    theForm.pcIDDropshipper.focus();
		    return (false);
		    }
	}
	<%end if
	set rs=nothing
	'End SDBA%>
	
	if (theForm.CP1.value == "11")
  	{
			if ((theForm.pcv_prdtype1.checked == false)<%if scBTO=1 then%> && (theForm.pcv_prdtype2.checked == false) && (theForm.pcv_prdtype3.checked == false)<%end if%> && (theForm.pcv_prdtype4.checked == false) && (theForm.pcv_prdtype5.checked == false))
			{
		    alert("Please select a product type.");
		    theForm.pcv_prdtype1.focus();
		    return (false);
		    }
	}
	
	if (theForm.UP1.value == "0")
 	{
		    alert("Please select one of the available product change options before proceeding.");
		    return (false);
	}
	
	
	if (theForm.UP1.value == "1")
  	{
			if (theForm.cprice.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.cprice.focus();
		    return (false);
		    }
			if (allDigit(theForm.cprice.value) == false)
			{
		    alert("Please enter a right value for this field.");
		    theForm.cprice.focus();
		    return (false);
		   }		    
	}
	if (theForm.UP1.value == "2")
  	{
			if (theForm.wprice.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.wprice.focus();
		    return (false);
		    }
			if (allDigit(theForm.wprice.value) == false)
			{
		    alert("Please enter a right value for this field.");
		    theForm.wprice.focus();
		    return (false);
		    }		    
	}
	if (theForm.UP1.value == "3")
  	{
			if (theForm.coption.value == "")
			{
			    alert("Please select one option.");
			    theForm.coption.focus();
			    return (false);
		    }
		    if (theForm.coption.value == "4")
			{
				if (theForm.nfsmsg.value == "")
			    {
				    alert("Please enter a value for this field.");
				    theForm.nfsmsg.focus();
				    return (false);
			    }
			    var tmpStr=theForm.nfsmsg.value;
			    if (tmpStr.length > 250)
			    {
				    alert("The text you have entered exceeds the character limitation for that field, which is 250 characters. Please enter a shorter description and submit the form again.");
				    theForm.nfsmsg.focus();
				    return (false);
			    }
		    }
		    if (theForm.coption.value == "12")
			{
				if (theForm.pcv_height.value == "")
			    {
				    alert("Please enter a value for this field.");
				    theForm.pcv_height.focus();
				    return (false);
			    }
			    if (theForm.pcv_width.value == "")
			    {
				    alert("Please enter a value for this field.");
				    theForm.pcv_width.focus();
				    return (false);
			    }
			    if (theForm.pcv_length.value == "")
			    {
				    alert("Please enter a value for this field.");
				    theForm.pcv_length.focus();
				    return (false);
			    }
			    if (allDigit1(theForm.pcv_height.value) == false)
				{
					alert("Please enter a right integer value for this field.");
					theForm.pcv_height.focus();
					return (false);
				}
				if (allDigit1(theForm.pcv_width.value) == false)
				{
					alert("Please enter a right integer value for this field.");
					theForm.pcv_width.focus();
					return (false);
				}
				if (allDigit1(theForm.pcv_length.value) == false)
				{
					alert("Please enter a right integer value for this field.");
					theForm.pcv_length.focus();
					return (false);
				}
		    }

	}
	if (theForm.UP1.value == "4")
  	{
			if (theForm.roption.value == "")
			{
		    alert("Please select one option.");
		    theForm.roption.focus();
		    return (false);
		    }
	}					
	<%If scBTO=1 then %>
	if (theForm.UP1.value == "5")
  	{
			if (theForm.ptoption.value == "")
			{
		    alert("Please select one option.");
		    theForm.ptoption.focus();
		    return (false);
		    }
	}	
	<%end if%>
	
	if (theForm.UP1.value == "6")
  	{
	<% If scShipFromWeightUnit="KGS" then %>
		if ((theForm.weight_units.value == "") || (theForm.weight_units.value == "0"))
		{
			if (theForm.weight_kg.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.weight_kg.focus();
		    return (false);
		    }
			if (allDigit1(theForm.weight_kg.value) == false)
			{
		    alert("Please enter a right integer value for this field.");
		    theForm.weight_kg.focus();
		    return (false);
		   }
			if (theForm.weight_g.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.weight_g.focus();
		    return (false);
		    }
			if (allDigit1(theForm.weight_g.value) == false)
			{
		    alert("Please enter a right integer value for this field.");
		    theForm.weight_g.focus();
		    return (false);
		   }
		    if ((theForm.weight_kg.value == "0") && (theForm.weight_g.value == "0") )
			{
		    alert("Please enter a value greater than zero for product weight.");
		    theForm.weight_kg.focus();
		    return (false);
		    }
		}
		else
		{
			if (allDigit1(theForm.weight_units.value) == false)
			{
				alert("Please enter a right integer value for this field.");
				theForm.weight_units.focus();
				return (false);
			}
		}
	<%else%>
		if ((theForm.weight_units.value == "") || (theForm.weight_units.value == "0"))
		{
			if (theForm.weight.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.weight.focus();
		    return (false);
		    }
			if (allDigit1(theForm.weight.value) == false)
			{
		    alert("Please enter a right integer value for this field.");
		    theForm.weight.focus();
		    return (false);
		   }
			if (theForm.weight_oz.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.weight_oz.focus();
		    return (false);
		    }
			if (allDigit1(theForm.weight_oz.value) == false)
			{
		    alert("Please enter a right integer value for this field.");
		    theForm.weight_oz.focus();
		    return (false);
		   }
		   if ((theForm.weight.value == "0") && (theForm.weight_oz.value == "0") )
			{
		    alert("Please enter a value greater than zero for product weight.");
		    theForm.weight.focus();
		    return (false);
		    }
		}
		else
		{
			if (allDigit1(theForm.weight_units.value) == false)
			{
				alert("Please enter a right integer value for this field.");
				theForm.weight_units.focus();
				return (false);
			}
		}
	<%end if%>		    
	} 
	if (theForm.UP1.value == "7")
  	{
			if (theForm.numoptions.value == "")
			{
		    alert("Please select a product setting.");
		    theForm.numoptions.focus();
		    return (false);
		    }
			if (theForm.numvalue.value == "")
			{
		    alert("Please enter a numeric value for this field.");
		    theForm.numvalue.focus();
		    return (false);
		    }
			if (allDigit(theForm.numvalue.value) == false)
			{
		    alert("Please enter a numeric value for this field.");
		    theForm.numvalue.focus();
		    return (false);
		   }		    
	}
	
	if (theForm.UP1.value == "8")
  	{
			if (theForm.minimumqty.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.minimumqty.focus();
		    return (false);
		    }
			if (allDigit1(theForm.minimumqty.value) == false)
			{
		    alert("Please enter a right integer value for this field.");
		    theForm.minimumqty.focus();
		    return (false);
			}
	}
	<%'Start SDBA
	'Get Suppliers List
	query="Select pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers order by pcSupplier_Company asc"
	set rs=connTemp.execute(query)
	if not rs.eof then%>
	if (theForm.UP1.value == "9")
  	{
			if (theForm.pcToIDSupplier.value == "")
			{
		    alert("Please select a supplier.");
		    theForm.pcToIDSupplier.focus();
		    return (false);
		    }
	}
	<%end if
	set rs=nothing
	'End SDBA%>
	
	<%'Start SDBA
	'Get Drop-Shippers List
	query="SELECT pcDropShipper_ID,pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName,0 FROM pcDropShippers UNION (SELECT pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName,1 FROM pcSuppliers WHERE pcSupplier_IsDropShipper=1) ORDER BY pcDropShipper_Company ASC"
	set rs=connTemp.execute(query)
	if not rs.eof then%>
	if (theForm.UP1.value == "10")
  	{
			if (theForm.pcToIDDropshipper.value == "")
			{
		    alert("Please select a drop-shipper.");
		    theForm.pcToIDDropshipper.focus();
		    return (false);
		    }
	}
	<%end if
	set rs=nothing
	'End SDBA%>
	
	if (theForm.UP1.value == "11")
  	{
			if (theForm.stroptions.value == "")
			{
		    alert("Please select a product setting.");
		    theForm.stroptions.focus();
		    return (false);
		    }
			if (theForm.strvalue.value == "")
			{
		    alert("Please enter a string value for this field.");
		    theForm.strvalue.focus();
		    return (false);
		    }
	}
	
<%query="Select IDBrand,BrandName from Brands order by BrandName asc"
set rs=connTemp.execute(query)
if not rs.eof then%>
	if (theForm.UP1.value == "12")
  	{
			if (theForm.pcToIDBrand.value == "")
			{
		    alert("Please select a brand name.");
		    theForm.pcToIDBrand.focus();
		    return (false);
		    }
	}
<%end if
set rs=nothing
%>

if (theForm.UP1.value == "14")
  	{
			if ((theForm.ToIDCategory.value == "0") || (theForm.ToIDCategory.value == ""))
			{
		    alert("Please select a category.");
		    theForm.ToIDCategory.focus();
		    return (false);
		    }
	}
	
if (theForm.UP1.value == "15")
  	{
			if (theForm.goSett.value == "")
			{
		    alert("Please select a Google Shopping Setting");
		    theForm.goSett.focus();
		    return (false);
		    }
			if (theForm.goValue.value == "")
			{
		    alert("Please enter Google Shopping Setting value");
		    theForm.goValue.focus();
		    return (false);
		    }
	}
	
return (true);
}
//-->
</script>
      

<form name="UpdateForm" action="globalChanges.asp?action=update" method="post" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" name="nav" value="<%=request("nav")%>">
	<input type="hidden" name="CP1" value="0">
	<input type="hidden" name="UP1" value="0">
<table class="pcCPcontent">
<tr> 
	<td colspan="2">
		<p>This feature allows you to apply changes that affect multiple products in your store. Please note that changes cannot be undone, and therefore this feature should be used with caution. For more information, please <a href="http://wiki.earlyimpact.com/productcart/products_global_changes" target="_blank">refer to the ProductCart User Guide</a>.</p>
		<% If scDB="SQL" Then %>
            <div class="pcCPmessageInfo">	
            Products that are currently included in a running Sale will not be affected by Global Changes.</div>
        <% end If %>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Select the products to which the change will apply:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="5%" align="right">                          
		<input type="radio" name="CP" value="1" onClick="UpdateForm.CP1.value='1';" class="clearBorder">
	</td>
	<td>All</td>
</tr>
<tr valign="top">
	<td width="5%" align="right">                          
		<input type="radio" name="CP" value="2" onClick="UpdateForm.CP1.value='2';" class="clearBorder">
	</td>
	<td>All products in the following category
		<%
		cat_DropDownName="idcategory"
		cat_Type="1"
		cat_DropDownSize="1"
		cat_MultiSelect="0"
		cat_ExcBTOHide="0"
		cat_StoreFront="0"
		cat_ShowParent="1"
		cat_DefaultItem="All"
		cat_SelectedItems="0,"
		cat_ExcItems=""
		cat_ExcSubs="0"
		cat_EventAction=""
		%>
		<!--#include file="../includes/pcCategoriesList.asp"-->
		<%call pcs_CatList()%>
		<br>
		<input type="checkbox" name="incSubCats" value="1" class="clearBorder"> Include sub-categories
	</td>
</tr>
<tr> 
	<td width="5%" align="right">              
		<input type="radio" name="CP" value="3" onClick="UpdateForm.CP1.value='3';" class="clearBorder">
	</td>
	<td> Products whose part number (SKU) contains: 
		<input name="sku" type="text" size="15" maxlength="150" value="">
	</td>
</tr>
<tr>
<td width="5%" align="right">
	<input type="radio" name="CP" value="4" onClick="UpdateForm.CP1.value='4';" class="clearBorder">
</td>
	<td>Products whose name or description contains: 
		<input name="nd" type="text" size="15" maxlength="150">
	</td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="5" onClick="UpdateForm.CP1.value='5';" class="clearBorder">
	</td>
	<td>Products whose <select name="hpType">
		<option value="" selected></option>
		<option value="1">Online Price</option>
		<option value="2">List Price</option>                          
		</select> is higher than: 
		<input name="hprice" type="text" size="15" maxlength="150">
	</td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="6" onClick="UpdateForm.CP1.value='6';" class="clearBorder"></td>
	<td>Products whose <select name="lpType">
		<option value="" selected></option>
		<option value="1">Online Price</option>
		<option value="2">List Price</option>                          
		</select> is lower than:
		<input name="lprice" type="text" size="15" maxlength="150">
	</td>
</tr>
<%query="Select IDBrand,BrandName from Brands order by BrandName asc"
set rstemp4=connTemp.execute(query)
if not rstemp4.eof then%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="7" onClick="UpdateForm.CP1.value='7';" class="clearBorder">
	</td>
	<td>All products belonging to the following brand:
		<select name="IDBrand">
			<option value="" selected>Select one...</option>
			<option value="0">All</option>
			<%do while not rstemp4.eof%>
			<option value="<%=rstemp4("IDBrand")%>"><%=rstemp4("BrandName")%></option>
			<%rstemp4.MoveNext
			loop%>
		</select>
	</td>
</tr>
<%end if%>
<%'Start SDBA
'Get Suppliers List
query="Select pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers order by pcSupplier_Company asc"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="8" onClick="UpdateForm.CP1.value='8';" class="clearBorder">
	</td>
	<td>All products belonging to the following supplier:
		<select name="pcIDSupplier">
			<option value="" selected>Select one...</option>
			<option value="0">All</option>
			<%For i=0 to intCount%>
				<option value="<%=pcArray(0,i)%>"><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
			<%Next%>
		</select>
	</td>
</tr>
<%end if
set rs=nothing
'End SDBA%>
<%'Start SDBA
'Get Drop-Shippers List
query="SELECT pcDropShipper_ID,pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName,0 FROM pcDropShippers UNION (SELECT pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName,1 FROM pcSuppliers WHERE pcSupplier_IsDropShipper=1) ORDER BY pcDropShipper_Company ASC"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="9" onClick="UpdateForm.CP1.value='9';" class="clearBorder">
	</td>
	<td>All products belonging to the following drop-shipper:
		<select name="pcIDDropshipper">
			<option value="" selected>Select one...</option>
			<option value="0">All</option>
			<%For i=0 to intCount%>
				<option value="<%=pcArray(0,i)%>_<%=pcArray(4,i)%>"><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
			<%Next%>
		</select>
	</td>
</tr>
<%end if
set rs=nothing
'End SDBA%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="10" onClick="UpdateForm.CP1.value='10';" class="clearBorder">
	</td>
	<td>
		All products that are: 
<input name="pcv_instock" type="radio" value="1" checked  class="clearBorder"> In Stock&nbsp;&nbsp;&nbsp;&nbsp;<input name="pcv_instock" type="radio" value="0" class="clearBorder"> Out of stock
	</td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="11" onClick="UpdateForm.CP1.value='11';" class="clearBorder">
	</td>
	<td>
		All... 
		<input name="pcv_prdtype1" type="checkbox" value="1"<% if section<>"services" then %> checked<% end if %> class="clearBorder">Standard Products
		<%if scBTO=1 then%>
		&nbsp;&nbsp;&nbsp;<input name="pcv_prdtype2" type="checkbox" value="1" <% if section="services" then %> checked<% end if %> class="clearBorder">BTO Products
		&nbsp;&nbsp;&nbsp;<input name="pcv_prdtype3" type="checkbox" value="1" class="clearBorder">BTO Items
		<%end if%>
		&nbsp;&nbsp;&nbsp;<input name="pcv_prdtype4" type="checkbox" value="1" class="clearBorder">Downloadable Products
		&nbsp;&nbsp;&nbsp;<input name="pcv_prdtype5" type="checkbox" value="1" class="clearBorder">Gift Certificates
</td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="CP" value="13" onClick="UpdateForm.CP1.value='13';" class="clearBorder">
	</td>
	<td>Products that are not assigned to any category</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th  colspan="2">Select the change that you would like to apply:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="1" onClick="UpdateForm.UP1.value='1';" class="clearBorder">
	</td>
	<td>Change the&nbsp;            
		<select name="priceSelect" id="priceSelect" size="1">
			<option value="1" selected>Online Price</option>
			<option value="2">List Price</option>
			<option value="3">Wholesale Price</option>
			<%tmp_HavePricingCAT=0
			query="Select idcustomerCategory, pcCC_Name FROM pcCustomerCategories order by pcCC_Name asc"
			set rstemp4=connTemp.execute(query)
			if not rstemp4.eof then
			tmp_HavePricingCAT=1%>
			<%do while not rstemp4.eof%>
			<option value="CC_<%=rstemp4("idcustomerCategory")%>"><%=rstemp4("pcCC_Name")%></option>
			<%rstemp4.MoveNext
			loop%>
			<%end if
			set rstemp4=nothing%>
		</select>
		&nbsp;by:&nbsp;              
		<input name="cprice" type="text" id="priceChange" size="8" maxlength="150">
		<select name="cpriceType" id="cpriceType" size="1">
			<option value="1" selected>% change</option>
			<option value="2"># change</option>
		</select>
		&nbsp;&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=403')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input name="cpriceRound" type="radio" id="cpriceRound" value="1" class="clearBorder">&nbsp;Round updated price to the nearest integer
		<br>
		<input name="cpriceRound" type="radio" id="cpriceRound" value="2" checked class="clearBorder">&nbsp;Round updated price to the nearest hundredth
	</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td width="5%" align="right" valign="top">
		<input type="radio" name="UP" value="2" onClick="UpdateForm.UP1.value='2';" class="clearBorder">
	</td>
	<td valign="top">Recalculate the&nbsp; 
		<select name="priceSelect1" size="1">
			<option value="1" selected>Online Price</option>
			<option value="2">List Price</option>
			<option value="3">Wholesale Price</option>
			<%tmp_HavePricingCAT=0
			query="Select idcustomerCategory, pcCC_Name FROM pcCustomerCategories WHERE pcCC_CategoryType<>'ATB' ORDER BY pcCC_Name asc"
			set rstemp4=connTemp.execute(query)
			if not rstemp4.eof then
			tmp_HavePricingCAT=1%>
			<%do while not rstemp4.eof%>
			<option value="CC_<%=rstemp4("idcustomerCategory")%>"><%=rstemp4("pcCC_Name")%></option>
			<%rstemp4.MoveNext
			loop%>
			<%end if
			set rstemp4=nothing%>
		</select>
		&nbsp;as&nbsp;
		<input name="wprice" type="text" id="listPriceChange" size="8" maxlength="150">
		% of the&nbsp;
		<select name="priceSelect2" size="1">
			<option value="1" selected>Online Price</option>
			<option value="2">List Price</option>
			<option value="3">Wholesale Price</option>
			<%'Start SDBA%>
			<option value="4">Product Cost</option>
			<%'End SDBA%>
			<%tmp_HavePricingCAT=0
			query="Select idcustomerCategory, pcCC_Name FROM pcCustomerCategories WHERE pcCC_CategoryType<>'ATB' ORDER BY pcCC_Name asc"
			set rstemp4=connTemp.execute(query)
			if not rstemp4.eof then
			tmp_HavePricingCAT=1%>
			<%do while not rstemp4.eof%>
			<option value="CC_<%=rstemp4("idcustomerCategory")%>"><%=rstemp4("pcCC_Name")%></option>
			<%rstemp4.MoveNext
			loop%>
			<%end if
			set rstemp4=nothing%>
		</select>
		&nbsp;&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=402')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input name="cpriceRound1" type="radio" id="cpriceRound1" value="1" class="clearBorder">&nbsp;Round updated price to the nearest integer
		<br>
		<input name="cpriceRound1" type="radio" id="cpriceRound1" value="2" checked class="clearBorder">&nbsp;Round updated price to the nearest hundredth
	</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="3" onClick="UpdateForm.UP1.value='3';" class="clearBorder">
	</td>
	<td>Assign the following option:
		<select name="coption" id="optionSelect" size="1">
			<option value="" selected>Select one...</option>
			<option value="6">Active</option>
			<option value="1">Show Savings</option>
			<option value="28">Featured Product</option>
			<option value="2">Special</option>
			<option value="3">Non Taxable</option>
			<option value="4">Not for Sale</option>
			<option value="5">Free or No Shipping</option>
			<option value="8">Display No Shipping Text</option>
			<option value="29">Hide SKU on Product Details page</option>
			<option value="7">Disregard Stock</option>
			<%'Start SDBA%>
			<option value="9">Back-Ordering</option>
			<option value="10">Low Inventory Notification</option>
			<option value="11">Drop-Shipped</option>
			<%'End SDBA%>
			<option value="12">Oversized</option>
			<option value="13">Downloadable Product</option>
			<option value="14">Make download URL expire (Downloadable Product)</option>
			<option value="15">Deliver license with order confirmation (Downloadable Product)</option>
			<option value="16">Gift Certificate</option>
			<option value="17">Does not expire (Gift Certificate)</option>
			<option value="18">Expires on the Date (Gift Certificate)</option>
			<option value="19">Expires N days after purchase (Gift Certificate)</option>
			<option value="20">Electronic Only (Gift Certificate)</option>
			<option value="21">Use default generator (Gift Certificate)</option>
			<option value="22">Use custom generator (Gift Certificate)</option>
			<%if scBTO=1 then%>
			<option value="23">Hide BTO Price (BTO Product)</option>
			<option value="24">Hide Default Configuration (BTO Product)</option>
			<option value="25">Disallow purchasing - Show Prices (BTO Product)</option>
			<option value="26">Disallow purchasing - Hide Prices (BTO Product)</option>
			<option value="27">Skip Product Details Page (BTO Product)</option>
			<%end if%>
			<option value="30">Image Magnifier (MojoZoom)</option>            
		</select>
	</td>
</tr>
<tr>
	<td width="5%" align="right">&nbsp;</td>
	<td>If you select &quot;Not for Sale&quot;, you enter a message here:</td>
</tr>
<tr>
	<td width="5%" align="right">&nbsp;</td>
	<td>
        <textarea name="nfsmsg" id="nfsmsg" rows="4" cols="40" tabindex="711" onkeyup="javascript:testchars(this,'1',250); javascript:document.getElementById('emailTextCounter').style.display='';"></textarea>
        <div id="emailTextCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar1" name="countchar1" style="font-weight: bold"><%=maxlength%></span> characters left.</div>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="5%" align="right">&nbsp;</td>
	<td>If you select &quot;Oversized&quot;, set the size below in inches:</td>
</tr>
<tr>
	<td width="5%" align="right">&nbsp;</td>
	<td>
		Height: <input name="pcv_height" type="text" size="4" maxlength="10">&nbsp;&nbsp;&nbsp;&nbsp; Width: <input name="pcv_width" type="text" size="4" maxlength="10">&nbsp;&nbsp;&nbsp;&nbsp;Length: <input name="pcv_length" type="text" size="4" maxlength="10">
	</td>
</tr>
<tr>
	<td width="5%" align="right">&nbsp;</td>
	<td>The selected option will be assigned to the products chosen above.</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="4" onClick="UpdateForm.UP1.value='4';" class="clearBorder">
	</td>
	<td>Unassign the following option:
		<select name="roption" id="optionSelect" size="1">
			<option value="" selected>Select one...</option>
			<option value="6">Active</option>
			<option value="22">Featured Product</option>
			<option value="1">Show Savings</option>
			<option value="2">Special</option>
			<option value="23">Hide SKU on Product Details page</option>
			<option value="3">Non Taxable</option>
			<option value="4">Not for Sale</option>
			<option value="5">Free or No Shipping</option>
			<option value="8">Display No Shipping Text</option>
			<option value="7">Disregard Stock</option>
			<%'Start SDBA%>
			<option value="9">Back-Ordering</option>
			<option value="10">Low Inventory Notification</option>
			<option value="11">Drop-Shipped</option>
			<%'End SDBA%>
			<option value="12">Oversized</option>
			<option value="13">Downloadable Product</option>
			<option value="14">Make download URL expire (Downloadable Product)</option>
			<option value="15">Deliver license with order confirmation (Downloadable Product)</option>
			<option value="16">Gift Certificate</option>
			<option value="17">Electronic Only (Gift Certificate)</option>
			<%if scBTO=1 then%>
			<option value="18">Hide BTO Price (BTO Product)</option>
			<option value="19">Hide Default Configuration (BTO Product)</option>
			<option value="20">Disallow purchasing (BTO Product)</option>
			<option value="21">Skip Product Details Page (BTO Product)</option>
			<%end if%>
			<option value="24">Image Magnifier (MojoZoom)</option>            
		</select>
	</td>
</tr>
<tr>
	<td width="5%" align="right">&nbsp;</td>
	<td>The selected option will be removed from the products chosen above.</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<%If scBTO=1 then %>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="5" onClick="UpdateForm.UP1.value='5';" class="clearBorder">
	</td>
	<td>Change Product Type:
		<select name="ptoption" id="optionSelect" size="1">
			<option value="" selected>Select one...</option>
			<option value="1">Standard Product</option>
			<option value="2">BTO Product</option>
			<option value="3">BTO Only Item</option>
		</select>
	</td>
</tr>
<%end if%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="6" onClick="UpdateForm.UP1.value='6';" class="clearBorder">
	</td>
	<td>Set the weight to: 
	<% If scShipFromWeightUnit="KGS" then %>
	<input type="text" name="weight_kg" value="0" size="4">
			kg
	<input type="text" name="weight_g" value="0" size="4">
	g
	&nbsp;&nbsp;&nbsp;Units to make 1 kg: <input type="text" name="weight_units" value="0" size="4">								
	<% else %>
	<input type="text" name="weight" value="0" size="4">
                            lbs. 
	<input type="text" name="weight_oz" value="0" size="4">
                            ozs.
	&nbsp;&nbsp;&nbsp;Units to make 1 lb: <input type="text" name="weight_units" value="0" size="4">										
	<% end if %>
	</td>
</tr>
<%'Start SDBA%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="7" onClick="UpdateForm.UP1.value='7';" class="clearBorder">
	</td>
	<td>Set the <select name="numoptions" id="numoptionSelect" size="1">
			<option value="" selected>Select one...</option>
			<option value="1">Stock Level</option>
			<option value="2">Product Cost</option>
			<option value="3">Reorder Level</option>
			<option value="4">Ship within N Days</option>
			<% 'RP ADDON-S
			If RewardsActive <> 0 Then %>
			<option value="5"><%=RewardsLabel%></option>
			<%End if
			'RP ADDON-E%>
			<option value="6">URL will expire after N days (Downloadable Product)</option>
			<option value="7">Expires after N Days (Gift Certificate)</option>
			</select>
			to the value: <input name="numvalue" type="text" id="numvalueChange" size="8" maxlength="150"> (Units / $ / % / days / points)</td>
</tr>
<%'End SDBA%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="11" onClick="UpdateForm.UP1.value='11';" class="clearBorder">
	</td>
	<td>Set the <select name="stroptions" id="stroptionSelect" size="1">
			<option value="" selected>Select one...</option>
			<option value="1">Downloadable File Location (Downloadable Product)</option>
			<option value="2">Local license generator (Downloadable Product)</option>
			<option value="3">Remote license generator (Downloadable Product)</option>
			<option value="4">License Field (1) (Downloadable Product)</option>
			<option value="5">License Field (2) (Downloadable Product)</option>
			<option value="6">License Field (3) (Downloadable Product)</option>
			<option value="7">License Field (4) (Downloadable Product)</option>
			<option value="8">License Field (5) (Downloadable Product)</option>
			<option value="9">Custom generator file name (Gift Certificate)</option>
			<option value="10">Expiration Date (Gift Certificate)</option>
			</select>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		Value: <input name="strvalue" type="text" id="strvalueChange" size="60" maxlength="255">
	</td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="8" onClick="UpdateForm.UP1.value='8';" class="clearBorder">
	</td>
	<td>Minimum quantity customers can buy: <input name="minimumqty" type="text" size="8" maxlength="150"> Units&nbsp;&nbsp; 
	|&nbsp;&nbsp; Force purchase of multiples of minimum: <input type="checkbox" name="qtyvalidate" value="1" class="clearBorder"></td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<%'Start SDBA
'Get Suppliers List
query="Select pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers order by pcSupplier_Company asc"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
	set rs=nothing%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="9" onClick="UpdateForm.UP1.value='9';" class="clearBorder">
	</td>
	<td>Move filtered products to the following supplier:
		<select name="pcToIDSupplier">
			<option value="" selected>Select one...</option>
			<%For i=0 to intCount%>
				<option value="<%=pcArray(0,i)%>"><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
			<%Next%>
		</select>
	</td>
</tr>
<%end if
set rs=nothing
'End SDBA%>
<%'Start SDBA
'Get Drop-Shippers List
query="SELECT pcDropShipper_ID,pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName,0 FROM pcDropShippers UNION (SELECT pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName,1 FROM pcSuppliers WHERE pcSupplier_IsDropShipper=1) ORDER BY pcDropShipper_Company ASC"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
	set rs=nothing%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="10" onClick="UpdateForm.UP1.value='10';" class="clearBorder">
	</td>
	<td>Move filtered products to the following drop-shipper:
		<select name="pcToIDDropshipper">
			<option value="" selected>Select one...</option>
			<%For i=0 to intCount%>
				<option value="<%=pcArray(0,i)%>_<%=pcArray(4,i)%>"><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
			<%Next%>
		</select>
	</td>
</tr>
<%end if
set rs=nothing
'End SDBA%>

<%query="Select IDBrand,BrandName from Brands order by BrandName asc"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
	set rs=nothing%>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="12" onClick="UpdateForm.UP1.value='12';" class="clearBorder">
	</td>
	<td>Move filtered products to the following brand:
		<select name="pcToIDBrand">
			<option value="" selected>Select one...</option>
			<%For i=0 to intCount%>
				<option value="<%=pcArray(0,i)%>"><%=pcArray(1,i)%></option>
			<%Next%>
		</select>
	</td>
</tr>
<%end if
set rs=nothing%>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td width="5%" align="right">                          
		<input type="radio" name="UP" value="14" onClick="UpdateForm.UP1.value='14';" class="clearBorder">
	</td>
	<td>Assign selected products to the following category:
		<%
		cat_DropDownName="ToIDCategory"
		cat_Type="0"
		cat_DropDownSize="1"
		cat_MultiSelect="0"
		cat_ExcBTOHide="0"
		cat_StoreFront="0"
		cat_ShowParent="1"
		cat_DefaultItem="Select one..."
		cat_SelectedItems="0,"
		cat_ExcItems=""
		cat_ExcSubs="0"
		cat_EventAction=""
		pcv_havecats1=0
		pcv_havecats2=0
		%>
		<%call pcs_CatList()%>
	</td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="13" onClick="UpdateForm.UP1.value='13';" class="clearBorder">
	</td>
    <td>Page Layout:
    	<select name="pcv_displayLayout" id="displayLayout">
	        <option value="" selected>Use Default</option>
	        <option value="c">Two Columns-Image on Right</option>
	        <option value="l">Two Columns-Image on Left</option>
	        <option value="o">One-Column</option>
        </select>
    </td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td><b>Google Shopping Settings</b></td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="15" onClick="UpdateForm.UP1.value='15';" class="clearBorder">
	</td>
    <td>Set the:
    	<select name="goSett" id="goSett">
	        <option value="" selected>Select one... </option>
	        <option value="1">Google Product Category</option>
	        <option value="2">Google Shopping - Gender</option>
	        <option value="3">Google Shopping - Age</option>
			<option value="4">Google Shopping - Color</option>
			<option value="5">Google Shopping - Size</option>
			<option value="6">Google Shopping - Pattern</option>
			<option value="7">Google Shopping - Material</option>
        </select>
		&nbsp;to the value: <input type="text" name="goValue" id="goValue" size="20" value="">
    </td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr align="center">
	<td colspan="2">
		<input type="submit" name="Submit" value="Preview" class="submit2">&nbsp;
		<input type="button" name="back" value="Back" onClick="javascript:history.back()">
	</td>
</tr>
</table>
</form>
<% 
End if
call closeDb()
set rstemp= nothing
set rstemp4= nothing
set rstemp5= nothing
%>
<!--#include file="AdminFooter.asp"-->