<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=false
Server.ScriptTimeout = 54000%>
<% pageTitle="Generate Google Shopping Data Feed" %>
<% Section="genRpts" %>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="frooglecurrencyformatinc.asp"--> 
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())

dim rstemp, rstemp2, conntemp, mysql, strtext, strtext1, File1, pcv_rootCat, strParent, intIdProduct, pcIntIncludeWeight, pcIntUniqueIdent

'/////////////////////////////////////////////////////////////////////
'// START: GEN FILE
'/////////////////////////////////////////////////////////////////////
IF request("action")="gen" THEN	

	'// Get Filename
	File1=request("pcv_filename")
	pcv_rootCat=request("pcv_rootCat")
	
	'// Unique identifiers?
	pcIntUniqueIdent=request("UniqueIdent")
	if pcIntUniqueIdent = "" then
		pcIntUniqueIdent = 1
	end if
	
	'// Get Currency
	strCurrency=request("idCurrency")
	if strCurrency = "" then
		strCurrency = "USD"
	end if

	'// Include weight
	pcIntIncludeWeight=request("includeWeight")
	if not validNum(pcIntIncludeWeight) then pcIntIncludeWeight=0
	
	'// Include Apparel
	pcIntIncludeApparel=request("includeApparel")
	if not validNum(pcIntIncludeApparel) then pcIntIncludeApparel=0
	
	'// Get Condition
	strCondition=request("idCondition")
	if strCondition = "" then
		strCondition = "new"
	end if
	
	'// Get Date
	strExpDate=request("ExpirationDate")
	if strExpDate="" then
		strExpDate=Date()+30
		strExpDate=Year(strExpDate) & "-" & FixDate(Month(strExpDate)) & "-" & FixDate(Day(strExpDate))
	end if
	
	'// Get Brand
	strCustomBrand=request("CustomBrand")

	'// Get Path
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
	if Right(SPathInfo,1)<>"/" then
		SPathInfo=SPathInfo & "/"
	end if

	'// Get Category List
	pcv_strIncSubCats=request("incSubCats")
	pcv_IdCategory=request("idcategory")	
	if pcv_IdCategory="" then
	pcv_IdCategory=0
	end if
	pcv_IdCategory=trim(pcv_IdCategory)	
	pcList1=split(pcv_IdCategory,",")
	
	'// Set Header Rows
    strtext1="title" & chr(9) & "description" & chr(9) & "product_type" & chr(9) & "link" & chr(9) & "image_link" & chr(9) & "id" & chr(9) & "availability" & chr(9) & "price" & chr(9) & "condition" & chr(9) & "brand" & chr(9) & "expiration_date" & chr(9)
	if pcIntUniqueIdent=1 then
		strtext1 = strtext1  & "mpn"  & chr(9) & "gtin" & chr(9)
	end if
	strtext1 = strtext1  & "currency" & chr(9)
	if pcIntIncludeWeight=1 then
		strtext1 = strtext1 & "shipping_weight" & chr(9)
	end if
	strtext1 = strtext1 & "google_product_category" & chr(9)
	if pcIntIncludeApparel=1 then
		strtext1 = strtext1 & "gender" & chr(9)
		strtext1 = strtext1 & "age_group" & chr(9)
		strtext1 = strtext1 & "color" & chr(9)
		strtext1 = strtext1 & "size" & chr(9)
		strtext1 = strtext1 & "pattern" & chr(9)
		strtext1 = strtext1 & "material" & chr(9)
	end if
	strtext1 = strtext1 & "featured_product" & vbcrlf
	%>
	<div id="ea">
	<b><span id="co">0</span></b> product(s) were exported successfully<br />
	Exporting...
	</div>
	<%
	Count=0
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start:  Do For Each Category
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		tmpCats=""
		For lk=lbound(pcList1) to ubound(pcList1)
		'// Filter By Category
			If trim(pcList1(lk))<>"0" then
				if tmpCats<>"" then
					tmpCats=tmpCats & ","
				end if
				tmpCats=tmpCats & trim(pcList1(lk))
				'// Add on the Sub-Categories for this Parent
				If len(pcv_strIncSubCats)>0 Then
					Dim TmpCatList
					TmpCatList=""
					call opendb()
					call pcs_GetSubCats(trim(pcList1(lk))) '// get sub cats
					call closedb()
					tmpCats=tmpCats & TmpCatList
				End If
			End If
		Next
		If tmpCats<>"" then		
			query1=" AND categories_products.idcategory IN (" & tmpCats & ") "		
		Else		
			query1=""		
		End if
		
		if request("excWCats")="1" then
			ExcWCatsStr=" AND categories.pccats_RetailHide<>1 "
		else
			ExcWCatsStr=""
		end if
		
		if request("excNFSPrds")="1" then
			ExcNFSPrds=" AND products.formQuantity=0 "
		else
			ExcNFSPrds=""
		end if
		
		'// Select Products
		call opendb()
		query="SELECT categories_products.idcategory,categories.categoryDesc,categories.pccats_BreadCrumbs,products.idProduct,products.description,products.serviceSpec,products.price,products.imageUrl,products.sku,products.stock,products.hotDeal,products.showInHome,products.IDBrand,products.noStock,products.pcProd_BackOrder,products.pcProd_BTODefaultPrice,products.weight,products.pcProd_GoogleCat,products.pcProd_GoogleGender,products.pcProd_GoogleAge,products.pcProd_GoogleSize,products.pcProd_GoogleColor,products.pcProd_GooglePattern,products.pcProd_GoogleMaterial,products.pcProd_GoogleGroup,products.details,products.sDesc FROM categories_products,categories,products WHERE ((products.price>0) OR ((products.price=0) AND (products.serviceSpec<>0) AND (products.pcProd_BTODefaultPrice>0))) AND products.active = -1 " & ExcWCatsStr &  ExcNFSPrds & " and products.removed=0 and products.configOnly=0 and products.idproduct=categories_products.idproduct AND categories.idcategory=categories_products.idcategory " & query1 & " order by categories_products.idcategory asc,products.description;"
		if request("excWCats")="1" then
		querytest="SELECT DISTINCT products.idproduct FROM categories_products,categories,products WHERE ((products.price>0) OR ((products.price=0) AND (products.serviceSpec<>0) AND (products.pcProd_BTODefaultPrice>0))) AND products.active = -1 AND categories.pccats_RetailHide=1" &  ExcNFSPrds & " and products.removed=0 and products.configOnly=0 and products.idproduct=categories_products.idproduct AND categories.idcategory=categories_products.idcategory " & query1 & ";"
		end if
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)		
		intIDCategory=0
		strCatPath=""		
		Count1=-1		
		if not rs.eof then
			pcArray=rs.getRows()
			Count1=ubound(pcArray,2)
		end if		
		set rs=nothing
		
		'// Summary of Array Items
		' 0 = category ID
		' 1 = category description
		' 2 = category breadcrumb
		' 3 = product ID
		' 4 = product name
		' 5 = whether it is a BTO product
		' 6 = price
		' 7 = image file name
		' 8 = sku
		' 9 = stock
		' 10 = whether it's a special
		' 11 = whether it's a featured item
		' 12 = brand
		' 13 = disregard stock
		' 14 = backorder
		' 15 = BTO default price
		' 16 = weight
		' 17 = pcProd_GoogleCat
		' 18 = pcProd_GoogleGender
		' 19 = pcProd_GoogleAge
		' 20 = pcProd_GoogleSize
		' 21 = pcProd_GoogleColor
		' 22 = pcProd_GooglePattern
		' 23 = pcProd_GoogleMaterial
		' 24 = pcProd_GoogleGroup
		' 25 = long description
		' 26 = short description
				
		query="SELECT idBrand,BrandName FROM Brands;"
		set rsBrand=connTemp.execute(query)
		CountB=-1
		if not rsBrand.eof then
			pcArr2=rsBrand.getRows()
			CountB=ubound(pcArr2,2)
		end if
		set rsBrand=nothing
		
		call closedb()
		
		Count2=0
		Count3=0
		
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set a=fso.CreateTextFile(server.MapPath(".") & "\" & File1,True)
		a.Write(strtext1)
		strtext1=""
		
		set StringBuilderObj = new StringBuilder
		
		if Count1>-1 then
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start:  Do For Each Product In Category
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			For k=0 to Count1
				
				intIdProduct=pcArray(3,k)
				
				'IF session("tmp" & intIdProduct)<>"1" THEN
				
					'session("tmp" & intIdProduct)=1
					
					'If it's a BTO product, the price is the Default BTO Price
					If (pcArray(5,k)=-1) then
						If pcArray(15,k)>"0" then
							pcArray(6,k) = pcArray(15,k)
						End if
					ENd if
					dblPrice=pcArray(6,k)
					
					if isNull(dblPrice) OR dblPrice="" then
						dblPrice=0
					end if
					
					IF dblPrice>0 THEN
				
						if pcIntUniqueIdent=1 then
							tmpMPN=""
							tmpUPC=""
							tmpISBN=""
							tmpGTIN=""
					
							tmpMPN=pcf_FillByName("f","MPN")
							tmpUPC=pcf_FillByName("f","UPC")
							tmpISBN=pcf_FillByName("f","ISBN")

							if tmpUPC<>"" then
								tmpGTIN=tmpUPC
								else
								tmpGTIN=tmpISBN
							end if
						end if
				
						tmpIDCategory=pcArray(0,k)
						
						strPrdType=pcArray(1,k)
						'// Load category breadcrumb, if it exists
							pccats_BreadCrumbs=pcArray(2,k)
							IF pccats_BreadCrumbs<>"" AND instr(pccats_BreadCrumbs,"||") THEN
								pcArrayBreadCrumbs=split(pccats_BreadCrumbs,"|,|")
								strBreadCrumb=""
								for i=0 to ubound(pcArrayBreadCrumbs)
									pcArrayCrumb=split(pcArrayBreadCrumbs(i),"||")							
									if i=0 then
										strBreadCrumb=strBreadCrumb & pcArrayCrumb(1)
									else
										strBreadCrumb=strBreadCrumb & " > " & pcArrayCrumb(1)
									end if
								next
							strPrdType=strBreadCrumb
							End If
						'// END - Load breadcrumb
												
						strProductName=pcArray(4,k)
						strProductName1=pcArray(4,k)
						strBTO=pcArray(5,k)
							if IsNull(strBTO) or trim(strBTO)="" then strBTO=0
						strProductImg=pcArray(7,k)
						strSKU=pcArray(8,k)
						strStock=pcArray(9,k)
						strSpecial=pcArray(10,k)
						strSpecial2=pcArray(11,k)
							if strSpecial=-1 or strSpecial2=-1 then
								strFeatured="y"
								else
								strFeatured="n"
							end if
						strIDBrand=pcArray(12,k)
						if strIDBrand<>"" and strIDBrand<>"0" then
							strBrandName=strCustomBrand
							For m=0 to CountB
								if Clng(strIDBrand)=Clng(pcArr2(0,m)) then
									strBrandName=pcArr2(1,m)
									exit for
								end if
							Next
						else
							strBrandName=strCustomBrand
						end if
						
						strDisregardStock=pcArray(13,k)
							if IsNull(strDisregardStock) or trim(strDisregardStock)="" then strDisregardStock=0
						strBackOrderOK=pcArray(14,k)
							if IsNull(strBackOrderOK) or trim(strBackOrderOK)="" then strBackOrderOK=0
						'// Determine availability
						if scOutOfStockPurchase=0 then
							strAvailability = "in stock"
						else
							if CLng(strStock)<1 then
								If (scOutofStockPurchase=-1 AND strBTO=0 AND strDisregardStock=0 AND strBackOrderOK=1) OR (strBTO<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND strDisregardStock=0 AND strBackOrderOK=1) Then
									strAvailability = "available for order"
								else
									strAvailability = "out of stock"		
								end if
							else
								strAvailability = "in stock"
							end if
						end if
						'// END - Determine Availability
							
							
						strWeight=pcArray(16,k)
							if scShipFromWeightUnit="LBS" then
								strWeight = strWeight & " ounces"
								else
								strWeight = strWeight & " grams"								
							end if
						'//Google Shopping new attributes	
						If pcArray(17,k)<>"" then
							pcv_GCat=pcArray(17,k)
						Else
							pcv_GCat=strPrdType
						End if
						pcv_GGen = pcArray(18,k)
						pcv_GAge = pcArray(19,k)
						pcv_GSize = pcArray(20,k)
						pcv_GColor = pcArray(21,k)
						pcv_GPat = pcArray(22,k)
						pcv_GMat = pcArray(23,k)
						pcv_GGroup = pcArray(24,k)
						
						strProductDesc=pcArray(25,k)
						strShortProductDesc=pcArray(26,k)
								
						strParent=""
				
						'// SEO Links
						'// Build Product Link
						if scSeoURLs<>1 then
							strProductURL=SPathInfo & "pc/viewPrd.asp?idproduct=" & intIdProduct & "&idcategory=" & tmpIDCategory
						else
							strProductURL=SPathInfo & "pc/" & removeChars(strProductName1) & "-" & "p" & intIdProduct & ".htm"
						end if
						'//
				
						'// Check Short Description Text
						if trim(strShortProductDesc)<>"" then
							strProductDesc=trim(strShortProductDesc)
						end if
				
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Pricing
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
						dblPrice=money(dblPrice)
						
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Product Image
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' If not available, must be left blank
						if trim(strProductImg)<>"" and lcase(strProductImg) <> "no_image.gif" then
							strProductImgURL=SPathInfo & "pc/catalog/" & strProductImg
						else
							strProductImgURL=""
						end if
						
						Count3=Count3+1
						
						StringBuilderObj.append ClearHTMLTagsGB(strProductName & "/*/" & strProductDesc & "/*/" & strPrdType,0) & "/*/" & strProductURL & "/*/" & strProductImgURL & "/*/" & strSKU & "/*/" & strAvailability & "/*/" & dblPrice & "/*/" & strCondition & "/*/" & strBrandName & "/*/" & strExpDate
						if pcIntUniqueIdent=1 then
							StringBuilderObj.append "/*/" & tmpMPN & "/*/" & tmpGTIN
						end if
						StringBuilderObj.append "/*/" & strCurrency
						if pcIntIncludeWeight=1 then
							StringBuilderObj.append "/*/" & strWeight	
						end if
						StringBuilderObj.append "/*/" & pcv_GCat
						if pcIntIncludeApparel=1 then
							StringBuilderObj.append "/*/" & pcv_GGen
							StringBuilderObj.append "/*/" & pcv_GAge
							StringBuilderObj.append "/*/" & pcv_GColor
							StringBuilderObj.append "/*/" & pcv_GSize
							StringBuilderObj.append "/*/" & pcv_GPat
							StringBuilderObj.append "/*/" & pcv_GMat
						end if
						StringBuilderObj.append "/*/" & strFeatured & "/-/"	
												
						if Count3=250 then
							strtext = StringBuilderObj.toString
							strtext=replace(strtext,"/*/",chr(9))
							strtext=replace(strtext,"/-/",VBCrlf)
							a.Write(strtext)
							Count2=Count2+Count3
							response.write "<script>document.getElementById('co').innerHTML='" & Count2 & "';</script>"
							Count3=0
							strtext=""
							set StringBuilderObj = nothing
							set StringBuilderObj = new StringBuilder
						end if
						
					END IF
					
                'END IF
		
			Next
			
			strtext = StringBuilderObj.toString
			if strtext<>"" then
					strtext=replace(strtext,"/*/",chr(9))
					strtext=replace(strtext,"/-/",VBCrlf)
					a.Write(strtext)
					Count2=Count2+Count3
					response.write "<script>document.getElementById('co').innerHTML='" & Count2 & "';</script>"
					Count3=0
					strtext=""
			end if			
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End:  Do For Each Product In Category
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		end if '// '/ if Count1>-1 then
	
		set StringBuilderObj = nothing
		Count=Count2
		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// End:  Do For Each Category
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
	
	a.Close
	Set fso=Nothing
	response.write "<script>document.getElementById('ea').style.display='none';</script>"%>
	<table class="pcCPcontent">
		<tr> 
			<td>
			The Google Shopping data feed was created successfully.
				<br>
				<br>
				<b><%=Count%></b> products have been exported succefully to the Google Shopping data feed: <a href="<%=File1%>"><%=File1%></a>. To download the file, either right-click on the file name and select '<strong>Save Target As...</strong>' or '<strong>Save Link as...</strong>' or FTP into the <em><%=scAdminFolderName%></em> folder and download it. <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_google_base_file#uploading_or_scheduling" target="_blank">Learn more about submitting your Google Shopping file</a>.
				<br>
				<br>
				<br>
				<form class="pcForms">
					<input type=button name=back value="Create Another File" onclick="location='exportFroogle.asp'">&nbsp;
					<input type=button name=back value="Start Page" onclick="location='menu.asp'">
				</form>
				<br>
				<br>
			</td>
		</tr>
		<%IF tmpCats="" AND request("showReports")="1" THEN%>
		<tr>
			<th>Report(s):</th>
		</tr>
		<tr>
			<td>
				<b><%=Count%></b> products included in the Google Shopping feed<br />
				<%call opendb()
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.active = 0 AND products.removed=0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products are inactive and not included in the Google Shopping feed<br />"
					end if
				end if
				set rsQ=nothing
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.removed <> 0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products are designated as removed from the store and not included in the Google Shopping feed<br />"
					end if
				end if
				set rsQ=nothing
				
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.active <> 0 AND products.removed = 0 AND products.configOnly<>0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products are BTO items and not included in the Google Shopping feed<br />"
					end if
				end if
				set rsQ=nothing
				
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.active <> 0 AND products.removed = 0 AND products.configOnly=0 AND price<=0 AND pcProd_BTODefaultPrice<=0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products have a price equal to zero and not included in the Google Shopping feed<br />"
					end if
				end if
				set rsQ=nothing
				
				if request("excNFSPrds")="1" then
					query="SELECT Count(*) As tmpTotal FROM products WHERE products.active <> 0 AND products.removed = 0 AND products.configOnly=0 AND (price>0 OR pcProd_BTODefaultPrice>0) AND products.formQuantity<>0;"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						if rsQ("tmpTotal")>"0" then
						Count=Count+Clng(rsQ("tmpTotal"))
						response.write rsQ("tmpTotal") & " products are 'Not for Sale' products and not included in the Google Shopping feed<br />"
						end if
					end if
					set rsQ=nothing
				end if
				
				if request("excWCats")="1" then
					set rsQ=connTemp.execute(querytest)
					if not rsQ.eof then
						pcArr=rsQ.getRows()
						countM=ubound(pcArr,2)
						CountB=CountM+1
						For m=0 to CountM
							if session("tmp" & pcArr(0,m))="1" then
								CountB=CountB-1
							end if
						Next
						if CountB>0 then
							Count=Count+CountB
							response.write CountB & " products are assigned to wholesale categories and not included in the Google Shopping feed<br />"
						end if
					end if
					set rsQ=nothing
				end if
				
				query="SELECT Count(*) As tmpTotal FROM products;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					CountTotal=rsQ("tmpTotal")
					response.write rsQ("tmpTotal") & " total products in the database<br />"
					end if
				end if
				set rsQ=nothing
				
				if CountTotal<>Count and (CountTotal-Count>1) then
					response.write "<i>Note: there appear to be " & CountTotal-Count & " orphaned products that were not included in the Google Shopping feed. You can <a href=""javascript:location='srcFreePrds.asp';"">view them</a> here. Assign them to a category to include them in the data feed.</i>"
				end if
				
				%>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
			</td>
		</tr>
		<%END IF
		For k=0 to Count1
			session("tmp" & pcArray(2,k))=""
		Next%>
	</table>



<% END IF
'/////////////////////////////////////////////////////////////////////
'// END: GEN FILE
'/////////////////////////////////////////////////////////////////////




'/////////////////////////////////////////////////////////////////////
'// START: DISPLAY FORM
'/////////////////////////////////////////////////////////////////////
IF request("action")<>"add" AND request("action")<>"gen" THEN
	%>
	<script language="JavaScript">
	<!--
		
	function isTXT(s)
		{
			var test=""+s ;
			test1="";
			
			if (test.length<=4)
			{
			return (false);
			}
			for (var k=test.length-4; k <test.length; k++)
			{
				var c=test.substring(k,k+1);
				test1 += c
			}
			if (test1==".TXT"||test1==".Txt"||test1==".txt"||test1==".TXt"||test1==".TxT"||test1==".txT"||test1==".tXt"||test1==".tXT")
				{
					return (true);
				}
			
			return (false);
		}
		
	
	function Form1_Validator(theForm)
	{
	
		if (theForm.pcv_filename.value == "")
		{
			alert("Please enter file name.");
			theForm.pcv_filename.value == ""
			theForm.pcv_filename.focus();
			return (false);
		}
		else
		{	if (isTXT(theForm.pcv_filename.value)==false)
		{
			alert("Invalid TEXT file type (*.txt) is not allowed on Google Shopping.");
			theForm.pcv_filename.value == ""
			theForm.pcv_filename.focus();
			return (false);
			}
		}
		if (theForm.idcategory.value == "")
		{
			alert("Please select at least one category.");
			theForm.idcategory.focus();
			return (false);
		}
		validateRadioButton = -1;
		for (i=theForm.UniqueIdent.length-1; i > -1; i--) {
		if (theForm.UniqueIdent[i].checked) {
		validateRadioButton = i; i = -1;
		}
		}
		if (validateRadioButton == -1) {
		alert("Please specify whether to use unique product identifiers or not.");
		return false;
		}
	pcf_Open_GoogleBase2();
	return (true);
	}
	//-->
	</script>
    <form name="form1" method="post" action="exportFroogle.asp?action=gen" onSubmit="return Form1_Validator(this)" class="pcForms">
        <table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>	
        <tr>
            <th colspan="2">Notes about the data feed content</th>
        </tr>	
        <tr>
            <td colspan="2">			
                <ul>
                    <li>Data feed content: please review the <a href="http://www.google.com/support/merchants/bin/answer.py?hl=en&answer=188484" target="_blank">Google Shopping policies</a></li>
                    <li>Google Shopping now requires <a href="#unique">unique product identifiers</a>. See <a href="http://wiki.earlyimpact.com/productcart/managing_search_fields#mapping_custom_search_fields_to_export_fields" target="_blank">how to leverage Customer Search Fields</a> for this purpose.</li>
                    <li>More: <a href="http://www.google.com/support/merchants/bin/answer.py?hl=en_US&answer=188494" target="_blank">Product Search Feed specifications</a> (Google Web site) | <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_google_base_file" target="_blank">How to use this feature</a> (ProductCart WIKI).</li>
              </ul>
            </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>	
        <tr>
            <th colspan="2">Products to include in the data feed</th>
        </tr>	
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="2" align="left" valign="top">Select the <strong>categories</strong> that contain the products to be included in the Google Shopping data feed. Press down the CTRL key on your keyboard to select multiple categories. Use the text links below to limit the number of categories displayed in the list.
            	<%
				pcv_strCatType = request("Type")
				pcv_strExcSubs = request("ExcSubs")
				%>
                <div style="margin: 10px 0;">
                <a href="exportFroogle.asp">Only Categories with Products</a>&nbsp;|&nbsp;
                <a href="exportFroogle.asp?Type=2&ExcSubs=1">Only Parent Categories</a>&nbsp;|&nbsp;
                <a href="exportFroogle.asp?Type=2">Show All Categories</a>
				</div>
                <%
                cat_DropDownName="idcategory"
				If len(pcv_strCatType)>0 Then
                	cat_Type="0"
				Else
	                cat_Type="1"			
				End If
                cat_DropDownSize="10"
                cat_MultiSelect="1"
                cat_ExcBTOHide="1"
                cat_StoreFront="0"
                cat_ShowParent="1"
                cat_DefaultItem="All categories"
                cat_SelectedItems="0,"
                cat_ExcItems=""
				If len(pcv_strExcSubs)>0 Then
                	cat_ExcSubs="1"
				Else
	                cat_ExcSubs="0"			
				End If
                cat_ExcBTOItems="1"
                cat_EventAction=""
                %>
                <!--#include file="../includes/pcCategoriesList.asp"-->
        		<%call pcs_CatList()%>
                <div style="margin: 10px 0;">
                <input type="checkbox" name="incSubCats" value="1" class="clearBorder"> <strong>Include sub-categories</strong>. Select the parent category above and all of its sub-categories will be included.
                </div>
        	</td>
          </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"><hr></td>
        </tr>
        <tr>
            <td colspan="2">Other product selection settings:</td>
        </tr>	
        <tr>
            <td colspan="2"><input type="checkbox" name="excWCats" value="1" checked class="clearBorder"> Exclude wholesale-only categories</td>
        </tr>
        <tr>
            <td colspan="2"><input type="checkbox" name="excNFSPrds" value="1" class="clearBorder"> Exclude 'not for sale' products</td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"><a name="unique"></a></td>
        </tr>	
        <tr>
            <th colspan="2">Unique product identifiers</th>
        </tr>	
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer">Google Shopping now requires the use of <strong>unique product identifiers</strong> (<a href="http://www.google.com/support/merchants/bin/answer.py?answer=160161">learn more</a>). If unique identifiers are not available (e.g. custom products), you will need to <a href="http://www.google.com/support/merchants/bin/request.py?contact_type=unique_id_exemption">request an exemption</a>.</td>
        </tr>
        <tr>
            <td nowrap align="right" valign="top"><input type="radio" id="UniqueIdent" name="UniqueIdent" value="1" class="clearBorder"></td>
        	<td>Yes, I have to use unique product identifiers - <a href="http://wiki.earlyimpact.com/productcart/managing_search_fields#mapping_custom_search_fields_to_export_fields" target="_blank" class="pcSmallText">Learn how to associate these values to products</a></td>
        </tr>
        <tr>
            <td nowrap align="right" valign="top"><input type="radio" id="UniqueIdent" name="UniqueIdent" value="2" class="clearBorder"></td>
        	<td>Not needed, I have been granted an exemption - <a href="http://www.google.com/support/merchants/bin/request.py?contact_type=unique_id_exemption" target="_blank" class="pcSmallText">Request an exemption</a>.</td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <th colspan="2">Data feed file name and other settings</th>
        </tr>	
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
        	<td colspan="2">The file name must be a TEXT file (*.txt) and must <u>not</u> include the date. The file will be saved to the <strong><%=scAdminFolderName%></strong> folder.</td>
        </tr>
        <tr>
            <td nowrap align="right" width="20%">File name:</td>
            <td width="80%">
                <input name="pcv_filename" type="text" size="40" value="Products.txt">
                <input name="pcv_rootCat" type="hidden" value="Our Products">
            </td>
        </tr>
        <tr>
          <td valign="top" align="right">Currency:</td>
          <td>
            <select name="idCurrency" id="idCurrency">
              <option value="USD" selected>USD</option>
              <option value="AUD">AUD</option>
              <option value="CAD">CAD</option>
              <option value="CHF">CHF</option>
              <option value="CZK">CZK</option>
              <option value="DKK">DKK</option>
              <option value="EUR">EUR</option>
              <option value="GBP">GBP</option>
              <option value="GRD">GRD</option>
              <option value="HKD">HKD</option>
              <option value="HUF">HUF</option>
              <option value="INR">INR</option>
              <option value="MXN">MXN</option>
              <option value="MYR">MYR</option>
              <option value="NOK">NOK</option>
              <option value="NZD">NZD</option>
              <option value="PLN">PLN</option>
              <option value="SEK">SEK</option>
              <option value="SGD">SGD</option>
              <option value="THB">THB</option>
              <option value="TWD">TWD</option>
            </select>
            select the <em>Currency</em> that applies to your product prices.
          </td>
        </tr>
        
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
          <td valign="top" align="right">Product Condition:</td>
          <td>
            <select name="idCondition" id="idCondition">
              <option value="new" selected>new</option>
              <option value="used">used</option>
              <option value="refurbished">refurbished</option>
            </select>
          </td>
        </tr>
        
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
          <td valign="top" align="right">Brand:</td>
          <td>
            We will use the brand associated with each product. If you are not using <em>Brands</em> in your store, enter a generic brand here (e.g. your company name):<br>
            <input type="text" name="CustomBrand" value="<%=scCompanyName%>" size="30">
          </td>
        </tr>
        
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
          <td valign="top" align="right">Expiration Date:</td>
          <td>The date on which the product will no longer be available (defaults to 30 days from now):<br>
			<%strExpDate=Date()+30%>
            <input type="text" name="ExpirationDate" value="<%=Year(strExpDate) & "-" & FixDate(Month(strExpDate)) & "-" & FixDate(Day(strExpDate))%>" size="30"> Format: YYYY-MM-DD
          </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
		 <tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="includeApparel" value="1" checked class="clearBorder"></td>
        	<td>Include Google Shopping Apparel Product Variants<br />
		    <span class="pcSmallText"><i>Such as: gender, age group, size, color,etc.</i></span></td>
        </tr>
		<tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="includeWeight" value="1" class="clearBorder"></td>
        	<td>Include shipping weight<br />
		    <span class="pcSmallText">Note: you need to <a href="http://www.google.com/support/merchants/bin/answer.py?answer=160162" target="_blank">setup shipping options</a> in your Google Merchant account</span></td>
        </tr>        
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>	
		 <tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="showReports" value="1" class="clearBorder"></td>
        	<td>Display reports about included and excluded products<br />
		    <span class="pcSmallText">Note: This option can slow down the data feed generation process, and it only works when exporting all categories</span></td>
        </tr>
        <tr>					
            <td colspan="2" valign="top">
                <hr color="#e1e1e1" width="100%" size="1" noshade>
            </td>
        </tr>

        <tr>
            <td colspan="2" align="center">
                <input type="submit" name="Submit" value=" Generate Google Shopping Data Feed" class="submit2">
                <%
                '// Loading Window
                '	>> Call Method with OpenHS();
                response.Write(pcf_ModalWindow("Gathering information for Google Shopping Data Feed...", "GoogleBase2", 300))
                %>
                <input type="button" name="back" value="Back" onClick="javascript:history.back()">
            </td>
        </tr>
        </table>
    </form>
	<%
END IF
'/////////////////////////////////////////////////////////////////////
'// END: DISPLAY FORM
'/////////////////////////////////////////////////////////////////////




'/////////////////////////////////////////////////////////////////////
'// START: FUNCTIONS
'/////////////////////////////////////////////////////////////////////
Function FixDate(datevalue)
	Dim Tmp1,Tmp2
	Tmp2=datevalue
	Tmp1=Cstr(Tmp2)
	if cint(Tmp1)<10 then
		FixDate="0" & Tmp1
	else
		FixDate="" & Tmp1
	end if
End Function


function ClearHTMLTagsGB(strHTML2, intWorkFlow2)

	'Variables used in the function
	
	dim regEx2, strTagLess2
	
	'---------------------------------------
	strTagLess2 = strHTML2
	'Move the string into a private variable
	'within the function
	'---------------------------------------
	
	'---------------------------------------
	'NetSource Commerce codes
	IF strTagLess2<>"" THEN
		strTagLess2=replace(strTagLess2,"<br>"," ")
		strTagLess2=replace(strTagLess2,"<BR>"," ")
		strTagLess2=replace(strTagLess2,"<p>"," ")
		strTagLess2=replace(strTagLess2,"<P>"," ")
		strTagLess2=replace(strTagLess2,"</p>"," ")
		strTagLess2=replace(strTagLess2,"</P>"," ")
		strTagLess2=replace(strTagLess2,vbcrlf," ")
		strTagLess2=replace(strTagLess2,"™","&trade;")
		strTagLess2=replace(strTagLess2,"©","&copy;")
		strTagLess2=replace(strTagLess2,"®","&reg;")
		strTagLess2=replace(strTagLess2,chr(9)," ")
		strTagLess2=replace(strTagLess2,VbCr," ")
		strTagLess2=replace(strTagLess2,VbLf," ")
		strTagLess2=replace(strTagLess2,"&nbsp;"," ")
	END IF
	'Modify the string to a friendly ONLY 1 LINE string
	'---------------------------------------
	
	IF strTagLess2<>"" THEN

	'regEx2 initialization

		'---------------------------------------
		set regEx2 = New regExp 
		'Creates a regEx2p object		
		regEx2.IgnoreCase = True
		'Don't give frat about case sensitivity
		regEx2.Global = True
		'Global applicability
		'---------------------------------------
		'Phase I
		'	"bye bye html tags"
		if intWorkFlow2 <> 1 then
			'---------------------------------------
			regEx2.Pattern = "<[^>]*>"
			'this pattern mathces any html tag
			strTagLess2 = regEx2.Replace(strTagLess2, "")
			'all html tags are stripped
			'---------------------------------------
		end if

		'Phase II
		'	"bye bye rouge leftovers"
		'	"or, I want to render the source"
		'	"as html."

		'---------------------------------------
		'We *might* still have rouge < and > 
		'let's be positive that those that remain
		'are changed into html characters
		'---------------------------------------	
		if intWorkFlow2 > 0 and intWorkFlow2 < 3 then
			regEx2.Pattern = "[<]"
			'matches a single <
			strTagLess2 = regEx2.Replace(strTagLess2, "&lt;")

			regEx2.Pattern = "[>]"
			'matches a single >
			strTagLess2 = regEx2.Replace(strTagLess2, "&gt;")
			'---------------------------------------
		end if
		
		'Clean up
		'---------------------------------------
		set regEx2 = nothing
		'Destroys the regEx2p object
		'---------------------------------------	
	END IF 'vefiry strTagLess2 (null strings)
	
	'---------------------------------------
	'NetSource Commerce codes
	IF strTagLess2<>"" THEN
		strTagLess2=replace(strTagLess2,chr(34),"&quot")
		strTagLess2=trim(strTagLess2)
		do while instr(strTagLess2,"  ")>0
			strTagLess2=replace(strTagLess2,"  "," ")
		loop
	END IF
	'Remove white spaces
	'---------------------------------------
	
	'---------------------------------------
	ClearHTMLTagsGB = strTagLess2
	'The results are passed back
	'---------------------------------------
end function

Function pcf_FillByName(fileid,sfn)
	call openDb()
	Dim pcv_strFillByName
	pcv_strFillByName=""
	query="SELECT pcSearchData.pcSearchDataName "
	query = query & "FROM pcSearchFields_Products INNER JOIN "
	query = query & "( "
	query = query & "pcSearchData INNER JOIN "
	query = query & "( "
	query = query & "pcSearchFields INNER JOIN pcSearchFields_Mappings ON pcSearchFields.idSearchField  = pcSearchFields_Mappings.idSearchField "
	query = query & ") "
	query = query & "ON pcSearchData.idSearchField = pcSearchFields.idSearchField "				
	query = query & ") "
	query = query & "ON pcSearchFields_Products.idSearchData = pcSearchData.idSearchData "
	query=query&"WHERE pcSearchFields_Products.idproduct=" & intIdProduct & " "
	query=query&"AND pcSearchFields_Mappings.pcSearchFieldsColumn='" & sfn & "' "
	query=query&"AND pcSearchFields_Mappings.pcSearchFieldsFileID='" & fileid & "';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_strFillByName=rs("pcSearchDataName")
	end if
	pcf_FillByName = pcv_strFillByName
	set rs=nothing
	call closeDb()
End Function

'/////////////////////////////////////////////////////////////////////
'// END: FUNCTIONS
'/////////////////////////////////////////////////////////////////////
%><!--#include file="AdminFooter.asp"-->