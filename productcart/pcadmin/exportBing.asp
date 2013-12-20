<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=false
Server.ScriptTimeout = 54000%>
<% pageTitle="Generate Bing Shopping data feed" %>
<% Section="genRpts" %>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
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

dim rstemp, rstemp2, conntemp, mysql, strtext, strtext1, File1, pcv_rootCat, strParent, intIdProduct, pcIntIncludeWeight

'/////////////////////////////////////////////////////////////////////
'// START: ADD DETAILS
'/////////////////////////////////////////////////////////////////////
IF request("action")="add" then

	call opendb()
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Global Attributes
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'// Filename
	File1=request("pcv_filename")
	pcv_rootCat=request("pcv_rootCat")
	
	'// Condition
	strCondition=request("idCondition")
	if strCondition = "" then
		strCondition = "New"
	end if
	
	'// Brand
	strCustomBrand=request("CustomBrand")
	
	'// Include weight
	pcIntIncludeWeight=request("includeWeight")
	if not validNum(pcIntIncludeWeight) then pcIntIncludeWeight=0

	'// Get Path Info
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
	pcv_IdCategory=request("idcategory")	
	if pcv_IdCategory="" then
		pcv_IdCategory=0
	end if
	pcv_IdCategory=trim(pcv_IdCategory)	
	pcList1=split(pcv_IdCategory,",")	
	%>
    <form name="form2" method="post" action="exportBing.asp?action=gen" class="pcForms">
        <table class="pcCPcontent">
            <tr> 
                <td>
                    <div style="padding: 4px">
                        Microsoft Bing Shopping recommends that you specify this additional information, if available.
                    </div> 
                    <% if request("showDetails")="1" then %>        	
                        <div class="pcCPsearch" style="padding: 4px">
                            <img src="images/note.gif" align="left" vspace="8" hspace="4">
                            <strong>Use your existing custom search fields to auto fill the MPN, UPC, or ISBN.</strong> For example, you can map your custom field named "MPN" to the export column below named &quot;MPN&quot;. Then ProductCart will use the value for &quot;MPN&quot; saved with each product to auto fill the fields below. <a href="ManageSearchFields.asp">Manage Custom Fields</a>&nbsp;|&nbsp;<a href="SearchFields_Export.asp?export=f">Add/Modify Mappings</a>&nbsp;|&nbsp;<a href="http://wiki.earlyimpact.com/productcart/managing_search_fields#mapping_custom_search_fields_to_export_fields" target="_blank">Help</a>
                        </div>
					<% 
                    end if		 
                    Count=0
	
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Start:  Do For Categories
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						tmpCats=""
						For lk=lbound(pcList1) to ubound(pcList1)
							'// Filter By Category
							If trim(pcList1(lk))<>"0" then
								if tmpCats<>"" then
									tmpCats=tmpCats & ","
								end if
								tmpCats=tmpCats & trim(pcList1(lk))
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
						query="SELECT categories_products.idcategory,categories.categoryDesc,products.idProduct,products.description,products.serviceSpec,products.price,products.imageUrl,products.sku,products.IDBrand,products.details,products.sDesc FROM categories_products,categories,products WHERE ((products.price>0) OR ((products.price=0) AND (products.serviceSpec=-1))) AND products.active = -1 " & ExcWCatsStr &  ExcNFSPrds & " and products.removed=0 and products.configOnly=0 and products.idproduct=categories_products.idproduct AND categories.idcategory=categories_products.idcategory " & query1 & " order by categories_products.idcategory asc;"
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
						
						Count2=0
						
						if Count1>-1 then
							if request("showDetails")="1" then			
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start:  Do For Each Product In Category
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							For k=0 to Count1		
								intIdProduct=pcArray(2,k)
								strPrdName=pcArray(3,k)
								strPrdSKU=pcArray(7,k)		
								
								IF (Session("add_" & intIdProduct)<>"1") or (pcv_IdCategory="0") or (pcv_IdCategory="0,") THEN
									
									Count2=Count2+1
											
									if (pcv_IdCategory<>"0") and (pcv_IdCategory<>"0,") then
										Session("add_" & intIdProduct)="1"
									end if
							
									query="SELECT pcExpG_MPN,pcExpG_UPC,pcExpG_ISBN FROM pcExportGoogle WHERE idproduct=" & intIdProduct & ";"
									set rs=connTemp.execute(query)
									tmpMPN=""
									tmpUPC=""
									tmpISBN=""
									if not rs.eof then
										tmpMPN=rs("pcExpG_MPN")
										tmpUPC=rs("pcExpG_UPC")
										tmpISBN=rs("pcExpG_ISBN")
									end if
									set rs=nothing
									
									if Count=0 and Count2=1 then%>	
                                    <table class="pcCPcontent" width="100%">
					                  <tr>
					                    <td colspan="4" class="pcCPspacer"></td>
					                  </tr>					
									<tr>
										<th width="40%">Product</th>
										<th nowrap="nowrap">MPN</th>
										<th nowrap="nowrap">UPC</th>
										<th nowrap="nowrap">ISBN</th>
									</tr>
					                  <tr>
					                    <td colspan="4" class="pcCPspacer"></td>
					                  </tr>	
									<%end if
									Count=Count+1%>
									<tr>
										<td valign="top"><%=strPrdName%><br><i>(<%=strPrdSKU%>)</i> 
										<input type="hidden" name="idprd<%=Count%>" value="<%=intIdProduct%>"></td>
										<td valign="top"><input type="text" name="MPN<%=Count%>" size="20" value="<%=pcf_FillByName("f","MPN",tmpMPN)%>"></td>
										<td valign="top"><input type="text" name="UPC<%=Count%>" size="20" value="<%=pcf_FillByName("f","UPC",tmpUPC)%>"></td>
										<td valign="top"><input type="text" name="ISBN<%=Count%>" size="20" value="<%=pcf_FillByName("f","ISBN",tmpISBN)%>"></td>
									</tr>					
									<%
									if Count=0 and Count2=1 then
									%>
									</table>
									<%
									end if
								END IF '// IF (Session("add_" & intIdProduct)<>"1") or (pcv_IdCategory="0") or (pcv_IdCategory="0,") THEN
								
							Next
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start:  Do For Each Product In Category
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							else
								pcv_strSkipDetails = 1
								Count=Count+Count1+1
							end if
						end if '// if Count1>-1 then
						
					If pcv_strSkipDetails = 1 Then 
						%>
                        <table class="pcCPcontent" width="100%">
                            <tr>
                                <td valign="top" colspan="4">
                                    <div class="pcCPsearch" style="padding: 4px">
                                        <img src="images/note.gif" align="left" vspace="8" hspace="4">
                                        You have chosen not to map additional fields.  The MPN, UPC, and ISBN columns will be left blank. To enter the MPN, UPC, and/or ISBN please use the back button and check the "Map Additional Fields" option.			
                                    </div> 
                                </td>
                            </tr>
                        </table>
					  <%
                    End If
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    '// End:  Do For Each Category
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
                    call closedb()
                    
                    Function pcf_FillByName(fileid,sfn,orig)
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
                        else
                            pcv_strFillByName=orig		
                        end if
                        pcf_FillByName = pcv_strFillByName
                        set rs=nothing
                    End Function
                    %>
	
					<% If Count=0 Then %>
                        <table class="pcCPcontent">
                            <tr> 
                                <td>
                                    <div class="pcCPmessage">
                                        Cannot find any products. Please run a new search.
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <a href="exportBing.asp">Return to Bing Shopping data feed generator</a>
                                </td>
                            </tr>
                        </table>
                    <% Else %>
                        <table class="pcCPcontent">
                            <tr>
                                <td colspan="4" class="pcCPspacer"></td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <input type="submit" name="Submit1" value="Generate Bing Shopping Data Feed" onClick="pcf_Open_BingShopping();" class="submit2">
                                    &nbsp;
                                    <input type="button" name="Back" value="Back" onClick="document.location.href='exportBing.asp'">
									<%
									'// Loading Window
									'	>> Call Method with OpenHS();
									response.Write(pcf_ModalWindow("Generating Bing Shopping Data Feed...", "BingShopping", 300))
									%>
																		
                                    <input type="hidden" name="pcv_filename" value="<%=request("pcv_filename")%>">
                                    <input type="hidden" name="pcv_rootCat" value="<%=request("pcv_rootCat")%>">
                                    <input type="hidden" name="idCondition" value="<%=request("idCondition")%>">
                                    <input type="hidden" name="CustomBrand" value="<%=request("CustomBrand")%>">
                                    <input type="hidden" name="idcategory" value="<%=request("idcategory")%>">
                                    <input type="hidden" name="excNFSPrds" value="<%=request("excNFSPrds")%>">
                                    <input type="hidden" name="excWCats" value="<%=request("excWCats")%>">
                                    <input type="hidden" name="showDetails" value="<%=request("showDetails")%>">
									<input type="hidden" name="showReports" value="<%=request("showReports")%>">
									<input type="hidden" name="includeWeight" value="<%=request("includeWeight")%>">
                                    <input type="hidden" name="Count" value="<%=Count%>">
                                <td>
                            </tr>
                        </table>	
                    <% End If %>
    
				</td>
			</tr>
		</table>
	</form>
	<%
END IF
'/////////////////////////////////////////////////////////////////////
'// END: ADD DETAILS
'/////////////////////////////////////////////////////////////////////



'/////////////////////////////////////////////////////////////////////
'// START: GEN FILE
'/////////////////////////////////////////////////////////////////////
IF request("action")="gen" THEN
	
	call opendb()

	Count=request("Count")
	if Count<>"" then
	For i=1 to Count
		tmpID=request("idprd" & i)
		tmpMPN=request("MPN" & i)
		tmpUPC=request("UPC" & i)
		tmpISBN=request("ISBN" & i)
		if trim(tmpID)<>"" then
			query="DELETE FROM pcExportGoogle WHERE idproduct=" & tmpID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			if tmpMPN<>"" then
				tmpMPN=replace(tmpMPN,"'","''")
			end if
			if tmpUPC<>"" then
				tmpUPC=replace(tmpUPC,"'","''")
			end if
			if tmpISBN<>"" then
				tmpISBN=replace(tmpISBN,"'","''")
			end if
			if tmpMPN & tmpUPC & tmpISBN<>"" then
				query="INSERT INTO pcExportGoogle (idproduct,pcExpG_MPN,pcExpG_UPC,pcExpG_ISBN) VALUES (" & tmpID & ",'" & tmpMPN & "','" & tmpUPC & "','" & tmpISBN & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
		end if
	Next
	end if

	'// Get Filename
	File1=request("pcv_filename")
	pcv_rootCat=request("pcv_rootCat")
	
	'// Include weight
	pcIntIncludeWeight=request("includeWeight")
	if not validNum(pcIntIncludeWeight) then pcIntIncludeWeight=0
	
	'// Get Condition
	strCondition=request("idCondition")
	if strCondition = "" then
		strCondition = "New"
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
	pcv_IdCategory=request("idcategory")	
	if pcv_IdCategory="" then
	pcv_IdCategory=0
	end if
	pcv_IdCategory=trim(pcv_IdCategory)	
	pcList1=split(pcv_IdCategory,",")
	
	'// Set Header Rows - Bing Shopping
    strtext1="MerchantProductID" & chr(9) & "Title" & chr(9) & "Brand" & chr(9) & "MPN" & chr(9) & "UPC" & chr(9) & "ISBN" & chr(9) & "SKU" & chr(9) & "ProductURL" & chr(9) & "Price" & chr(9) & "Availability" & chr(9) & "Description" & chr(9) & "ImageURL" & chr(9) & "MerchantCategory" & chr(9)
	if pcIntIncludeWeight=1 then
		strtext1 = strtext1 & "ShippingWeight" & chr(9) & "Condition" &  vbcrlf
		else
		strtext1 = strtext1 & "Condition" & vbcrlf		
	end if
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
		query="SELECT categories_products.idcategory,categories.categoryDesc,categories.pccats_BreadCrumbs,products.idProduct,products.description,products.serviceSpec,products.price,products.imageUrl,products.sku,products.stock,products.weight,products.IDBrand,products.noStock,products.pcProd_BackOrder,products.pcProd_BTODefaultPrice,products.details,products.sDesc FROM categories_products,categories,products WHERE ((products.price>0) OR ((products.price=0) AND (products.serviceSpec=-1) AND (products.pcProd_BTODefaultPrice>0))) AND products.active = -1 " & ExcWCatsStr &  ExcNFSPrds & " and products.removed=0 and products.configOnly=0 and products.idproduct=categories_products.idproduct AND categories.idcategory=categories_products.idcategory " & query1 & " order by categories_products.idcategory asc,products.description;"
		if request("excWCats")="1" then
		querytest="SELECT DISTINCT products.idproduct FROM categories_products,categories,products WHERE ((products.price>0) OR ((products.price=0) AND (products.serviceSpec=-1) AND (products.pcProd_BTODefaultPrice>0))) AND products.active = -1 AND categories.pccats_RetailHide=1" &  ExcNFSPrds & " and products.removed=0 and products.configOnly=0 and products.idproduct=categories_products.idproduct AND categories.idcategory=categories_products.idcategory " & query1 & ";"
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
		' 9 = inventory level (stock)
		' 10 = weight in ounces
		' 11 = brand ID
		' 12 = disregard stock property
		' 13 = backorder property
		' 14 = BTO default price
		' 15 = long description
		' 16 = short description

		
		query="SELECT idproduct,pcExpG_MPN,pcExpG_UPC,pcExpG_ISBN FROM pcExportGoogle;"
		set rs=connTemp.execute(query)
		tmpMPN=""
		tmpUPC=""
		tmpISBN=""
		CountGF=-1
		if not rs.eof then
			pcArr1=rs.getRows()
			CountGF=ubound(pcArr1,2)
		end if
		set rs=nothing
		
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
		
		tmpAddedPrds="*"
		
		set StringBuilderObj = new StringBuilder
		
		if Count1>-1 then
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start:  Do For Each Product In Category
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			For k=0 to Count1
				
				intIdProduct=pcArray(3,k)
				
				IF InStr(tmpAddedPrds,"*" & intIdProduct & "*")=0 THEN
				
				'IF session("tmp" & intIdProduct)<>"1" THEN
				
					tmpAddedPrds=tmpAddedPrds & intIdProduct & "*"
					'session("tmp" & intIdProduct)=1
					
					'If it's a BTO product, check if there is a Default Price and use it
					If (pcArray(5,k)=-1) then
						If pcArray(14,k)>"0" then
							pcArray(6,k) = pcArray(14,k)
						End if
					ENd if
					dblPrice=pcArray(6,k)
					
					if isNull(dblPrice) OR dblPrice="" then
						dblPrice=0
					end if
					
					IF dblPrice>0 THEN
				
						tmpMPN=""
						tmpUPC=""
						tmpISBN=""
				
						For m=0 to CountGF
							if Clng(intIdProduct)=Clng(pcArr1(0,m)) then
								tmpMPN=pcArr1(1,m)
								tmpUPC=pcArr1(2,m)
								tmpISBN=pcArr1(3,m)
								exit for
							end if
						Next
				
						tmpIDCategory=pcArray(0,k)
						
						strPrdCategory=pcArray(1,k) ' Category description
						
						'// Load breadcrumb, if it exists
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
							strPrdCategory=strBreadCrumb
							End If
								
		
								
						'// END - Load breadcrumb
						
						strProductName=pcArray(4,k)
						strProductName1=pcArray(4,k)
						strBTO=pcArray(5,k)
							if IsNull(strBTO) or trim(strBTO)="" then strBTO=0
						strProductImg=pcArray(7,k)
						strSKU=pcArray(8,k)
						strStock=pcArray(9,k)
						strWeight=pcArray(10,k)
							if IsNull(strWeight) then strWeight=0
							if not validNum(strWeight) then strWeight=0
							if strWeight>0 then
								strWeight=round((strWeight/16),2)
							end if
						strIDBrand=pcArray(11,k)
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
						strDisregardStock=pcArray(12,k)
							if IsNull(strDisregardStock) or trim(strDisregardStock)="" then strDisregardStock=0
						strBackOrderOK=pcArray(13,k)
							if IsNull(strBackOrderOK) or trim(strBackOrderOK)="" then strBackOrderOK=0
						strProductDesc=pcArray(15,k)
						strShortProductDesc=pcArray(16,k)
								
						strParent=""
				
						'// SEO Links
						'// Build Product Link
						if scSeoURLs<>1 then
							strProductURL=SPathInfo & "pc/viewPrd.asp?idproduct=" & intIdProduct & "&idcategory=" & tmpIDCategory
						else
							strProductURL=SPathInfo & "pc/" & removeChars(strProductName1) & "-" & tmpIDCategory & "p" & intIdProduct & ".htm"
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
						
						if strProductImg<>"" and strProductImg <> "no_image.gif" then
							strProductImgURL=SPathInfo & "pc/catalog/" & strProductImg
						else
							strProductImgURL=""
						end if
						
						'// Determine availability
						if scOutOfStockPurchase=0 then
							strAvailability = "In Stock"
						else
							if CLng(strStock)<1 then
								If (scOutofStockPurchase=-1 AND strBTO=0 AND strDisregardStock=0 AND strBackOrderOK=1) OR (strBTO<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND strDisregardStock=0 AND strBackOrderOK=1) Then
									strAvailability = "Back-Order"
								else
									strAvailability = "Out of Stock"		
								end if
							else
								strAvailability = "In Stock"
							end if
						end if
						'// END - Determine Availability				
						
						
						Count3=Count3+1	
						'// Product weight
						if pcIntIncludeWeight=1 then
							addStringWeight = "/*/" & strWeight
						else
							addStringWeight = ""
						end if
						
						StringBuilderObj.append strSKU & "/*/" & left(ClearHTMLTagsGB(strProductName,0),255) & "/*/" & strBrandName & "/*/" & tmpMPN & "/*/" & tmpUPC & "/*/" & tmpISBN & "/*/" & strSKU & "/*/" & strProductURL & "/*/" & dblPrice & "/*/" & strAvailability & "/*/" & left(ClearHTMLTagsGB(strProductDesc,0),5000) & "/*/" & strProductImgURL & "/*/" & left(ClearHTMLTagsGB(strPrdCategory,0),255) & addStringWeight & "/*/" & strCondition & "/-/"

												
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
				END IF
		
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
	tmpAddedPrds=""
	response.write "<script>document.getElementById('ea').style.display='none';</script>"%>
	<table class="pcCPcontent">
		<tr> 
			<td>
			The Bing Shopping data feed was created successfully.
				<br>
				<br>
				<b><%=Count%></b> products have been exported succefully to the Bing Shopping data feed: <a href="<%=File1%>"><%=File1%></a>. To download the file, either right-click on the file name and select '<strong>Save Target As...</strong>' or '<strong>Save Link as...</strong>' or FTP into the <em><%=scAdminFolderName%></em> folder and download it. <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_bing_shopping_file" target="_blank">Learn more about submitting your data feed to Bing Shopping</a>.
				<br>
				<br>
				<br>
				<form class="pcForms">
					<input type=button name=back value="Create Another File" onclick="location='exportBing.asp'">&nbsp;
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
				<b><%=Count%></b> products included in the Bing Shopping data feed<br />
				<%call opendb()
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.active = 0 AND products.removed=0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products are inactive and not included in the Bing Shopping data feed<br />"
					end if
				end if
				set rsQ=nothing
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.removed <> 0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products are designated as removed from the store and not included in the Bing Shopping data feed<br />"
					end if
				end if
				set rsQ=nothing
				
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.active <> 0 AND products.removed = 0 AND products.configOnly<>0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products are BTO items and not included in the Bing Shopping data feed<br />"
					end if
				end if
				set rsQ=nothing
				
				query="SELECT Count(*) As tmpTotal FROM products WHERE products.active <> 0 AND products.removed = 0 AND products.configOnly=0 AND price<=0 AND pcProd_BTODefaultPrice<=0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("tmpTotal")>"0" then
					Count=Count+Clng(rsQ("tmpTotal"))
					response.write rsQ("tmpTotal") & " products have a price equal to zero and not included in the Bing Shopping data feed<br />"
					end if
				end if
				set rsQ=nothing
				
				if request("excNFSPrds")="1" then
					query="SELECT Count(*) As tmpTotal FROM products WHERE products.active <> 0 AND products.removed = 0 AND products.configOnly=0 AND (price>0 OR pcProd_BTODefaultPrice>0) AND products.formQuantity<>0;"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						if rsQ("tmpTotal")>"0" then
						Count=Count+Clng(rsQ("tmpTotal"))
						response.write rsQ("tmpTotal") & " products are 'Not for Sale' products and not included in the Bing Shopping data feed<br />"
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
							response.write CountB & " products are assigned to wholesale-only categories and not included in the Bing Shopping data feed<br />"
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
					response.write "<i>Note: there appear to be " & CountTotal-Count & " orphaned products that were not included in the Bing Shopping data feed. You can <a href=""javascript:location='srcFreePrds.asp';"">view them</a> here. Assign them to a category to include them in the data feed.</i>"
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
		if (theForm.idcategory.value == "")
		{
			alert("Please select at least one category.");
			theForm.idcategory.focus();
			return (false);
		}
	return (true);
	}
	//-->
	</script>
    <form name="form1" method="post" action="exportBing.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
        <input name="pcv_filename" type="hidden" value="bingshopping.txt">
        <input name="pcv_rootCat" type="hidden" value="Our Products">
        <table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>	
        <tr>
            <th colspan="2">Notes about generating a Bing Shopping Data Feed</th>
        </tr>	
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>	
        <tr>
            <td colspan="2">			
                <ul>
                <li>Review the User Guide for <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_bing_shopping_file" target="_blank">information about this feature</a>.</li>
				<li>Take advantage of Custom Search Fields to pre-fill the MPN, UPC, and ISBN codes (if applicable). <a href="http://wiki.earlyimpact.com/productcart/managing_search_fields#mapping_custom_search_fields_to_export_fields" target="_blank">More information</a>.</li>
                </ul>
            </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>	
        <tr>
            <th colspan="2">First Step: General Settings</th>
        </tr>	
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td valign="top" align="right">Category List:</td>
            <td>
                Please <strong>select the categories</strong> that you would like to <strong>include</strong> in the Bing Shopping data feed. Press down the CTRL key on your keyboard to select multiple entries.<br>
                <br>
                <%
                cat_DropDownName="idcategory"
                cat_Type="1"
                cat_DropDownSize="5"
                cat_MultiSelect="1"
                cat_ExcBTOHide="1"
                cat_StoreFront="0"
                cat_ShowParent="1"
                cat_DefaultItem="All categories"
                cat_SelectedItems="0,"
                cat_ExcItems=""
                cat_ExcSubs="0"
                cat_ExcBTOItems="1"
                cat_EventAction=""
                %>
                <!--#include file="../includes/pcCategoriesList.asp"-->
                <%call pcs_CatList()%>							
                </td>
        </tr>        
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
          <td valign="top" align="right" nowrap>Product Condition:</td>
          <td>
            <select name="idCondition" id="idCondition">
              <option value="New" selected>New</option>
              <option value="Used">Used</option>
              <option value="Collectable">Collectable</option>
              <option value="Open Box">Collectable</option>
              <option value="Refurbished">Refurbished</option>
              <option value="Remanufactured">Remanufactured</option>
            </select>
          </td>
        </tr>
        
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
          <td valign="top" align="right">Brand Name:</td>
          <td>
            We will use the brand name of each product in your store database automatically. If you did not specify brand name of products. Please enter here:<br>
            <input type="text" name="CustomBrand" value="<%=scCompanyName%>" size="30">
          </td>
        </tr>
        
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>	
        <tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="includeWeight" value="1" checked class="clearBorder"></td>
        	<td>Include product weight<br />This must be included if you indicated that shipping costs are determined by weight in the Bing Shopping Merchant Center.</td>
        </tr>
        <tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="excWCats" value="1" checked class="clearBorder"></td>
        	<td>Exclude wholesale categories</td>
        </tr>
        <tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="excNFSPrds" value="1" class="clearBorder"></td>
        	<td>Exclude 'not for sale' products</td>
        </tr>   
        <tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="showDetails" value="1" checked class="clearBorder"></td>
        	<td>Map Additional Fields (MPC, UPC, ISBN)<br />This option should not be used with large exports &gt;1000 products</td>
        </tr>
		 <tr>
            <td nowrap align="right" valign="top"><input type="checkbox" name="showReports" value="1" checked class="clearBorder"></td>
        	<td>Display reports about included and excluded products<br />
			<i>Note: This option can slow down the process and it only works when exporting all categories</i></td>
        </tr>
        <tr>					
            <td colspan="2" valign="top">
                <hr color="#e1e1e1" width="100%" size="1" noshade>
            </td>
        </tr>

        <tr>
            <td colspan="2" align="center">
                <input type="submit" name="Submit" value=" Generate Bing Shopping Data Feed" onClick="pcf_Open_BingShopping2();" class="submit2">
                <%
                '// Loading Window
                '	>> Call Method with OpenHS();
                response.Write(pcf_ModalWindow("Gathering information for Bing Shopping Data Feed...", "BingShopping2", 300))
                %>
                &nbsp;
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

'/////////////////////////////////////////////////////////////////////
'// END: FUNCTIONS
'/////////////////////////////////////////////////////////////////////
%><!--#include file="AdminFooter.asp"-->