<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true
Server.ScriptTimeout = 5400%>
<% pageTitle="Generate Bing Cashback Product Data Feed" %>
<% Section="genRpts" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
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

dim rstemp, rstemp2, conntemp, mysql, strtext, File1, pcv_rootCat, strParent, intIdProduct

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
	
	'// Currency (always will default to USD with Cashback)
	strCurrency=request("idCurrency")
	if strCurrency = "" then
		strCurrency = "USD"
	end if
	
	'// Condition
	strCondition=request("idCondition")
	if strCondition = "" then
		strCondition = "New"
	end if

	'// Expiration Date
	strExpDate=request("ExpirationDate")
	if strExpDate="" then
		strExpDate=Year(Date())+1 & "-" & FixDate(Month(Date())) & "-" & FixDate(Day(Date()))
	end if
	
	'// Brand
	strCustomBrand=request("CustomBrand")
	
	'// Commission
	strCommission=request("Commission")
	
	'// Stock
	strStock=request("Stock")
	
	'// Shipping
	strShipping=request("Shipping")

	'// Path Info
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
    <form name="form2" method="post" action="exportCashBack.asp?action=gen" class="pcForms">
        <table class="pcCPcontent">
            <tr> 
                <td>
                	MPN, UPC, and ISBN information, when applicable, is required. As mentioned on the previous page, you can use Custom Search Fields to automatically fill-in this information (<a href="ManageSearchFields.asp">Manage Custom Fields</a>&nbsp;|&nbsp;<a href="SearchFields_Export.asp?export=c">Add/Modify Mappings</a>). There are many products for which this information is not available. If that is the case, you can leave them blank.
                    <%  
                    Count=0
	
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Start:  Do For Each Category
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					For lk=lbound(pcList1) to ubound(pcList1)
					
						'// Filter By Category
						If trim(pcList1(lk))<>"0" then		
							query1=" AND categories_products.idcategory=" & trim(pcList1(lk)) & " "		
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
							
									query="SELECT pcExpCB_MPN,pcExpCB_UPC,pcExpCB_ISBN,pcExpCB_SHIPPING,pcExpCB_COMMISSION FROM pcExportCashback WHERE idproduct=" & intIdProduct & ";"
									set rs=connTemp.execute(query)
									tmpMPN=""
									tmpUPC=""
									tmpISBN=""
									tmpSHIPPING=""
									tmpCOMMISSION=""
									if not rs.eof then
										tmpMPN=rs("pcExpCB_MPN")
										tmpUPC=rs("pcExpCB_UPC")
										tmpISBN=rs("pcExpCB_ISBN")
										tmpSHIPPING=rs("pcExpCB_SHIPPING")
										tmpCOMMISSION=rs("pcExpCB_COMMISSION")
									end if
									set rs=nothing
									
									if Count=0 and Count2=1 then%>	
									<table class="pcCPcontent" width="100%">
					                  <tr>
					                    <td colspan="6" class="pcCPspacer"></td>
					                  </tr>		
									<tr>
										<th width="50%">Product</th>
										<th nowrap="nowrap" width="10%">MPN</th>
										<th nowrap="nowrap" width="10%">UPC</th>
										<th nowrap="nowrap" width="10%">ISBN</th>
					                   	<th nowrap="nowrap" width="10%">Shipping Cost</th>
					                    <th nowrap="nowrap" width="10%">Commission %</th>
									</tr>
					                  <tr>
					                    <td colspan="6" class="pcCPspacer"></td>
					                  </tr>	
									<%end if
									Count=Count+1%>
									<tr>
										<td valign="top"><%=strPrdName%><br><i>(<%=strPrdSKU%>)</i> 
										<input type="hidden" name="idprd<%=Count%>" value="<%=intIdProduct%>"></td>
										<td valign="top"><input type="text" name="MPN<%=Count%>" size="15" value="<%=pcf_FillByName("c","MPN",tmpMPN)%>"></td>
									  	<td valign="top"><input type="text" name="UPC<%=Count%>" size="15" value="<%=pcf_FillByName("c","UPC",tmpUPC)%>"></td>
										<td valign="top"><input type="text" name="ISBN<%=Count%>" size="15" value="<%=pcf_FillByName("c","ISBN",tmpISBN)%>"></td>
					                    <td valign="top"><input type="text" name="SHIPPING<%=Count%>" size="5" value="<%=pcf_FillByName("c","SHIPPING",tmpSHIPPING)%>"></td>
					                   	<td valign="top"><input type="text" name="COMMISSION<%=Count%>" size="5" value="<%=pcf_FillByName("c","COMMISSION",tmpCOMMISSION)%>"></td>
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
						
						If trim(pcList1(lk))="0" then
							exit for
						End if
					
					Next

					If pcv_strSkipDetails = 1 Then 
						%>
                        <table class="pcCPcontent" width="100%">
                            <tr>
                                <td valign="top" colspan="4">
                                    <div class="pcCPsearch" style="padding: 4px">
                                        <img src="images/note.gif" align="left" vspace="8" hspace="4">
                                        You have chosen to hide the "Additional Details" section.  The MPN, UPC, and ISBN will be left blank.
                                        To enter the MPN, UPC, or ISBN please use the back button and check the "Enter Additional Details" option.			
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
                                        Can not find any products. Please do a new search.
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <a href="exportCashBack.asp">Return to Bing Cashback Export page</a>
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
                                    <input type="submit" name="Submit1" value="Generate Bing Cashback Bulk Upload File" onClick="pcf_Open_GoogleBase();" class="submit2">
									<%
									'// Loading Window
									'	>> Call Method with OpenHS();
									response.Write(pcf_ModalWindow("Generate Bing Cashback Bulk Upload File...", "GoogleBase", 300))
									%>
																		
                                    <input type="hidden" name="pcv_filename" value="<%=request("pcv_filename")%>">
                                    <input type="hidden" name="pcv_rootCat" value="<%=request("pcv_rootCat")%>">
                                    <input type="hidden" name="Title1" value="<%=request("Title1")%>">
                                    <input type="hidden" name="Title2" value="<%=request("Title2")%>">
                                    <input type="hidden" name="Title3" value="<%=request("Title3")%>">
                                    <input type="hidden" name="idCurrency" value="<%=request("idCurrency")%>">
                                    <input type="hidden" name="idCondition" value="<%=request("idCondition")%>">
                                    <input type="hidden" name="Stock" value="<%=request("Stock")%>">                                    
                                    <input type="hidden" name="ExpirationDate" value="<%=request("ExpirationDate")%>">
                                    <input type="hidden" name="CustomBrand" value="<%=request("CustomBrand")%>">
                                    <input type="hidden" name="idcategory" value="<%=request("idcategory")%>">
                                    <input type="hidden" name="excNFSPrds" value="<%=request("excNFSPrds")%>">
                                    <input type="hidden" name="excWCats" value="<%=request("excWCats")%>">
                                    <input type="hidden" name="showDetails" value="<%=request("showDetails")%>">
                                    <input type="hidden" name="Commission" value="<%=request("Commission")%>">
                                    <input type="hidden" name="Shipping" value="<%=request("Shipping")%>">
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
		tmpSHIPPING=request("SHIPPING" & i)
		tmpCOMMISSION=request("COMMISSION" & i)
		if trim(tmpID)<>"" then
			query="DELETE FROM pcExportCashback WHERE idproduct=" & tmpID & ";"
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
			if tmpSHIPPING<>"" then
				tmpSHIPPING=replace(tmpSHIPPING,"'","''")
			end if
			if tmpCOMMISSION<>"" then
				tmpCOMMISSION=replace(tmpCOMMISSION,"'","''")
			end if
			if tmpMPN & tmpUPC & tmpISBN & tmpSHIPPING & tmpCOMMISSION<>"" then
				query="INSERT INTO pcExportCashback (idproduct,pcExpCB_MPN,pcExpCB_UPC,pcExpCB_ISBN,pcExpCB_SHIPPING,pcExpCB_COMMISSION) VALUES (" & tmpID & ",'" & tmpMPN & "','" & tmpUPC & "','" & tmpISBN & "','" & tmpSHIPPING & "','" & tmpCOMMISSION & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
		end if
	Next
	end if

	'// Get Filename
	File1=request("pcv_filename")
	pcv_rootCat=request("pcv_rootCat")
	
	'// Get Currency
	strCurrency=request("idCurrency")
	if strCurrency = "" then
		strCurrency = "USD"
	end if
	
	'// Get Condition
	strCondition=request("idCondition")
	if strCondition = "" then
		strCondition = "new"
	end if
	
	'// Get Date
	strExpDate=request("ExpirationDate")
	if strExpDate="" then
		strExpDate=Year(Date())+1 & "-" & FixDate(Month(Date())) & "-" & FixDate(Day(Date()))
	end if
	
	'// Get Brand
	strCustomBrand=request("CustomBrand")
	
	'// Commission
	strCommission=request("Commission")
	
	'// Stock
	strStock=request("Stock")
	
	'// Shipping
	strShipping=request("Shipping")

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
	
	'// No headers required for Cashback
	'strtext="title" & chr(9) & "description" & chr(9) & "link" & chr(9) & "image_link" & chr(9) & "id" & chr(9) & "price" & chr(9) & "condition" & chr(9) & "product_type" & chr(9) & "brand" & chr(9) & "expiration_date" & chr(9) & "currency"  & chr(9) & "mpn"  & chr(9) & "upc"  & chr(9) & "isbn" & vbcrlf
	
	
	Count=0
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start:  Do For Each Category
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	For lk=lbound(pcList1) to ubound(pcList1)
	
		'// Filter By Category
		If trim(pcList1(lk))<>"0" then		
			query1=" AND categories_products.idcategory=" & trim(pcList1(lk)) & " "		
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
		query="SELECT categories_products.idcategory,categories.categoryDesc,products.idProduct,products.description,products.serviceSpec,products.price,products.imageUrl,products.sku,products.IDBrand,products.details,products.sDesc,products.stock,products.pcProd_BackOrder,products.noStock,products.serviceSpec,products.weight,products.pcProd_BTODefaultPrice FROM categories_products,categories,products WHERE ((products.price>0) OR ((products.price=0) AND (products.serviceSpec=-1))) AND products.active = -1 " & ExcWCatsStr &  ExcNFSPrds & " and products.removed=0 and products.configOnly=0 and products.idproduct=categories_products.idproduct AND categories.idcategory=categories_products.idcategory " & query1 & " order by categories_products.idcategory asc;"
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
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start:  Do For Each Product In Category
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			For k=0 to Count1

				intIdProduct=pcArray(2,k)

				query="SELECT pcExpCB_MPN,pcExpCB_UPC,pcExpCB_ISBN,pcExpCB_SHIPPING,pcExpCB_COMMISSION FROM pcExportCashback WHERE idproduct=" & intIdProduct & ";"
				set rs=connTemp.execute(query)
				tmpMPN=""
				tmpUPC=""
				tmpISBN=""
				tmpSHIPPING=""
				tmpCOMMISSION=""
				if not rs.eof then
					tmpMPN=rs("pcExpCB_MPN")
					tmpUPC=rs("pcExpCB_UPC")
					tmpISBN=rs("pcExpCB_ISBN")
					tmpSHIPPING=rs("pcExpCB_SHIPPING")
					tmpCOMMISSION=rs("pcExpCB_COMMISSION")
				end if
				set rs=nothing


				'If it's a BTO product, the price is the Default BTO Price
				If (pcArray(4,k)=-1) then pcArray(5,k) = pcArray(16,k)

				
				'IF (Session(intIdProduct)<>"1") THEN
				
					Count2=Count2+1
				
					'Session(intIdProduct)="1"
				
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Category ID
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					tmpIDCategory=pcArray(0,k)
							
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Type
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strPrdType=pcArray(1,k)
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Name
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strProductName=pcArray(3,k)
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Price
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					dblPrice=pcArray(5,k)
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Image
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strProductImg=pcArray(6,k)
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product SKU
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strSKU=pcArray(7,k)
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Brand Name
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strIDBrand=pcArray(8,k)
					if strIDBrand<>"" and strIDBrand<>"0" then
						strBrandName=pcf_GetBrandName(strIDBrand)
					else
						strBrandName=strCustomBrand
					end if
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Description
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strProductDesc=pcArray(9,k)
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Short Desc
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strShortProductDesc=pcArray(10,k)
	
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product URL
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// SEO Links
					'// Build Product Link
					pcStrPrdLink=strProductName & "-" & tmpIDCategory & "p" & intIdProduct & ".htm"
					pcStrPrdLink=removeChars(pcStrPrdLink)
					strProductURL=SPathInfo & "pc/" & pcStrPrdLink & "?cashback=yes"
					if scSeoURLs<>1 then
						strProductURL=SPathInfo & "pc/viewPrd.asp?idproduct=" & intIdProduct & "&idcategory=" & tmpIDCategory & "&cashback=yes"
					end if
					'//
		
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Name
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strProductName=ClearHTMLTags2(strProductName,0)
					strProductName=replace(strProductName,chr(34),"&quot")
					strProductName=replace(strProductName,chr(9)," ")
					strProductName=replace(strProductName,VbCrLf," ")
					strProductName=replace(strProductName,VbCr," ")
					strProductName=replace(strProductName,VbLf," ")
					strProductName=replace(strProductName,"&nbsp;"," ")
					do while InStr(strProductName,"  ")>0
						strProductName=replace(strProductName,"  "," ")
					loop
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Title
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pcv_strTitle=""
					pcv_intCont = 0
					
					'// Title Part One
					Title1=request("Title1")
					Select Case Title1
						Case "Product": Title1 = strProductName
						Case "Brand": Title1 = strBrandName
						Case "SKU": Title1 = strSKU
						Case "Category": Title1= pcf_GetCategoryName(tmpIDCategory)
					End Select 
					If len(Title1)>0 Then
						pcv_strTitle = Title1
						pcv_intCont = 1		
					Else
						pcv_strTitle=strProductName
						pcv_intCont = 0
					End If
					
					'// Title Part Two
					If pcv_intCont = 1 Then	
						
						Title2=request("Title2")
						Select Case Title2
							Case "Product": Title2 = strProductName
							Case "Brand": Title2 = strBrandName
							Case "SKU": Title2 = strSKU
							Case "Category": Title2= pcf_GetCategoryName(tmpIDCategory)
						End Select 
						If len(Title2)>0 Then
							pcv_strTitle = pcv_strTitle & " >>> " & Title2	
						Else
							pcv_intCont = 0
						End If
					
					End If
					
					'// Title Part Three
					If pcv_intCont = 1 Then	
						
						Title3=request("Title3")
						Select Case Title3
							Case "Product": Title3 = strProductName
							Case "Brand": Title3 = strBrandName
							Case "SKU": Title3 = strSKU
							Case "Category": Title3= pcf_GetCategoryName(tmpIDCategory)
						End Select 
						If len(Title3)>0 Then
							pcv_strTitle = pcv_strTitle & " >>> " & Title3	
						Else
							pcv_intCont = 0
						End If
					
					End If


					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Stock
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					If strStock="Auto" Then
						pStock = pcArray(11,k)
						pcv_intBackOrder = pcArray(12,k)
						if isNull(pcv_intBackOrder) OR pcv_intBackOrder="" then
							pcv_intBackOrder = 0
						end if
						pNoStock = pcArray(13,k)
						pserviceSpec = pcArray(14,k)
						if CLng(pStock) > 0 then
							strInStock = "In Stock"
						else
							strInStock = "Out of Stock"
							If scOutofStockPurchase=-1 Then							
								If (pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=1) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=1) Then
									If clng(pcv_intShipNDays)>0 then
										strInStock = "Back-Order"
									End if
								End If
							Else
								strInStock = "In Stock"
							End If
						End if						
						
					Else
						strInStock = strStock
					End If
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Product Type
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					strPrdType=ClearHTMLTags2(strPrdType,0)
					strPrdType=replace(strPrdType,chr(34),"&quot")
					strPrdType=replace(strPrdType,chr(9)," ")
					strPrdType=replace(strPrdType,VbCrLf," ")
					strPrdType=replace(strPrdType,VbCr," ")
					strPrdType=replace(strPrdType,VbLf," ")
					strPrdType=replace(strPrdType,"&nbsp;"," ")
					do while InStr(strPrdType,"  ")>0
						strPrdType=replace(strPrdType,"  "," ")
					loop
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Fix Strings
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
					'// Clear HTML on Short Description
					if trim(strShortProductDesc)<>"" then
						strShortProductDesc=ClearHTMLTags2(strShortProductDesc,0)	
					end if		
				
					'// Check Short Description Text
					if trim(strShortProductDesc)<>"" then
						strProductDesc=strShortProductDesc
					else
					'// No Short Description Text, Use Long Description
						strProductDesc=ClearHTMLTags2(strProductDesc,0)				
					end if
					
					strProductDesc=trim(strProductDesc)
					strProductDesc=replace(strProductDesc,chr(34),"&quot")
					strProductDesc=replace(strProductDesc,chr(9)," ")
					strProductDesc=replace(strProductDesc,VbCrLf," ")
					strProductDesc=replace(strProductDesc,VbCr," ")
					strProductDesc=replace(strProductDesc,VbLf," ")
					strProductDesc=replace(strProductDesc,"&nbsp;"," ")
					do while InStr(strProductDesc,"  ")>0
						strProductDesc=replace(strProductDesc,"  "," ")
					loop
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Commission
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					If tmpCOMMISSION="" Then
						tmpCOMMISSION=strCommission
					End If

					tmpCOMMISSION=replace(tmpCOMMISSION,"%","")
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Shipping Cost
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					If tmpSHIPPING="" Then
						tmpSHIPPING=strShipping
					End If
					If tmpSHIPPING="" Then
						tmpSHIPPING=0
					End If
					tmpSHIPPING=replace(tmpSHIPPING,"$","")
					

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Merchant Category
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					srtMerchantCategory=pcf_GetCategoryBC(tmpIDCategory)
					IF len(srtMerchantCategory)>0 THEN
						pcArrayBreadCrumbs=split(srtMerchantCategory,"|,|")
						srtMerchantCategory=""
						for i=0 to ubound(pcArrayBreadCrumbs)
							pcArrayCrumb=split(pcArrayBreadCrumbs(i),"||")
							if i=0 then
								srtMerchantCategory=srtMerchantCategory& pcArrayCrumb(1)
							else
								srtMerchantCategory=srtMerchantCategory&" > " & pcArrayCrumb(1)
							end if
						next
						
					END IF
				
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Shipping Weight
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pWeight = pcArray(15,k)
					if pWeight>0 then
						srtShipWeight=round((pWeight/16),2)
					else
						srtShipWeight=""			
					end if


					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// Pricing
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if NOT isNull(dblPrice) AND dblPrice<>"" then
					else
						dblPrice=0
					end if
					
					IF dblPrice>0 THEN
						dblPrice=money(dblPrice)
						
						'// If an image does not exist for a product, the image field should be left blank and an image file such as "image not available" should not be used. If you want to change this behavior and use a default image, remove the comment from the following 4 lines.
						'if strProductImg<>"" then
						'else
							'strProductImg="no_image.gif"
						'end if
						
						if strProductImg<>"" and strProductImg <> "no_image.gif" then
							strProductImgURL=SPathInfo & "pc/catalog/" & strProductImg
							else
							strProductImgURL=""
						end if
						
						' 1 Merchant Product ID
						' 2 Reserved
						' 3 Title
						' 4 Manufacturer
						' 5 Manufacturer Part #
						' 6 UPC
						' 7 ISBN
						' 8 SKU
						' 9 Product URL
						'10 Price
						'11 In Stock
						'12 Description
						'13 Image URL
						'14 Shipping Costs
						'15 Merchant’s Category
						'16 Shipping Weight
						'17 Commission %
						'18 Condition
						
						strtext=strtext _
						& 			intIdProduct _
						& chr(9) & 	"" _
						& chr(9) &	pcv_strTitle _						
						& chr(9) & 	strBrandName _
						& chr(9) & 	tmpMPN _
						& chr(9) & 	tmpUPC _
						& chr(9) & 	tmpISBN _
						& chr(9) & 	strSKU _
						& chr(9) & 	strProductURL _
						& chr(9) & 	dblPrice _
						& chr(9) & 	strInStock _
						& chr(9) & 	strProductDesc _
						& chr(9) & 	strProductImgURL _
						& chr(9) & 	tmpSHIPPING _
						& chr(9) & 	srtMerchantCategory _
						& chr(9) & 	srtShipWeight _
						& chr(9) & 	tmpCOMMISSION _
						& chr(9) & 	strCondition _
						& vbcrlf
						
					END IF
					
				'END IF
		
			Next
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End:  Do For Each Product In Category
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		end if '// '/ if Count1>-1 then
	
		Count=Count+Count2
		
		If trim(pcList1(lk))="0" then
			exit for
		End if
	
	Next
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// End:  Do For Each Category
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
	
	call closedb()
	
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fso.CreateTextFile(server.MapPath(".") & "\" & File1,True)
	a.Write(strtext)	
	a.Close
	Set fso=Nothing
	%>
	<table class="pcCPcontent">
		<tr> 
			<td>
				<div class="pcCPmessageSuccess">The Bing Cashback data feed was created successfully.</div>
				<br>
				<b><%=Count%></b> products have been exported successfully to the Cashback data feed: <a href="<%=File1%>"><%=File1%></a>. To download the file, either right-click on the file name and select '<strong>Save Target As...</strong>' or '<strong>Save Link as...</strong>' or FTP into the <em><%=scAdminFolderName%></em> folder and download it. Upload the file to the location provided by Microsoft in your welcome email.  Remember to update your cashback file any time your product pricing changes.
				<br>
				<br>
				<br>
				<form class="pcForms">
					<input type=button name=back value="Create Another File" onclick="location='exportCashBack.asp'">&nbsp;
					<input type=button name=back value="Start Page" onclick="location='menu.asp'">
				</form>
				<br>
				<br>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
			</td>
		</tr>
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
			alert("Invalid TEXT file type (*.txt) is not allowed on Cashback.");
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
	return (true);
	}
	//-->
	</script>
    <form name="form1" method="post" action="exportCashBack.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
        <input name="pcv_filename" type="hidden" value="cashback.txt">
        <input name="pcv_rootCat" type="hidden" value="Our Products">
        <table class="pcCPcontent">
        	<tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<th>Take Advantage of Custom Search Fields</th>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td>ProductCart can help you upload product information to your Microsoft Bing Cashback merchant account. For more information about submitting your products to Bing Cashback, please <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_cashback_file" target="_blank">consult our documentation</a>.
            <br /><br />
            To add additional, product specific information to your data feed such as the <a href="http://en.wikipedia.org/wiki/Manufacturer_part_number" target="_blank">MPN</a>, <a href="http://en.wikipedia.org/wiki/UPC_code" target="_blank">UPC</a>, or <a href="http://en.wikipedia.org/wiki/Isbn" target="_blank">ISBN</a> code, you can take advantage of <a href="ManageSearchFields.asp" target="_blank">Custom Search Fields</a>. For example, you can map an existing custom search field named "MPN" to the export column named &quot;MPN&quot;. Since custom search field values can be imported, you can quickly add that information for many products at once. 
            <ul class="pcListIcon">
            	<li><a href="ManageSearchFields.asp">Manage Custom Fields</a></li>
                <li><a href="SearchFields_Export.asp?export=c">Add/Modify Custom Search Field Mappings</a></li>
            </ul>          
            When generating the data feed, ProductCart will use the value for &quot;MPN&quot; saved with each product to <em>auto-fill</em> that specific field. You can then review/edit the values automatically assigned by ProductCart on the next page of this export utility.
            </td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<th>Create the Data Feed</th>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Product Title</td>
          </tr>
					<tr>
            <td>        
                You need to create your own <strong>Title</strong> using the 3 options below. Each option will be separated by ">>>". Click here to learn more &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=213')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>. 
                <br />
                <br />
                <select name="Title1" id="Title1">
                  <option value="Product">Product Name</option>
                  <option value="SKU">SKU Name</option>
                  <option value="Brand">Brand</option>                  
                  <option value="Category">Category Name</option>                  
                </select>
                &nbsp;>>>&nbsp;
                <select name="Title2" id="Title2">
                  <option value=""></option>
                  <option value="Product">Product Name</option>
                  <option value="SKU">SKU Name</option>
                  <option value="Brand">Brand</option>                  
                  <option value="Category">Category Name</option> 
                </select>
                &nbsp;>>>&nbsp;
                <select name="Title3" id="Title3">
                  <option value=""></option>
                  <option value="Product">Product Name</option>
                  <option value="SKU">SKU Name</option>
                  <option value="Brand">Brand</option>                  
                  <option value="Category">Category Name</option> 
                </select>
               </td>
        </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Product Categories</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
            <td>
<strong>Select the categories</strong> that you would like to include in the Cashback Base Bulk Upload file. Press down the CTRL key on your keyboard to select multiple entries. DO NOT select &quot;All Categories&quot; if your store has thousands of products.<br /><br />
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
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Currency</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
          <td>
            Please <strong>select the Currency</strong> that you would like to include in the file.
            <br />
            <br />
            <select name="idCurrency" id="idCurrency">
              <option value="USD" selected>USD</option>
            </select> <i>(note: Cashback is only USD at this time)</i>
          </td>
        </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Product Availability</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
          <td>
            <select name="Stock" id="Stock">
              <option value="Auto" selected>Use ProductCart Settings</option>
              <option value="In Stock">In Stock</option>
              <option value="Out of Stock">Out of Stock</option>
              <option value="Pre-Order">Pre-Order</option>
              <option value="Back-Order">Back-Order</option>
            </select>
          </td>
        </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Commission (Cashback Amount)</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
          <td>
          	Do not enter text or symbols (%). You may optionally specify the commission per product on the next page. Note, setting some or all of your commission in your data feed will prevent you from using the Bidding Center (Bid Management in the Merchant Center). The Bidding Center is located in the Merchant Center. With this feature you’re able to set a base commission based on the categories of your products. This way, you do not have to enter them separately and manually.
            <br />
            <br />
          	<input type="text" name="Commission" size="10">
          	%  
          </td>
        </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Shipping Costs</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
          <td>
          	Do not enter text or symbols ($). You may optionally specify the shipping cost per product on the next page. The Shipping cost is not displayed at Bing Cashback and will be calculated at the time of checkout.
            <br />
            <br />            
          	$<input type="text" name="Shipping" value="0" size="10"></td>
        </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Product Condition</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
          <td>
            <select name="idCondition" id="idCondition">
              <option value="New" selected>New</option>
              <option value="Used">Used</option>
              <option value="Refurbished">Refurbished</option>
            </select>
           </td>
        </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Brand</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
          <td>
            If a product has already been associated with a brand, that brand name will be used. If not, the following, default value will be used:
            <br />
            <br />
            <input type="text" name="CustomBrand" value="<%=scCompanyName%>" size="30">
          </td>
        </tr>
        <!--<tr>
          <td align="right" valign="top" nowrap><div class="pcCPsearch"><strong>Expiration Date:</strong></div></td>
          <td>
            Please enter the date that the product listing expires:<br>
            <input type="text" name="ExpirationDate" value="<%=Year(Date())+1 & "-" & FixDate(Month(Date())) & "-" & FixDate(Day(Date()))%>" size="30"> <i>(format: YYYY-MM-DD)</i>
          </td>
        </tr>-->      
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
          <tr>
          	<td class="pcCPsectionTitle">Other Options</td>
          </tr>
          <tr>
          	<td class="pcCPspacer"></td>
          </tr>
        <tr>
            <td><input type="checkbox" name="excWCats" value="1" checked class="clearBorder"> Exclude wholesale categories</td>
        </tr>
        <tr>
          <td><input type="checkbox" name="excNFSPrds" value="1" class="clearBorder"> Exclude 'Not For Sale' products</td>
        </tr>   
        <tr>
          <td><input type="checkbox" name="showDetails" value="1" checked class="clearBorder"> Show Additional Details (e.g. MPC, UPC, ISBN). This option should not be used with large exports &gt;1000 products</td>
        </tr>
        <tr>					
            <td><hr color="#e1e1e1" width="100%" size="1" noshade></td>
        </tr>
        <tr>
            <td class="pcCPspacer"></td>
        </tr>	
        <tr>
            <td align="center">
                <input type="submit" name="Submit" value="Generate Bing Cashback Data Feed" onClick="pcf_Open_GoogleBase2();" class="submit2">
				<%
				'// Loading Window
				'	>> Call Method with OpenHS();
				response.Write(pcf_ModalWindow("Gathering information for Bing Cashback Data Feed...", "GoogleBase2", 300))
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


'[ClearHTMLTags2]
	
'Coded by Jóhann Haukur Gunnarsson
'joi@innn.is

'	Purpose: This function clears all HTML tags from a string using Regular Expressions.
'	 Inputs: strHTML2;	A string to be cleared of HTML TAGS
'		 intWorkFlow2;	An integer that if equals to 0 runs only the regEx2p filter
'							  .. 1 runs only the HTML source render filter
'							  .. 2 runs both the regEx2p and the HTML source render
'							  .. >2 defaults to 0
'	Returns: A string that has been filtered by the function


function ClearHTMLTags2(strHTML2, intWorkFlow2)

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
	strTagLess2=trim(strTagLess2)
	do while instr(strTagLess2,"  ")>0
	strTagLess2=replace(strTagLess2,"  "," ")
	loop
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
	ClearHTMLTags2 = strTagLess2
	'The results are passed back
	'---------------------------------------
	
end function


Function pcf_GetBrandName(strIDBrand)
	query="SELECT BrandName FROM Brands WHERE idBrand=" & strIDBrand & ";"
	set rsBrand=server.CreateObject("ADODB.RecordSet")
	set rsBrand=connTemp.execute(query)
	if not rsBrand.eof then
		pcf_GetBrandName=rsBrand("BrandName")
	else
		pcf_GetBrandName=strCustomBrand
	end if
	set rsBrand=nothing
End Function

Function pcf_GetCategoryName(strIDCategory)
	query="SELECT categories.categoryDesc FROM categories WHERE idCategory=" & strIDCategory & ";"
	set rsCategory=server.CreateObject("ADODB.RecordSet")
	set rsCategory=connTemp.execute(query)
	if not rsCategory.eof then
		pcf_GetCategoryName=rsCategory("categoryDesc")
	else
		pcf_GetCategoryName=""
	end if
	set rsCategory=nothing
End Function

Function pcf_GetCategoryBC(strIDCategory)
	query="SELECT categories.pccats_BreadCrumbs FROM categories WHERE idCategory=" & strIDCategory & ";"
	set rsCategoryBC=server.CreateObject("ADODB.RecordSet")
	set rsCategoryBC=connTemp.execute(query)
	if not rsCategoryBC.eof then
		pcf_GetCategoryBC=rsCategoryBC("pccats_BreadCrumbs")
	else
		pcf_GetCategoryBC=""
	end if
	set rsCategoryBC=nothing
End Function
'/////////////////////////////////////////////////////////////////////
'// END: FUNCTIONS
'/////////////////////////////////////////////////////////////////////
%><!--#include file="AdminFooter.asp"-->