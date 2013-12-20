<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
Dim strShowBTO
strShowBTO=""
Dim showAddtoCart,showCustomize
showAddtoCart=0
showCustomize=0
Dim bCounter


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Enhanced Views Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim pcv_strUseEnhancedViews, pcv_strHighSlide_Align, pcv_strHighSlide_Template
Dim pcv_strHighSlide_Eval, pcv_strHighSlide_Effects, pcv_strHighSlide_MinWidth, pcv_strHighSlide_MinHeight

pcv_strUseEnhancedViews = True '// Turn Enhanced Views ON or OFF
pcv_strHighSlide_Align = "center" '// Align Images from anchor or screen
pcv_strHighSlide_Template = "rounded-white" '// Template
pcv_strHighSlide_Eval = "this.thumb.alt"
pcv_strHighSlide_Effects = "'expand', 'crossfade'"
pcv_strHighSlide_MinWidth = 250
pcv_strHighSlide_MinHeight = 250
pcv_strHighSlide_Fade = "true"
pcv_strHighSlide_Dim = 0.3
pcv_strHighSlide_Interval = 3500
pcv_strHighSlide_Heading = "highslide-caption" '// "highslide-heading"
pcv_strHighSlide_Hide = "true"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Enhanced Views Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcCheckPricingCats.asp"-->
<%

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PRODUCT ID - Retrieve and validate product ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
pIdProduct=session("idProductRedirect")
if not validNum(pIdProduct) then
	pIdProduct=request("idProduct")
	if not validNum(pIdProduct) then
		'// Set Privacy Settings Test Cookie		
		Response.Cookies("pcC_detect") = "PASS"
		Response.Cookies("pcC_detect").Expires = Date() + 1
		call closedb()
		response.redirect "msg.asp?message=207"
	end if
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
' ADMIN PREVIEW: Check to see if this is a store manager preview
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim pcv_intAdminPreview
pcv_intAdminPreview=0
pcv_intAdminPreview=getUserInput(Request("adminPreview"),10)
	if validNum(pcv_intAdminPreview) and session("admin") <> 0 then
		session("pcv_intAdminPreview")=pcv_intAdminPreview
	else
		session("pcv_intAdminPreview")=0
	end if
pcv_IDProduct=pIdProduct

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' CATEGORY ID - Retrieve, validate and lookup category ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	'// Retrieve category ID from querystring and validate
	pIdCategory=session("intTempCatId")
	session("intTempCatId")=""
	if not validNum(pIdCategory) then
	pIdCategory=request.QueryString("idCategory")
		if not validNum(pIdCategory) then
			pIdCategory=0
		end if
	end if
	pcv_IDCategory=pIdCategory
		
	'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
	if pIdCategory=0 then		
		' If customer is not wholesale, disallow wholesale-only categories
		if not session("customerType")="1" then
			queryW = " AND categories.pccats_RetailHide<>1"
		end if
		' If admin preview, ignore hidden categories
		if session("pcv_intAdminPreview")<>1 then
			queryHC = " AND categories.iBTOhide<>1" & queryW
		end if		
		query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& pIdProduct & queryHC &";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.EOF then
			pIdCategory=rs("idCategory")
		else
			set rs=nothing
			call closeDb()
			response.redirect "msg.asp?message=86"   
		end if
		set rs=nothing
	end if


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Previous and Next Buttons
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_newNextButtons
Dim i,pcArr,pcv_strPreviousPage,pcv_strNextPage
	pcArr=split(session("pcstore_prdlist"),"*****")
	pcv_strPreviousPage=0
	pcv_strNextPage=0
	
    ' Feedback ID: #13290 - begin
    ' More than one product, then only show Prev & Next buttons
    if(ubound(pcArr) > 2) then	
	    For i=lbound(pcArr) to ubound(pcArr)
		    if trim(pcArr(i))<>"" then
			    if clng(pcArr(i))=clng(pIDProduct) then
				    pcv_strPreviousPage=i-1
				    if pcv_strPreviousPage=0 then
					    pcv_strPreviousPage=ubound(pcArr)-1
				    end if
				    pcv_strPreviousPage=pcArr(pcv_strPreviousPage)
				    pcv_strNextPage=i+1
				    if pcv_strNextPage=ubound(pcArr) then
					    pcv_strNextPage=1
				    end if
				    pcv_strNextPage=pcArr(pcv_strNextPage)
				    exit for
			    end if
		    end if
	    Next
	    %>
	    <div align="center">
		    <a rel="nofollow" href="viewPrd.asp?idcategory=&idproduct=<%=pcv_strPreviousPage%>&frmsrc=1" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%> onmouseover="javascript:document.getPrd.idproduct.value='<%=pcv_strPreviousPage%>'; sav_callxml='1'; return runXML1('prd_<%=pcv_strPreviousPage%>');" onmouseout="javascript: sav_callxml=''; hidetip();" <%end if%>><img src="<%=pcv_tmpNewPath%><%=rslayout("pcLO_previous")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_3")%>"></a>
		    &nbsp;
		    <a rel="nofollow" href="viewPrd.asp?idcategory=&idproduct=<%=pcv_strNextPage%>&frmsrc=1" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%> onmouseover="javascript:document.getPrd.idproduct.value='<%=pcv_strNextPage%>'; sav_callxml='1'; return runXML1('prd_<%=pcv_strNextPage%>');" onmouseout="javascript: sav_callxml=''; hidetip();" <%end if%>><img src="<%=pcv_tmpNewPath%><%=rslayout("pcLO_next")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_4")%>"></a>
		    <br/>
	    </div>
	    <%
	end if
	' Feedback ID: #13290 - end
	
End Sub

Public Sub pcs_NextButtons

IF ((session("pcstore_newsrc")="OK") or (request("frmsrc")="1")) AND (session("pcstore_prdlist")<>"") THEN
	session("pcstore_newsrc")=""
	call pcs_newNextButtons
ELSE
	session("pcstore_newsrc")=""
	session("pcstore_prdlist")=""
	'// We can only display this section is the category is greater than 0
	If pIdCategory>1 Then		
		'// Get our array
		'// Unfortunately we have to generate this everytime, since the admin may deactivate a product at any time.
		'// We can NOT use a session or save to the database for this reason.
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Decide Order By
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		query="Select POrder from categories_products where idCategory="& pIdCategory &";"
		set rsCatOrder=Server.CreateObject("ADODB.Recordset")     
		set rsCatOrder=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsCatOrder=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		UONum=0
		do while not rsCatOrder.eof
			pcv_strCatOrder=rsCatOrder("POrder")
			if pcv_strCatOrder<>"" AND isNULL(pcv_strCatOrder)=False then
				UONum=UONum+CLng(pcv_strCatOrder)
			end if
			rsCatOrder.MoveNext
		loop
		SET rsCatOrder=nothing		
		
		ProdSort=""
		if UONum>0 then
			ProdSort="19"
		else
			ProdSort="" & PCOrd
		end if			
		if ProdSort="" then
			ProdSort="0"
		end if
		
		select case ProdSort
			Case "19": query1 = " ORDER BY categories_products.POrder Asc"
			Case "0": query1 = " ORDER BY products.SKU Asc"
			Case "1": query1 = " ORDER BY products.description Asc" 	
			Case "2": 
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
					else
						query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) DESC"
					end if
				else
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
					else
						query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) DESC"
					end if
				End if
			Case "3":
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
					else
						query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) ASC"
					end if
				else
					if Ucase(scDB)="SQL" then
						query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
					else
						query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) ASC"
					end if
				End if
		end select
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Decide Order By
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		' SELECT DATA SET
		' TABLES: products, categories_products
		query = 		"SELECT products.idProduct, products.description, products.pcProd_BTODefaultWPrice, products.bToBprice, products.pcProd_BTODefaultPrice, categories_products.idCategory, categories_products.POrder "
		query = query & "FROM products "
		query = query & "INNER JOIN categories_products "
		query = query & "ON products.idProduct = categories_products.idProduct "
		query = query & "WHERE categories_products.idCategory=" & pIdCategory &" "
		query = query & "AND products.active=-1 AND products.removed=0 AND products.configOnly=0 "
		query = query & "" & query1
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)			
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		pcv_strNextProductID = ""		
			
		if NOT rs.eof then
			Do until rs.eof
				'response.write rs("idProduct") & " " & rs("description") & "<br />"
				'// We need to form our Array
				xProductArrayCount = xProductArrayCount + 1
				pcv_strTmpProductID = rs("idProduct")
				if pcv_strTmpProductID <> "" then			
					pcv_strNextProductID = pcv_strNextProductID & pcv_strTmpProductID & chr(124)	
				end if	
			rs.movenext
			Loop
			
			'// Trim the last pipe if there is one
			xStringLength = len(pcv_strNextProductID)
			pcv_strShowButtons = 0
			if xStringLength>0 then
				pcv_strNextProductID = left(pcv_strNextProductID,(xStringLength-1))
				'// If there are no other pipes left then we only have one product in this category, so we can exit.
				if instr(pcv_strNextProductID,chr(124))>0 then
					pcv_strShowButtons = 1 '// show buttons
				else
					pcv_strShowButtons = 0
				end if
			end if
			
		end if
		set rs=nothing
		
		If pcv_strShowButtons = 1 Then
		
			'// Set Up Our Array
			pcArrayNextProductID = split(pcv_strNextProductID,chr(124))		
			pcv_intLBound = LBound(pcArrayNextProductID)
			pcv_intUBound = UBound(pcArrayNextProductID)
			
			'// Now find our place in the array
			For i = pcf_IDMaximum(pcv_intLBound, intStartIndex) To pcv_intUBound
				If CStr(pcArrayNextProductID(i)) = CStr(pIdProduct) Then
					pcv_intCurrentPosition = i
					Exit For
				End If
			Next
			
			'// Previous Product	
			if (pcv_intCurrentPosition-1) < pcv_intLBound then
				pcv_strPreviousPage=pcArrayNextProductID(pcv_intUBound)
			else
				pcv_strPreviousPage=pcArrayNextProductID(pcv_intCurrentPosition-1)
			end if
			
			'// Next Product
			if (pcv_intCurrentPosition+1) > pcv_intUBound then
				pcv_strNextPage=pcArrayNextProductID(pcv_intLBound)
			else
				pcv_strNextPage=pcArrayNextProductID(pcv_intCurrentPosition+1)
			end if

			'// Generate SEO Links
			'// Get product description
			query = "SELECT description FROM products WHERE idProduct="&pcv_strPreviousPage
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pcv_strPreviousPageDesc=rs("description")
			query = "SELECT description FROM products WHERE idProduct="&pcv_strNextPage
			set rs=conntemp.execute(query)			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pcv_strNextPageDesc=rs("description")
			set rs=nothing
			Call pcGenerateSeoLinks
			%>
			<div align="center">
			<a href="<%=pcStrPrdPreLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%> onmouseover="javascript:document.getPrd.idproduct.value='<%=pcv_strPreviousPage%>'; sav_callxml='1'; return runXML1('prd_<%=pcv_strPreviousPage%>');" onmouseout="javascript: sav_callxml=''; hidetip();" <%end if%>><img src="<%=pcv_tmpNewPath%><%=rslayout("pcLO_previous")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_3")%>"></a>
			&nbsp;
			<a href="<%=pcStrPrdNextLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%> onmouseover="javascript:document.getPrd.idproduct.value='<%=pcv_strNextPage%>'; sav_callxml='1'; return runXML1('prd_<%=pcv_strNextPage%>');" onmouseout="javascript: sav_callxml=''; hidetip();" <%end if%>><img src="<%=pcv_tmpNewPath%><%=rslayout("pcLO_next")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_4")%>"></a>
			<br/>
			</div>
			<%
		End If '// If pcv_strShowButtons = 1 Then
	End If
END IF

End Sub


Function pcf_IDMaximum(ByVal x, ByVal y) 
  If x > y Then 
    pcf_IDMaximum = x 
  Else 
    pcf_IDMaximum = y 
  End If 
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Previous and Next Buttons
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Check if we should hide the options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Function pcf_VerifyShowOptions
	If pserviceSpec<>0 AND (pPrice=0 OR scConfigPurchaseOnly=1) Then
		pcf_VerifyShowOptions = false
	Else
		pcf_VerifyShowOptions = true
	End If
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Check if we should hide the options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Start SDBA
' START:  Display Back-Order Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_DisplayBOMsg
	If (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=1) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=1) Then
		If clng(pcv_intShipNDays)>0 then
			response.write "<div>"&dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b") & "</div>"
		End if
	End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Display Back-Order Message
'End SDBA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  BTOisConfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Function pcf_BTOisConfig

	query="SELECT categories.categoryDesc, products.description, configSpec_products.configProductCategory, configSpec_products.price, categories_products.idCategory, categories_products.idProduct, products.weight FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NOT rstemp.eof then
		pcf_BTOisConfig = true
	else
		pcf_BTOisConfig = false
	end if 
	Set rstemp = nothing
	
	query="SELECT * FROM configSpec_Charges WHERE specProduct="&pIdProduct
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if	
	BTOCharges=0
	if not rstemp.eof then
		BTOCharges=1
	end if
	set rstemp=nothing

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  BTOisConfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  CATEGORY TREE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_CategoryTree
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: get category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pIdCategory > 0 then
	%>  <!--#include file="pcBreadCrumbs.asp"-->  <% 
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  get category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  show breadcrumbs - category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If strBreadCrumb<>"" then %>
		<div class="pcPageNav">
			<%=dictLanguage.Item(Session("language")&"_viewCat_P_2") %>
			<%=strBreadCrumb %>
            <%
            intIdCategory=pIdCategory

            '// Load category discount icon
            %>
            <!--#Include File="pcShowCatDiscIcon.asp" -->
		</div>
	<% 
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  show breadcrumbs - category tree array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  CATEGORY TREE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show product name 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductName 
'// PC v4.5 AddThis integration
if scAddThisDisplay=1 then pcs_AddThis
%>
<h1><%=pDescription%></h1>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show product name 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show SKU
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowSKU
IF pHideSKU<>"1" THEN%>
	<div class="pcShowProductSku">
		<%=dictLanguage.Item(Session("language")&"_viewCat_P_8")%>: <%=pSku%>
	</div>
<%END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show SKU
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' PRV41 start
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Average Rating (from reviews)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowRating
IF pRSActive And pcv_ShowRatSum And pNumRatings>0 THEN%>
	<div class="pcShowProductRating">
<%
			IF pcv_RatingType="0" then
				query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pcv_IDProduct
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)

				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)

				query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " AND pcRev_Active=1 AND pcRev_MainRate>0"
				set rs=connTemp.execute(query)

				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				intCount = clng(rs("ct"))

				set rs=Nothing
				%>
				<%if pcv_tmpRating>"0" then%><a href="#productReviews" style="text-decoration: none;"><%=dictLanguage.Item(Session("language")&"_prv_2")%></a><img src="catalog/<%=pcv_Img1%>" align="absbottom"><%=pcv_tmpRating%>% <%=pcv_MainRateTxt1%> (<%=intCount%>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)<%end if%>
				<%
			ELSE
				if pcv_CalMain="1" then     ' Can be set independently of sub-ratings 
					query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pcv_IDProduct
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)

					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if

					pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)

					set rs=nothing
					if CDbl(pcv_tmpRating)>0 then 
				    %>
					    <a href="#productReviews" style="text-decoration: none;"><%=dictLanguage.Item(Session("language")&"_prv_39")%></a> 
				        <% Call WriteStar(pcv_tmpRating,1) 
			        end if
				    %>
		        <% else 'Will be calculated automatically by averaging sub-ratings
					Call CreateList()
				    pcv_tmpRating=CalRating()
					if CDbl(pcv_tmpRating)>0 then %>
				    <a href="#productReviews" style="text-decoration: none;"><%=dictLanguage.Item(Session("language")&"_prv_2")%></a>
				    <% Call WriteStar(pcv_tmpRating,1)
					end if %>
		        <% end if
		    END IF 'Main Rating
		 %>

	</div>
<%END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Average Rating
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PRV41 end


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Custom Search Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_CustomSearchFields
Dim query,rs,pcArr,intCount,i
	query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pIdProduct & " AND pcSearchFieldShow=1 ORDER BY pcSearchFields.pcSearchFieldOrder ASC,pcSearchFields.pcSearchFieldName ASC;"
	set rs=connTemp.execute(query)
	IF not rs.eof THEN
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		response.Write("<div style='padding-top: 5px;'></div>")
		For i=0 to intCount
				response.write "<div class='pcShowProductCustSearch'>"&pcArr(1,i)&": <a href='showsearchresults.asp?customfield="&pcArr(0,i)&"&SearchValues="&Server.URLEncode(pcArr(2,i))&"'>"&pcArr(3,i)&"</a></div>"
		Next
	END IF
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Custom Search Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Weight (If admin turned on)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_DisplayWeight
if scShowProductWeight="-1" then
		if int(pWeight)>0 then
			response.write "<div class='pcShowProductWeight'>"
			response.write ship_dictLanguage.Item(Session("language")&"_viewCart_c")
			if scShipFromWeightUnit="KGS" then
				pKilos=Int(pWeight/1000)
				pWeight_g=pWeight-(pKilos*1000)
				pWeight=pKilos
				if pWeight_g>0 then
					response.write dictLanguage.Item(Session("language")&"_viewCart_c")&pWeight&" kg "&pWeight_g&" g" & "<br />"
				else
					response.write dictLanguage.Item(Session("language")&"_viewCart_c")&pWeight&" kg" & "<br />"
				end if
			else
				pPounds=Int(pWeight/16)
				pWeight_oz=pWeight-(pPounds*16)
				pWeight=pPounds
				if pWeight_oz>0 then
					response.write dictLanguage.Item(Session("language")&"_viewCart_c")&pWeight&" lbs "&pWeight_oz&" ozs" & "<br />"
				else
					response.write dictLanguage.Item(Session("language")&"_viewCart_c")&pWeight&" lbs" & "<br />"
				end if
			end if
			response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Weight
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Brand (If assigned)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowBrand
	if sBrandPro="1" then
		if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
		response.write "<div class='pcShowProductBrand'>"
		response.write dictLanguage.Item(Session("language")&"_viewPrd_brand")
		%>
			<a href="showsearchresults.asp?IDBrand=<%=pIDBrand%>">
				<%=BrandName%>
			</a>
		<% 
		response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Brand
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Units in Stock (if on, show the stock level here)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_UnitsStock
	if scdisplayStock=-1 AND pNoStock=0 then
		if pstock > 0 then
			response.write "<div class='pcShowProductStock'>"
			response.write dictLanguage.Item(Session("language")&"_viewPrd_19") & " " & pStock
			response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Units in Stock (if on, show the stock level here)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductDescription
	if psDesc <> "" then 
		response.write "<div class='pcShowProductSDesc' style='padding-top: 5px'>"
		response.Write psDesc & " <a href='#details'>" & dictLanguage.Item(Session("language")&"_viewPrd_21") & "</a>"
		response.write "</div>"
	else
		response.write "<div class='pcShowProductSDesc' style='padding-top: 5px'>"
		response.Write pDetails
		response.write "</div>"
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_GetBTOConfiguration
Dim query,rs
	pcv_BTORP=Clng(0)
	strShowBTO=""		
	if pserviceSpec=true then
	 '// Product is BTO
		
		' Get data
		query="SELECT categories.categoryDesc, products.description, products.iRewardPoints,configSpec_products.configProductCategory, configSpec_products.price, configSpec_products.Wprice, categories_products.idCategory, categories_products.idProduct, products.weight, products.pcprod_minimumqty FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		iAddDefaultPrice=Cdbl(0)
		iAddDefaultWPrice=Cdbl(0)
		iAddDefaultPrice1=Cdbl(0)
		iAddDefaultWPrice1=Cdbl(0)
		
		if NOT rs.eof then 
			Dim FirstCnt
			FirstCnt=0
			if intpHideDefConfig="0" then
				strShowBTO= strShowBTO & "<div class='pcShowProductBTOConfig' style='padding-top: 10px; padding-bottom: 2px;'>"
				strShowBTO= strShowBTO & "<b>"&dictLanguage.Item(Session("language")&"_viewPrd_25")&"</b>"
				strShowBTO= strShowBTO & "</div>"
			end if
			do until rs.eof
				FirstCnt=FirstCnt+1
				strCategoryDesc=rs("categoryDesc")
				strDescription=rs("description")
				strConfigProductCategory=rs("configProductCategory")
				dblPrice=rs("price")
				dblWPrice=rs("Wprice")
				intIdCategory=rs("idCategory")
				intIdProduct=rs("idProduct")
				intReward=rs("iRewardPoints")
				if (intReward<>"") and (intReward<>"0") then
				else
				intReward=0
				end if
				intWeight=rs("weight")
				pcv_iminqty=rs("pcprod_minimumqty")
				if IsNull(pcv_iminqty) or pcv_iminqty="" then
					pcv_iminqty=1
				end if
				if pcv_iminqty="0" then
					pcv_iminqty=1
				end if
				pcv_BTORP=pcv_BTORP+clng(intReward*pcv_iminqty)
				
				dblPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,0)
				dblWPrice1=CheckPrdPrices(pIdProduct,intIdProduct,dblPrice,dblWPrice,1)
				iAddDefaultPrice=Cdbl(iAddDefaultPrice+dblPrice*pcv_iminqty)
				iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+dblWPrice*pcv_iminqty)
				iAddDefaultPrice1=Cdbl(iAddDefaultPrice1+dblPrice1*pcv_iminqty)
				iAddDefaultWPrice1=Cdbl(iAddDefaultWPrice1+dblWPrice1*pcv_iminqty)
				ItemPrice=0
				if Session("CustomerType")=1 then
					if (dblWPrice<>0) then
						ItemPrice=dblWPrice1
					else
						ItemPrice=dblPrice1
					end if
				else
					ItemPrice=dblPrice1
				end if
				if intpHideDefConfig="0" then
					strShowBTO= strShowBTO & "<div class='pcShowProductBTOConfig'>"
					strShowBTO= strShowBTO & "<b>"&strCategoryDesc&"</b>: "&strDescription
					strShowBTO= strShowBTO & "</div>"
				end if
				response.write "<input name=""CAT"&FirstCnt&""" type=""HIDDEN"" value=""CAG"&intIdCategory&""">"
				response.write "<input name=""CAG"&intIdCategory&"QF"" type=""HIDDEN"" value=""" & pcv_iminqty & """>"
				response.write "<input type=""hidden"" name=""CAG"&intIdCategory&""" value="""&intIdProduct&"_0_"&intWeight&"_" & ItemPrice & """>"
				rs.moveNext
			loop			
			response.write "<input type=""hidden"" name=""FirstCnt"" value="""&FirstCnt&""">"
		end if 
		set rs=nothing
	end if
End Sub

Public Sub pcs_BTOConfiguration
	if strShowBTO<>"" then
		response.write strShowBTO
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show BTO Default Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Reward Points
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_RewardPoints
	If RewardsActive=1 then
		' Show Reward Points associated with this product, if any
		' By default, Reward Points are not shown to Wholesale Customers
		if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")<>"1" then
			response.write "<div style='padding-top: 5px;'>"&dictLanguage.Item(Session("language")&"_viewPrd_50")&Clng(iRewardPoints+clng(pcv_BTORP))&"&nbsp;"&RewardsLabel&dictLanguage.Item(Session("language")&"_viewPrd_51")&"</div>"
		else
			' If the system is setup to include Wholesale Customers, then show Reward Points to them too
			if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")="1" and RewardsIncludeWholesale=1 then
				response.write "<div style='padding-top: 5px;'>"&dictLanguage.Item(Session("language")&"_viewPrd_50")&Clng(iRewardPoints+clng(pcv_BTORP))&"&nbsp;"&RewardsLabel&dictLanguage.Item(Session("language")&"_viewPrd_51")&"</div>"
			end if 
		end If
	End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Reward Points
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show product prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductPrices
Dim rs,query,pcTestPrice,pcHidePricesIfNFS
Dim ShowSaleIcon,rsS,pcSCID,pcSCName,pcSCIcon,pcTargetPrice

ShowSaleIcon=0
pcTestPrice=0

	'// If product is "Not for Sale", should prices be hidden or shown?
	'// Set pcHidePricesIfNFS = 1 to hide, 0 to show.
	'// Here we leverage the "pnoprices" variable to change the behavior (a Control Panel setting could be added in the future)
	pcHidePricesIfNFS = 0
	if (pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0) and pcHidePricesIfNFS=1 then
		pnoprices=2
	end if

	' Don't show prices if the BTO product has been set up to hide prices (pnoprices)
	If pnoprices<2 Then
	
		if UCase(scDB)="SQL" then
			query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & pidproduct & " AND Products.pcSC_ID>0;"
			set rsS=Server.CreateObject("ADODB.Recordset")
			set rsS=conntemp.execute(query)
					
			if not rsS.eof then
				ShowSaleIcon=1
				pcSCID=rsS("pcSC_ID")
				pcSCName=rsS("pcSC_SaveName")
				pcSCIcon=rsS("pcSC_SaveIcon")
				pcTargetPrice=rsS("pcSales_TargetPrice")
				
				query="SELECT pcSB_Price FROM pcSales_BackUp WHERE idProduct=" & pIdProduct & " AND pcSC_ID=" & pcSCID & ";"
				set rsQ=connTemp.execute(query)
				pcOrgPrice=0
				if not rsQ.eof then
					pcOrgPrice=rsQ("pcSB_Price")
				end if
				set rsQ=nothing
				%>
				<script language="JavaScript">
				<!--
				function winSale(fileName)
				{
					myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=450,height=300')
					myFloater.location.href=fileName;
				}
				//-->
				</script>
			<%end if
			set rsS=nothing
		end if
	
		' If this is a BTO product, calculate the base price as the sum of price + default prices
		pPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
		pBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
		if pserviceSpec=true then
			pPrice=Cdbl(pPrice+iAddDefaultPrice)
			pBtoBPrice=Cdbl(pBtoBPrice+iAddDefaultWPrice)
			pPrice1=Cdbl(pPrice1+iAddDefaultPrice1)
			pBtoBPrice1=Cdbl(pBtoBPrice1+iAddDefaultWPrice1)
		end if
		
		'START - Bing Cashback Gleaming
		If LSCB_STATUS = "1" AND LSCB_KEY <>"" Then
			pcv_strIsCashback = getUserInput(request("cashback"),1)
			If len(pcv_strIsCashback)>0 Then
				response.Write("<script language=""javascript"" type=""text/javascript"" ")
				response.Write("src=""http://search.live.com/cashback/products/gleam/javascript.ashx")
				response.Write("?merchantId="& LSCB_KEY &"&type=1&bgcolor=FFFFFF&version=1.00""")
				response.Write("></script>")
			End If
		End If
		'END - Bing Cashback Gleaming
        	
		'START - Visually separate prices from other information. Don't use if layout is One Column.
		if pcv_strViewPrdStyle <> "o" then
			response.write "<div class='pcShowPrices'>"
		end if
	 
		' Display the online price if it's not zero
		if (pPrice>Cdbl(0)) and (pcv_intHideBTOPrice<>"1") then
		
			' If the List Price is not zero and higher than the online price, display striken through
			if ((pListPrice-pPrice)>0) and (pcv_intHideBTOPrice<>"1") then
				response.write "<div class='pcShowProductPrice'>"
				response.write dictLanguage.Item(Session("language")&"_viewPrd_20")
				response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pListPrice) & "</span>"
				response.write "</div>"
			end if
			
			if (ShowSaleIcon=1) AND (pcTargetPrice="0") then
				response.write "<div class='pcShowProductPrice'>"
				response.write dictLanguage.Item(Session("language")&"_Sale_3")
				response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pcOrgPrice) & "</span>"
				response.write "</div>"
			end if
		
			' Display online price
			response.write "<div class='pcShowProductPrice'><span class='pcShowProductMainPrice'>"
			response.write dictLanguage.Item(Session("language")&"_viewPrd_3")
			response.write scCurSign & money(pPrice)
			response.write "</span>"
			if (ShowSaleIcon=1) AND (pcTargetPrice="0") then
				response.write " <span class=""pcSaleIcon""><a href=""javascript:winSale('sm_showdetails.asp?id=" & pcSCID & "')""><img src=""catalog/" & pcSCIcon & """ title=""" &  pcSCName & """ alt=""" & pcSCName & """></a></span>"
			end if
			response.write "</div>"
			
			' If the product is setup to use the Show Savings feature, show the savings if they exist and the customer is retail
			if ((pListPrice-pPrice)>0) AND (plistHidden<0) AND (session("customerType")<>1) and (pcv_intHideBTOPrice<>"1") then
				'response.write " - "
				response.write "<div class='pcShowProductSavings'>"
				response.write dictLanguage.Item(Session("language")&"_viewPrd_4")
				response.write scCurSign & money((pListPrice-pPrice))
				response.write " (" & round(((pListPrice-pPrice)/pListPrice)*100) & "%)"
				response.write "</div>"
			end if

			' If the store is using and showing VAT, show the VAT included message and price without VAT
			if ptaxVAT="1" and ptaxdisplayVAT="1" and pnotax <> "-1" then
				if session("customerType")="1" AND ptaxwholesale="0" then
				else
					response.write "<div class='pcSmallText'>"
					response.write dictLanguage.Item(Session("language")&"_viewPrd_26") & "<br>"
					response.write dictLanguage.Item(Session("language")&"_viewPrd_27") & scCurSign & _
					money(pcf_RemoveVAT(pPrice,pIdProduct)) & ""
					response.write "</div>"
				end if
			end if
		
		end if 'this is the IF statement regarding the online price being > zero
	
		' If this is a wholesale customer and the wholesale price is > zero, display it here
		if pcv_intHideBTOPrice<>"1" then
			if session("customertype")=1 and pBtoBPrice1>0 then
				pPrice1=pBtoBPrice1
			end if
			if session("customerCategory")<>0 then
				if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=clng(session("customerCategory"))) then
					response.write "<div class='pcShowProductPriceW'>"
					response.write session("customerCategoryDesc") & " " & dictLanguage.Item(Session("language")&"_Sale_3")
					response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pcOrgPrice) & "</span>"
					response.write "</div>"
				end if
				response.write "<div class='pcShowProductPriceW'>"
				response.write session("customerCategoryDesc")&": "
				response.write scCurSign & money(pPrice1)
				if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=clng(session("customerCategory"))) then
					response.write " <span class=""pcSaleIcon""><a href=""javascript:winSale('sm_showdetails.asp?id=" & pcSCID & "')""><img src=""catalog/" & pcSCIcon & """ title=""" &  pcSCName & """ alt=""" & pcSCName & """></a></span>"
				end if
				response.write "</div>"
			else
				if (pBtoBPrice1>"0") and (session("customerType")=1) then
					if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then
						response.write "<div class='pcShowProductPrice'>"
						response.write dictLanguage.Item(Session("language")&"_Sale_4")
						response.write "<span class='pcShowProductListPrice'>" & scCurSign & money(pcOrgPrice) & "</span>"
						response.write "</div>"
					end if 
					response.write "<div class='pcShowProductPrice'>"
					response.write dictLanguage.Item(Session("language")&"_viewPrd_15") &" "
					response.write scCurSign & money(pBtoBPrice1)
					if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then
						response.write " <span class=""pcSaleIcon""><a href=""javascript:winSale('sm_showdetails.asp?id=" & pcSCID & "')""><img src=""catalog/" & pcSCIcon & """ title=""" &  pcSCName & """ alt=""" & pcSCName & """></a></span>"
					end if
					response.write "</div>"
				end if
			end if
		end if
		
		' END - Visually separate prices from rest of product information
		if pcv_strViewPrdStyle <> "o" then
			response.write "</div>"
		end if
		
	end if 'this is the IF statement regarding the BTO product being setup not to show prices	
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show product prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get Additional Images Array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
function pcf_GetAdditionalImages
	' // SELECT DATA SET
	' TABLES: pcProductsImages
	' COLUMNS ORDER: pcProductsImages.pcProdImage_Url, pcProductsImages.pcProdImage_LargeUrl, pcProductsImages.pcProdImage_Order
	
	query = 		"SELECT pcProductsImages.pcProdImage_Url, pcProductsImages.pcProdImage_LargeUrl, pcProductsImages.pcProdImage_Order "
	query = query & "FROM pcProductsImages "
	query = query & "WHERE pcProductsImages.idProduct=" & pidProduct &" "
	query = query & "ORDER BY pcProductsImages.pcProdImage_Order;"	
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	If rs.EOF Then
		pcf_GetAdditionalImages = ""
	Else
		pcf_GetAdditionalImages = ""
		Dim xCounter '// declare a temporary counter
		xCounter = 0
		do while NOT rs.EOF
		
		pcv_strProdImage_Url = ""
		pcv_strProdImage_LargeUrl = ""
		pcv_strProdImage_Url = rs("pcProdImage_Url")
		pcv_strProdImage_LargeUrl = rs("pcProdImage_LargeUrl")
			
			if len(pcv_strProdImage_Url)>0 then
			xCounter = xCounter + 1
				if xCounter > 1 then
					pcf_GetAdditionalImages = pcf_GetAdditionalImages & ","
				end if
				'// Add a sorted item onto the end of the string
				pcf_GetAdditionalImages = pcf_GetAdditionalImages & pcv_strProdImage_Url & "," & pcv_strProdImage_LargeUrl
			end if

		rs.movenext 
		loop		
	End If
	set rs=nothing
end function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Get Additional Images Array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductImage

	'// If display option is One Column (ideal for products without images), and there is no image  
	'// don't show anything
	IF NOT (pcv_strViewPrdStyle = "o" AND (len(pImageUrl) = 0 OR pImageUrl="no_image.gif")) THEN  

		if len(pImageUrl) > 0 then
		'// A)  The image exists
		%>
            <div class="pcShowMainImage">
				<%
                '// If this is the pop window swap out the image for the selection  
                if pcv_strPopWindowOpen = 1 then
                    pcv_strVariableImage = pcv_strCurrentUrl  
                else
                    pcv_strVariableImage = pImageUrl
                end if	

				Dim pcv_strZoomLink, pcv_strZoomLocation  			
 
				if pcv_strUseEnhancedViews = True then
					pcv_strZoomLink = "javascript:;"
					pcv_strZoomLocation = "onclick=""pcf_initEnhancement(this,'"&pcv_tmpNewPath&"catalog/"&pLgimageURL&"')"" class=""highslide"""
				else
					pcv_strZoomLink="javascript:enlrge('"&pcv_tmpNewPath&"catalog/"&pLgimageURL&"')"
					pcv_strZoomLocation = ""
				end if 
                %>
                <% if len(pLgimageURL)>0 then %>
					<a id="mainimgdiv" href="<%=pcv_strZoomLink%>" <%=pcv_strZoomLocation%>><img id='mainimg' name='mainimg' src='<%=pcv_tmpNewPath%>catalog/<%=pImageUrl%>' alt="<%=replace(pDescription,"""","&quot;")%>" <%if pcv_IntMojoZoom="1" then%>data-zoomsrc="<%=pcv_tmpNewPath&"catalog/"&pLgimageURL%>" <%end if%>/></a>
                <% else %>
					<img id='mainimg' name='mainimg' src='<%=pcv_tmpNewPath%>catalog/<%=pImageUrl%>' alt="<%=replace(pDescription,"""","&quot;")%>" />
                <% end if %>
                <% if pcv_strUseEnhancedViews = True then %>
                	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                <% end if %>
            </div>
	
			<% if len(pLgimageURL)>0 and pcv_strUseEnhancedViews = False then %>
				<div style="width:100%; text-align:right;">
					<a href="<%=pcv_strZoomLink%>" <%=pcv_strZoomLocation%>><img src="<%=pcv_tmpNewPath%><%=rsIconObj("zoom")%>" border="0" hspace="10" alt="<%=dictLanguage.Item(Session("language")&"_altTag_5")%>"></a>
                    <% if pcv_strUseEnhancedViews = True then %>
                    	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                    <% end if %>
				</div>
			<% end if %>
		<%
		else
		'// B)  The image DOES NOT exist (show no_image.gif)
		%>		
			<div class="pcShowMainImage">
				<img name='mainimg' src='<%=pcv_tmpNewPath%>catalog/no_image.gif' alt="<%=replace(pDescription,"""","&quot;")%>">
			</div>
		<% 
		end if
		
	END IF

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_MakeAdditionalImage

	'// Make the popup link, but dont set large image preference if the large image doesnt exist
	If len(pcv_strShowImage_LargeUrl)>0 Then		
		pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('catalog/"&pcv_strShowImage_LargeUrl&"','"&pidProduct&"')" 
	Else
		pcv_strShowImage_LargeUrl = pcv_strShowImage_Url '// we dont have one, show the regular size
		pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('catalog/"&pcv_strShowImage_Url&"','"&pidProduct&"')" 
	End If
	
	if pcv_strPopWindowOpen = 1 then
		%>
		<a href="#">	
			<img onmouseover='javascript:window.document.mainimg.src="<%=pcv_tmpNewPath%>catalog/<%=pcv_strShowImage_LargeUrl%>";' src='catalog/<%=pcv_strShowImage_Url%>' alt="<%=replace(pDescription,"""","&quot;")%>" />		
		</a> 
	<% else %>
			<%
			'// Use Enhanced Views
			If pcv_strUseEnhancedViews = True Then 
			%>
				<a href="catalog/<%=pcv_strShowImage_LargeUrl%>" class="highslide" onclick="javascript:return hs.expand(this, { slideshowGroup: 'slides' })" id="<%=bcounter%>"><img onmouseover='CurrentImg=<%=bcounter%>;javascript:window.document.mainimg.src="<%=pcv_tmpNewPath%>catalog/<%=pcv_strShowImage_Url%>"; linkChanger("<%=pcv_tmpNewPath%>catalog/<%=pcv_strShowImage_LargeUrl%>","<%=bCounter%>")' src='catalog/<%=pcv_strShowImage_Url%>' alt="<%=replace(pDescription,"""","&quot;")%>" /></a>
                <% if pcv_strUseEnhancedViews = True then %>
                	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                <% end if %>
        	<%
			'// Use Pop Window 
			Else 
				%>	
                <a href="<%=pcv_strLargeUrlPopUp%>"><img onmouseover='javascript:window.document.mainimg.src="<%=pcv_tmpNewPath%>catalog/<%=pcv_strShowImage_Url%>"; linkChanger("<%=pcv_tmpNewPath%>catalog/<%=pcv_strShowImage_LargeUrl%>","<%=bCounter%>")' src='catalog/<%=pcv_strShowImage_Url%> )' src='catalog/<%=pcv_strShowImage_Url%>' alt="<%=replace(pDescription,"""","&quot;")%>" /></a> 
        <% End If		
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Additional Product Images (If there are any)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_AdditionalImages

if len(pImageUrl) > 0 then ' // only if there is a main image can there be additional images.
	pcv_strAdditionalImages = pcf_GetAdditionalImages '// set variable to array of images, if there are any
	if len(pcv_strAdditionalImages)>0 then '// there is a main, are there additionals?
	%>
	<script language="javascript">
	<!--
	function linkChanger(NewLink, Pos) {
		<% if pcv_IntMojoZoom="1" then %>
		MojoZoom.makeZoomable(document.getElementById("mainimg"), NewLink, '', '', '', false);  
		<% end if %>
		document.getElementById('mainimgdiv').onclick = function() {document.getElementById(Pos).onclick();} 
	}
	// end -->
	</script>	
    
	<table class="pcShowAdditional">
	<tr>
	<%
	'// the main image to the first place in the image set
	pcv_strAdditionalImages = pImageUrl & "," & pLgimageURL & "," & pcv_strAdditionalImages
	
	Dim pcArray_AdditionalImages '// declare a temporary array
	pcArray_AdditionalImages = Split(pcv_strAdditionalImages,",")	
	
	bCounter = 1
	
	'// When the product has additional images, this variable defines how many thumbnails are shown per row, below the main product image
	if pcv_intProdImage_Columns="" then
		pcv_intProdImage_Columns = 3
	end if
	
	modnum = pcv_intProdImage_Columns '// Get this from the db
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START Loop
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	For cCounter = LBound(pcArray_AdditionalImages) TO UBound(pcArray_AdditionalImages)
	
	'// Check if we have a normal image
	Dim pcv_strTempAssignment	
	pcv_strTempAssignment = ""
	pcv_strTempAssignment = pcArray_AdditionalImages(cCounter)
	pcv_strShowImage_Url = pcv_strTempAssignment '// we have one, set it
	
	cCounter = cCounter + 1 '// now get the large image
		
	'// Do Not generate an additional image if there is not one
	If len(pcv_strShowImage_Url)>0 Then
	
			'// Check if we have a large image
			pcv_strTempAssignment = ""	
			pcv_strTempAssignment = pcArray_AdditionalImages(cCounter)
			pcv_strShowImage_LargeUrl = pcv_strTempAssignment '// we have one
			
			if not bCounter mod modnum = 0 then
			%>
				<td width="<%=cint(100/modnum)%>%" class="pcShowAdditionalImage">
					<%pcs_MakeAdditionalImage%>
				</td>
			<% Else %>
				<td width="<%=cint(100/modnum)%>%" class="pcShowAdditionalImage">
					<%pcs_MakeAdditionalImage%>
				</td>
			</tr>
			<%
			end if		
			bCounter = bCounter + 1
	End If	
	
	Next
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END Loop
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	%>
	<% if pcv_strPopWindowOpen <> 1 then %>
	<tr><td colspan="<%=modnum%>" align="center" style="padding-bottom: 10px;">(<i><%=dictLanguage.Item(Session("language")&"_viewPrd_28")%></i>)</td></tr>
	<% end if %>
	</table>
	
	<%
	end if '// end if len(pcf_GetAdditionalImages)>0 then
end if	'// end if len(pImageUrl) > 0 then
%>
<% if pcv_strUseEnhancedViews = True then %>
	<script type="text/javascript">	
		var CurrentImg=1;
        hs.align = '<%=pcv_strHighSlide_Align%>';
        hs.transitions = [<%=pcv_strHighSlide_Effects%>];
        hs.outlineType = '<%=pcv_strHighSlide_Template%>';
        hs.fadeInOut = <%=pcv_strHighSlide_Fade%>;
        hs.dimmingOpacity = <%=pcv_strHighSlide_Dim%>;
        //hs.numberPosition = 'caption';
        <% if bCounter>0 then %>
            if (hs.addSlideshow) hs.addSlideshow({
                slideshowGroup: 'slides',
                interval: <%=pcv_strHighSlide_Interval%>,
                repeat: true,
                useControls: true,
                fixedControls: false,
                overlayOptions: {
                    opacity: .75,
                    position: 'top center',
                    hideOnMouseOut: <%=pcv_strHighSlide_Hide%>
                }
            });	
        <% end if %>
        function pcf_initEnhancement(ele,img) {
            if (document.getElementById('1')==null) {
                hs.expand(ele, { src: img, minWidth: <%=pcv_strHighSlide_MinWidth%>, minHeight: <%=pcv_strHighSlide_MinHeight%> }); 
            } else {
                document.getElementById(CurrentImg).onclick();			
            }
        }
    </script>
<% end if %>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Additional Product Images (If there are any)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Free Shipping Text
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_FreeShippingText
	if scorderlevel <> "0" then
	else
		' Check to see if the product is set for Free Shipping and display message if product is for sale
		if pnoshipping="-1" and (pFormQuantity <> "-1" or NotForSaleOverride(session("customerCategory"))=1) and pnoshippingtext="-1" then 
			response.write "<div class='pcShowProductShipping'>"
			response.write dictLanguage.Item(Session("language")&"_viewPrd_8")
			response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Free Shipping Text
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  BTO ADDON
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_BTOADDON
	if pserviceSpec<>0 then
		if Cdbl(pBtoBPrice)>0 and session("customerType")="1" then
			response.write "<input type=""hidden"" name=""GrandTotal"" value="""&scCurSign&money(pBtoBPrice1)&""">"
			response.write "<input type=""hidden"" name=""TLPriceDefault"" value="""&money(pBtoBPrice1)&""">"
			response.write "<input type=""hidden"" name=""TLPriceDefaultVP"" value="""">"
		else
			response.write "<input type=""hidden"" name=""GrandTotal"" value="""&scCurSign&money(pPrice1)&""">"
			response.write "<input type=""hidden"" name=""TLPriceDefault"" value="""">"
			response.write "<input type=""hidden"" name=""TLPriceDefaultVP"" value="""&money(pPrice1)&""">"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  BTO ADDON
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Out of Stock Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_OutStockMessage
	' if out of stock and show message is enabled (-1) then show message unless stock is ignored
	if (scShowStockLmt=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=0) OR (pserviceSpec<>0 AND scShowStockLmt=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=0) then
		response.write "<div>"&dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_viewPrd_7")& "</div>"
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Out of Stock Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show quantity discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_QtyDiscounts

	if pDiscountPerQuantity=-1 then
		'if customer is retail, check if there are discounts with retail <> 0
		VardiscGo=0
		if session("customerType")=1 then
			query="SELECT discountPerWUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerWUnit>0"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if rs.eof then
				VardiscGo=1
			end if
			set rs=nothing
		else
			query="SELECT discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" AND discountPerUnit>0"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if rs.eof then
				VardiscGo=1
			end if
			set rs=nothing
		end if
	
		if VardiscGo=0 then
			query="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" ORDER BY num"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query) 
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if NOT rs.eof then '// Quick Loop - there will not be too many discounts
				pcv_intTotalDiscounts = 0
				do until rs.eof
					pcv_intTotalDiscounts=pcv_intTotalDiscounts+1
				rs.moveNext		
				loop
				rs.moveFirst
			end if
			%>			
						
						<table class="pcShowList" align="center" style="margin-top: 6px;">
							<tr> 
								<th width="65%"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_1")%></th>
								<th width="35%" nowrap="nowrap"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_2")%>&nbsp;<a href="javascript:optwin('<%=pcv_tmpNewPath%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&SIArray=<%=pIdProduct%>')"><img src="<%=rsIconObj("discount")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_6")%>"></a>
								</th>
							</tr>
							<tr>
								<td colspan="2" class="pcSpacer"></td>
							</tr>
							<% 
							pc_intCounterQ = 0
							do until rs.eof
								pc_intCounterQ = pc_intCounterQ + 1 '// count Discount Rows
								dblQuantityFrom=rs("quantityFrom")
								dblQuantityUntil=rs("quantityUntil")
								dblPercentage=rs("percentage")
								dblDiscountPerWUnit=rs("discountPerWUnit")
								dblDiscountPerUnit=rs("discountPerUnit")
								%>
								<tr>
									<% if dblQuantityFrom=dblQuantityUntil then %>
										<td style="padding-left: 4px;">
											<%=dblQuantityUntil%>&nbsp;<% response.write dictLanguage.Item(Session("language")&"_pricebreaks_4") %>
										</td>
									<% else %>
										<td style="padding-left: 4px;">
											<%=dblQuantityFrom%>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_pricebreaks_3")%>&nbsp;<%=dblQuantityUntil%>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
										</td>
									<% end if %>
										<td>
											<div align="center">
											<% If session("customerType")=1 Then
												If dblPercentage="0" then
													response.write scCurSign & money(dblDiscountPerWUnit)
												else
													response.write dblDiscountPerWUnit & "%"
												End If
												else
												If dblPercentage="0" then
													response.write scCurSign & money(dblDiscountPerUnit)
												else
													response.write dblDiscountPerUnit & "%"
												End If
												end If
											%>
											</div>
									</td>
								</tr>
							<% 							
							if pc_intCounterQ = 6 then '// limit to 6 Rows
								exit do
							end if							
							
							rs.moveNext		
							loop
							set rs=nothing 
							
							'// Display link to full chart
							if pcv_intTotalDiscounts > pc_intCounterQ then 
							%>
							<tr>
								<td colspan="2">
									<div align="left">
											<a href="javascript:optwin('<%=pcv_tmpNewPath%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&SIArray=<%=pIdProduct%>')"><%response.write dictLanguage.Item(Session("language")&"_mainIndex_9")%></a>
									</div>
								</td>
							</tr>
							<% end if %>
						</table>
						
		<% end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show quantity discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: INPUT FIELDS (X)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_OptionsX


							xrequired="0"
							xfieldCnt=0
							xfieldArrCnt=0
							reqstring="" 

							dim isArrCount,tmpCount
							isArrCount=0
							tmpCount=0

				            if tIndex<>0 then ' Check they are updating the product after adding it to the shopping cart
					            pcCartArray=session("pcCartSession")
					            tempIdOpt = ""
					            tempIdOpt = pcCartArray(tIndex,21)
					            if tempIdOpt = "" then
					            else
						            tempIdOpt = Split(trim(tempIdOpt),"<br>")
						            xfieldArrCnt=Ubound(tempIdOpt)
						            isArrCount=xfieldArrCnt
						            if xfieldArrCnt=0 then isArrCount=1
					                for xfieldCnter = 0 to Ubound(tempIdOpt)
					                    tempIdOpt(xfieldCnter) = mid(tempIdOpt(xfieldCnter),instr(1,tempIdOpt(xfieldCnter),": ")+2)
					                    tempIdOpt(xfieldCnter) = replace(tempIdOpt(xfieldCnter),"''","'")
					                    tempIdOpt(xfieldCnter) = replace(tempIdOpt(xfieldCnter),"<BR>",vbcrlf)
					                next
					            end if	
					        end if
						
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start pxfield #1
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if pxfield1<>"0" then
								'select from the database more info 
								query= "SELECT xfield,textarea,widthoffield,rowlength,maxlength FROM xfields WHERE idxfield="&pxfield1
								set rs=server.createobject("adodb.recordset")
								set rs=conntemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rs=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if

								if not rs.EOF then '// Check for no field in DB, although the field is referenced by the product
									xField=rs("xfield")
									TextArea=rs("textarea")
									widthoffield=rs("widthoffield")
									rowlength=rs("rowlength")
									maxlength=rs("maxlength")
									set rs=nothing
									tmpCount=tmpCount+1
									
									if px1req="-1" then
										xfieldCnt=xfieldCnt+1
										xrequired="1"
										reqstring=reqstring&"additem.xfield1.value,'"&replace(xField,"'","")&"'"
									end if
									%>
									
                                    <input type="hidden" name="xf1" value="<%=pxfield1%>">
                                    <%=xField%>
                                    
                                    <% if TextArea="-1" then %>
                                        <br>
                                        <textarea name="xfield1" cols="<%=widthoffield%>" rows="<%=rowlength%>" style="margin-top: 6px" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'1',<%=maxlength%>);"<%end if%>><%if tIndex<>0 and (xfieldArrCnt > 0 or isArrCount>0) then response.write tempIdOpt(0) end if %></textarea>
                                        <%if maxlength>"0" then%>
                                        <br>
                                        <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar1" name="countchar1" style="font-weight: bold"><%=maxlength%></span> <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
                                        <%end if%>
                                    <% else %>
                                        <br>
                                        <input type="text" name="xfield1" size="<%=widthoffield%>" maxlength="<%=maxlength%>" style="margin-top: 6px" <%if tIndex<>0 and (xfieldArrCnt > 0 or isArrCount>0) then%> value="<% response.write tempIdOpt(0) %>" <%end if %> <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'1',<%=maxlength%>);"<%end if%>>
                                        <%if maxlength>"0" then%>
                                        <br>
                                        <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar1" name="countchar1" style="font-weight: bold"><%=maxlength%></span> <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
                                        <%end if%>
                                    <% end if %>
								<br> <br>
							<% 
								end if ' rs.eof
							end if ' pxfield1<>"0"
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// End pxfield #1
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~							
							'// Start pxfield #2
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if pxfield2<>"0" then 
								'select from the database more info 
								query= "SELECT xfield,textarea,widthoffield,rowlength,maxlength FROM xfields WHERE idxfield="&pxfield2
								set rs=server.createobject("adodb.recordset")
								set rs=conntemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rs=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								
								if not rs.EOF then '// Check for no field in DB, although the field is referenced by the product
									xField=rs("xfield")
									TextArea=rs("textarea")
									widthoffield=rs("widthoffield")
									rowlength=rs("rowlength")
									maxlength=rs("maxlength")
									set rs=nothing
									tmpCount=tmpCount+1
									
									if px2req="-1" then
										xfieldCnt=xfieldCnt+1
										if xrequired="1" then
											reqstring=reqstring&","
										end if
										xrequired="1"
										reqstring=reqstring&"additem.xfield2.value,'"&replace(xField,"'","")&"'"
									end if 
									%>
									 
									<input type="hidden" name="xf2" value="<%=pxfield2%>">
									<%=xField%>
									<% if TextArea="-1" then%>
										<br> 
										<textarea name="xfield2" cols="<%=widthoffield%>" rows="<%=rowlength%>" style="margin-top: 6px" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'2',<%=maxlength%>);"<%end if%>><%if tIndex<>0 and xfieldArrCnt > 0  then response.write tempIdOpt(1) end if %></textarea>
										<%if maxlength>"0" then%>
										<br>
										<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar2" name="countchar2" style="font-weight: bold"><%=maxlength%></span> <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
										<%end if%>
									<% else %>
										<br> 
										<input type="text" name="xfield2" size="<%=widthoffield%>" maxlength="<%=maxlength%>" style="margin-top: 6px" <%if tIndex<>0 and xfieldArrCnt > 0 then%> value="<% response.write tempIdOpt(1) %>" <%end if %> <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'2',<%=maxlength%>);"<%end if%>>
										<%if maxlength>"0" then%>
										<br>
										<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar2" name="countchar2" style="font-weight: bold"><%=maxlength%></span> <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
										<%end if%>
									<% end if %>
									<br><br>
							<% 
								end if ' rs.eof
							end if ' pxfield2<>"0"
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// End pxfield #2
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~							
							'// Start pxfield #3
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if pxfield3<>"0" then 
								'select from the database more info 
								query= "SELECT xfield,textarea,widthoffield,rowlength,maxlength FROM xfields WHERE idxfield="&pxfield3
								set rs=server.createobject("adodb.recordset")
								set rs=conntemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rs=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								
								if not rs.EOF then '// Check for no field in DB, although the field is referenced by the product
									xField=rs("xfield")
									TextArea=rs("textarea")
									widthoffield=rs("widthoffield")
									rowlength=rs("rowlength")
									maxlength=rs("maxlength")
									set rs=nothing
									tmpCount=tmpCount+1
									
									if px3req="-1" then
										xfieldCnt=xfieldCnt+1
										if xrequired="1" then
											reqstring=reqstring&","
										end if
										xrequired="1"
										reqstring=reqstring&"additem.xfield3.value,'"&replace(xField,"'","")&"'"
									end if
									%>
									<input type="hidden" name="xf3" value="<%=pxfield3%>"> 
									<%=xField%>
									<% if TextArea="-1" then %>
										<br>
										<textarea name="xfield3" cols="<%=widthoffield%>" rows="<%=rowlength%>" style="margin-top: 6px" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'3',<%=maxlength%>);"<%end if%>><%if tIndex<>0 and xfieldArrCnt > 1 then response.write tempIdOpt(2) end if %></textarea>
										<%if maxlength>"0" then%>
										<br>
										<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar3" name="countchar3" style="font-weight: bold"><%=maxlength%></span> <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
										<%end if%>
									<% else %>
										<br>
										<input type="text" name="xfield3" size="<%=widthoffield%>" maxlength="<%=maxlength%>" style="margin-top: 6px" <%if tIndex<>0 and xfieldArrCnt > 1 then%> value="<% response.write tempIdOpt(2) %>" <%end if %> <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'3',<%=maxlength%>);"<%end if%>>
										<%if maxlength>"0" then%>
										<br>
										<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar3" name="countchar3" style="font-weight: bold"><%=maxlength%></span> <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
										<%end if%>
									<% end if %>
									<br><br>
							<% 
								end if ' rs.eof
							end if ' pxfield3<>"0"
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// End pxfield #3
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if tmpCount>0 then%>
					<script>
					function testchars(tmpfield,idx,maxlen)
					{
						var tmp1=tmpfield.value;
						if (tmp1.length>maxlen)
						{
							alert("<%response.write dictLanguage.Item(Session("language")&"_CheckTextField_1")%>" + maxlen + "<%response.write dictLanguage.Item(Session("language")&"_CheckTextField_1a")%>");
							tmp1=tmp1.substr(0,maxlen);
							tmpfield.value=tmp1;
							document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
							tmpfield.focus();
						}
						document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
					}
					</script>
					<%end if		
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: INPUT FIELDS (X)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show WishList
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_WishList
	if scWL=-1 then
		if pserviceSpec=0 then 
		
			'// Form the link that gets attached to the wishlist button
			pcv_strWishListLink =  "location='Custwl.asp?OptionGroupCount="&pcv_intOptionGroupCount&"&idproduct="&pIdProduct
			Dim bCounter
			Do until bCounter = pcv_intOptionGroupCount
				bCounter = bCounter + 1
				pcv_strWishListLink = pcv_strWishListLink & "&idOption"&bCounter&"='+"
				'if optionN <> "" then
				pcv_strWishListLink = pcv_strWishListLink & "document.additem.idOption"&bCounter&".value+'"
				'end if	
			Loop
			pcv_strWishListLink = pcv_strWishListLink & "';"
			pcv_strFuntionCall = "cdDynamic"
			if xOtionrequired = "1" then '// If there are any required options at all.
			'// figure some stuff out
			%>
				<a href="javascript: if (checkproqty(document.additem.quantity)) { if (<%=pcv_strFuntionCall%>(<%=pcv_strReqOptString%>,1)) {<%=pcv_strWishListLink%>;}};"><img src="<%=pcv_tmpNewPath%><%=rslayout("addtowl")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_7")%>"></a>
			<% else %>
				<a href="javascript:<%=pcv_strWishListLink%>"><img src="<%=pcv_tmpNewPath%><%=rslayout("addtowl")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_7")%>"></a>
			<%end if 
		end if 
	end if 
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show WishList
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show tell-a-friend
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_TellaFriend
	if scTF<>0 then
	%>
		<a href="tellafriend.asp?idproduct=<%=pIdProduct%>" rel="nofollow"><img src="<%=pcv_tmpNewPath%><%=rslayout("checkoutbtn")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_8")%>"></a>
	<%
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show tell-a-friend
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Customize Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CustomizeButton
Dim rsQ,queryQ
	queryQ="SELECT TOP 1 configProduct FROM configSpec_products WHERE specProduct=" & pIdProduct & ";"
	set rsQ=connTemp.execute(queryQ)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsQ=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rsQ.eof then
		showCustomize=1%>
		<a href="javascript:document.additem.action='configurePrd.asp?idproduct=<%=pIdProduct%>';document.additem.submit();"><img src="<%=pcv_tmpNewPath%><%=rslayout("customize")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_altTag_9")%>"></a>
	<%end if
	set rsQ=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Customize Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Required Cross Selling Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim cs_imageheight, cs_imagewidth, cs_ViewCnt, pcv_strHaveResults, pcv_intProductCount, pcArray_CSRelations, pcv_intCategoryActive, pcv_intAccessoryActive

Public Sub pcs_RequiredCrossSelling

    xCSCnt = 0
    pcv_strCSString = ""
    pcv_strReqCSString = ""
    cs_RequiredIds = ""
    pcv_strPrdDiscounts = ""
    pcv_strCSDiscounts = ""
	
	Dim pcv_strGetSitewide
	Dim pcv_strIsBundleActiveFlag

	'// Get Cross Sell Settings - Product Level
	query= "SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,cs_ImageHeight,cs_ImageWidth,crossSellText,cs_ProductViewCnt FROM crossSelldata WHERE id=" & pIdProduct
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	pcv_strGetSitewide=0
	if NOT rs.eof then
		scCS=rs("cs_status")
		cs_showprod=rs("cs_showprod")
		cs_showcart=rs("cs_showcart")
		cs_showimage=rs("cs_showimage")
		cs_imageheight=rs("cs_imageheight")
		cs_imagewidth=rs("cs_imagewidth")
		crossSellText=rs("crossSellText")
		cs_ViewCnt=rs("cs_ProductViewCnt")
	else
		pcv_strGetSitewide=1			
	end if	
	set rs=nothing	
	
	'// Get Cross Sell Settings - Sitewide 
	If pcv_strGetSitewide=1 Then			
		query= "SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,cs_ImageHeight,cs_ImageWidth,crossSellText,cs_ProductViewCnt FROM crossSelldata WHERE id=1;"
		set rs=server.createobject("adodb.recordset")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if NOT rs.eof then
			scCS=rs("cs_status")
			cs_showprod=rs("cs_showprod")
			cs_showcart=rs("cs_showcart")
			cs_showimage=rs("cs_showimage")
			cs_imageheight=rs("cs_imageheight")
			cs_imagewidth=rs("cs_imagewidth")
			crossSellText=rs("crossSellText")
			cs_ViewCnt=rs("cs_ProductViewCnt")
		end if
		set rs=nothing
	 End If 
	
	If scCS=-1 AND cs_showProd="-1" Then		
	
		If cs_ViewCnt < 1 then
			cs_ViewCnt = 2
		End if			
		
		query="SELECT cs_relationships.idproduct, cs_relationships.idrelation, cs_relationships.cs_type, cs_relationships.discount, cs_relationships.ispercent,cs_relationships.isRequired, products.servicespec, products.price, products.description, products.bToBprice, products.serviceSpec, products.noprices FROM cs_relationships INNER JOIN products ON cs_relationships.idrelation=products.idProduct WHERE (((cs_relationships.idproduct)="&pidproduct&") AND ((products.active)=-1) AND ((products.removed)=0)) ORDER BY cs_relationships.num,cs_relationships.idrelation;"		
		set rs=server.createobject("adodb.recordset")
		set rs=conntemp.execute(query)	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		pcv_strHaveResults=0
		if NOT rs.eof then
			pcArray_CSRelations = rs.getRows()
			pcv_intProductCount = UBound(pcArray_CSRelations,2)+1
			pcv_strHaveResults=1
		end if
		set rs=nothing		
		
		tCnt=Cint(0)	
		
		if pcv_strHaveResults=1 then
			do while (tCnt < pcv_intProductCount)					

				pidrelation=pcArray_CSRelations(1,tCnt) '// rs("idrelation")
				pcsType=pcArray_CSRelations(2,tCnt) '// rs("cs_type")			
				pDiscount=pcArray_CSRelations(3,tCnt) '// rs("discount")
				pIsPercent=pcArray_CSRelations(4,tCnt) '// rs("isPercent")
				pcv_strIsRequired=pcArray_CSRelations(5,tCnt) '// rs("isRequired")
				cs_pserviceSpec=pcArray_CSRelations(6,tCnt) '// rs("servicespec")
				
				ppPrice=pcArray_CSRelations(7,tCnt) '// rs("price")
				
				if pcArray_CSRelations(9,tCnt)>"0" then
					ppBPrice=pcArray_CSRelations(9,tCnt)
				else
					ppBPrice=ppPrice
				end if
				
				cs_pserviceSpec=pcArray_CSRelations(10,tCnt)
				if cs_pserviceSpec="" OR IsNull(cs_pserviceSpec) then
					cs_pserviceSpec=0
				end if
				cs_pnoprices=pcArray_CSRelations(11,tCnt)
				if cs_pnoprices="" OR IsNull(cs_pnoprices) then
					cs_pnoprices=0
				end if
				
				pPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
				pBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
				if session("customertype")=1 and pBtoBPrice1>0 then
					pPrice1=pBtoBPrice1
				end if
				
				tmp_pidProduct=pidProduct
				tmp_pPrice=pPrice
				tmp_pPrice1=pPrice1
				tmp_pBtoBPrice=pBtoBPrice
				tmp_pBtoBPrice1=pBtoBPrice1
				tmp_pnoprices=pnoprices
				tmp_pserviceSpec=pserviceSpec
				
				pidProduct=pidrelation
				pPrice=ppPrice
				pBtoBPrice=ppBPrice
				pnoprices=cs_pnoprices
				pserviceSpec=cs_pserviceSpec
				
				ppPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
				ppBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
				if session("customertype")=1 and ppBtoBPrice1>0 then
					ppPrice1=ppBtoBPrice1
				end if
				
				pidProduct=tmp_pidProduct
				pPrice=tmp_pPrice
				pPrice1=tmp_pPrice1
				pBtoBPrice=tmp_pBtoBPrice
				pBtoBPrice1=tmp_pBtoBPrice1
				pnoprices=tmp_pnoprices
				pserviceSpec=tmp_pserviceSpec
				
				tCnt=tCnt+1
				
				'Store ALL Ids
				pcv_strCSString = pcv_strCSString & pidrelation & ","
				xCSCnt = xCSCnt + 1
				
				if pIsPercent<>0 then
					pcv_strPrdDiscounts = pcv_strPrdDiscounts & CDbl(pPrice1*(pDiscount/100)) & ","
					pcv_strCSDiscounts = pcv_strCSDiscounts & CDbl(ppPrice1*(pDiscount/100)) & ","
				else
					pcv_strCSDiscounts = pcv_strCSDiscounts & CDbl(pDiscount) & ","
					pcv_strPrdDiscounts = pcv_strPrdDiscounts & "0,"
				end if
				
				cs_RequiredIds = cs_RequiredIds & pcv_strIsRequired & ","
				
				'// Clear Variables
				cs_pserviceSpec=""
				pidrelation=""
				pcsType=""
			
			loop		
			
			
			if len(pcv_strCSString) > 0 then
				pcv_strCSString = left(pcv_strCSString,len(pcv_strCSString)-1)
			end if
			if len(pcv_strPrdDiscounts) > 0 then
				pcv_strPrdDiscounts = left(pcv_strPrdDiscounts,len(pcv_strPrdDiscounts)-1)
			end if
			if len(pcv_strCSDiscounts) > 0 then
				pcv_strCSDiscounts = left(pcv_strCSDiscounts,len(pcv_strCSDiscounts)-1)
			end if
			if len(cs_RequiredIds) > 0 then
				cs_RequiredIds = left(cs_RequiredIds,len(cs_RequiredIds)-1)
			end if
			
		end if
		
	End if '// If cint(cs_pOptCnt) <> cint(cs_pCnt) Then
	%>
	<input name="pCSCount" type="hidden" value="<%=xCSCnt%>">
	<input name="pCrossSellIDs" type="hidden" value="<%=pcv_strCSString%>">
	<input name="pPrdDiscounts" type="hidden" value="<%=pcv_strPrdDiscounts%>">
	<input name="pCSDiscounts" type="hidden" value="<%=pcv_strCSDiscounts%>">
	<input name="pRequiredIDs" type="hidden" value="<%=cs_RequiredIds%>">
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Required Cross Selling Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Cross Selling With Discounts (Bundles)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CrossSellingDiscounts
	
	'// Only run when having "Add to Cart" button
	IF scCS=-1 AND cs_showProd="-1" THEN 
								
		dim cs_count,cs_pCnt, cs_pOptCnt,cs_pAddtoCart		
		
		tCnt=Cint(0)	
		if pcv_strHaveResults=1 then
			
			cs_pCnt=Cint(0)
			cs_pOptCnt=Cint(0)
			cs_pAddtoCart=Cint(0)
			pcv_intCategoryActive=2	'// set bundle group to inactive
			pcv_intAccessoryActive=2 '// set accessories group to inactive
			cs_count=Cint(0)
			session("listcross")=""
			
			do while ( (tCnt < pcv_intProductCount) AND (tCnt < cs_ViewCnt))				
				
				pidrelation=pcArray_CSRelations(1,tCnt) '// rs("idrelation")
				pcsType=pcArray_CSRelations(2,tCnt) '// rs("cs_type")			
				pDiscount=pcArray_CSRelations(3,tCnt) '// rs("discount")
				cs_pserviceSpec=pcArray_CSRelations(6,tCnt)				
				pcArray_CSRelations(8,tCnt) = 1
				
				If (pcsType="Accessory") OR ((pcsType="Bundle") AND (pDiscount>0)) Then
					
					'// CHECK IF BUNDLES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY		
					'// CHECK IF ACCESSORIES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY  						
					If Session("customerType")=1 Then
						pcv_strCSTemp=""
					else
						pcv_strCSTemp=" AND pccats_RetailHide<>1 "
					end if									
					query="SELECT categories_products.idProduct "
					query=query+"FROM categories_products " 
					query=query+"INNER JOIN categories "
					query=query+"ON categories_products.idCategory = categories.idCategory "
					query=query+"WHERE categories_products.idProduct="& pidrelation &" AND iBTOhide=0 " & pcv_strCSTemp & " "
					query=query+"ORDER BY priority, categoryDesc ASC;"	
					set rsCheckCategory=server.CreateObject("ADODB.RecordSet")
					set rsCheckCategory=conntemp.execute(query)									
					If NOT rsCheckCategory.eof Then
						If pcsType="Accessory" Then
							pcv_intAccessoryActive=1
						End If
						If pcsType="Bundle" Then							
							pcv_intCategoryActive=1
						End If	
					Else
						session("listcross")=session("listcross") & "," & pidrelation					
					End If	
					set rsCheckCategory=nothing
					
				End If '// If (pcsType="Bundle") AND (pDiscount>0) Then	

				pcv_intOptionsExist=0
				
				'// CHECK FOR REQUIRED OPTIONS							
				pcv_intOptionsExist=pcf_CheckForReqOptions(pidrelation) '// check options function (1=YES, 2=NO)			


				'// CHECK FOR REQUIRED INPUT FIELDS
				if pcv_intOptionsExist=2 then
					pcv_intOptionsExist=pcf_CheckForReqInputFields(pidrelation)
				end if				


				'// VALIDATE
				if (cs_pserviceSpec=true) OR (pcv_intOptionsExist = 1) then
					If pcsType<>"Accessory" Then
						cs_pOptCnt=cs_pOptCnt+1
					End If
					pcArray_CSRelations(8,tCnt) = 0					
				End If	
				If pcsType<>"Accessory" Then
					cs_pCnt=cs_pCnt+1 
				End If
				tCnt=tCnt+1				
			loop
		
		end if '// if pcv_strHaveResults=1 then		

					
		'// If ALL items are either BTO or have options or inactive, do not show items
		if (cint(cs_pOptCnt) <> cint(cs_pCnt)) AND (pcv_intCategoryActive=1) then
			
			cs_DisplayCheckBox=-1
			cs_Bundle=-1
			
			%>
			<table class="pcShowContent">
				<tr> 
					<td colspan="2" class="pcSpacer"></td>
				</tr>
				<tr> 
					<td colspan="2" class="pcSectionTitle">
						<%=dictLanguage.Item(Session("language")&"_viewPrd_cs1")%>
					</td>
				</tr>
				<tr> 
					<td colspan="2">
						<%=dictLanguage.Item(Session("language")&"_viewPrd_cs2")%>
					</td>
				</tr>
				<tr> 
					<td colspan="2">
					<% if cs_showImage="-1" then %>
						<!--#include file="cs_img.asp"-->
					<% else %>
						<!--#include file="cs.asp"-->
					<% end if %></td>
				</tr>
			</table>
			<% 
		end if
	
	END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Cross Selling With Discounts (Bundles)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Cross Selling With Accessories
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_CrossSellingAccessories
	
	'// Only run when having "Add to Cart" button
	IF scCS=-1 AND cs_showProd="-1" THEN 
								
		dim cs_pAddtoCart
		
		if pcv_strHaveResults=1 then

			cs_DisplayCheckBox=-1
			cs_Bundle=0
			
			if pcv_intAccessoryActive=1 then 
			%>
			<table class="pcShowContent">
				<tr> 
					<td colspan="2" class="pcSpacer"></td>
				</tr>
				<tr> 
					<td colspan="2" class="pcSectionTitle">
						<%=crossSellText%>
					</td>
				</tr>
				<tr> 
					<td colspan="2">
                    	<%
						if showAddtoCart=1 then
							response.write dictLanguage.Item(Session("language")&"_viewPrd_cs3")
						%>&nbsp;(<img src="<%=pcv_tmpNewPath%><%=rsIconObj("requiredicon")%>">)<%=dictLanguage.Item(Session("language")&"_viewPrd_cs4")%>
                        <%
						else
							if (pserviceSpec<>0) AND (showCustomize=1) then
							else
								response.write pDescription & dictLanguage.Item(Session("language")&"_viewPrd_cs10")
							end if
						end if
						%>
					</td>
				</tr>
				<tr> 
					<td colspan="2" class="pcSpacer"></td>
				</tr>
				<tr> 
					<td colspan="2">
					<% if cs_showImage="-1" then %>
						<!--#include file="cs_img.asp"-->
					<% else %>
						<!--#include file="cs.asp"-->
					<% end if %>
					</td>
				</tr>
			</table>
			<% 
			end if
			
		end if
	
	END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Cross Selling With Accessories
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Long Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_LongProductDescription 
' Display long product description if there is a short description
	if psDesc <> "" then %>
		<table class="pcShowContent">
			<tr> 
				<td><a name="details">&nbsp;</a></td>
			</tr>
			<tr> 
				<td class="pcSectionTitle">
					<%=dictLanguage.Item(Session("language")&"_viewPrd_22")%>
				</td>
			</tr>
			<tr> 
				<td style="padding:8px;">
					<%=pDetails%>
				</td>
			</tr>
			<tr>
				<td>
					<%
					response.write "<div align='right'><a href='#top'>" & dictLanguage.Item(Session("language")&"_viewPrd_23") & "</a></div>"
					%>
				</td>
			</tr>
		</table>
<%
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Long Product Description
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>



<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Options (N)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_OptionsN
	' SELECT DATA SET
	' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
	query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
	query = query & "FROM products "
	query = query & "INNER JOIN ( "
	query = query & "pcProductsOptions INNER JOIN ( "
	query = query & "optionsgroups "
	query = query & "INNER JOIN options_optionsGroups "
	query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
	query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
	query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
	query = query & "WHERE products.idProduct=" & pidProduct &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order, optionsGroups.OptionGroupDesc;"
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	' If we have data	
	if NOT rs.eof then
		pcv_intOptionGroupCount = 0 '// keeps count of the number of options
		xOptionsCnt = 0 '// keeps count of the number of required options
		do until rs.eof				
			
			'if pcv_intOptionGroupCount <= 5  then ' // start limit to 5 options
				'// Get the Group Name
				pcv_strOptionGroupDesc=rs("OptionGroupDesc")
				'// Get the Group ID
				pcv_strOptionGroupID=rs("idOptionGroup")
				'// Is it required
				pcv_strOptionRequired=rs("pcProdOpt_Required")			
		
				'// Start: Do Option Count
				pcv_intOptionGroupCount = pcv_intOptionGroupCount + 1 
				'// End: Do Option Count
				
				'// Get the number of the Option Group
				pcv_strOptionGroupCount = pcv_intOptionGroupCount
				
				'// Start: Do Required Option Count AND generate validation string
				if IsNull(pcv_strOptionRequired) OR pcv_strOptionRequired="" then
						pcv_strOptionRequired=0 '// not required // else it is "1"
				end if			
				if pcv_strOptionRequired=1 then
					
					' Keep Tally
					xOptionsCnt = xOptionsCnt + 1
					
					' Generate String
					if xOtionrequired="1" then
						pcv_strReqOptString = pcv_strReqOptString & ","
					end if
				
					xOtionrequired="1"
					pcv_strOptionGroupDesc2=pcv_strOptionGroupDesc
					pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"'","")
					pcv_strOptionGroupDesc2=replace(pcv_strOptionGroupDesc2,"""","\'\'")
					pcv_strReqOptString = pcv_strReqOptString & "document.additem.idOption" & pcv_strOptionGroupCount & ".selectedIndex,'"& pcv_strOptionGroupDesc2 &"'"
				
				end if
				'// End: Do Required Option Count
			
				'// Make the Option Box
				pcs_makeOptionBox							
		
			'end if ' // end limit to 5 options
		rs.movenext
		loop		
	end if
	set rs=nothing
%>
<input type="hidden" name="OptionGroupCount" value="<%=pcv_intOptionGroupCount%>">
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Options (N)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Options Box
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_makeOptionBox
	' SELECT DATA SET
	' TABLES: options_optionsGroups, options
	query = 		"SELECT options_optionsGroups.InActive, options_optionsGroups.price, options_optionsGroups.Wprice, "
	query = query & "options_optionsGroups.idoptoptgrp, options.idoption, options.optiondescrip "
	query = query & "FROM options_optionsGroups "
	query = query & "INNER JOIN options "
	query = query & "ON options_optionsGroups.idOption = options.idOption "
	query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_strOptionGroupID &" "
	query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
	query = query & "ORDER BY options_optionsGroups.sortOrder, options.optiondescrip;"	
	set rs2=server.createobject("adodb.recordset")
	set rs2=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs2=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	' If we have data
	if NOT rs2.eof then
		
		'// clean up the option group description
		if pcv_strOptionGroupDesc<>"" then
			pcv_strOptionGroupDesc=replace(pcv_strOptionGroupDesc,"""","&quot;")
		end if 
		
		'// START SELECT
		pcv_isOptionSelected="" '// Is this option box selected? Fill variable to "1" during the following loop. %>
		<div><%=pcv_strOptionGroupDesc%>:</div>
		<select name="idOption<%=pcv_strOptionGroupCount%>" style="margin-top: 3px;">
			<% '// Only execute when the Remove Option Feature is activated.
            if pcv_strRemoveFeature<>"1" then %>
                <option value=""><%=dictLanguage.Item(Session("language")&"_viewPrd_61")%></option>
            <% end if %>
			<%
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Start Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			do until rs2.eof			
				
				OptInActive=rs2("InActive") ' Is it active?
				if IsNull(OptInActive) OR OptInActive="" then
					OptInActive="0"
				end if
				
				dblOptPrice=rs2("price") '// Price
				dblOptWPrice=rs2("Wprice") '// WPrice
				intIdOptOptGrp=rs2("idoptoptgrp") '// The Id of the Option Group
				intIdOption=rs2("idoption") '// The Id of the Option
				strOptionDescrip=rs2("optiondescrip") '// A description of the Option
		
				'**************************************************************************************************
				' START: Dispay the Options
				'**************************************************************************************************
				if OptInActive="0" then
					If session("customerType")=1 then 
						optPrice=dblOptWPrice
					Else
						optPrice=dblOptPrice
					End If %>
					<option value="<%=intIdOptOptGrp%>" 
						<% 
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: Check if Option should be Selected
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						Dim xIdOptCounter
						
						if tIndex<>0 then ' Check they are updating the product after adding it to the shopping cart
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
							pcCartArray=session("pcCartSession")
							tempIdOpt = ""
							tempIdOpt = pcCartArray(tIndex,11)
							
							if tempIdOpt = "" then
								response.write ">"
							else
								tempIdOpt = Split(trim(tempIdOpt),chr(124))							
								for xIdOptCounter = 0 to Ubound(tempIdOpt)
									if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
										response.write " selected"								
									end if
								next
								response.write ">"
							end if						
						else
							tempIdOpt = ""
							tempIdOpt = request.querystring("idOptionArray")
							
							if tempIdOpt = "" then
								response.write ">"
							else
								tempIdOpt = Split(trim(tempIdOpt),chr(124))
								for xIdOptCounter = 0 to Ubound(tempIdOpt)
									if clng(intIdOptOptGrp) = clng(tempIdOpt(xIdOptCounter)) then
										response.write " selected"
										pcv_isOptionSelected="1"								
									end if
								next
								response.write ">"
							end if						
						end if
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: Check if Option should be Selected
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: Display Option Name
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						response.write strOptionDescrip & "&nbsp;"
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: Display Option Name
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: Display Pricing
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						if optPrice>0 then '// If there is a price thats greater than zero %>
							<%=" - " &dictLanguage.Item(Session("language")&"_prodOpt_1")&" "&scCurSign& money(optPrice)%>  
						<% end if %>
						<% if optPrice<0 then  '// If there is not a price %>
							<%=" - " &dictLanguage.Item(Session("language")&"_prodOpt_2")&" "&scCurSign& money(optPrice)%> 
						<% end if 
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: Display Pricing
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						%>										
					</option>	
				<% end if
				'**************************************************************************************************
				' END: Dispay the Options
				'**************************************************************************************************
				rs2.movenext 
			loop
			set rs2=nothing	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			'// Only execute when the Remove Option Feature is activated.
			if pcv_strAdminPrefix="1" AND pcv_strRemoveFeature="1" then %>		
				<% if pcv_isOptionSelected="1" then %>
					<option value=""></option>
					<option value="">----- Remove Option -----</option>
				<% else %>
					<option value="" selected><%=dictLanguage.Item(Session("language")&"_viewPrd_61")%></option>
				<% end if %>
			<% end if %>
        </select>
	<% end if %>
    <br />
    <br />
<% End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Options Box
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Add to Cart (Dynamic)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddtoCart
pcv_strFuntionCall = "cdDynamic"
%>
<table>
	<tr> 
		<td valign="middle">
			<%
			if tIndex<>0 then '// Check they are updating the product after adding it to the shopping cart
				pcCartArray=session("pcCartSession")
				tempQty = ""
				tempQty = pcCartArray(tIndex,2)
				if tempQty<>"" then
					pcv_intQuantityField=tempQty
				else
					pcv_intQuantityField=1
				end if	
			else
				if pcv_lngMinimumQty <> 0 then
					pcv_intQuantityField=pcv_lngMinimumQty
				else
					pcv_intQuantityField=1
				end if
			end if
			%>
			<%
			'SB S
			if pSubscriptionID > 0 and pSubType <> 2 Then%>
				<input  type="hidden" name="quantity" value="1"> 
			<%else%>
			<input type="text" name="quantity" size="5" maxlength="10" value="<%=pcv_intQuantityField%>">
			<%
			End if 
			'SB E %>
		</td>
			
		<td valign="middle">
		<input type="hidden" name="idproduct" value="<%=pidProduct%>">
		
<% 
'// there is at least one reuqired custom field	
if xrequired="1" then
%>

		<% 
		If BTOCharges=0 then
			if xOtionrequired = "1" then '// If there are any required options at all.
			'// figure some stuff out
			%>
			<a href="" onClick="javascript: if (CheckRequiredCS('<%=pcv_strReqCSString%>')) {if (checkproqty(document.additem.quantity)) {<%=pcv_strFuntionCall%>(<%=pcv_strReqOptString%>,<%=reqstring%>,0);}} return false"><%showAddtoCart=1%><img border="0" src="<%=pcv_tmpNewPath%><%=rslayout("addtocart")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_2b")%>"></a>
			<%
			else ' There are no required options at all.
			%>
			<a href="" onClick="javascript: if (CheckRequiredCS('<%=pcv_strReqCSString%>')) {if (checkproqty(document.additem.quantity)) {<%=pcv_strFuntionCall%>(<%=reqstring%>,0);}} return false"><%showAddtoCart=1%><img border="0" src="<%=pcv_tmpNewPath%><%=rslayout("addtocart")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_2b")%>"></a> 
			<% 
			end if
		end if 
		%>
		
<% 
' There are no required custom fields.
else
%>
		
		<% 
		If BTOCharges=0 then
			if xOtionrequired = "1" then '// If there are any required options at all.
			'// figure some stuff out
			%>
			<a href="" onClick="javascript: if (CheckRequiredCS('<%=pcv_strReqCSString%>')) {if (checkproqty(document.additem.quantity)) {<%=pcv_strFuntionCall%>(<%=pcv_strReqOptString%>,0);}} return false"><%showAddtoCart=1%><img border="0" src="<%=pcv_tmpNewPath%><%=rslayout("addtocart")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_2b")%>"></a>
			<% else %>
			<%showAddtoCart=1%>
			<%
			'SB S
			'// Don't show if subscription and subscriptions disabled
			if (pSubscriptionID = "0" OR scSBStatus="1") OR (pcv_strAdminPrefix="1") then %> 
			<input alt=Add src="<%=pcv_tmpNewPath%><%=rslayout("addtocart")%>" type=image name="add" border="0" id="submit">
			<% 
			end if
			'SB E
			%>
			<% 
			end if
		End if 
		%>
		

<% 
'// End ADD TO CART SECTION
end if 
%>
		
		<% If pserviceSpec<>0 then
		Dim rsQ,queryQ
		queryQ="SELECT TOP 1 configProduct FROM configSpec_products WHERE specProduct=" & pIdProduct & ";"
		set rsQ=connTemp.execute(queryQ)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsQ=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rsQ.eof then
			showCustomize=1%>
			<a href="javascript:document.additem.action='configurePrd.asp?idproduct=<%=pIdProduct%>&qty='+document.additem.quantity.value; document.additem.submit();">
				<img src="<%=pcv_tmpNewPath%><%=rslayout("customize")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_altTag_9")%>">
			</a> 
		<%End if
		set rsQ=nothing
		End If %>
		<!--#include file="inc_addPinterest.asp"-->
		</td>
	</tr>
</table>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Add to Cart (Dynamic)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'SB S
Public Sub pcs_SubscriptionProduct

	If pSubscriptionID <> 0  then
		
	  	If pIsLinked="1" Then
			%> <!--#include file="inc_sb_widget.asp"--> <%
		End If	  

	 	response.write "<input type=""hidden"" name=""pSubscriptionID"" id=""pcSubId"" value="""&pSubscriptionID&""">"
		
	End If
	
End Sub
'SB S


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Product Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_ProductPromotionMsg
	Dim rs,rsQ,query,tmpStr

	query="SELECT pcPrdPro_id,idproduct,pcPrdPro_QtyTrigger,pcPrdPro_DiscountType,pcPrdPro_DiscountValue,pcPrdPro_ApplyUnits,pcPrdPro_PromoMsg,pcPrdPro_ConfirmMsg,pcPrdPro_SDesc,pcPrdPro_IncExcCust,pcPrdPro_IncExcCPrice,pcPrdPro_RetailFlag,pcPrdPro_WholesaleFlag FROM pcPrdPromotions WHERE pcPrdPro_Inactive=0 AND idproduct=" & pIDProduct & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		pcv_HavePrdPromotions=1
		PrdPromoArr=rsQ.getRows()
		set rsQ=nothing
		PrdPromoCount=ubound(PrdPromoArr,2)
		
		tmpIDCode=PrdPromoArr(0,0)
		tmpIDProduct=PrdPromoArr(1,0)
		tmpQtyTrigger=clng(PrdPromoArr(2,0))
		tmpDiscountType=PrdPromoArr(3,0)
		tmpDiscountValue=PrdPromoArr(4,0)
		tmpApplyUnits=PrdPromoArr(5,0)
		tmpConfirmMsg=PrdPromoArr(7,0)
		tmpDescMsg=PrdPromoArr(8,0)
		pcIncExcCust=PrdPromoArr(9,0)
		pcIncExcCPrice=PrdPromoArr(10,0)
		pcv_retail=PrdPromoArr(11,0)
		pcv_wholeSale=PrdPromoArr(12,0)
		
		pcv_Filters=0
		pcv_FResults=0
		'Filter by Customers
		pcv_CustFilter=0
		query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			pcv_Filters=pcv_Filters+1
			pcv_CustFilter=1
		end if
		set rs=nothing
		
		if pcv_CustFilter=1 then
				
		query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode & " and IDCustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			if (pcIncExcCust="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCust="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing
		
		end if
		'End of Filter by Customers
		
		
		'Filter by Customer Categories
		pcv_CustCatFilter=0
		
		query="select idCustomerCategory from pcPPFCustPriceCats where pcPrdPro_id=" & tmpIDCode
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			pcv_Filters=pcv_Filters+1
			pcv_CustCatFilter=1
		end if
		set rs=nothing
		
		if pcv_CustCatFilter=1 then
				
		query="select pcPPFCustPriceCats.idCustomerCategory from pcPPFCustPriceCats, Customers where pcPPFCustPriceCats.pcPrdPro_id=" & tmpIDCode & " and pcPPFCustPriceCats.idCustomerCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if not rs.eof then
			if (pcIncExcCPrice="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCPrice="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing
		
		end if
		'End of Filter by Customer Categories
		
		' Check to see if promotion is filtered by reatil or wholesale.
		if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
			pcv_Filters=pcv_Filters+1
			if pcv_wholeSale = "1" and session("customertype") = 1 then
				pcv_FResults=pcv_FResults+1		
			end if 
			if pcv_retail = "1" and session("customertype") <> 1 Then
				pcv_FResults=pcv_FResults+1
			end if    
		end if
		
		if (pcv_Filters=pcv_FResults) AND PrdPromoArr(6,0)<>"" then%>
			<div class="pcPromoMessage">
				<%=PrdPromoArr(6,0)%>
	    	</div>
		<%end if
	end if
	set rsQ=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Product Promotion
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
