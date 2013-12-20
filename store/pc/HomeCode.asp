<% 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
Dim pIdProduct, pDescription, iAddDefaultWPrice, iAddDefaultPrice, pBtoBPrice, pPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<%

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Product of the Month
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductOfTheMonth

    if pcIntHPFirst<>0 then
  
    %>
    <!-- Product of the Month -->
	<tr>
		<td class="pcSectionTitle"><%response.write dictLanguage.Item(Session("language")&"_mainIndex_12")%></td>
	</tr>

	<tr>
		<td>
			<table class="pcShowProducts">
			   <tr>
					<%
					'Set the product count to zero
					count=0
					
					tCnt=Cint(0)
						
					do while (tCnt < pcv_intProductCount) and (count < 1)
	
						pidProduct=pcArray_Products(0,tCnt)
						pSku=pcArray_Products(1,tCnt)
						pDescription=pcArray_Products(2,tCnt)  
						pPrice=pcArray_Products(3,tCnt)
						pListHidden=pcArray_Products(4,tCnt)
						pListPrice=pcArray_Products(5,tCnt)						   
						pserviceSpec=pcArray_Products(6,tCnt)
						pBtoBPrice=pcArray_Products(7,tCnt)   
						pSmallImageUrl=pcArray_Products(8,tCnt)   
						pnoprices=pcArray_Products(9,tCnt)
						if isNULL(pnoprices) OR pnoprices="" then
							pnoprices=0
						end if
						pStock=pcArray_Products(10,tCnt)
						pNoStock=pcArray_Products(11,tCnt)
						pcv_intHideBTOPrice=pcArray_Products(12,tCnt)
						if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
							pcv_intHideBTOPrice="0"
						end if
						if pnoprices=2 then
							pcv_intHideBTOPrice=1
						end if
						pFormQuantity=pcArray_Products(14,tCnt)
						pcv_intBackOrder=pcArray_Products(15,tCnt)
						pidrelation=pcArray_Products(0,tCnt)						
						'SB S
						Dim objSB 
						Set objSB = New pcARBClass
						pSubscriptionID = objSB.getSubscriptionID(pidProduct)
						if isNull(pSubscriptionID) OR pSubscriptionID="" then
							pSubscriptionID = "0"
						end if					
						'SB E
												
						'// Get sDesc
						query="SELECT sDesc FROM products WHERE idProduct="&pidrelation&";"
						set rsDescObj=server.CreateObject("ADODB.RecordSet")
						set rsDescObj=conntemp.execute(query)
						psDesc=rsDescObj("sDesc")
						set rsDescObj=nothing
						
						if pcPageStyle = "m" then
							pCnt=pCnt+1
						end if
						tCnt=tCnt+1
						%>
						<!--#include file="pcGetPrdPrices.asp"-->
						<%
					
						'*******************************
						' Show product information
						'*******************************
						%>				    
						<td> 
							<!--#include file="pcShowProductP.asp" -->
						</td>
						<%	
						count=count + 1
	
					loop
					%>
                </tr>			
	        </table>
        </td>
    </tr>
		<tr>
			<td class="pcSpacer"></td>
		</tr>

	<%
    end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Product of the Month
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Featured Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_FeaturedProducts
  
  if pcIntHPFeaturedCount > 0 then
 
    %>
    <!-- Featured Products -->
		<tr>
			<td class="pcSectionTitle"><%response.write dictLanguage.Item(Session("language")&"_mainIndex_7")%></td>
		</tr>
	
		<tr>
			<td>
				<% if pcPageStyle = "m" then %>
					<form action="instPrd.asp" method="post" name="m" id="m" class="pcForms">
				<% end if %>
				<table class="pcShowProducts">
				<%
						'*******************************
						' Add table headers for display
						' styles L and M
						'*******************************
				%>
				<% if pcPageStyle = "l" then	%>
						<tr class="pcShowProductsLheader">
						<% if pShowSmallImg <> 0 then %>
							<td>&nbsp;</td>
						<% end if %>
							<td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %></td>
						<% if pShowSku <> 0 then %>
							<td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %></td>
						<% end if %>
							<td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_10") %></td>
						</tr>
					<% elseif pcPageStyle = "m" then %>
						<tr class="pcShowProductsMheader">
							<td colspan="<%if iShow=1 then%>5<%else%>4<%end if%>">
								<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_12") %>
							</td>
						</tr>
						<tr>
						<tr class="pcShowProductsMheader">
							<% if iShow=1 then %> 
								<% if pAddtoCart = 1 then %>
									<td width="8%">
										<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_7") %>
									</td>
								<% end if %>
							<% end if %>
							<% if pShowSmallImg <> 0 then %>
							<td width="8%">&nbsp;</td>
							<% end if %>
							<% if pShowSku <> 0 then %>
							<td width="11%">
								<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %>
							</td>
							<% end if %>
							<td width="47%">
								<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %>
							</td>
							<td width="16%" align="center">
								<% If session("customerType")="1" then
										response.write dictLanguage.Item(Session("language")&"_viewCat_P_11")
									 else
										response.write dictLanguage.Item(Session("language")&"_viewCat_P_10")
								end if %>
							</td>
						</tr>
			 			<% else
						if pcv_intProductCount>0 then %>
						<tr>
						<%end if%>
					<% end if %>
					<%
					'*******************************
					' End table headers
					'*******************************
				
					'*******************************
					' Load product information
					' Loop through the products
					'*******************************
				
					'Set the product count to zero
					count=0
						
					if pcPageStyle = "m" then
						pCnt=Cint(0)
						pSQty=0
						pAllCnt=Cint(0)
					end if
	
					tCnt=Cint(0)					

					'Loop until the total number of products to show
					if pcIntHPFirst<>0 then
						tCnt=tCnt+1
						pcIntHPFeaturedCount=pcIntHPFeaturedCount+1
						count=count + 1
					end if
				
                    do while (tCnt < pcv_intProductCount) and (count < pcIntHPFeaturedCount)

						pidProduct=pcArray_Products(0,tCnt)
						pSku=pcArray_Products(1,tCnt)
						pDescription=pcArray_Products(2,tCnt)   
						pPrice=pcArray_Products(3,tCnt)
						pListHidden=pcArray_Products(4,tCnt)
						pListPrice=pcArray_Products(5,tCnt)						   
						pserviceSpec=pcArray_Products(6,tCnt)
						pBtoBPrice=pcArray_Products(7,tCnt)   
						pSmallImageUrl=pcArray_Products(8,tCnt)   
						pnoprices=pcArray_Products(9,tCnt)
						if isNULL(pnoprices) OR pnoprices="" then
							pnoprices=0
						end if
						pStock=pcArray_Products(10,tCnt)
						pNoStock=pcArray_Products(11,tCnt)
						pcv_intHideBTOPrice=pcArray_Products(12,tCnt)
						if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
							pcv_intHideBTOPrice="0"
						end if
						if pnoprices=2 then
							pcv_intHideBTOPrice=1
						end if
						pFormQuantity=pcArray_Products(14,tCnt)
						pcv_intBackOrder=pcArray_Products(15,tCnt)
						pidrelation=pcArray_Products(0,tCnt)						
												
						'// Get sDesc
						query="SELECT sDesc FROM products WHERE idProduct="&pidrelation&";"
						set rsDescObj=server.CreateObject("ADODB.RecordSet")
						set rsDescObj=conntemp.execute(query)
						psDesc=rsDescObj("sDesc")
						set rsDescObj=nothing
						
						if pcPageStyle = "m" then
							pCnt=pCnt+1
						end if
						tCnt=tCnt+1
						%>
						<!--#include file="pcGetPrdPrices.asp"-->
						<%
   				
						'*******************************
						' Show product information
						' depending on the page style
						'*******************************
							
						' FIRST STYLE - Show products horizontally, with images
						if pcPageStyle = "h" then	%>
							<td> 
								<!--#include file="pcShowProductH.asp" -->
							</td>
							<% i=i + 1
							If i > (scPrdRow-1) then 
								response.write "</TR><TR>"
								i=0
							End If
						end if
					
						' SECOND STYLE - Show products vertically, with images 
						if pcPageStyle = "p" then	%>
							<td> 
								<!--#include file="pcShowProductP.asp" -->
							</td>
						</tr>
						<% end if
					
						' THIRD STYLE - Show a list of products, with a small image 
						if pcPageStyle = "l" then	%>
								<!--#include file="pcShowProductL.asp" -->
						<% end if
					
						' FOURTH STYLE - Show a list of products, with multiple add to cart 
						if pcPageStyle = "m" then	%>
								<!--#include file="pcShowProductM.asp" -->
						<% end if
					
						iRecordsShown=iRecordsShown + 1
						count=count + 1
					loop
					
					if count < cint(pTotalCount) then
						Dim intColSpan
						intColSpan=0
						if pcPageStyle = "l" then
							intColSpan=4
						elseif pcPageStyle = "m" then
							if iShow=1 then
							intColSpan=5
							else
							intColSpan=4
							end if
						else
							intColSpan=scPrdRow
						end if
					%>
						<tr>
							<td colspan="<%=intColspan%>"><a href="showfeatured.asp"><%response.write (dictLanguage.Item(Session("language")&"_mainIndex_13") & dictLanguage.Item(Session("language")&"_mainIndex_7"))%> &gt;&gt;</a></td>
						</tr>	        
					<% end if %>
				</table>

				<% 
				' If page style is M, show the Add to Cart button when
				' products can be added to the cart from this page.	
				if pcPageStyle = "m" then %>
					<input type="hidden" name="pCnt" value="<%=pCnt%>">
					<% if iShow=1 and clng(pSQty)<>0 then %>
						<div style="padding-top: 5px;">
							<input name="submit" type="image" src="<%=rslayout("addtocart")%>" id="submit">
						</div>
					<% end if %>
				</form>
				<% end if %>
			</td>
		</tr>
<% end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Featured Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Best sellers, new arrivals, specials
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowProducts	
	
    dim sArray(3,2), strArray(3,3), iCount(3), i, count, iWidth
    
    'Specials
    sArray(0,0)=cint(pcIntHPSpcOrder)
    sArray(0,1)=cint(pcIntHPSpcCount)
    strArray(0,0)="Specials"      
    strArray(0,1)="_mainIndex_4"    
    strArray(0,2)="showspecials.asp" 
	strArray(0,3)=cint(0)        
    
    'New Arrivals
    sArray(1,0)=cint(pcIntHPNewOrder)
    sArray(1,1)=cint(pcIntHPNewCount)
    strArray(1,0)="New Arrivals" 
    strArray(1,1)="_mainIndex_10" 
    strArray(1,2)="shownewarrivals.asp"
	strArray(1,3)=cint(0)
    
    'Best Sellers
    sArray(2,0)=cint(pcIntHPBestOrder)
    sArray(2,1)=cint(pcIntHPBestCount)
    strArray(2,0)="Best Sellers" 
    strArray(2,1)="_mainIndex_6"  
    strArray(2,2)="showbestsellers.asp" 
	strArray(2,3)=cint(0)

    i = 0
    count = 0
    do while true
        if sArray(i,0) > sArray(i+1,0) then
            tmp = sArray(i,0)
            tmpc = sArray(i,1)
            temp = strArray(i,0)
            tempc = strArray(i,1)
            temph = strArray(i,2)

            sArray(i,0) = sArray(i+1,0)
            sArray(i,1) = sArray(i+1,1)
            strArray(i,0) = strArray(i+1,0)
            strArray(i,1) = strArray(i+1,1)
            strArray(i,2) = strArray(i+1,2)

            sArray(i+1,0) = tmp
            sArray(i+1,1) = tmpc
            strArray(i+1,0) = temp
            strArray(i+1,1) = tempc
            strArray(i+1,2) = temph
        end if
        i=i+1
        count=count+1
        if i=2 then i=0
        if count=5 then exit do
    loop
    
    iWidth = 0
    for i= 0 to 2
        if sArray(i,1) > 0 then iWidth=iWidth+1
    next

    if iWidth > 0 then

%>
        <tr>
            <td>
		        <table class="pcShowProducts">
			        <tr class="pcSectionTitle">
			        <% 
			            for i = 0 to 2
			                if sArray(i,1) > 0 then
			        %>
			                    <td width='<%=Round(100/iWidth,0)%>%' align=center><strong><%response.write dictLanguage.Item(Session("language")&strArray(i,1))%></strong></td>
		            <%     
			                end if
			            next
			        %>
			        </tr>
			        <tr>			        
			        <% 
			            for i = 0 to 2
			                if sArray(i,1) > 0 then
			        %>
								<td style="vertical-align: top;">
									<table>
			        <%                                    
                                Dim rsProducts, queryNFS, pagesize

                                'New Arrivals
                                if strArray(i,0)="New Arrivals" then 
                                    Dim pcIntNewArrNFS, pcIntNewArrInStock, queryInStock
                                    query="SELECT pcNAS_NDays, pcNAS_NotForSale, pcNAS_OutOfStock FROM pcNewArrivalsSettings;"
                                    set rs=Server.CreateObject("ADODB.RecordSet")
                                    set rs=connTemp.execute(query)                                    
									if not rs.eof then
	                                    pcNDays=rs("pcNAS_NDays")
	                                    pcIntNewArrNFS=rs("pcNAS_NotForSale")
	                                    pcIntNewArrInStock=rs("pcNAS_OutOfStock")
                                    end if
                                    set rs=nothing
                                    if isNULL(pcNDays) OR (pcNDays="0") OR (pcNDays="") then
	                                    pcNDays=15
                                    end if
                                    if pcIntNewArrNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 then
	                                    queryNFS = "((products.formQuantity)=0) AND"
                                    else
	                                    queryNFS = " "
                                    end if
                                    '*******************************
                                    ' GET new arrivals from DB
                                    '*******************************
									pcTodayDate=Date()
									if SQL_Format="1" then
										pcTodayDate=Day(pcTodayDate)&"/"&Month(pcTodayDate)&"/"&Year(pcTodayDate)
									else
										pcTodayDate=Month(pcTodayDate)&"/"&Day(pcTodayDate)&"/"&Year(pcTodayDate)  
									end if
									if scdb="SQL" then
										y="'"
									else
										y="#"
									end if
                                    if session("CustomerType")<>"1" then
	                                    query1= " AND ((categories.pccats_RetailHide)=0)"
                                    else
	                                    query1=""
                                    end if
									
									'// START v4.1 - Not For Sale override
										if NotForSaleOverride(session("customerCategory"))=1 then
											queryNFS=""
										else
											queryNFS="((products.formQuantity)=0) AND "
										end if
									'// END v4.1
									
									if pcIntNewArrInStock <> 0 then
										query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl,  products.formQuantity, products.pcprod_EnteredOn FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.stock)>0) AND " & queryNFS & "((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-[products].[pcprod_EnteredOn])<="& pcNDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) OR (((products.noStock)=-1) AND " & queryNFS & "((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-[products].[pcprod_EnteredOn])<="& pcNDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) ORDER BY products.pcprod_EnteredOn DESC;"
									else
										query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl,  products.formQuantity, products.pcprod_EnteredOn FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE ("&queryNFS&"((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-[products].[pcprod_EnteredOn])<="& pcNDays &") AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.pcprod_EnteredOn DESC;"
									end if
                                end if
                                
                                'Specials
                                if strArray(i,0)="Specials" then
                                    Dim pcIntSpecialsNFS
                                    pcIntSpecialsNFS = 0 ' Not for sale items are shown
                                    pcIntSpecialsNFS = -1 ' Not for sale items are not shown
	                                if pcIntSpecialsNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 then
		                                queryNFS = "AND formQuantity=0 "
		                            else
		                                queryNFS = ""
	                                end if
                                    '*******************************
                                    ' GET sorting criteria
                                    '*******************************
                                    Dim querySort
                                    querySort = " ORDER BY products.description Asc" 	
                                    '*******************************
                                    ' GET Specials from DB
                                    '*******************************
                                    if session("CustomerType")<>"1" then
	                                    query1= " AND categories.pccats_RetailHide=0"
                                    else
	                                    query1=""
                                    end if
                                    query="SELECT distinct products.idProduct,products.sku,products.description,products.smallImageUrl FROM products,categories_products,categories WHERE products.active=-1 AND products.hotdeal=-1 AND products.configOnly=0 AND products.removed=0 " & queryNFS & " AND categories_products.idProduct=products.idProduct AND categories.idCategory=categories_products.idCategory AND categories.iBTOhide=0 " & query1 & querySort
                                end if

                                'Best Sellers
                                if strArray(i,0)="Best Sellers" then
                                    Dim pcIntBestSellNFS, pcIntBestSellInStock, pcIntBestSellSales
                                    pcIntBestSellSales=0
                                    query="SELECT pcBSS_BestSellCount,pcBSS_Style,pcBSS_PageDesc,pcBSS_NSold,pcBSS_NotForSale,pcBSS_OutOfStock,pcBSS_SKU,pcBSS_ShowImg FROM pcBestSellerSettings;"
                                    set rs=connTemp.execute(query)
                                    if not rs.eof then
	                                    pcIntBestSellSales=rs("pcBSS_NSold")
	                                    pcIntBestSellNFS=rs("pcBSS_NotForSale")
	                                    pcIntBestSellInStock=rs("pcBSS_OutOfStock")
                                    end if
                                    set rs=nothing
                                    if isNULL(pcIntBestSellSales) or (pcIntBestSellSales="0") then
	                                    pcIntBestSellSales=2
                                    end if
                                    if pcIntBestSellNFS<> 0 and NotForSaleOverride(session("customerCategory"))=0 then
                                        queryNFS = " AND ((products.formQuantity)=0)"
                                    else
                                        queryNFS = " "
                                    end if
                                    '*******************************
                                    ' GET Best Sellers from DB
                                    '*******************************
                                    if session("CustomerType")<>"1" then
	                                    query1= " AND ((categories.pccats_RetailHide)=0)"
                                    else
	                                    query1=""
                                    end if
                                    if pcIntBestSellInStock<> 0 then
	                                    query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl, products.sales, products.formQuantity FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.stock)>0) AND ((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") OR (((products.noStock)=-1) AND ((products.sales)>"&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
                                    else
	                                    query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl, products.sales, products.formQuantity FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
                                    end if
                                end if
								set rsProducts=server.CreateObject("ADODB.Recordset")
								set rsProducts=conntemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsProducts=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								pcv_intProductCount=-1
								if NOT rsProducts.eof then
									pcArray_Products = rsProducts.getRows()
									pcv_intProductCount = UBound(pcArray_Products,2)+1
									strArray(i,3)=cLng(pcv_intProductCount)
								end if
								set rsProducts = nothing

				                'Loop until the total number of products to show
				                count=0
				                iCount(i)=0
								
								tCnt=Cint(0)
								
								do while (tCnt < pcv_intProductCount)
					                
									pIdProduct=pcArray_Products(0,tCnt)
					                pSku=pcArray_Products(1,tCnt)					                
					                pDescription=pcArray_Products(2,tCnt)
									pSmallImageUrl=pcArray_Products(3,tCnt)
					                pDesc=pDescription
									
									tCnt=tCnt+1

													
					                if count < cint(sArray(i,1)) then        
										'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
										%>
										<!--#include file="pcSeoFirstCat.asp"-->
										<%
							
										'// Call SEO Routine
										pcGenerateSeoLinks
										'//

			        					%>
										<tr class="pcShowProductsL"> 
											<% if pShowSmallImg <> 0 then%>
												<td class="pcShowProductsLCell">
													<%if pSmallImageUrl<>"" then%>
														<a href="<%=pcStrPrdLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pIdProduct%>'; sav_callxml='1'; return runXML1('prd_<%=pIdProduct%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><img src="catalog/<%response.write pSmallImageUrl%>" <%if scStoreUseToolTip<>"1" and scStoreUseToolTip<>"2" then%>alt="<%=pDescription%>"<%end if%> class="pcShowProductImageL"></a>
													<% else %>
														&nbsp;
													<%end if %>
												</td>
											<% end if %>
											<td>
												<div class="pcShowProductName">
													<a href="<%=pcStrPrdLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pIdProduct%>'; sav_callxml='1'; return runXML1('prd_<%=pIdProduct%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><%=pDesc%></a>
												</div>
												<%if pShowSKU <> 0 then%>
													<div>
														<%=pSku%>
													</div>
												<% end if %>
											</td>
										</tr>
					 				<%

                                    end if
					                iRecordsShown=iRecordsShown + 1
					                count=count + 1
					                iCount(i) = iCount(i)+1
				                loop
                     %>
		                        
									</table>
								</td>
		            <%     
			                end if			                    
			            next
			        %>
                    </tr>
                    <tr>
			        <% 
			            for i = 0 to 2
			                if (sArray(i,1)>0) then
			                    if (sArray(i,1)<strArray(i,3)) AND (iCount(i)>0) then
			        %>
			                        <td><a href='<%=strArray(i,2)%>'><%response.write (dictLanguage.Item(Session("language")&"_mainIndex_13") & dictLanguage.Item(Session("language")&strArray(i,1)))%> &gt;&gt;</a></td>
		            <%
		                        else     
			        %>
			                        <td>&nbsp;</td>
		            <%     
			                    end if
			                end if
			            next
			        %>
                    </tr>
	            </table>
		    </td>
	    </tr>

<%
    end if
    
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Best sellers, new arrivals, specials
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


%>