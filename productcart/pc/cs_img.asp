<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

	pcv_strIsBundleActiveFlag = False

	'// Check for Bundled Cross Sell Products
	if cs_Bundle=-1 then
		'// Get Discount Bundles
		pcsFilterType="Bundle"
    else
		'// Get Accessories
		pcsFilterType="Accessory"
    end if
%>
	
	<table class="pcShowProducts">
		<tr>

<%
	IF pcv_strHaveResults<>1 THEN
		' No items to display
		%><td><%response.write dictLanguage.Item(Session("language")&"_advSrcb_2")%></td><%
	ELSE

		tCnt=Cint(0)
		
		'// See if Admin set Thumbnail sizes
		if (cs_imageheight > 0) AND (cs_imagewidth > 0) then 
			iWidth=cs_imagewidth
			iHeight=cs_imageheight
		else
			iWidth=""
			iHeight=""
		end if
		
		pcv_intDisplayCounter = Cint(0)
		
		if len(tmp_PList)>0 then
			session("listcross")=session("listcross") & "," & tmp_PList
		else
			session("listcross")=session("listcross")
		end if 
		
		'// Count cells in this row
		Dim pcIntCellCount
		pcIntCellCount=1
		
		do while ((tCnt < pcv_intProductCount) AND (pcv_intDisplayCounter < cs_ViewCnt))
			
			pidrelation=pcArray_CSRelations(1,tCnt) '// rs("idrelation")
			pcsType=pcArray_CSRelations(2,tCnt) '// rs("cs_type")
			pDiscount=pcArray_CSRelations(3,tCnt) '// rs("discount")
			pIsPercent=pcArray_CSRelations(4,tCnt) '// rs("isPercent")
			pcv_strIsRequired=pcArray_CSRelations(5,tCnt) '// rs("isRequired")
			cs_pserviceSpec=pcArray_CSRelations(6,tCnt) '// rs("servicespec")
			ppPrice=pcArray_CSRelations(7,tCnt) '// rs("price")
			if pcsFilterOverRide<>"" then
				pcArray_CSRelations(8,tCnt) = 1
			end if
			
			If (pcsType=pcsFilterType) OR pcsFilterOverRide<>"" Then	
							
				if InStr(","& session("listcross") &",",","& pidrelation &",")=0 then
					
					session("listcross")=session("listcross") & "," & pidrelation 

					'// Query Product
					query="SELECT products.idProduct, products.description, products.sku, products.price,products.listhidden,products.listprice, "
					query=query+"products.serviceSpec,products.bToBPrice,products.noprices,products.pcprod_HideBTOPrice, products.smallImageUrl, products.formQuantity "
					query=query+"FROM products WHERE products.idProduct="&pidrelation
					query=query+"AND active=-1 AND configOnly=0 and removed=0 "
					set cs_rs=Server.CreateObject("ADODB.Recordset")   			 			
					set cs_rs=conntemp.execute(query)	
					if err.number<>0 then
						call LogErrorToDatabase()
						set cs_rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if not cs_rs.eof then
						cs_pidProduct=cs_rs("idProduct")
						pDescription=cs_rs("description")
						psku=cs_rs("sku")
						cs_pPrice=cs_rs("price")
						cs_pListPrice=cs_rs("listprice")
						cs_pListHidden=cs_rs("listhidden")   
						cs_pserviceSpec=cs_rs("serviceSpec")
						cs_pbToBPrice=cs_rs("bToBPrice") 
						cs_pnoprices=cs_rs("noprices")
						cs_pcv_intHideBTOPrice=cs_rs("pcprod_HideBTOPrice")
						pSmallImageUrl=cs_rs("smallImageUrl")
						pNotForSale=cs_rs("formQuantity")
					end if	
					set cs_rs=nothing					
 
					
					'/////////////////////////////////////////////////////////////////////////////////////////
					'// START: PRICING
					'/////////////////////////////////////////////////////////////////////////////////////////
					
					if cdbl(cs_pBtoBPrice)=0 then
						cs_pBtoBPrice=cs_pPrice
					end if  					
					
					if isNULL(cs_pnoprices) OR cs_pnoprices="" then  
						cs_pnoprices=0
					end if					
					
					if isNULL(cs_pcv_intHideBTOPrice) OR cs_pcv_intHideBTOPrice="" then  
						cs_pcv_intHideBTOPrice="0"
					end if
					
					tmp_pidProduct=pidProduct
					tmp_pPrice=pPrice
					tmp_pPrice1=pPrice1
					tmp_pBtoBPrice=pBtoBPrice
					tmp_pBtoBPrice1=pBtoBPrice1
					tmp_pnoprices=pnoprices
					tmp_pserviceSpec=pserviceSpec
					
					pidProduct=cs_pidProduct
					pPrice=cs_pPrice
					pBtoBPrice=cs_pBtoBPrice
					pnoprices=cs_pnoprices
					pserviceSpec=cs_pserviceSpec
					
					cs_dblpcCC_Price=0
					
					%><!--#include file="pcGetPrdPrices.asp"--><%
					
					cs_dblpcCC_Price=dblpcCC_Price
					cs_pPrice=pPrice
					cs_pPrice1=pPrice
					
					pidProduct=tmp_pidProduct
					pPrice=tmp_pPrice
					pPrice1=tmp_pPrice1
					pBtoBPrice=tmp_pBtoBPrice
					pBtoBPrice1=tmp_pBtoBPrice1
					pnoprices=tmp_pnoprices
					pserviceSpec=tmp_pserviceSpec
					
					
					if cs_pnoprices=0 then				
						'// Check for discount per quantity
						query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &cs_pidProduct
						if session("CustomerType")<>"1" then
							query=query & " and discountPerUnit<>0"
						else
							query=query & " and discountPerWUnit<>0"
						end if
						set cs_rsDisc=Server.CreateObject("ADODB.Recordset")
						set cs_rsDisc=conntemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set cs_rsDisc=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if						
						if not cs_rsDisc.eof then
							cs_pDiscountPerQuantity=-1
						else
							cs_pDiscountPerQuantity=0
						end if
						set cs_rsDisc = nothing
					end if 
					'/////////////////////////////////////////////////////////////////////////////////////////
					'// END: PRICING
					'/////////////////////////////////////////////////////////////////////////////////////////




					'/////////////////////////////////////////////////////////////////////////////////////////
					'// START: DISPLAY CROSS SELLING
					'/////////////////////////////////////////////////////////////////////////////////////////
					
					' If item is either BTO, has required accessories, or has required options,
					' do not show item (bundle) or 
					' do not show checkbox (accessory)
					cs_pAddtoCart=Cint(0)
					if pcArray_CSRelations(8,tCnt) = 1 then 
						cs_pAddtoCart = 1
					end if
					
					if ((cs_pAddtoCart=1 AND pcsType<>"Accessory") OR (pcsType="Accessory")) then
									
						'// Call SEO Routine
						pcGenerateSeoLinks
						'//
						%>
						<td> 
							<table class="pcShowProductsHCS">
								<tr>
									<td class="pcShowProductImageH">
									<%if pSmallImageUrl<>"" then%>
										<p><a href="<%=pcStrPrdCSLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pidrelation%>'; sav_callxml='1'; return runXML1('prd_<%=pidrelation%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><img src="catalog/<%response.write pSmallImageUrl%>" <% if trim(iWidth)<>"" then %>width="<%=iWidth%>"<% end if %> <% if trim(iHeight)<>"" then %>height="<%=iHeight%>"<%end if%><%if scStoreUseToolTip<>"1" and scStoreUseToolTip<>"2" then%> alt="<%=pDescription%>"<%end if%>></a></p>
									<% else %>
										<p><a href="<%=pcStrPrdCSLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pidrelation%>'; sav_callxml='1'; return runXML1('prd_<%=pidrelation%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><img src="catalog/no_image.gif" width="<%=iWidth%>" height="<%=iHeight%>" <%if scStoreUseToolTip<>"1" and scStoreUseToolTip<>"2" then%>alt="<%=pDescription%>"<%end if%>></a></p>
									<%end if %>
									</td>
								</tr>
								<tr>
									<td class="pcShowProductInfoH">
										<%
										pcv_intDisplayCounter=pcv_intDisplayCounter+1
										if showAddtoCart=1 then '0
										if cs_DisplayCheckBox=-1 then '1
											if pcsType="Accessory" then '2
												if cs_pAddtoCart=1 then '3
													if pcv_strIsRequired<>0 then '4
														%><img src="<%=rsIconObj("requiredicon")%>"> <%response.write dictLanguage.Item(Session("language")&"_alert_18")%><br /><%
													end if '4
													%><input name="bundle<%=pidrelation%>" type="checkbox" value="<%=pDescription%>" class="clearBorder" <%if pcv_strIsRequired<>0 then%>checked<%end if%> <%if pcv_strIsRequired<>0 then%> disabled <%end if%> ><%
												end if '3
											else '2
												%>
												<input name="rdoBundle" type="radio" value="<%=pidrelation%>" id="<%=pDescription%>" class="clearBorder">
												<%
												pcv_strIsBundleActiveFlag = True '// Set a Flag that we have at least one bundle
											end if '2
												%><input name="bundlePrd<%=pidrelation%>" type="hidden" value="<%=pidrelation%>"><%
										end if '1
										end if '0
										%>
											
										<p class="pcShowProductName">
											<%
											if pcsType<>"Accessory" and trim(pMainProductName)<>"" then
												response.write pMainProductName & " + "
											end if
											%>
											<a href="<%=pcStrPrdCSLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pidrelation%>'; sav_callxml='1'; return runXML1('prd_<%=pidrelation%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><%=pDescription%></a>
										</p>
			
										<% 
										'//////////////////////////////////////////
										'// Start: Not For Sale
										'//////////////////////////////////////////
										
										'The following conditional statement was altered in v4.5 to allow the purchase
										'of Not for Sale items when they are part of a bundle or an accessory
										'if pNotForSale = 0 or NotForSaleOverride(session("customerCategory"))=1 then 
										   
										   if cs_Bundle=-1 then
											   '// Calculate Discounts and savings
											   if pIsPercent<>0 then
												   pSavings=CDbl(cs_dblpcCC_Price+pPrice1)*CDbl(pDiscount/100)
											   else
												   pSavings=pDiscount
											   end if
											   cs_pPrice1=CDbl(cs_dblpcCC_Price+pPrice1)-pSavings
										   end if
			 
											if (cs_pPrice1>0) and (cs_pcv_intHideBTOPrice<>"1") then %>
												<%ShowSaleIcon=0
			
												if UCase(scDB)="SQL" then	
													if pnoprices=0 then
														query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & cs_pidProduct & " AND Products.pcSC_ID>0;"
														set rsS=Server.CreateObject("ADODB.Recordset")
														set rsS=conntemp.execute(query)
					
														if not rsS.eof then
															ShowSaleIcon=1
															pcSCID=rsS("pcSC_ID")
															pcSCName=rsS("pcSC_SaveName")
															pcSCIcon=rsS("pcSC_SaveIcon")
															pcTargetPrice=rsS("pcSales_TargetPrice") %>
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
												end if%>
												<p class="pcShowProductPrice">
												<%response.write dictLanguage.Item(Session("language")&"_prdD1_1") & " " & scCursign & money(cs_pPrice1)%>
												<%if (ShowSaleIcon=1) AND (session("customerCategory")=0) AND (pcTargetPrice="0") AND (session("customerType")="0") then%>
												<span class="pcSaleIcon"><a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
												<%end if%>
												<!-- Load quantity discount icon -->
												<!--#Include File="pcShowQtyDiscIconCS.asp" -->
												</p>
												<%if cs_Bundle=-1 then %>
													<%if (pSavings)>0 AND session("customerType")<>1 then %>
														<p class="pcShowProductSavings">
														<% response.write dictLanguage.Item(Session("language")&"_prdD1_2") & scCursign & money(pSavings)%>
														</p>
													<% end if%>
												<% end if
											end if 
			
											'if customer category type logged in - show pricing
											if session("customerCategory")<>0 and (cs_dblpcCC_Price>"0") and (cs_pcv_intHideBTOPrice<>"1") then %>
												<p class="pcShowProductPriceW">
													<% response.write session("customerCategoryDesc")& " " & scCursign & money(cs_dblpcCC_Price)%>
													<%if (ShowSaleIcon=1) AND (clng(session("customerCategory"))=clng(pcTargetPrice)) AND (clng(pcTargetPrice)>0) then%>
													<span class="pcSaleIcon"><a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
													<%end if%>
												</p>
											<% else %>
												<% if (cs_dblpcCC_Price>"0") and (session("customerType")="1") and (cs_pcv_intHideBTOPrice<>"1") then %>
													<p class="pcShowProductPriceW">
														<% response.write dictLanguage.Item(Session("language")&"_prdD1_4")& " " & scCursign & money(cs_dblpcCC_Price)%>
														<%if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then%>
														<span class="pcSaleIcon"><a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
														<%end if%>
													</p>
												<% end if %>	            
											<% 
											end if 
											
										'end if '// Not for sale - Commented out in v4.5
										'//////////////////////////////////////////
										'// End: Not For Sale
										'//////////////////////////////////////////
										%>
									</td>
								</tr>
							</table>
						</td>  
						<%
					end if
					'/////////////////////////////////////////////////////////////////////////////////////////
					'// END: DISPLAY CROSS SELLING
					'/////////////////////////////////////////////////////////////////////////////////////////
					
				end if '// if InStr(","& session("listcross") &",",","& pidrelation &",")=0 then
			End If '// If (pcsType=pcsFilterType) Then
			
			' Close current table row and open a new one, if there are still items to display
			If (pcIntCellCount = scPrdRow) And (tCnt+1 < pcv_intProductCount) Then
				pcIntCellCount=1
				response.write "</tr><tr>"
			Else
				pcIntCellCount=pcIntCellCount+1
			End If
			
			tCnt=tCnt+1

		loop	
		
		' Add missing table cells, if any
		if ((tCnt = pcv_intProductCount) OR (pcv_intDisplayCounter = cs_ViewCnt)) AND pcIntCellCount<(scPrdRow+1) then
			do until pcIntCellCount=(scPrdRow+1)
				response.write "<td></td>"
				pcIntCellCount=pcIntCellCount+1
			loop
		end if
	
	END IF
%>
</tr>
<%		
'// Add an extra row to faciliate the deselecting of radio boxes
If pcv_strIsBundleActiveFlag = True Then
	%>
    <tr>
    	<td colspan="<%=scPrdRow%>">
            <div style="padding-bottom:5px;" align="right">
                <a href="JavaScript:;" onClick="pcf_clearRadioBox();"><%=dictLanguage.Item(Session("language")&"_prdD1_6")%></a>
                <input name="rdoBundle" type="radio" value="" id="deselect" style="display:none">
				<script language="javascript">
                    function pcf_clearRadioBox() {		
                        el = document.getElementById("deselect")
                        el.checked=true
                    }
                </script>
            </div>
    	</td>
	</tr>
	<%
End if
%>
</table>