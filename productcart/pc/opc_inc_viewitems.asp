<% On Error Resume Next
pSubTotal=ccur(calculateCartTotal(pcCartArray, ppcCartIndex))
'SB S
Dim pcIsSubscription , StrandSub 
pcIsSubscription = session("pcIsSubscription")
'SB E

'GET CUSTOMER SESSION DATA
call openDb()

%>
	<table class="pcShowContent">
				<tr> 
					<th width="4%"><p><%=dictLanguage.Item(Session("language")&"_orderverify_25")%></p></th>
					<th width="62%"><p><%=dictLanguage.Item(Session("language")&"_orderverify_27")%></p></th>
					<th width="12%" nowrap align="right"><p><%=dictLanguage.Item(Session("language")&"_orderverify_32")%></p></th>
					<th width="12%" nowrap align="right"><p><%=dictLanguage.Item(Session("language")&"_orderverify_28")%></p></th>
				</tr>
                <tr>
                	<td colspan="4" class="pcSpacer"></td>
                </tr>
				
				<% 'START GET PRODUCTS ORDERING
				strBundleArray=""
				pSFstrBundleArray=""
				Dim pcProductList(100,5)
				for f=1 to ppcCartIndex
					pcProductList(f,0)=pcCartArray(f,0)
					pcProductList(f,1)=pcCartArray(f,10)
					pcProductList(f,3)=pcCartArray(f,2)
					pcProductList(f,4)=0
					'SB S
					if (pcCartArray(f,38)) > 0  then
						'// Get the Sub Details
						pSubscriptionID = (pcCartArray(f,38)) %>				
						<!--#include file="../includes/pcSBDataInc.asp" --> 	
				  	<% end if 
					'SB E
					if pcCartArray(f,10)=0 then
							
						'BTO ADDON-S
						pBTOValues=0
						if trim(pcCartArray(f,16))<>"" then 
							
							query="SELECT stringProducts, stringValues, stringCategories,stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rs=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							stringProducts=rs("stringProducts")
							stringValues=rs("stringValues")
							stringCategories=rs("stringCategories")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							Qstring=rs("stringQuantity")
							ArrQuantity=Split(Qstring,",")
							Pstring=rs("stringPrice")
							ArrPrice=split(Pstring,",")
							set rs=nothing
							
							
							if ArrProduct(0)="na" then
							else
								for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
									query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
									set rsQ=connTemp.execute(query)
									tmpMinQty=1
									if not rsQ.eof then
										tmpMinQty=rsQ("pcprod_minimumqty")
										if IsNull(tmpMinQty) or tmpMinQty="" then
											tmpMinQty=1
										else
											if tmpMinQty="0" then
												tmpMinQty=1
											end if
										end if
									end if
									set rsQ=nothing
									tmpDefault=0
									query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
									set rsQ=connTemp.execute(query)
									if not rsQ.eof then
										tmpDefault=rsQ("cdefault")
										if IsNull(tmpDefault) or tmpDefault="" then
											tmpDefault=0
										else
											if tmpDefault<>"0" then
											 	tmpDefault=1
											end if
										end if
									end if
									set rsQ=nothing
									
									if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
									if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
										if tmpDefault=1 then
											UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
										else
											UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
										end if
									else
										UPrice=0
									end if
									pBTOValues=pBTOValues+ccur((ArrValue(i)+UPrice)*pcCartArray(f,2))
									end if
									set rsObj=nothing
								next
							end if						
						End if
						'BTO ADDON-E
					End if
											
					if pcCartArray(f,10)=0 then
						
						if pcv_IsEUMemberState=0 then
							tmpRowPrice=ccur( pcCartArray(f,2) * pcCartArray(f,17) )
						end if

						pRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))
						pExtRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))-ccur(pBTOvalues) %>
						<% 'Validate for multiple of N
						query="SELECT pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty FROM products WHERE idproduct=" & pcCartArray(f,0)
						set rs=server.CreateObject("ADODB.RecordSet") 									
						set rs=connTemp.execute(query)
								
						if err.number<>0 then
							call LogErrorToDatabase()
							'set rs=nothing
							'call closedb()
							'response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
								
						pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
						if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
							pcv_intHideBTOPrice="0"
						end if
						pcv_intQtyValidate=rs("pcprod_QtyValidate")
						if isNULL(pcv_intQtyValidate) OR pcv_intQtyValidate="" then
							pcv_intQtyValidate="0"
						end if				
						pcv_lngMinimumQty=rs("pcprod_MinimumQty")
						if isNULL(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
							pcv_lngMinimumQty="0"
						end if
						set rs=nothing 
						
						%>
						<tr valign="top"> 
							<td>
								<p><%=pcCartArray(f,2)%></p>
							</td>
							<td>
								<p><%=pcCartArray(f,1) %>&nbsp;<span class="opcSku">(<%=pcCartArray(f,7)%>)</span></p>
							</td>
							<td align="right">
							<% if pcv_intHideBTOPrice<>"1" then
								if pcCartArray(f,17) > 0 then %>
									<%=scCurSign & money(pcCartArray(f,17)-ccur(ccur(pBTOvalues)/pcCartArray(f,2)))%>
								<% 	end if
							end if %>
							</td>
							<td align="right" nowrap>
								<p><% if pExtRowPrice > 0 then response.write(scCurSign & money(pExtRowPrice)) end if %></p>
							</td>
						</tr>
					
						<% 'BTO ADDON-S
						if trim(pcCartArray(f,16))<>"" then
							
							query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rs=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
						
							stringProducts=rs("stringProducts")
							stringValues=rs("stringValues")
							stringCategories=rs("stringCategories")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							Qstring=rs("stringQuantity")
							ArrQuantity=Split(Qstring,",")
							Pstring=rs("stringPrice")
							ArrPrice=split(Pstring,",")
							set rs=nothing
							%>
								
							<tr> 
								<td>&nbsp;</td>
								<td colspan="3" class="pcShowBTOconfiguration"> 
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr> 
											<td colspan="2"><p><%=bto_dictLanguage.Item(Session("language")&"_viewcart_1")%></p></td>
										</tr>
                      
										<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
											query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
											set rs=server.CreateObject("ADODB.RecordSet") 
											set rs=conntemp.execute(query)
											
											if err.number<>0 then
												call LogErrorToDatabase()
												'set rs=nothing
												'call closedb()
												'response.redirect "techErr.asp?err="&pcStrCustRefID
											end if
											
											strCategoryDesc=rs("categoryDesc")
											strDescription=rs("description")
											set rs=nothing
													
											query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pcCartArray(f,0) 
											set rs=server.CreateObject("ADODB.RecordSet") 
											set rs=conntemp.execute(query)
														
											if err.number<>0 then
												call LogErrorToDatabase()
												'set rs=nothing
												'call closedb()
												'response.redirect "techErr.asp?err="&pcStrCustRefID
											end if
												
											btDisplayQF=rs("displayQF")
											set rs=nothing
											
											query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
											set rsQ=connTemp.execute(query)
											tmpMinQty=1
											if not rsQ.eof then
												tmpMinQty=rsQ("pcprod_minimumqty")
												if IsNull(tmpMinQty) or tmpMinQty="" then
													tmpMinQty=1
												else
													if tmpMinQty="0" then
														tmpMinQty=1
													end if
												end if
											end if
											set rsQ=nothing
											tmpDefault=0
											query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
											set rsQ=connTemp.execute(query)
											if not rsQ.eof then
												tmpDefault=rsQ("cdefault")
												if IsNull(tmpDefault) or tmpDefault="" then
													tmpDefault=0
												else
													if tmpDefault<>"0" then
													 	tmpDefault=1
													end if
												end if
											end if
											set rsQ=nothing %>
											<tr> 
												<td width="85%" valign="top">
													<p><%=strCategoryDesc%>:&nbsp;
													<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%>(<%=ArrQuantity(i)%>)&nbsp;<%end if%>
													<%=strDescription%>
													</p>
												</td>
												<td width="15%" valign="top">
													<p align="right">
													<%if (ArrValue(i)<>"") and (ArrQuantity(i)<>"") and (ArrPrice(i)<>"") then 
														if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
															if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
																if tmpDefault=1 then
																	UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
																else
																	UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
																end if
															else
																UPrice=0
															end if %>
															<%=scCurSign & money(ccur((ArrValue(i)+UPrice)*pcCartArray(f,2)))%>
														<%else
															if tmpDefault=1 then%>
																<%=dictLanguage.Item(Session("language")&"_defaultnotice_1")%>
															<%end if
														end if
													end if%>
													</p>
												</td>
											</tr>
										 <% next %>
									</table>
								</td>
							</tr>
						<% End if 
						'BTO ADDON-E
								
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: SHOW PRODUCT OPTIONS
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

						if trim(pcCartArray(f,4))<>"" then
						
							Dim pcv_strOptionsArray, pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice, tAprice
							Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions

							pcv_strOptionsArray = trim(pcCartArray(f,4))
						
							if len(pcv_strOptionsArray)>0 then %>
								<tr valign="top">
									<td>&nbsp;</td>
									<td colspan="2">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<%
										'#####################
										' START LOOP
										'#####################	
										
										'// Generate Our Local Arrays from our Stored Arrays  
										
										' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
										pcArray_strSelectedOptions = ""					
										pcArray_strSelectedOptions = Split(trim(pcCartArray(f,11)),chr(124))
										
										' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
										pcArray_strOptionsPrice = ""
										pcArray_strOptionsPrice = Split(trim(pcCartArray(f,25)),chr(124))
										
										' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
										pcArray_strOptions = ""
										pcArray_strOptions = Split(trim(pcv_strOptionsArray),chr(124))
										
										' Get Our Loop Size
										pcv_intOptionLoopSize = 0
										pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
										
										' Start in Position One
										pcv_intOptionLoopCounter = 0
										
										' Display Our Options
										For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize %>
											<tr>
												<td width="67%"><p><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></p></td>
												<td align="right" width="33%">									
												<% tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
												
												if tempPrice="" or tempPrice=0 then
													response.write "&nbsp;"
												else %>
													<table width="100%" cellpadding="0" cellspacing="0" border="0">
														<tr>
															<td align="left" width="60%">
																<%=scCurSign&money(tempPrice)%>
															</td>
															<td align="right" width="40%">
																<%									
																tAprice=(tempPrice*ccur(pcCartArray(f,2)))
																response.write scCurSign&money(tAprice) 
																%>
															</td>
														</tr>
													</table>
												<% end if %>			
												</td>
											</tr>
										<% Next
										'#####################
										' END LOOP
										'#####################	
									
										%>
										</table>
									</td>
									<td>&nbsp;</td>
								</tr>															
							<% end if
						end if
								
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: SHOW PRODUCT OPTIONS
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
						pRowPrice=pRowPrice + ccur(pcCartArray(f,2) * pcCartArray(f,5)) %>
								
						<% if trim(pcCartArray(f,21))<>"" then %>
							<tr> 
								<td>&nbsp;</td>
								<td colspan="2"><p><% response.write(replace(pcCartArray(f,21),"''","'"))%></p></td>
								<td>&nbsp;</td>
							</tr>
						<%end if %>
							
						<% 'if items quantities discounts apply to this product, show the total applied amount here
						if trim(pcCartArray(f,16))<>"" then
							if ccur(pcCartArray(f,30))>0 then
								pRowPrice=pRowPrice-ccur(pcCartArray(f,30)) %>
								<tr> 							
									<td>&nbsp;</td>
									<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_showcart_23")%></p></td>
									<td nowrap align="right">
										<p>- 
										<% response.write scCurSign &  money(ccur(pcCartArray(f,30))) %>
										</p>
									</td>
								</tr>
							<% end if
						End if%>
							
						<% 'BTO Additional Charges
						if trim(pcCartArray(f,16))<>"" then
							query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
									
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rs=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							stringCProducts=rs("stringCProducts")
							stringCValues=rs("stringCValues")
							stringCCategories=rs("stringCCategories")
							ArrCProduct=Split(stringCProducts, ",")
							ArrCValue=Split(stringCValues, ",")
							ArrCCategory=Split(stringCCategories, ",")
							set rs=nothing
									
							if ArrCProduct(0)<>"na" then
								pRowPrice=pRowPrice+ccur(pcCartArray(f,31))%>
								<tr> 
									<td>&nbsp;</td>
									<td colspan="3" valign="top" class="pcShowBTOconfiguration"> 
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr> 
												<td><p><b><%=bto_dictLanguage.Item(Session("language")&"_viewcart_3")%></b></p></td>
												<td></td>
											</tr>
											<% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
												query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))"
												set rs=server.CreateObject("ADODB.RecordSet") 
												set rs=conntemp.execute(query)
												
												if err.number<>0 then
													call LogErrorToDatabase()
													'set rs=nothing
													'call closedb()
													'response.redirect "techErr.asp?err="&pcStrCustRefID
												end if
												
												strCategoryDesc=rs("categoryDesc") 
												strDescription=rs("description") 
												set rs=nothing %>
												<tr> 
													<td width="85%" valign="top">
														<p><%=strCategoryDesc%>:&nbsp;<%=strDescription%></p>
													</td>
													<td width="15%" align="right" valign="top">
													<p> 
													<%if (ccur(ArrCValue(i))>0)then %>
														<%=scCurSign & money(ArrCValue(i))%>
													<%end if%>
													</p>
													</td>
												</tr>
											<% next %>
										</table>
									</td>
								</tr>
							<% End if
							'Have Charges 
							
						End if 
						'BTO Additional Charges %>
														
						<% 'if quantity discounts apply to this product, show the total applied amount here
						if trim(pcCartArray(f,15))<>"" AND trim(pcCartArray(f,15))>0 then
							pRowPrice=pRowPrice-ccur(pcCartArray(f,15)) %>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2">
									<%=dictLanguage.Item(Session("language")&"_showcart_20")%>
									<%=dictLanguage.Item(Session("language")&"_showcart_20b")%>
								</td>
								<td nowrap align="right">
									<p>-<% response.write scCurSign & money(pcCartArray(f,15)) %></p>
								</td>
							</tr>
						<% End if 
								
						if pExtRowPrice<>pRowPrice then %>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right"><p><%=dictLanguage.Item(Session("language")&"_showcart_22")%></td>
								<td nowrap align="right"><p><%=scCurSign & money(pRowPrice) %></p></td>
							</tr>
						<% end if %>
						<% 
						'SB S
						if (pcCartArray(f,38)) > 0  then
						
					 		'// Get the data 
					  		pSubscriptionID = (pcCartArray(f,38)) 

                            '// If there's a trial set the line total to the trial price
                            if pcv_intIsTrial = "1" Then
                            	pRowPrice = "8" '// pcv_curTrialAmount
                            else
                            	pRowPrice = "8" '// pExtRowPrice
                            end if  
							
						end if 
						'SB E 
						%>
						
						<% 'START 10th Row - Cross Sell Bundle Discount %>	
						<% if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then 	%>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
								<%=dictLanguage.Item(Session("language")&"_showcart_26")%>
								</td>
								<td align="right">
								<% =scCurSign &  money( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) %>
								</td>
							</tr>
							<% strBundleArray=strBundleArray&pcCartArray(f,0)&","&pcCartArray(f,27)&","&pcCartArray(f,28)&","&((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)&"||"
						end if %>
						<% 'END 10th Row - Cross Sell Bundle Discount %>	
						
						<% 'START 11th Row - Cross Sell Bundle Subtotal %>
						<% if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then 
						    pRowPrice = ( ccur(pRowPrice) + ccur(pcProductList(cint(pcCartArray(f,27)),2)) ) - ( ( ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28) ) ) * pcCartArray(f,2) )%>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
								<%= dictLanguage.Item(Session("language")&"_showcart_22")%>
								</td>
								<td align="right">
								<%= scCurSign &  money(pRowPrice) %>
								</td>
							</tr>
						<% end if %>
						<% 'END 11th Row - Cross Sell Bundle Subtotal %>	


						<%'GGG Add-on start
						if Session("Cust_GW")="1" then
							
							GWmsg="<u>" & dictLanguage.Item(Session("language")&"_orderverify_36a") & "</u>: "
							gIDPro=pcCartArray(f,0)
							gMode=1
							query="select pcPE_IDProduct from pcProductsExc where pcPE_IDProduct=" & gIDPro
							set rsG=server.CreateObject("ADODB.RecordSet")
							set rsG=connTemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rsG=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
										
							if not rsG.eof then
								GWmsg=GWmsg & dictLanguage.Item(Session("language")&"_orderverify_38a")
								gMode=0
							else
								if (pcCartArray(f,34)="") or (pcCartArray(f,34)="0") then
									GWmsg=GWmsg & dictLanguage.Item(Session("language")&"_orderverify_37a")
									gMode=1
								else
									gIDOpt=pcCartArray(f,34)
									query="select pcGW_OptName,pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & gIDOpt
									set rsG=server.CreateObject("ADODB.RecordSet")
									set rsG=connTemp.execute(query)
									
									if err.number<>0 then
										call LogErrorToDatabase()
										'set rsG=nothing
										'call closedb()
										'response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
												
									if NOT rsG.eof then
										pcv_strOptName = rsG("pcGW_OptName")
										pcv_strOptPrice = rsG("pcGW_OptPrice")
										GWmsg=GWmsg & pcv_strOptName & " - " & scCurSign & money(pcv_strOptPrice)
										GiftWrapPaymentTotal=GiftWrapPaymentTotal+pcv_strOptPrice
									end if 

									gMode=1
								end if
							end if %>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="3"><p><%=GWmsg%>&nbsp;</p></td>
							</tr> 
						<%end if
						'GGG end%>
						<tr> 
							<td colspan="4"><hr></td>
						</tr>
					<% end if %>
					
					<% 
					if pcv_IsEUMemberState = 0 then
						pcProductList(f,2) = tmpRowPrice
					else
						pcProductList(f,2) = pRowPrice
					end if
				next
				pSFstrBundleArray=strBundleArray %>
				<%
				'// Discounts by Categories
				Dim pcv_strApplicableProducts
				pcv_strApplicableProducts=""
				CatDiscTotal=0

				query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				
				if err.number<>0 then
					call LogErrorToDatabase()
					'set rs=nothing
					'call closedb()
					'response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				Do While not rs.eof
					CatSubQty=0
					CatSubTotal=0
					CatSubDiscount=0
					ApplicableCategoryID = rs("IDCat")
					CanNotRun=0
					IDCat=rs("IDCat")
					query="SELECT categories_products.idcategory FROM categories_products INNER JOIN pcPrdPromotions ON categories_products.idproduct=pcPrdPromotions.idproduct WHERE categories_products.idcategory=" & IDCat & ";"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						CanNotRun=1
					end if
					set rsQ=nothing
					
					IF CanNotRun=0 THEN
							
					For f=1 to ppcCartIndex
						if (pcProductList(f,1)=0) and (pcProductList(f,4)=0) then 
							query="select idproduct from categories_products where idcategory=" & rs("IDCat") & " and idproduct=" & pcProductList(f,0)
							set rstemp=server.CreateObject("ADODB.RecordSet")
							set rstemp=connTemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rstemp=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
								
							if not rstemp.eof then
								CatSubQty=CatSubQty+pcProductList(f,3)
								CatSubTotal=CatSubTotal+pcProductList(f,2)
								pcProductList(f,4)=1
								pcv_strApplicableProducts = pcv_strApplicableProducts & pcProductList(f,0) & chr(124) &  ApplicableCategoryID & ","								
							end if
							set rstemp=nothing
						end if
						
					Next
					
					pcv_strrApplicableCategories = pcv_strrApplicableCategories & CatSubTotal & chr(124) &  ApplicableCategoryID & ","
					
					if CatSubQty>0 then

						query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & IDCat & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
								
						if err.number<>0 then
							call LogErrorToDatabase()
							'set rstemp=nothing
							'call closedb()
							'response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
								
						if not rstemp.eof then
							'// There are quantity discounts defined for that quantity 
							pDiscountPerUnit=rstemp("pcCD_discountPerUnit")
							pDiscountPerWUnit=rstemp("pcCD_discountPerWUnit")
							pPercentage=rstemp("pcCD_percentage")
							pbaseproductonly=rstemp("pcCD_baseproductonly")

							if session("customerType")<>1 then  'customer is a normal user
								if pPercentage="0" then 
									CatSubDiscount=pDiscountPerUnit*CatSubQty
								else
									CatSubDiscount=(pDiscountPerUnit/100) * CatSubTotal
								end if
							else  'customer is a wholesale customer
								if pPercentage="0" then 
									CatSubDiscount=pDiscountPerWUnit*CatSubQty
								else
									CatSubDiscount=(pDiscountPerWUnit/100) * CatSubTotal
								end if
							end if							
						end if
						
					set rstemp=nothing						
					end if '// if CatSubQty>0 then

					CatDiscTotal=CatDiscTotal+CatSubDiscount
					
					END IF 'CanNotRun
					rs.MoveNext
				loop
				set rs=nothing

				'// Round the Category Discount to two decimals
				if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
					CatDiscTotal = RoundTo(CatDiscTotal,.01)
				end if

			
			'Display Applied Product Promotions (if any)
			TotalPromotions=0
			if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
				PromoArr1=Session("pcPromoSession")
				PromoIndex=Session("pcPromoIndex")
				For m=1 to PromoIndex%>
				<tr>
					<td colspan="3" align="right">
					<%=PromoArr1(m,1)%>
					</td>
					<td align="right">
						-<%=scCurSign  & money(PromoArr1(m,2))%>
						<%TotalPromotions=TotalPromotions+cdbl(PromoArr1(m,2))%>
					</td>
				</tr>
				<%Next
			end if
			
			' Calculate & display order total
				pSubTotal=pSubTotal-CatDiscTotal-TotalPromotions %>
				<tr> 
					<td colspan="3" align="right"><p><b><%=dictLanguage.Item(Session("language")&"_orderverify_15")%></b></p></td>
					<td nowrap align="right">
						<p><%response.write scCurSign & money(pSubTotal)%></p>
						<script>
						 $("#pcOPCtotalAmount").text("<%=scCurSign & money(pSubTotal)%>");
						 </script>
					</td>
				</tr>
				<%' Display category-based quantity discounts
				if CatDiscTotal>0 then%>
				<tr> 
					<td colspan="3" align="right"><%response.write dictLanguage.Item(Session("language")&"_catdisc_1")%></td>
					<td nowrap align="right">
						<% response.write scCurSign & money(CatDiscTotal) %>
					</td>
				</tr>
				<% end if %>
</table>
<%call closedb() %>