<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/ShipFromSettings.asp" -->
<html>
<head>
	<title>Print Pick List</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
	<STYLE TYPE="text/css">
	.text {
		font-family: Arial, Helvetica, sans-serif;
		font-size: 12px;
	}
	</STYLE>
</head>
<body style="margin: 10px; background-image: none;">
	<% dim connTemp, query, rs, qry_ID
 
	TmpStr=""
	Count=request("count")
	if (Count="") or (Count="0") then
		response.redirect "menu.asp"
	end if

	For k=1 to Count
		if (request("check" & k)="1") and (request("idord" & k)<>"") then
			TmpStr=TmpStr & request("idord" & k) & "***"
		end if
	Next

	if TmpStr="" then
		response.redirect "menu.asp"
	end if
	
	call openDb()
	A=split(TmpStr,"***")
	For k=lbound(A) to ubound(A)
		IF A(k)<>"" then
			if k<>lbound(A) then%>
				<P CLASS="breakhere">&nbsp;</p>
			<%end if%>
			<%qry_ID=A(k)
			query="SELECT idcustomer, orderdate FROM orders WHERE idOrder=" & qry_ID & ";"

			Set rs=Server.CreateObject("ADODB.Recordset")
			Set rs=connTemp.execute(query)
			Dim pidcustomer, porderdate

			pidcustomer=rs("idcustomer")
			porderdate=rs("orderdate")
			porderdate=ShowDateFrmt(porderdate)
	
			query="SELECT [name],lastname,customerCompany FROM customers WHERE idCustomer=" & pidcustomer
			Set rsCustObj=Server.CreateObject("ADODB.Recordset")
			Set rsCustObj=connTemp.execute(query)
			CustomerName=rsCustObj("name")& " " & rsCustObj("lastname")
			CustomerCompany=rsCustObj("customerCompany")
			if CustomerCompany<> "" then
				CustomerCompany= " - " & CustomerCompany
			end if
			set rsCustObj=nothing
						
			While Not rs.EOF %>
				
						<table width="100%" cellpadding="5" cellspacing="0" border="1" class="invoice">
								<tr> 
									<td class="invoice" colspan="2" bgcolor="#e1e1e1">
										<span style="font-size:13px; font-weight: bold;"><%= (scpre+int(qry_ID))  & " - " & porderdate & " - " & CustomerName & CustomerCompany %></span></td>
									<td class="invoice" colspan="2" bgcolor="#e1e1e1"><div align="right"><strong>Done <input type="checkbox"></strong></div></td>
</td>
								</tr>
								<tr> 
									<td class="invoice" width="70%">SKU - DESCRIPTION</td>
									<td class="invoice" width="18%">OPTIONS</td>
									<td class="invoice" width="2%">QTY</td>
									<td class="invoice" width="5%">BACKORDER</td>
								</tr>
								<tr>

                		<% query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.unitPrice, ProductsOrdered.xfdetails"
										'BTO ADDON-S
										if scBTO=1 then
											query=query&", ProductsOrdered.idconfigSession"
										end if
										'BTO ADDON-E
										query=query&",ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts,pcPO_GWOpt,pcPO_GWNote,pcPO_GWPrice FROM ProductsOrdered WHERE ProductsOrdered.idOrder=" & qry_ID & ";"
										Set rsTemp=Server.CreateObject("ADODB.Recordset")
										set rsTemp=connTemp.execute(query)
										dim intTotalWeight
										intTotalWeight=int(0)
										
										Do until rsTemp.eof
											pidProduct=rstemp("idProduct")
											pquantity=rstemp("quantity")
											
											'// Product Options Arrays
											pcv_strSelectedOptions = rsTemp("pcPrdOrd_SelectedOptions") ' Column 11
											pcv_strOptionsPriceArray = rsTemp("pcPrdOrd_OptionsPriceArray") ' Column 25
											pcv_strOptionsArray = rsTemp("pcPrdOrd_OptionsArray") ' Column 4
											
											punitPrice=rstemp("unitPrice")
											pxdetails=rstemp("xfdetails")
											if scBTO=1 then
												pidConfigSession=rstemp("idConfigSession")
											end if
											QDiscounts=rstemp("QDiscounts")
											ItemsDiscounts=rstemp("ItemsDiscounts")
											
											'GGG Add-on start
											pGWOpt=rstemp("pcPO_GWOpt")
											if pGWOpt<>"" then
											else
											pGWOpt="0"
											end if
											pGWText=rstemp("pcPO_GWNote")
											pGWPrice=rstemp("pcPO_GWPrice")
											if pGWPrice<>"" then
											else
											pGWPrice="0"
											end if
											'GGG Add-on end
											
											query="SELECT sku,description,weight,pcprod_QtyToPound FROM products WHERE idproduct="& pidProduct
											Set rsTemp2=Server.CreateObject("ADODB.Recordset")
											set rsTemp2=connTemp.execute(query)
											psku=rsTemp2("sku")
											pDescription=rsTemp2("description")
											pWeight=rsTemp2("weight")
											pcv_QtyToPound=rsTemp2("pcprod_QtyToPound")
											if pcv_QtyToPound>0 then
												pWeight=(16/pcv_QtyToPound)
												if scShipFromWeightUnit="KGS" then
													pWeight=(1000/pcv_QtyToPound)
												end if
											end if
											intTotalWeight=intTotalWeight+(pWeight*pquantity)
											%>

											<% 'BTO ADDON-S
											err.number=0
											TotalUnit=0
											If scBTO=1 then
												pIdConfigSession=trim(pidconfigSession)
												if pIdConfigSession<>"0" then 
													query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
													set rsConfigObj=conntemp.execute(query)
													if err.number <> 0 then
														set rsConfigObj=nothing
														call closedb()
														response.redirect "techErr.asp?error="& Server.Urlencode("Error in BatchPrint: "&err.description) 
													end if
													stringProducts=rsConfigObj("stringProducts")
													stringValues=rsConfigObj("stringValues")
													stringCategories=rsConfigObj("stringCategories")
													stringQuantity=rsConfigObj("stringQuantity")
													stringPrice=rsConfigObj("stringPrice")
													ArrProduct=Split(stringProducts, ",")
													ArrValue=Split(stringValues, ",")
													ArrCategory=Split(stringCategories, ",")
													ArrQuantity=Split(stringQuantity, ",")
													ArrPrice=Split(stringPrice, ",")
													set rsConfigObj=nothing
													for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
														query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
														set rsConfigObj=conntemp.execute(query)
														if NOT isNumeric(ArrQuantity(i)) then
															pIntQty=1
														else
															pIntQty=ArrQuantity(i)
														end if
														if (CDbl(ArrValue(i))<>0) or (((ArrQuantity(i)-1)*pQuantity>0) and (ArrPrice(i)>0)) then
															if (ArrQuantity(i)-1)>=0 then
																UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
															else
																UPrice=0
															end if
															TotalUnit=TotalUnit+((ArrValue(i)+UPrice)*pQuantity)
														end if
														set rsConfigObj=nothing
													next
												end if 
											End If 
			'BTO ADDON-E
			
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Get the total Price of all options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pOpPrices=0
		dim pcv_tmpOptionLoopCounter, pcArray_TmpCounter
		
		If len(pcv_strOptionsPriceArray)>0 then
		
			pcArray_TmpCounter = split(pcv_strOptionsPriceArray,chr(124))
			For pcv_tmpOptionLoopCounter = 0 to ubound(pcArray_TmpCounter)
				pOpPrices = pOpPrices + pcArray_TmpCounter(pcv_tmpOptionLoopCounter)
			Next
			
		end if				

		if NOT isNumeric(pOpPrices) then
			pOpPrices=0
		end if	
		
		'// Apply Discounts to Options Total
		'   >>> call function "pcf_DiscountedOptions(OriginalOptionsTotal, ProductID, Quantity, CustomerType)" from stringfunctions.asp
		Dim pcv_intDiscountPerUnit
		pOpPrices = pcf_DiscountedOptions(pOpPrices, pidProduct, pquantity, CustomerType)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Get the total Price of all options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			
											if TotalUnit>0 then
												punitPrice1=punitPrice
												if pIdConfigSession<>"0" then
													pRowPrice1=Cdbl(pquantity * ( punitPrice1 )) - TotalUnit
													punitPrice1=Round(pRowPrice1/pquantity,2)
												else
													pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
												end if
											else
												punitPrice1=punitPrice
												if pIdConfigSession<>"0" then
													pRowPrice1=Cdbl(pquantity * ( punitPrice1 ))
												else
													pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
					punitPrice1=Round(pRowPrice1/pquantity,2)
												end if
											end if
		
											%>
											
											<tr> 
												<td class="invoice"><strong><%=psku%> - <%=pDescription%></strong></td>
												<td class="invoice">
												
												
		<!-- start options -->
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: SHOW PRODUCT OPTIONS
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
						pcv_strSelectedOptions = ""
					end if
					
		if len(pcv_strSelectedOptions)>0 then %>

				<%
				'#####################
				' START LOOP
				'#####################	
							
							'// Generate Our Local Arrays from our Stored Arrays  
							
							' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
							pcArray_strSelectedOptions = ""					
							pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))
							
							' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
							pcArray_strOptionsPrice = ""
							pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))
							
							' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
							pcArray_strOptions = ""
							pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))
							
							' Get Our Loop Size
							pcv_intOptionLoopSize = 0
							pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
							
							' Start in Position One
							pcv_intOptionLoopCounter = 0
							
				' Display Our Options
				tempPrice=0
				For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize %>
				<div><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></div>
					<% Next
					'#####################
					' END LOOP
					'#####################					
					%>														
															
		<% end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: SHOW PRODUCT OPTIONS
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>
		<!-- end options -->	
												
												
												
												</td>
												<td class="invoice"><%=pquantity%></td>
												<td class="invoice"><div align="center">[&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</div></td>
											</tr>

											<% 'BTO ADDON-S
											if scBTO=1 then
												if pIdConfigSession<>"0" then 
							query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
													set rsConfigObj=connTemp.execute(query)
							
													if err.number <> 0 then
														response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails: "&err.description) 
													end if
													stringProducts=rsConfigObj("stringProducts")
													stringValues=rsConfigObj("stringValues")
													stringCategories=rsConfigObj("stringCategories")
													stringQuantity=rsConfigObj("stringQuantity")
													stringPrice=rsConfigObj("stringPrice")
													ArrProduct=Split(stringProducts, ",")
													ArrValue=Split(stringValues, ",")
													ArrCategory=Split(stringCategories, ",")
													ArrQuantity=Split(stringQuantity, ",")
													ArrPrice=Split(stringPrice, ",")

													'Hide this information if this is a packing slip 
													%>
													<tr> 
													<td class="invoice" colspan="3">
													<table width="100%" cellspacing="2" cellpadding="0" class="invoiceBto">
															<tr> 
																<td colspan="3" class="invoiceNob"><u>Customizations:</u></td>
															</tr>
																<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
																	query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
																	set rsConfigObj=connTemp.execute(query)
																	pcategoryDesc=rsConfigObj("categoryDesc")
																	pdescription=rsConfigObj("description")
																	psku=rsConfigObj("sku")
																	pItemWeight=rsConfigObj("weight")
																	if NOT isNumeric(ArrQuantity(i)) then
																		pIntQty=1
																	else
																		pIntQty=ArrQuantity(i)
																	end if %>
																	<tr> 
																<td width="20%" class="invoiceNob" valign="top"><%=pcategoryDesc%>:</td>
																<td width="70%" class="invoiceNob" valign="top"><%=psku%> - <%=pdescription%><%if pIntQty>1 then%>
																		- QTY: <%=ArrQuantity(i)%><%end if%></td>
																<%if (CDbl(ArrValue(i))<>0) or (((ArrQuantity(i)-1)*pQuantity>0) and (ArrPrice(i)>0)) then
																if (ArrQuantity(i)-1)>=0 then
																	UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
																else
																	UPrice=0
																end if
																'pfPrice=pfPrice+cdbl((ArrValue(i)+UPrice)*pQuantity) %> 
																<%end if%> 

															</tr>
															<% intItemWeight=int(pItemWeight)*pIntQTY*pquantity
																 intTotalWeight=intTotalWeight+intItemWeight
																 set rsConfigObj=nothing
																	next
																	set rsConfigObj=nothing %>
														</table>
														</td>
														<td class="invoice"></td>
													</tr>
												<% end if %>
											<% end if
											'BTO ADDON-E %>
									
											<%'BTO ADDON-S
											pRowPrice=(punitPrice)*(pquantity)
											If scBTO=1 then
												pidConfigSession=trim(pidConfigSession)

													'BTO Additional Charges
													if scBTO=1 then %>
														<% pIdConfigSession=trim(pidConfigSession)
														if pIdConfigSession<>"0" then 
															query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
															set rsConfigObj=connTemp.execute(query)
															
															stringCProducts=rsConfigObj("stringCProducts")
															stringCValues=rsConfigObj("stringCValues")
															stringCCategories=rsConfigObj("stringCCategories")
															ArrCProduct=Split(stringCProducts, ",")
															ArrCValue=Split(stringCValues, ",")
															ArrCCategory=Split(stringCCategories, ",")
															if ArrCProduct(0)<>"na" then
															%>
															<%
															' Hide if packing slip
															%>
															<tr> 
																<td class="invoice" colspan="3">

																<table width="100%" cellspacing="0" cellpadding="2" class="invoiceBto">
																	<tr class="small"> 
																	<td colspan="2" class="invoiceNob"><u>Additional Charges</u></td>
																	</tr>
			
																			<% Charges=0
																			for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
																				query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
																				set rsConfigObj=connTemp.execute(query)
																				pcategoryDesc=rsConfigObj("categoryDesc")
																				pdescription=rsConfigObj("description")
																				psku=rsConfigObj("sku")
																				pItemWeight=rsConfigObj("weight")
																				intTotalWeight=intTotalWeight+int(pItemWeight)
																				if (CDbl(ArrCValue(i))>0)then
																					Charges=Charges+cdbl(ArrCValue(i))
																				end if
																				%>
																				<tr> 
																				<td width="20%" class="invoiceNob" valign="top"><%=pcategoryDesc%>:</td>
																				<td width="70%" class="invoiceNob" valign="top"><%=psku%> - <%=pdescription%></td>
																				</tr>
																				<% set rsConfigObj=nothing
																				next
																				set rsConfigObj=nothing 
																				pRowPrice=pRowPrice+Cdbl(Charges)%>
																				</table>
																	</td>
													<td class="invoice">&nbsp;</td>
															</tr>
															<% end if
														end if %>
													<% end if
													'BTO Additional Charges
											end if 'BTO%>
											
											<% if len(pxdetails)>3 then %>
												<tr>
													<td colspan="4" class="invoice">Special Fields</td>
												</tr>
												<tr> 
													<td class="invoice" colspan="3"><%=replace(pxdetails,"|","<br>")%></td>
													<td class="invoice">&nbsp;</td>
												</tr>
											<% end if %>

										
											<% rstemp.moveNext
										loop
										set rstemp=nothing %>
                
	</table>


	
<%rs.MoveNext
Wend
set rs=nothing
%>
<%end if%>
<%
	Next
	closedb()
%>

</body>
</html>