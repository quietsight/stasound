<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages_ship.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/languagesCP.asp"-->
<% Dim pageTitle, Section
pageTitle="This is an Incomplete Order"
Section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<% 
Dim connTemp, query, rs, rstemp, qry_ID
call openDb()
If request.form("UpdateDO")<>"" then
	
	qry_ID=request.form("qry_ID")
	
	'// Update Order Status
	query="UPDATE orders SET orderstatus=2 WHERE idOrder=" & qry_ID & ";"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	set rs=nothing
	
	'// Update Inventory
	query="SELECT idProduct,quantity,idconfigSession FROM ProductsOrdered WHERE idOrder="& qry_ID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
		
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
		
	do while not rs.eof  
		pIdProduct=rs("idProduct")
		pQuantity=rs("quantity")
		idconfig=rs("idconfigSession")
		'check if stock is ignored or not
		query="SELECT noStock FROM products WHERE idProduct="&pIdProduct
		set rstemp=conntemp.execute(query)   
		pNoStock=rstemp("noStock")
	
		query="SELECT stock, sales, description FROM products WHERE idProduct="&pIdProduct
		set rstemp=conntemp.execute(query)   
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if        
		pDescription=rstemp("description")
		
		if pNoStock=0 then
			' decrease stock 
			if ppStatus=0 then
				query="UPDATE products SET stock=stock-" &pQuantity&" WHERE idProduct="&pIdProduct
				set rsTemp=conntemp.execute(query)  
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				'Update BTO Items & Additional Charges stock and sales 
				IF (idconfig<>"") and (idconfig<>"0") then
					query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
					set rs1=conntemp.execute(query)
					stringProducts=rs1("stringProducts")
					stringQuantity=rs1("stringQuantity")
					stringCProducts=rs1("stringCProducts")
					if (stringProducts<>"") and (stringProducts<>"na") then
						PrdArr=split(stringProducts,",")
						QtyArr=split(stringQuantity,",")
						
						for k=lbound(PrdArr) to ubound(PrdArr)
							if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
								query="UPDATE products SET stock=stock-" &QtyArr(k)*pQuantity&",sales=sales+" &QtyArr(k)*pQuantity&" WHERE idProduct=" &PrdArr(k)
								set rs1=conntemp.execute(query)
							end if
						next
					end if
					if (stringCProducts<>"") and (stringCProducts<>"na") then
						CPrdArr=split(stringCProducts,",")
						
						for k=lbound(CPrdArr) to ubound(CPrdArr)
							if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
								query="UPDATE products SET stock=stock-" &pQuantity&",sales=sales+" &pQuantity&" WHERE idProduct=" &CPrdArr(k)
								set rs1=conntemp.execute(query)
							end if
						next
					end if
				END IF
				'End Update BTO Items & Additional Charges
				
			end if
		end if
					 
		' update sales 
		if ppStatus=0 then  
			query="UPDATE products SET sales=sales+" &pQuantity&" WHERE idProduct=" &pIdProduct
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=conntemp.execute(query)  
			set rstemp=nothing  
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		end if 
		rs.movenext
	loop
	
	set rs=nothing
	call closeDb()
	
	response.redirect "Orddetails.asp?id="& qry_ID
	response.End()
End if

qry_ID=request.querystring("id")
If Not validNum(qry_ID) then
	response.redirect "techErr.asp?error="&Server.URLEncode("Not a valid order number")
End If

query="SELECT idcustomer, orderdate, Address, city, stateCode,zip,CountryCode,paymentDetails,shipmentDetails,shippingAddress,shippingCity,shippingStateCode,shippingZip,pcOrd_shippingPhone,shippingCountryCode,idAffiliate,affiliatePay,discountDetails,taxAmount,total,comments,orderStatus,processDate,shipDate,shipvia,trackingNum,returnDate,returnReason,iRewardPoints,iRewardPointsCustAccrued,iRewardValue,ord_VAT,pcOrd_CustomerIP FROM orders WHERE idOrder=" & qry_ID & ";"

Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)
Dim pidcustomer, porderdate, pAddress, pcity, pstateCode, pzip, pCountryCode, ppaymentDetails, pshipmentDetails, pshippingAddress, pshippingCity, pshippingStateCode, pshippingZip, pshippingPhone, pshippingCountryCode, pidAffiliate, paffiliatePay, pdiscountDetails, ptaxAmount, ptotal, pcomments, porderStatus, pprocessDate, pshipDate, pshipvia, ptrackingNum, preturnDate, preturnReason, pcv_strCustomerIP

pidcustomer=rs("idcustomer")
If Not validNum(pidcustomer) then
	response.redirect "techErr.asp?error="&Server.URLEncode("Not a valid customer number")
End If

porderdate=rs("orderdate")
porderdate=ShowDateFrmt(porderdate)
pAddress=rs("Address")
pcity=rs("city")
pstateCode=rs("stateCode")
pzip=rs("zip")
pCountryCode=rs("CountryCode")
ppaymentDetails=trim(rs("paymentDetails"))
pshipmentDetails=rs("shipmentDetails")
pshippingAddress=rs("shippingAddress")
pshippingCity=rs("shippingCity")
pshippingStateCode=rs("shippingStateCode")
pshippingZip=rs("shippingZip")
pshippingPhone=rs("pcOrd_shippingPhone")
pshippingCountryCode=rs("shippingCountryCode")
pidAffiliate=rs("idaffiliate")
paffiliatePay=rs("affiliatePay")
pdiscountDetails=rs("discountDetails")
ptaxAmount=rs("taxAmount")
ptotal=rs("total")
pcomments=rs("comments")
porderStatus=rs("orderStatus")
pprocessDate=rs("processDate")
pprocessDate=ShowDateFrmt(pprocessDate)
pshipDate=rs("shipDate")
pshipDate=ShowDateFrmt(pshipDate)
pshipvia=rs("shipvia")
ptrackingNum=rs("trackingNum")
preturnDate=rs("returnDate")
preturnDate=ShowDateFrmt(preturnDate)
preturnReason=rs("returnReason")
piRewardPoints=rs("iRewardPoints")
piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
piRewardValue=rs("iRewardValue")
pord_VAT=rs("ord_VAT")
if isNull(pord_VAT) or pord_VAT="" then
	pord_VAT=0
end if
pcv_strCustomerIP = rs("pcOrd_CustomerIP") 
%>
<form name="form2" method="post" action="OrdDetailsIncomplete.asp" class="pcForms">
<input type="hidden" name="qry_ID" value="<%=request.querystring("id")%>">
<input type="hidden" name="idcustomer" value="<%=pidcustomer%>">
	<table class="pcCPcontent">
    	<tr>
            <td colspan="2">
                <div>Order #: <b><%=(scpre+int(qry_ID))%></b>&nbsp;|&nbsp;Order Date: <strong><%=porderdate%></strong>&nbsp;|&nbsp;Order Total: <strong><%=scCurSign&money(ptotal)%></strong>&nbsp;&nbsp;&nbsp;<input type="submit" name="UpdateDO" value="Update Order Status" class="submit2"></div>
                <div style="margin-top: 10px;"><a href="http://wiki.earlyimpact.com/productcart/orders_status#incomplete_orders" target="_blank">Learn about <strong>incomplete orders</strong></a>. Incomplete orders <strong>are not displayed</strong> with other orders in the Control Panel. If you would like to change the status of this order so it is no longer treated as &quot;incomplete&quot;, click on &quot;Update Order Status&quot;. The order status will become <em>Pending</em>.</div>
            </td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>      
		<tr> 
			<th colspan="2">CUSTOMER INFORMATION</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<% 
		' Calculate customer number using sccustpre constant
        Dim pcCustomerNumber
        pcCustomerNumber = (sccustpre + int(pidcustomer))
		
		' Get customer information
		query="SELECT customerType,name,lastname,customerCompany,phone,email FROM customers WHERE idcustomer="& pidcustomer
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=connTemp.execute(query)
		pCustomerType=rs("customerType")
		pCustomerEmail=rs("email")
        %>
		<tr> 
			<td colspan="2">
            	Customer Number: <%=pcCustomerNumber%>
                &nbsp;|&nbsp;
				<%
				if pcv_strCustomerIP="" then
					pcv_strCustomerIP="Not Available"  
				end if
				%>
				Customer's IP Address: <%=pcv_strCustomerIP %> 
                <%
				if trim(pCustomerEmail)<>"" and not IsNull(pCustomerEmail) then
				%>
                &nbsp;|&nbsp;
                <a href="mailto:<%=pCustomerEmail%>">E-mail Customer</a>
                <%
				end if
				%>
				&nbsp;|&nbsp;
                <a href="modcusta.asp?idcustomer=<%=pidcustomer%>" target="_blank">Edit Customer</a>
                &nbsp;|&nbsp;
                <a href="viewCustOrders.asp?idcustomer=<%=pidcustomer%>" target="_blank">Other Orders by this Customer</a>
			</td>
		</tr>
        <tr> 
        <td width="20%">First Name:</td>
        <td width="80%"><%=rs("name")%></td>
        </tr>
        <tr> 
        <td>Last Name:</td>
        <td><%=rs("lastname")%></td>
        </tr>
        <tr> 
        <td>Company:</td>
        <td><%=rs("customerCompany")%></td>
        </tr>
        <tr> 
        <td>Address:</td>
        <td><%=pAddress%></td>
        </tr>
        <tr> 
        <td>City:</td>
        <td><%=pcity%></td>
        </tr>
        <tr> 
        <td>State/Province:</td>
        <td><%=pStateCode%></td>
        </tr>
        <tr> 
        <td>Postal Code:</td>
        <td><%=pzip%></td>
        </tr>
        <tr> 
        <td>Country:</td>
        <td><%=pCountryCode%></td>
        </tr>
        <tr> 
        <td>Telephone:</td>
        <td><%=rs("phone")%></td>
        </tr>	
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>			
<%
	rs.Close
	Set rs=Nothing
%>
					
        <tr> 
            <th colspan="2">PAYMENT INFORMATION</td>
        </tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>			
			<td colspan="2">Payment Method: 
								
				<% payment = split(ppaymentDetails,"||")
								PaymentType=payment(0)
								on error resume next
								If payment(1)="" then
									if err.number<>0 then
										PayCharge=0
									end if
									PayCharge=0
								else
									PayCharge=payment(1)
								end If
								err.number=0
								%>
								<%=PaymentType%>
                        </td>
                        </tr>
                                    
                        <% '=================================
						'Check for CC Order
						'====================================
						myquery="SELECT cardType,cardNumber,expiration,pcSecurityKeyID FROM creditCards WHERE idOrder=" & qry_ID & ";"
						Set rsCC=Server.CreateObject("ADODB.Recordset")
						set rsCC=connTemp.execute(myquery)
						if NOT rsCC.eof then 
						pcardType=rsCC("cardType")
						pcardNumber=rsCC("cardNumber")
						pexpiration=rsCC("expiration")
						pcardComments=rsCC("comments")
						pcv_SecurityKeyID=rsCC("pcSecurityKeyID")
								
						CCT=pcardType
						ccp="Y"  
						If CCT="M" then
						CCType="MasterCard"
						end if
						If CCT="V" then
						CCType="Visa"
						end if
						If CCT="D" then
						CCType="Discover"
						end if
						If CCT="A" then
						CCType="American Express"
						end if
						If CCT="DC" then
						CCType="Diner's Club"
						end if %>
						
                        <tr> 
                        <td>Card Type: 
                        <input type="hidden" name="ccp" value="<%=ccp%>">
                        <input type="hidden" name="CCT" value="<%=CCT%>">
                        </td>
                        <td><%=CCType%></td>
                        </tr>
                        <tr> 
                        <td>Card Number:</td>
                        <td>
                        <% 
						pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
						
						Dim VarCCNum
							VarCCNum=pcardNumber
							VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)
							%>
							<%=VarCCNum2%>
                        </td>
                        </tr>
                        <tr> 
                        <td>Expiration Date:</td>
                        <td>Month: <%=Month(pexpiration)%> - Year: <%=Year(pexpiration)%></td>
                        </tr>
                        <% end if
						rsCC.Close
						set rsCC=nothing
						'================
						'Check for offline payment Order
						'====================================
						query="SELECT idPayment, AccNum FROM offlinepayments WHERE idOrder=" & qry_ID & ";"
						Set rsCC=Server.CreateObject("ADODB.Recordset")
						Set rsCC=connTemp.execute(query)
						if rsCC.eof then
						else 
						query="SELECT CReq,Cprompt FROM Paytypes WHERE idPayment="& rsCC("idPayment")
						Set rstemp=Server.CreateObject("ADODB.Recordset")
						Set rstemp=connTemp.execute(query)
						If rstemp.eof then
							tempCReq="0"
						else
							tempCReq=rstemp("CReq")
							tempCprompt=rstemp("Cprompt")
						end if %>
						
                        <tr> 
                        <td colspan="2">Terms: 
                        <% payment = split(ppaymentDetails,"||")
								PaymentType=payment(0)
								on error resume next
								If payment(1)="" then
									if err.number<>0 then
										PayCharge=0
									end if
									PayCharge=0
								else
									PayCharge=payment(1)
								end If
								err.number=0
								%>
								<%=PaymentType%> </td>
                            </tr>
                            <% if tempCReq="-1" then %>
                            <tr> 
                            <td><%=tempCprompt%>:&nbsp;<%=rsCC("AccNum")%></td>
                            <td>&nbsp;</td>
                            </tr>
						
						<% end if %>
                                                
                        <% end if %>
                                                
                        <% if PayCharge>0 then %>
                        <tr> 
                        <td colspan="2">Additional Fee for Payment Type: <%=money(PayCharge)%> </td>
                        </tr>
                        <% end if %>
						
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
						
        <!--START - for ProductCart's Rewards Program -->
                                
        <% If RewardsActive <> 0 And piRewardPoints > 0 Then 
        iDollarValue = piRewardPoints * (RewardsPercent / 100)						
        %>
        <!-- Got Rewards -->
        <tr> 
        <td colspan="2">The customer used <%=piRewardPoints & " " & RewardsLabel%> on this purchase <br>for a dollar value of <%=scCurSign&money(iDollarValue)%>.</td>
        </tr>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
        <% End If %>
                                
        <!--END - for ProductCart's Rewards Program -->
                                
        <!--START - for ProductCart's Rewards Program -->
                                
        <% If RewardsActive <> 0 And piRewardPointsCustAccrued > 0 Then %>
        <!-- Got Rewards -->
        <tr> 
        <td colspan="2">The customer accrued <%=piRewardPointsCustAccrued & " " & RewardsLabel%> on this purchase.</td>
        </tr>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
        <% End If %>
                                
        <!--END - for ProductCart's Rewards Program -->
					
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">SHIPPING INFORMATION</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
						
<tr>
<td colspan="2">Shipping Method: 
							
<% shipping=split(pshipmentDetails,",")
							if ubound(shipping)>1 then
								if NOT isNumeric(trim(shipping(2))) then
									varShip="0"
									response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
								else
									Shipper=shipping(0)
									Service=shipping(1)
									Postage=trim(shipping(2))
									if ubound(shipping)=>3 then
										serviceHandlingFee=trim(shipping(3))
										if NOT isNumeric(serviceHandlingFee) then
											serviceHandlingFee=0
										end if
									else
										serviceHandlingFee=0
									end if
								end if
								response.write Service
							else
								varShip="0"
								response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
							end if  %>
							</td>
</tr>
					  
<% if varShip<>"0" then %>
<tr> 
<td>Shipping Address: </td>
<td> 
<% if pshippingAddress="" then %>
<% response.write "(Same as billing address)" %>
<% else %>
<%=pshippingAddress%>
<% end if %>
</td>
</tr>
<tr> 
<td>City:</td>
<td> 
<% if not pshippingAddress="" then %>
<%=pshippingcity%>
<% end if %>
&nbsp;</td>
</tr>
<tr> 
<td>State Code:</td>
<td> 
<% if not pshippingAddress="" then %>
<%=pshippingStateCode%>
<% end if %>
&nbsp;</td>
</tr>
<tr> 
<td>Postal Code:</td>
<td> 
<% if not pshippingAddress="" then %>
<%=pshippingZip%>
<% end if %>
&nbsp;</td>
</tr>
<tr> 
<td>Country Code:</td>
<td> 
<% if not pshippingAddress="" then %>
<%=pshippingCountryCode%>
<% end if %>
&nbsp;</td>
</tr>
<tr> 
<td>Phone:</td>
<td>
<% if not pshippingPhone="" then %>
<%=pshippingPhone%>
<% end if %>
</td>
</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
<% end if %>				

					
<% If pidaffiliate>"1" then %>
		<tr> 
			<th colspan="2">AFFILIATE INFORMATION</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td>Affiliate ID:</td>
		<td><%=pidaffiliate%> </td>
		</tr>
		<tr> 
			<td>Commission:</td>
			<td><%=money(paffiliatePay)%></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
<% end if %>

		<tr> 
		<th colspan="2">PRODUCT DETAILS</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2"> 
				<table class="pcCPcontent">
				<tr> 
				<td valign="top"><b>QTY </b></td>
				<td valign="top"><b>ITEM DESCRIPTION</b><b></b></td>
				<td valign="top"><div align="right"><b>UNIT PRICE</b></div></td>
				<td valign="top"><div align="right"><b>TOTAL</b></div></td>
				</tr>
																		
				<% query="SELECT ProductsOrdered.idProduct, ProductsOrdered.service, ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.unitPrice, ProductsOrdered.unitCost, ProductsOrdered.xfdetails"
												'BTO ADDON-S
												if scBTO=1 then
													query=query&", ProductsOrdered.idconfigSession"
												end if
												'BTO ADDON-E
												query=query&" FROM ProductsOrdered WHERE ProductsOrdered.idOrder=" & qry_ID & ";"
												Set rstemp=Server.CreateObject("ADODB.Recordset")
												set rstemp=connTemp.execute(query)
												Do until rstemp.eof
												
													'// Product Options Arrays
													'===================================================
													pcv_idProduct = rstemp("idproduct")
													pcv_service = rstemp("service")
													pcv_quantity = rstemp("quantity")
													pcv_strSelectedOptions = rsTemp("pcPrdOrd_SelectedOptions") ' Column 11
													pcv_strOptionsPriceArray = rsTemp("pcPrdOrd_OptionsPriceArray") ' Column 25
													pcv_strOptionsArray = rsTemp("pcPrdOrd_OptionsArray") ' Column 4
													pcv_unitPrice = rstemp("unitprice")
													pcv_unitCost
													pxdetails=trim(rstemp("xfdetails"))
													'BTO ADDON-S
													if scBTO=1 then
														pIdConfigSession=trim(rstemp("idconfigSession"))
													end if	
													'===================================================
													
													pxdetails=trim(rstemp("xfdetails"))
													query="SELECT * FROM products WHERE idproduct="& pcv_idProduct
													Set rstemp2=Server.CreateObject("ADODB.Recordset")
													set rstemp2=connTemp.execute(query)
													%>
																		
				<tr> 
																			
				<td valign="top"><%=pcv_quantity%></td>
				<td valign="top"><%=rstemp2("description")%></td>
				<td valign="top"> <div align="right"><%=money(pcv_unitPrice)%></div></td>
				<td valign="top"> <div align="right"><%=money(((pcv_unitPrice)*(pcv_quantity)))%></div></td>
				</tr>
																		
				<% 'BTO ADDON-S
													 if scBTO=1 then %>
																		
						<% 
															if pIdConfigSession<>"0" then 
																query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
																set rsConfigObj=connTemp.execute(query)
												
																'if err.number <> 0 then
																	'call closeDb()
																	'response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails: "&err.description) 
																'end if
																
																stringProducts=rsConfigObj("stringProducts")
																stringValues=rsConfigObj("stringValues")
																stringCategories=rsConfigObj("stringCategories")
																ArrProduct=Split(stringProducts, ",")
																ArrValue=Split(stringValues, ",")
																ArrCategory=Split(stringCategories, ",")
																%>
																		
				<tr> 
																			
				<td valign="top">&nbsp;</td>
				<td valign="top">
                <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFCC">
				<tr> 
				<td colspan="2"><u>Customized:</u></td>
				</tr>
																					
				<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
																	query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
																	set rsConfigObj=connTemp.execute(query)%>
																					
				<tr> 
																						
				<td width="30%"><%=rsConfigObj("categoryDesc")%>:</td>
				<td width="70%"><%=rsConfigObj("description")%></td>
				</tr>
																					
				<% set rsConfigObj=nothing
				next
				set rsConfigObj=nothing %>
                </table></td>
				<td valign="top">&nbsp;</td>
				<td valign="top">&nbsp;</td>
				</tr>
																		
				<% end if %>
																		
				<% 
				end if
				'BTO ADDON-E 
				%>
																		
				<!-- start options -->
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: SHOW PRODUCT OPTIONS
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
						pcv_strSelectedOptions = ""
					end if
					
					if len(pcv_strSelectedOptions)>0 then 
					%>
					<tr valign="top">
						<td>&nbsp;</td>
						<td colspan="3">							
							
							<table width="100%" cellspacing="0" cellpadding="2" bgcolor="#FFFFCC" border="0">
								<tr> 
									<td colspan="3" class="invoiceNob"><u>Options:</u></td>
								</tr>

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
							For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
							%>
							<tr>
								<td class="invoiceNob" width="70%"><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></td>
																
								<% 
								tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
								if tempPrice="" or tempPrice=0 then
									%><%
								else 
									%>
									<td width="20%" class="invoiceNob"><div align="right"><%=scCurSign&money(tempPrice)%></div></td>	
									<td width="10%" class="invoiceNob">	
									<div align="right">			 
									<%
									tAprice=(tempPrice*Cdbl(pquantity))
									response.write scCurSign&money(tAprice) 
									%>
									</div>
							  		</td>
									<%
								end if 
								%>
								
							</tr>
							<%
							Next
							'#####################
							' END LOOP
							'#####################					
							%>
					</table>
						
					</td>
				</tr>															
				<%					
				end if
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: SHOW PRODUCT OPTIONS
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%>
				<!-- end options -->							
									
				
				<% if pxdetails<>"" then %>
																		
				<tr> 
																			
				<td height="29" valign="top">&nbsp;</td>
				<td valign="top"><%=replace(pxdetails,"|","<br>")%></td>
				<td valign="top">&nbsp;</td>
				<td valign="top">&nbsp;</td>
				</tr>
																		
				<% end if %>
																		
				<!-- end of option descriptions -->																		
				<% 
					rstemp.moveNext
					loop 
					set rstemp=nothing %>
																		
				<% 'RP ADDON-S
						If RewardsActive <> 0 then
							if piRewardValue>0 then %>
				<tr> 
				<td height="29" valign="top">&nbsp;</td>
				<% if RewardsLabel="" then
							RewardsLabel="Rewards Program"
						end if %>
				<td valign="top"><%=RewardsLabel%></td>
				<td valign="top">&nbsp;</td>
				<td valign="top"><div align="right">-<%=money(piRewardValue)%></div></td>
				</tr>
				<% end if 
					end if 
				'RP ADDON-E %>
																		
				<% if Cstr(pdiscountDetails)="No discounts applied." then %>
																		
				<% else %>
																		
				<% end if %>
																		
				<tr> 
																			
				<td valign="top" colspan="2" rowspan="5">&nbsp;</td>
				<td valign="top"><div align="right"><b>SHIPPING</b></div></td>
				<td valign="top"><div align="right"><%=money(postage)%></div></td>
				</tr>
				
				<% if serviceHandlingFee<>0 then %>													
				<tr>																			
				<td valign="top"> <div align="right"><b>SHIPPING &amp;<br>
				HANDLING FEES:</b></div></td>
				<td valign="top"> <div align="right"><%=money(serviceHandlingFee)%></div></td>
				</tr>
				<% end if %>  
				
				<% if pord_VAT>0 then %> 
				<tr> 
					<td valign="top"> 
						<div align="right"><b>VAT</b></div></td>
					<td valign="top"> 
						<div align="right"><%=money(pord_VAT)%></div></td>
				</tr>
				<% else %>                          
				<tr> 
					<td valign="top"> 
						<div align="right"><b>TAXES</b></div></td>
					<td valign="top"> 
						<div align="right"><%=money(ptaxAmount)%></div></td>
				</tr>
				<% end if %>                            
				<% if instr(pdiscountDetails,",") then
					DiscountDetailsArry=split(pdiscountDetails,",")
					intArryCnt=ubound(DiscountDetailsArry)
				else
					intArryCnt=0
				end if
					
				dim discounts, discountType 
				
				for k=0 to intArryCnt
					if intArryCnt=0 then
						pTempDiscountDetails=pdiscountDetails
					else
						pTempDiscountDetails=DiscountDetailsArry(k)
					end if
					if instr(pTempDiscountDetails,"- ||") then
						discounts = split(pTempDiscountDetails,"- ||")
						discountType = discounts(0)
						discount = discounts(1)
						%>
						<tr> 
						<td valign="top"> <div align="right"><b>DISCOUNTS </b></div></td>
						<td valign="top"> <div align="right">
						<% response.write "-"&money(discount) %>
						</div></td>
						</tr>
					<% end if
				Next %>
																		
				<tr> 
				<td valign="top"> <div align="right"><b>TOTAL</b></div></td>
				<td valign="top"> <div align="right"><%=scCurSign&money(ptotal)%></div></td>
				</tr>
				</table>
	</td>
</tr>
<% if pcomments <> "" then %>		
<tr> 
<td colspan="2"><b>Additional Comments:</b> <%=pcomments%></td>
</tr>
<% end if %>
</table>
</form>
<!--#include file="AdminFooter.asp"-->