<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ShipFromSettings.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<html>
<head>
<title>Order Details - Printable Version</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body>
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td>    
			<% 
			Dim connTemp, query, rs, qry_ID
			
			call openDb() 

			qry_ID=getUserInput(request.querystring("id"),0)
			if not validNum(qry_ID) then
				response.redirect "msg.asp?message=35" 
			end if
			query="SELECT idcustomer, pcOrd_PaymentStatus,orderdate, Address, city, stateCode,state, zip,CountryCode,paymentDetails,shipmentDetails,shippingAddress,shippingCity,shippingStateCode,shippingState,shippingZip,shippingCountryCode,pcOrd_shippingPhone,idAffiliate,affiliatePay,discountDetails,taxAmount,total,comments,orderStatus,processDate,shipDate,shipvia,trackingNum,returnDate,returnReason, ShippingFullName, ord_DeliveryDate, ord_OrderName, OrdShipType, OrdPackageNum, iRewardPoints,iRewardPointsCustAccrued,iRewardValue,address2, shippingCompany, shippingAddress2,taxDetails,rmaCredit,SRF,ord_VAT,pcOrd_CatDiscounts,gwAuthCode,gwTransId,paymentCode FROM orders WHERE idOrder=" & qry_ID & ";"
			Set rs=Server.CreateObject("ADODB.Recordset")
			Set rs=connTemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rs=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
Dim pidcustomer, porderdate, pAddress, pAddress2, pcity, pstateCode, pstate, pzip, pCountryCode, ppaymentDetails, pshipmentDetails, pshippingCompany, pshippingAddress, pshippingAddress2, pshippingCity, pshippingStateCode, pshippingState,pshippingZip, pshippingCountryCode, pshippingPhone, pidAffiliate, paffiliatePay, pdiscountDetails, ptaxAmount, ptotal, pcomments, porderStatus, pprocessDate, pshipDate, pshipvia, ptrackingNum, preturnDate, preturnReason,ptaxDetails,pSRF, pord_DeliveryDate, pord_OrderName, pcgwAuthCode, pcgwTransId, pcpaymentCode
			
		' Show message is the customer is trying to view an order that is not his/hers
			if rs.eof then
				set rs=nothing
				call closeDb() 
				%>
				<table cellpadding="6" border="0">
				<tr> 
					<td class="invoice">
					<%=dictLanguage.Item(Session("language")&"_viewPostings_a")%>
					</td>
				</tr>
				</table>
		<% 
			end if  ' End show message
			
			pidCustomer=rs("idCustomer")
			query="SELECT ProductsOrdered.pcDropShipper_ID FROM pcDropShippersSuppliers INNER JOIN ProductsOrdered ON (pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & ") WHERE ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND ProductsOrdered.idorder=" & qry_ID & ";"
			set rsQ=connTemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsQ=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
	
			if rsQ.eof then
				set rsQ=nothing
				call closeDB()
				response.redirect "msg.asp?message=11"    
			end if
			
			if session("pc_sdsIsDropShipper")="1" then
				query="SELECT pcSupplier_NoticeType As A FROM pcSuppliers WHERE pcSupplier_ID=" & session("pc_idsds") & ";"
			else
				query="SELECT pcDropShipper_NoticeType As A FROM pcDropShippers WHERE pcDropShipper_ID=" & session("pc_idsds") & ";"
			end if
			
			Set rsQ=connTemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsQ=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			pcv_NoticeType=0
			if not rsQ.eof then
				pcv_NoticeType=rsQ("A")
				if IsNull(pcv_NoticeType) or pcv_NoticeType="" then
					pcv_NoticeType=0
				end if
			end if
			set rsQ=nothing
			
			pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
			if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
				pcv_PaymentStatus=0
			end if
			
			porderdate=rs("orderdate")
			porderdate=ShowDateFrmt(porderdate)
			pAddress=rs("Address")
			pcity=rs("city")
			pstateCode=rs("stateCode")
			pstate=rs("state")
			if pstateCode="" then
				pstateCode=pstate
			end if
			pzip=rs("zip")
			pCountryCode=rs("CountryCode")
			ppaymentDetails=trim(rs("paymentDetails"))
			pshipmentDetails=rs("shipmentDetails")
			pshippingAddress=rs("shippingAddress")
			pshippingCity=rs("shippingCity")
			pshippingStateCode=rs("shippingStateCode")
			pshippingState=rs("shippingState")
			if pshippingStateCode="" then
				pshippingStateCode=pshippingState
			end if
			pshippingZip=rs("shippingZip")
			pshippingCountryCode=rs("shippingCountryCode")
			pshippingPhone=rs("pcOrd_shippingPhone")
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
			pshipDate=ShowDateFrmt(pshipdate)
			pshipvia=rs("shipvia")
			ptrackingNum=rs("trackingNum")
			preturnDate=rs("returnDate")
			preturnDate=ShowDateFrmt(preturnDate)
			preturnReason=rs("returnReason")
			pshippingFullName=rs("ShippingFullName")
			pord_DeliveryDate=rs("ord_DeliveryDate")
			pord_OrderName=rs("ord_OrderName")
			pOrdShipType=rs("OrdShipType")
			pOrdPackageNum=rs("ordPackageNum")
			piRewardPoints=rs("iRewardPoints")
			piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
			piRewardValue=rs("iRewardValue")
			pAddress2=rs("address2")
			pshippingCompany=rs("shippingCompany")
			pshippingAddress2=rs("shippingAddress2")
			ptaxDetails=rs("taxDetails")
			pRmaCredit=rs("rmaCredit")
			pSRF=rs("SRF")
			pord_VAT=rs("ord_VAT")
			pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
			if pcv_CatDiscounts<>"" then
			else
			pcv_CatDiscounts="0"
			end if
			pcgwAuthCode=rs("gwAuthCode")
			pcgwTransId=rs("gwTransId")
			pcpaymentCode=rs("paymentCode")

			query="SELECT [name],lastname,customerCompany,phone,email,customertype FROM customers WHERE idCustomer=" & pidcustomer
			Set rsCustObj=Server.CreateObject("ADODB.Recordset")
			Set rsCustObj=connTemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rsCustObj=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			CustomerName=rsCustObj("name")& " " & rsCustObj("lastname")
			CustomerPhone=rsCustObj("phone")
			CustomerEmail=rsCustObj("email")
			CustomerCompany=rsCustObj("customerCompany")
			CustomerType=rsCustObj("customertype")
			set rsCustObj=nothing

			While Not rs.EOF %>
				<table border="0" cellspacing="0" cellpadding="0" width="100%">
					<tr valign="middle"> 
						<td width="18%" align="left"><img src="../pc/catalog/<%=scCompanyLogo%>"></td>
						<td width="39%" height="71" class="invoiceNob"><div align="center">
						<b><%=scCompanyName%></b><br>
						<%=scCompanyAddress%><br>
						<%=scCompanyCity%>, <%=scCompanyState%>&nbsp;<%=scCompanyZip%><br>
						<hr width=100 noshade align="center" color=SILVER>
						<%=scStoreURL%>
						</div>
						</td>
						<td width="43%" valign="bottom"> 
							<table width="50%" align="right" cellpadding="5" cellspacing="0" class="invoice">
								<tr> 
									<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_1")%>  
									<%
									if porderdate <> "" then
										response.write porderdate
									else
										response.write "N/A"
									end if
									%>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td colspan="3" valign="top">&nbsp;</td>
					</tr>
					<tr> 
						<td colspan="3" valign="top">&nbsp;</td>
					</tr>
				</table>
				<table border="0" cellspacing="0" cellpadding="0" width="100%">
					<tr>
						<td width="50%" valign="top" align="left">				
							<table width="95%" cellpadding="5" cellspacing="0" class="invoice">
								<tr>
									<td class="invoice"><strong><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_2")%></strong>:<br>
									
									<%=CustomerName%>
									<br>
									<% if CustomerCompany<>"" then 
										response.write CustomerCompany&"<BR>"
									end if %>
									<%=pAddress%>
									<br>
									<% if pAddress2<>"" then 
										response.write pAddress2&"<BR>"
									end if %>
									<% response.write pcity&", "&pStateCode&" "&pzip %>
									<% if pCountryCode <> scShipFromPostalCountry then
										response.write "<BR>" & pCountryCode
									end if %>
									<br><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_3") & CustomerPhone%>
									<br><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_4") & CustomerEmail%>
									</td>
								</tr>
							</table>
							<br>
						</td>
						<td rowspan="2" width="50%" valign="top">             
							<table align="right" width="95%" cellpadding="5" cellspacing="0" class="invoice">
								<tr> 
									<td class="invoice"><b><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_5") & (scpre+int(qry_ID))%></b></td>
								</tr>
								<tr>
								<% ' Calculate customer number using sccustpre constant
										Dim pcCustomerNumber
										if len(sccustpre)>0 then
											pcCustomerNumber = (sccustpre + int(pidcustomer))
										else
											pcCustomerNumber = (int(pidcustomer))
										end if
								%>
									<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_6") & pcCustomerNumber%></td>
								</tr>
								<%	if scOrderName="1" then
									if trim(pord_OrderName) <> "" Then%>
										<tr><td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_7") & pord_OrderName %><br></td></tr>
									<% 	end If
								end if %>
								<% If trim(pord_DeliveryDate) <> "1/1/1900" and trim(pord_DeliveryDate) <> "" Then
									if scDateFrmt="DD/MM/YY" then
										pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
										else
										pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
									end if
									pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
									%>
									<tr><td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_8") & pord_DeliveryDate & ", " & pord_DeliveryTime%><br></td></tr>
								<% End If %>
								<tr> 
									<td class="invoice">
									<%if pcv_NoticeType="1" then%>
										<%response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_7")%><br>
										<%=scCompanyName%><br>
										<%=scCompanyAddress%><br>
										<%=scCompanyCity & ", " & scCompanyState & " - " & scCompanyZip%><br>
										<%=scCompanyCountry%><br>
										<%=scFrmEmail%>
									<%else%>
										<%response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_8")%>
									<%end if%>
									</td>
								</tr>
			<tr>           
			<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_16")%>
				<% 
				Shipper=""
				Service=""
				Postage=""
				serviceHandlingFee=""
				If pSRF="1" then
					response.write ship_dictLanguage.Item(Session("language")&"_noShip_b")
				else
					'get shipping details...
					shipping=split(pshipmentDetails,",")
					if ubound(shipping)>1 then
						if NOT isNumeric(trim(shipping(2))) then
							varShip="0"
							response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
						else
							Shipper=shipping(0)
							Service=shipping(1)
							Postage=trim(shipping(2))
						end if
						if len(Service)>0 then
							response.write Service
						End If
					else
						varShip="0"
						response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
					end if
				end if
				%>
			</td>
		</tr>
	<%
		if pOrdShipType=0 then
			pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_18")
		else
			pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_19")
		end if
		if varShip<>"0" then
	%>
			<tr> 
				<td class="invoice"><%=dictLanguage.Item(Session("language")&"_sds_custviewpastD_17")%> <%=pDisShipType%></td>
			</tr>
	<%
	end if
	%>
								</table>
							</td>
						</tr>
						<tr> 
							<td width="50%" valign="top"> 
							<table width="95%" cellpadding="5" cellspacing="0" class="invoice">             
									<tr>             
										<td class="invoice"><strong><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_17")%></strong>:
										<br>
                    <% if pshippingAddress="" then %>
                    	
						<% response.write "(Same as billing address)" %>

                    <% ELSE %>
											
											<% 
											if pshippingFullName<>"" then
												response.write pshippingFullName
											else
												response.write CustomerName
											end if %>
											<br>									
											<% if pshippingCompany<>"" then 
												response.write pshippingCompany & "<br>"
											else
												if (pshippingFullName = "" or pshippingFullName = CustomerName) and customerCompany <> "" then
													response.write customerCompany & "<br>"
												end if											
											end if %>
											<%=pshippingAddress%><br>
											<% if pshippingAddress2<>"" then 
												response.write pshippingAddress2&"<BR>"
											end if %>
											<%=pshippingcity%>, <%=pshippingStateCode%>&nbsp;<%=pshippingZip%>
											<% if pShippingCountryCode <> scShipFromPostalCountry then
												response.write "<BR>" & pShippingCountryCode
											end if %>
											<br>
											<% 
												if pshippingPhone <> "" then
													response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_3") & pshippingPhone
												end if
											%>
                    <% END IF %>
										</td>
								</tr>
							</table>
					</td>
					</tr>
					<tr> 
						<td width="50%">&nbsp;</td>
						<td width="50%" valign="top">&nbsp;</td>
					</tr>
				</table>
      	<div align="center"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td>
						<table width="100%" cellpadding="5" cellspacing="0" border="1" class="invoice">
                <tr> 
                  <td class="invoice" width="7%"><b><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_18")%></b></td>
                  <td width="63%" align="left" valign="top" class="invoice"><b><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_19")%></b></td>
                  <td class="invoice" width="16%"></td>
                  <td class="invoice" width="14%"></td>
                </tr>
                <% 
				Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
				Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
				Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
			
				query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.unitPrice, ProductsOrdered.xfdetails,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts"
								'BTO ADDON-S
								if scBTO=1 then
						    	query=query&", ProductsOrdered.idconfigSession"
								end if
								'BTO ADDON-E
								query=query&" FROM pcDropShippersSuppliers,ProductsOrdered WHERE ProductsOrdered.idOrder=" & qry_ID & " AND pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & ";"
								Set rsTemp=Server.CreateObject("ADODB.Recordset")
								set rsTemp=connTemp.execute(query)
								if err.number<>0 then
									'//Logs error to the database
									call LogErrorToDatabase()
									'//clear any objects
									set rsTemp=nothing
									'//close any connections
									call closedb()
									'//redirect to error page
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								
								Do until rsTemp.eof
									pidProduct=rstemp("idProduct")
									pquantity=rstemp("quantity")

									'// Product Options Arrays
									pcv_strSelectedOptions = rsTemp("pcPrdOrd_SelectedOptions") ' Column 11
									pcv_strOptionsPriceArray = rsTemp("pcPrdOrd_OptionsPriceArray") ' Column 25
									pcv_strOptionsArray = rsTemp("pcPrdOrd_OptionsArray") ' Column 4

									punitPrice=rstemp("unitPrice")
									pxdetails=rstemp("xfdetails")
										pxdetails=replace(pxdetails,"|","<br>")
										pxdetails=replace(pxdetails,"::",":")

									QDiscounts=rstemp("QDiscounts")
									ItemsDiscounts=rstemp("ItemsDiscounts")
									if scBTO=1 then
										pidConfigSession=rstemp("idConfigSession")
									end if
									query="SELECT sku,description FROM products WHERE idproduct="& pidProduct
									Set rsTemp2=Server.CreateObject("ADODB.Recordset")
									set rsTemp2=connTemp.execute(query)
									if err.number<>0 then
										'//Logs error to the database
										call LogErrorToDatabase()
										'//clear any objects
										set rsTemp2=nothing
										'//close any connections
										call closedb()
										'//redirect to error page
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									
									psku=rsTemp2("sku")
									pDescription=rsTemp2("description")
									%>
									
		<% 'BTO ADDON-S
		err.number=0
		TotalUnit=0
		If scBTO=1 then
			pIdConfigSession=trim(pidconfigSession)
			if pIdConfigSession<>"0" then 
				query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
				set rsConfigObj=conntemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rsConfigObj=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
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
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rsConfigObj=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
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
										<td class="invoice" width="7%"><%=pquantity%></td>
										<td class="invoice"><%=psku%> - <%=pDescription%></td>
										<td class="invoice" width="16%"></td>
										<td class="invoice" width="14%"></td>
									</tr>
									<% 'BTO ADDON-S
									if scBTO=1 then
										if pIdConfigSession<>"0" then 
											query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
											set rsConfigObj=connTemp.execute(query)
											if err.number<>0 then
												'//Logs error to the database
												call LogErrorToDatabase()
												'//clear any objects
												set rsConfigObj=nothing
												'//close any connections
												call closedb()
												'//redirect to error page
												response.redirect "techErr.asp?err="&pcStrCustRefID
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
											%>
											<tr> 
												<td class="invoice">&nbsp;</td>
												<td class="invoice" colspan="3">
													<table width="100%" cellspacing="2" cellpadding="0" bgcolor="#FFFFCC" class="invoiceBto">
														<tr> 
                        			<td colspan="3" class="invoiceNob"> 
                          			<u><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_22")%></u>:</td>
                      			</tr>
                      			<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
									query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
									set rsConfigObj=connTemp.execute(query)
									if err.number<>0 then
										'//Logs error to the database
										call LogErrorToDatabase()
										'//clear any objects
										set rsConfigObj=nothing
										'//close any connections
										call closedb()
										'//redirect to error page
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									
									if NOT isNumeric(ArrQuantity(i)) then
										pIntQty=1
									else
										pIntQty=ArrQuantity(i)
									end if
									%>
                      				<tr> 
                       					<td width="20%" valign="top" class="invoiceNob"> 
                          				<%=rsConfigObj("categoryDesc")%>:</td>
                        				<td width="70%" valign="top" class="invoiceNob"> 
                         					<%=rsConfigObj("sku")%> - <%=rsConfigObj("description")%><%if pIntQty>1 then%>
														- <%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_18")%>: <%=ArrQuantity(i)%><%end if%></td>
									<%if pnoprices<2 then%>
									<%if (CDbl(ArrValue(i))<>0) or (((ArrQuantity(i)-1)*pQuantity>0) and (ArrPrice(i)>0)) then
									if (ArrQuantity(i)-1)>=0 then
										UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
									else
										UPrice=0
									end if
									'pfPrice=pfPrice+cdbl((ArrValue(i)+UPrice)*pQuantity) %> 
									<%end if%> 
									<% end if %>
									<td width="10%" valign="top" nowrap class="invoiceNob"></td>
                      				</tr>
                      				<% set rsConfigObj=nothing
									next
									set rsConfigObj=nothing %>
                    			</table>
							</td>
                			</tr>
                		<% 
						end if %>
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
						<td class="invoice">&nbsp;</td>
						<td colspan="3" class="invoice">							
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<!--								
								<tr>
									<td width="60%"><p><u></u></p></td>							
									<td align="right" width="40%">	</td>
								</tr>	 
								-->
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
							<td width="60%"><p><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></p></td>
							
							<td align="right" width="40%">									
							<% 
							tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
							
							if tempPrice="" or tempPrice=0 then
								response.write "&nbsp;"
							else
								'// Adjust for Quantity Discounts
								tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice)
								%>
								<table width="100%" cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td align="right" width="62%">&nbsp;</td>
										<td align="right" width="38%">&nbsp;</td>
									</tr>
								</table>
							<% 
							end if 
							%>			
							
							</td>
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
						<td class="invoice" width="7%">&nbsp;</td>
						<td class="invoice" style="padding-left:10px;"><%=pxdetails%></td>
						<td class="invoice" width="16%">&nbsp; 
						</td>
						<td class="invoice" width="14%">&nbsp; 
						</td>
					</tr>
				<% end if %>
				
                <% 'BTO ADDON-S
									pRowPrice=(punitPrice)*(pquantity)
									If scBTO=1 then
										pidConfigSession=trim(pidConfigSession)
										if pidConfigSession<>"0" then
											MyTest=0
											ItemsDiscounts=trim(ItemsDiscounts)
											if (ItemsDiscounts<>"") and (CDbl(ItemsDiscounts)<>"0") then
												MyTest=1%>
												<%
												pRowPrice=pRowPrice-Cdbl(ItemsDiscounts)
											end if
											%>
               			 					<% 'BTO Additional Charges-S
											if scBTO=1 then
												if pIdConfigSession<>"0" then 
													query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
													set rsConfigObj=conntemp.execute(query)
													if err.number<>0 then
														'//Logs error to the database
														call LogErrorToDatabase()
														'//clear any objects
														set rsConfigObj=nothing
														'//close any connections
														call closedb()
														'//redirect to error page
														response.redirect "techErr.asp?err="&pcStrCustRefID
													end if
									
													stringCProducts=rsConfigObj("stringCProducts")
													stringCValues=rsConfigObj("stringCValues")
													stringCCategories=rsConfigObj("stringCCategories")
													ArrCProduct=Split(stringCProducts, ",")
													ArrCValue=Split(stringCValues, ",")
													ArrCCategory=Split(stringCCategories, ",")
													if ArrCProduct(0)<>"na" then
														MyTest=1%>
														<tr> 
															<td class="invoice">&nbsp;</td>
															<td class="invoice" colspan="3">
																<table width="100%" border="0" cellspacing="2" cellpadding="0" bgcolor="#FFFFCC" class="invoiceBto">
                     				 								<tr> 
																		<td colspan="3" class="invoiceNob"> 
																			<u><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_24")%></u></td>
																	</tr>
																	<% 
																	Charges=0
																	for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
																		query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
																		set rsConfigObj=connTemp.execute(query)
																		if err.number<>0 then
																			'//Logs error to the database
																			call LogErrorToDatabase()
																			'//clear any objects
																			set rsConfigObj=nothing
																			'//close any connections
																			call closedb()
																			'//redirect to error page
																			response.redirect "techErr.asp?err="&pcStrCustRefID
																		end if
																		
																		if (CDbl(ArrCValue(i))>0)then
																		Charges=Charges+cdbl(ArrCValue(i))
																		end if
																		%>
																		<tr> 
																			<td width="20%" class="invoiceNob" valign="top"> 
																				<%=rsConfigObj("categoryDesc")%>:</td>
																			<td width="70%" class="invoiceNob" valign="top"> 
																				<%=rsConfigObj("sku")%> - <%=rsConfigObj("description")%></td>
																			<td width="10%" nowrap class="invoiceNob" valign="top"></td>
																		</tr>
																		<% set rsConfigObj=nothing
																	next
																	set rsConfigObj=nothing
																	pRowPrice=pRowPrice+Cdbl(Charges)%>
                    								</table>
													</tr>
                									<% 
                									end if 'Have Additional Charges
                								end if %>
                							<%end if
											'BTO Additional Charges %>
										<%end if
									end if 'BTO%>
                	<% rsTemp.moveNext
					loop %>
                
             		</table></td>
          	</tr>
        	</table>
        	<p>
          <%rs.MoveNext
				Wend
				Set rs=Nothing
				%>
			</p>
      </div>
    </td>
  </tr>
  <tr> 
    <td valign="top">&nbsp;</td>
  </tr>
</table>
</div>
</body>
</html>
<% call closeDB() %>