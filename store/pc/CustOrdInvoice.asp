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
<%'Allow Guest Account
AllowGuestAccess=1
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ShipFromSettings.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
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
			dim connTemp, query, rs, qry_ID
			
			call openDb() 

			qry_ID=getUserInput(request.querystring("id"),0)
			if not validNum(qry_ID) then
			   qry_ID=0
			end if
			query="SELECT orders.pcOrd_OrderKey, orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr, idcustomer, orderdate, Address, city, stateCode,state, zip,CountryCode, paymentDetails, shipmentDetails, shippingAddress, shippingCity, shippingStateCode, shippingState, shippingZip, shippingCountryCode, pcOrd_shippingPhone, idAffiliate, affiliatePay, discountDetails, pcOrd_GCDetails, pcOrd_GCAmount, taxAmount, total, comments, orderStatus, processDate, shipDate, shipvia, trackingNum, returnDate, returnReason, ShippingFullName, ord_DeliveryDate, ord_OrderName, iRewardPoints, iRewardPointsCustAccrued, iRewardValue, address2, shippingCompany, shippingAddress2, taxDetails, rmaCredit, SRF, ord_VAT, pcOrd_CatDiscounts, gwAuthCode, gwTransId, paymentCode, pcOrd_GCs, pcOrd_GcCode, pcOrd_GcUsed, pcOrd_IDEvent, pcOrd_GWTotal FROM orders WHERE idOrder=" & qry_ID & ";"
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
			
			Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
			Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
			Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
			
			'// Start: Show message - Is the customer is trying to view an order that is not his/hers			
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
				pidCustomer=0
			else
				pidCustomer=rs("idCustomer")
			end if  
			'// End: Show message
			if int(Session("idcustomer"))<=0 then
				if session("REGidCustomer")>"0" then
					testidCustomer=int(session("REGidCustomer"))
				end if
			else
				testidCustomer=int(Session("idcustomer"))
			end if
			if testidCustomer<>int(pidCustomer) then
				call closeDB()
				response.redirect "msg.asp?message=11"    
			end if
			
			pcOrdKey=rs("pcOrd_OrderKey")
			pshippingEmail=rs("pcOrd_ShippingEmail")
			pshippingFax=rs("pcOrd_ShippingFax")
			pcShowShipAddr=rs("pcOrd_ShowShipAddr")
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
			
				'// START - Test for existence of separate shipping address
				if IsNull(pcShowShipAddr) OR (pcShowShipAddr="") OR (pcShowShipAddr="0") then
					'This might be a v3 store, check another field
					if trim(pshippingAddress)="" then
						pcShowShipAddr=0
						else
						pcShowShipAddr=1
					end if
				end if
				'// END			

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
			GCDetails=rs("pcOrd_GCDetails")
			GCAmount=rs("pcOrd_GCAmount")
			if GCAmount="" OR IsNull(GCAmount) then
				GCAmount=0
			end if
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
			
			'GGG Add-on start
			pGCs=rs("pcOrd_GCs")
			pGiftCode=rs("pcOrd_GcCode")
			pGiftUsed=rs("pcOrd_GcUsed")
			gIDEvent=rs("pcOrd_IDEvent")
			if gIDEvent<>"" then
			else
			gIDEvent="0"
			end if
			pGWTotal=rs("pcOrd_GWTotal")
			if pGWTotal<>"" then
			else
			pGWTotal="0"
			end if
			'GGG Add-on end
			
			'// Check if the Customer is European Union 
			Dim pcv_IsEUMemberState
			pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)

			query="SELECT [name],lastname,customerCompany,phone,email,customertype,fax FROM customers WHERE idCustomer=" & pidcustomer
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
			CustomerFax=rsCustObj("fax")
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
									<%if CustomerPhone<>"" then%>
										<br><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_3") & CustomerPhone%>
									<%end if%>
									<%if CustomerEmail<>"" then%>
										<br><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_4") & CustomerEmail%>
									<%end if%>
									<%if CustomerFax<>"" then%>
										<br>Fax: <%=CustomerFax%>
									<%end if%>
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
								<%if pcOrdKey<>"" then%>
								<tr> 
									<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_opc_common_1") & " "%><b><%=pcOrdKey%></b></td>
								</tr>
								<%end if%>
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
								<%
								'GGG Add-on start
								if gIDEvent<>"0" then
									query="select pcEvents.pcEv_name, pcEvents.pcEv_Date, pcEvents.pcEv_HideAddress, customers.name, customers.lastname from pcEvents,Customers where Customers.idcustomer = pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
									set rs1=server.CreateObject("ADODB.RecordSet")
									set rs1=conntemp.execute(query)
									
									geName=rs1("pcEv_name")
									geDate=rs1("pcEv_Date")
									if year(geDate)="1900" then
									geDate=""
									end if
									if gedate<>"" then
										if scDateFrmt="DD/MM/YY" then
										gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
										else
										gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
										end if
									end if
									geHideAddress=rs1("pcEv_HideAddress")
									if geHideAddress="" then
										geHideAddress=0
									end if
									gReg=rs1("name") & " " & rs1("lastname")
									%>
									<tr><td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_1")%><%=geName %><br></td></tr>
									<tr><td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_2")%><%=geDate %><br></td></tr>
									<tr><td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_3")%><%=gReg %><br></td></tr>
							<% 	else
								geHideAddress=0
							End If
							'GGG Add-on end%>
								<tr> 
									<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_9")%> 
										<% If pSRF="1" then
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
													if ubound(shipping)=>3 then
														serviceHandlingFee=trim(shipping(3))
														if NOT isNumeric(serviceHandlingFee) then
															serviceHandlingFee=0
														end if
													else
														serviceHandlingFee=0
													end if
												end if
												if len(Service)>0 then
													response.write Service
												End If
											else
												varShip="0"
												response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
											end if
										end if %>
									</td>
								</tr>
								<% payment = split(ppaymentDetails,"||")
								PaymentType=trim(payment(0))
								
								'Get payment nickname
								query="SELECT paymentDesc,paymentNickName FROM paytypes WHERE paymentDesc = '" & replace(PaymentType,"'","''") & "';"
								Set rsTemp=Server.CreateObject("ADODB.Recordset")
								Set rsTemp=connTemp.execute(query)
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
										
								if not rsTemp.EOF then
									PaymentName=trim(rsTemp("paymentNickName"))
									else
									PaymentName=""
								end if
								
								Set rsTemp = nothing
								'End get payment nickname
								
								'Get authorization and transaction IDs, if any
								varTransID=""
								varTransName="Transaction ID"
								varAuthCode=""
								varAuthName="Authorization Code"

								if NOT isNull(pcpaymentCode) AND pcpaymentCode<>"" then 
									varShowCCInfo=0
									select case pcpaymentCode
									case "LinkPoint"
										varAry=split(pcgwAuthCode,":")
										varTransName="Approval Number"
										varAuthName="Reference Number"
										varTransID=left(varAry(1),6)
										varAuthCode=right(varAry(1),10)
									case "PFLink", "PFPro", "PFPRO", "PFLINK"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
										varShowCCInfo=1
										varGWInfo="P"
									case "Authorize"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
										varShowCCInfo=1
										if instr(ucase(PaymentType),"CHECK") then
											varShowCCInfo=0
										end if
										varGWInfo="A"
									case "twoCheckout"
										varTransName="2Checkout Order No"
										varTransID=pcgwTransId
									case "BOFA"
										varTransName="Order No"
										varAuthName="Authorization Code"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
									case "WorldPay"
										varTransID=""
										varAuthCode=""
									case "iTransact"
										varTransName="Transaction ID"
										varAuthName="Authorization Code"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
									case "PSI", "PSIGate"
										varTransName="Transaction ID"
										varAuthName="Authorization Code"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
									case "fasttransact", "FastTransact", "FAST","CyberSource"
										varTransName="Transaction ID"
										varAuthName="Authorization Code"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
									case "USAePay","FastCharge"
										varTransName="Transaction reference code"
										varAuthName="Authorization code"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
									case "PxPay"
										varTransName="DPS Transaction Reference Number"
										varAuthName="Authorization code"
										varTransID=pcgwTransId
										varAuthCode=pcgwAuthCode
									 case "Moneris2"					
										 Dim varIDEBIT_ISSCONF, varIDEBIT_ISSNAME,varRespName,varResponseCode
										   varTransName="Sequence Number"
										   varAuthName="Approval Code"
										   varRespName="Response / ISO Code"
										   varTransID=pcgwTransId
										   varAuthCode=pcgwAuthCode
										
										   query = "Select pcPay_MOrder_responseCode, pcPay_MOrder_ISOcode, pcPay_MOrder_IDEBIT_ISSCONF, pcPay_MOrder_IDEBIT_ISSNAME from pcPay_OrdersMoneris Where pcPay_MOrder_TransId='"& pcgwTransId &"';" 
										   set rstemp=server.CreateObject("ADODB.RecordSet")
										   set rstemp=conntemp.execute(query)											  
											if err.number<>0 then
												call LogErrorToDatabase()
												set rstemp=nothing
												call closedb()
												response.redirect "techErr.asp?err="&pcStrCustRefID
											end if
							
										if not rs.eof then
										   varResponseCode = RStemp("pcPay_MOrder_responseCode")
										   varISO_Code = RStemp("pcPay_MOrder_ISOcode")
										   varIDEBIT_ISSCONF = rstemp("pcPay_MOrder_IDEBIT_ISSCONF")
										   varIDEBIT_ISSNAME = rstemp("pcPay_MOrder_IDEBIT_ISSNAME")								 							
										end if
										set rstemp=nothing
									end select
								end if
								
								'End get authorization and transaction IDs
								
								If payment(1)="" then
								 if err.number<>0 then
									PayCharge=0
								 end if
									PayCharge=0
								else
									PayCharge=payment(1)
								end If
								err.number=0
								if instr(PaymentType,"FREE") AND len(PaymentType)<6 then
								else %>
									<tr> 
										<td class="invoice">&nbsp;</td>
									</tr>
									<tr> 
										<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_10")%>
										<%
											if PaymentName <> "" and PaymentName <> PaymentType then
												Dim pcv_strPaymentType
												Select Case PaymentType
													Case "PayPal Website Payments Pro": pcv_strPaymentType=PaymentName
													Case Else: pcv_strPaymentType=PaymentName & " (" & PaymentType & ")"
												End Select
												Response.Write pcv_strPaymentType
												else
												Response.Write PaymentType
											end if
										%>
										<% if PayCharge>0 then %>
											<br><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_11")%> 
											<%response.write scCurSign&money(PayCharge)%>
										<% end if %>
										<% if varTransID<>"" then %>
										<br><%=varTransName%>: <%=varTransID%>
										<% end if %>
										<% if varAuthCode<>"" then %>
										<br><%=varAuthName%>: <%=varAuthCode%>
										<% end if %>
										<%if varResponseCode <> "" and varISO_Code <> "" Then%>
										<BR><%=varRespName%>&nbsp;<%=varResponseCode%>/<%=varISO_Code%>
										<%end if %>
									    <% if varIDEBIT_ISSCONF <> ""  and varIDEBIT_ISSNAME <> "" then %>
										<br><%=dictLanguage.Item(Session("language")&"_CustviewOrd_48")%>
										<BR><%=dictLanguage.Item(Session("language")&"_CustviewOrd_49")%>&nbsp;<%=varIDEBIT_ISSNAME%>
										<BR><%=dictLanguage.Item(Session("language")&"_CustviewOrd_50")%>&nbsp;<%=varIDEBIT_ISSCONF%>						
										<% end if%>
										<br><br>
										</td>
									</tr>
								<% end if %>								
								<% If RewardsActive <> 0 And piRewardPoints > 0 Then 
									iDollarValue = piRewardPoints * (RewardsPercent / 100)%>
									<tr> 
										<td class="invoice"><%=ucase(RewardsLabel)%>:<br>
										<%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_12") & piRewardPoints & " " & RewardsLabel & dictLanguage.Item(Session("language")&"_custOrdInvoice_13") & scCurSign&money(iDollarValue)%></td>
									</tr>
									<% end if %>
									<% If RewardsActive <> 0 And piRewardPointsCustAccrued > 0 Then %>
										<tr> 
											<td class="invoice"><%=ucase(RewardsLabel)%>:<br>
										<%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_14") & piRewardPointsCustAccrued & " " & RewardsLabel & dictLanguage.Item(Session("language")&"_custOrdInvoice_15") %></td>
										</tr>
									<% end if %>
									<% 'if discount was present, show type here
									'Check if more then one discount code was utilized
									if instr(pdiscountDetails,",") then
										DiscountDetailsArry=split(pdiscountDetails,",")
										intArryCnt=ubound(DiscountDetailsArry)
										for k=0 to intArryCnt
											if (DiscountDetailsArry(k)<>"") AND (instr(DiscountDetailsArry(k),"- ||")=0) then
												DiscountDetailsArry(k+1)=DiscountDetailsArry(k)+"," + DiscountDetailsArry(k+1)
												DiscountDetailsArry(k)=""
											end if
										next
									else
										intArryCnt=0
									end if

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
											if discountType<>"" then %>
												<tr>
													<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_16") & discountType%></td>
												</tr>
											<% end if
										end if
									Next %>
									<%'start of gift certificates
									if GCDetails<>"" then
										GCArry=split(GCDetails,"|g|")
										intArryCnt=ubound(GCArry)
				
										for k=0 to intArryCnt
					
											if GCArry(k)<>"" then
												GCInfo = split(GCArry(k),"|s|")
												if GCInfo(2)="" OR IsNull(GCInfo(2)) then
													GCInfo(2)=0
												end if
												%>
												<tr>
													<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_15A") & GCInfo(1) & " (" & GCInfo(0) & ")"%></td>
												</tr>
											<% end if
										Next
									end if
									'end if gift certificates									
									%>
								</table>
							</td>
						</tr>
						<tr> 
							<td width="50%" valign="top"> 
								<% if geHideAddress=0 then %>

								<table width="95%" cellpadding="5" cellspacing="0" class="invoice">
									<tr>             
										<td class="invoice"><strong><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_17")%></strong>:
											<br>
											<% if pcShowShipAddr="0" then %>
												
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
												<% 
												if pshippingEmail <> "" then
													response.write "<br>" & "E-mail: " & pshippingEmail
												end if
												%>
												<% 
												if pshippingPhone <> "" then
													response.write "<br>" & dictLanguage.Item(Session("language")&"_custOrdInvoice_3") & pshippingPhone
												end if
												%>
												<% 
												if pshippingFax <> "" then
													response.write "<br>" & "Fax: " & pshippingFax
												end if
												%>
                    	<% END IF %>
											</td>
									</tr>
								</table>
								<% else %>
								&nbsp;
								<% end if %>
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
                  <td class="invoice" width="16%"> 
                    <div align="right"><b><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_20")%></b></div></td>
                  <td class="invoice" width="14%"> 
                    <div align="right"><b><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_21")%></b></div></td>
                </tr>
                <% 
				query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.unitPrice, ProductsOrdered.QDiscounts, ProductsOrdered.ItemsDiscounts"
				'BTO ADDON-S
				if scBTO=1 then
				query=query&", ProductsOrdered.idconfigSession"
				end if
				'BTO ADDON-E
				query=query&", ProductsOrdered.pcPO_GWOpt, ProductsOrdered.pcPO_GWNote, ProductsOrdered.pcPO_GWPrice, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.xfdetails, pcPrdOrd_BundledDisc FROM ProductsOrdered WHERE ProductsOrdered.idOrder=" & qry_ID & ";"

				Set rsTemp=Server.CreateObject("ADODB.Recordset")
				set rsTemp=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsTemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if		
				Do until rsTemp.eof
					pidProduct=rstemp("idProduct")
					pquantity=rstemp("quantity")
					punitPrice=rstemp("unitPrice")
					QDiscounts=rstemp("QDiscounts")
					ItemsDiscounts=rstemp("ItemsDiscounts")
					if scBTO=1 then
						pidConfigSession=rstemp("idConfigSession")
					end if
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
					'// Product Options Arrays
					pcv_strSelectedOptions = rsTemp("pcPrdOrd_SelectedOptions") ' Column 11
					pcv_strOptionsPriceArray = rsTemp("pcPrdOrd_OptionsPriceArray") ' Column 25
					pcv_strOptionsArray = rsTemp("pcPrdOrd_OptionsArray") ' Column 4

					pxdetails=rstemp("xfdetails")
					pxdetails=replace(pxdetails,"|","<br>")
					pxdetails=replace(pxdetails,"::",":")
					pcPrdOrd_BundledDisc=rstemp("pcPrdOrd_BundledDisc")
					
					query="SELECT sku,description FROM products WHERE idproduct="& pidProduct
					Set rsTemp2=Server.CreateObject("ADODB.Recordset")
					set rsTemp2=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsTemp2=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if		
					psku=rsTemp2("sku")
					pDescription=rsTemp2("description")
					set rsTemp2 = nothing
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
				query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
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
			
				if NOT isNumeric(ArrQuantity(i)) then
					pIntQty=1
				else
					pIntQty=ArrQuantity(i)
				end if
				
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
			if pIdConfigSession<>"0" AND pIdConfigSession<>"" then
				pRowPrice1=Cdbl(pquantity * ( punitPrice1 )) - TotalUnit
				punitPrice1=Round(pRowPrice1/pquantity,2)
			else
				pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
			end if
		else
			punitPrice1=punitPrice
			if pIdConfigSession<>"0" AND pIdConfigSession<>"" then
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
										<td class="invoice" width="16%"> 
											<div align="right"><%=scCurSign&money(punitPrice1)%></div></td>
										<td class="invoice" width="14%"> 
											<div align="right"><%=scCurSign&money(pRowPrice1)%></div></td>
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
                      				query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pidProduct 
									set rsQ=server.CreateObject("ADODB.RecordSet") 
									set rsQ=conntemp.execute(query)
														
									btDisplayQF=rsQ("displayQF")
									set rsQ=nothing
									err.clear 
											
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
									query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
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
                         					<%=rsConfigObj("sku")%> - <%=rsConfigObj("description")%>
											<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%> - <%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_18")%>: <%=ArrQuantity(i)%><%end if%>
										</td>
									<%if pnoprices<2 then%>
									<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
									if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
										if tmpDefault=1 then
											UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
										else
											UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
										end if
									else
										UPrice=0
									end if
									'pfPrice=pfPrice+cdbl((ArrValue(i)+UPrice)*pQuantity) %> 
									<%end if%> 
									<% end if %>
									<td width="10%" valign="top" nowrap class="invoiceNob">
										<div align="right">
										<%if pnoprices<2 then%>
											<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then%>
												<%=scCurSign & money((ArrValue(i)+UPrice)*pQuantity)%>
											<%else
												if tmpDefault=1 then%>
													<%=dictLanguage.Item(Session("language")&"_defaultnotice_1")%>
												<%end if
											end if%>
										<% end if %>
										</div>
									</td>
                      				</tr>
                      				<% set rsConfigObj=nothing
									next
									set rsConfigObj=nothing %>
                    			</table>
							</td>
                			</tr>
                		<% end if %>
                	<% end if
									'	BTO ADDON-E %>
                	
					
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
										<td align="right" width="62%">
											<%=scCurSign&money(tempPrice)%>									  </td>
										<td align="right" width="38%">
											<%									
											tAprice=(tempPrice*Cdbl(pquantity))
											response.write scCurSign&money(tAprice) 
											%>
									  </td>
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

                	<%'BTO ADDON-S
									pRowPrice=(punitPrice)*(pquantity)
				pExtRowPrice=pRowPrice
				Charges=0
									If scBTO=1 then
										pidConfigSession=trim(pidConfigSession)
										if pidConfigSession<>"0" then
											ItemsDiscounts=trim(ItemsDiscounts)
											if ItemsDiscounts="" then
												ItemsDiscounts=0
											end if
											if (ItemsDiscounts<>"") and (CDbl(ItemsDiscounts)<>"0") then
												%>
												<tr valign="top"> 
													<td width="5%" class="invoice">&nbsp;</td>
													<td class="invoice">&nbsp;</td>
													<td class="invoice"><div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_23")%></div></td>
													<td class="invoice"><div align="right"><%=scCurSign&money(-1*ItemsDiscounts)%></div></td>
												</tr>
												<%
												pRowPrice=pRowPrice-Cdbl(ItemsDiscounts)
											end if
											%>
               			 	<% 'BTO Additional Charges-S
												if pIdConfigSession<>"0" then 
													query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
													set rsConfigObj=conntemp.execute(query)
													if err.number<>0 then
														call LogErrorToDatabase()
														set rsConfigObj=nothing
														call closedb()
														response.redirect "techErr.asp?err="&pcStrCustRefID
													end if		
													stringCProducts=rsConfigObj("stringCProducts")
													stringCValues=rsConfigObj("stringCValues")
													stringCCategories=rsConfigObj("stringCCategories")
													ArrCProduct=Split(stringCProducts, ",")
													ArrCValue=Split(stringCValues, ",")
													ArrCCategory=Split(stringCCategories, ",")
							if ArrCProduct(0)<>"na" then %>
														<tr> 
															<td class="invoice">&nbsp;</td>
															<td class="invoice" colspan="3">
																<table width="100%" border="0" cellspacing="2" cellpadding="0" bgcolor="#FFFFCC" class="invoiceBto">
                     				 			<tr> 
																		<td colspan="3" class="invoiceNob"> 
																			<u><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_24")%></u></td>
																	</tr>
											<% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
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
																			<td width="10%" nowrap class="invoiceNob" valign="top"><%if pnoprices<2 then%><%if ArrCValue(i)>0 then%><div align="right"><%=scCurSign & money(ArrCValue(i))%></div><%end if%><%end if%></td>
																		</tr>
																		<% set rsConfigObj=nothing
																	next
																	set rsConfigObj=nothing
																	pRowPrice=pRowPrice+Cdbl(Charges)%>
                    						</table>
									</td>
											</tr>
							<% end if 'Have Additional Charges
						end if
						'BTO Additional Charges %>
                			<% end if
                end if 'BTO %>

                			<% QDiscounts=trim(QDiscounts)
											if QDiscounts="" then
												QDiscounts=0
											end if
				
                if (QDiscounts<>"") and (CDbl(QDiscounts)<>"0") then %>
					<tr> 
                        <td class="invoice">&nbsp;</td>
                        <td class="invoice" colspan="3">
                        	<table width="100%" cellspacing="2" cellpadding="0" bgcolor="#FFFFCC" class="invoiceBto">
                            	<tr> 
                                	<td colspan="3" class="invoiceNob"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_25")%></td>
                                    <td width="10%"><div align="right"><%=scCurSign&money(-1*QDiscounts)%></div></td>
                                </tr>
                            </table>
                        </td>
												</tr>
												<%
												pRowPrice=pRowPrice-Cdbl(QDiscounts)
                end if %>
											
                <% if pExtRowPrice<>pRowPrice then %>                    
												<tr valign="top"> 
													<td class="invoice">&nbsp;</td>
													<td class="invoice">&nbsp;</td>
													<td class="invoice"><div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_26")%></div></td>
													<td class="invoice"><div align="right"><%=scCurSign&money(pRowPrice)%></div></td>
												</tr>
				<% end if %>

                <% 'GGG Add-on start
									if pGWOpt<>"0" then
									query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
									set rsG=connTemp.execute(query)
									if not rsG.eof then%>
											<tr valign="top"> 
												<td class="invoice" width="8%">&nbsp;</td>
												<td class="invoice" colspan="3">
												<b><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_4")%></b> <%=rsG("pcGW_OptName")%><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_5")%>&nbsp;<%=scCurSign & money(pGWPrice)%>
												<%if pGWText<>"" then%>
												<br>
												<b><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_6")%></b><br><%=pGWText%>
												<%end if%>
												</td>
											</tr>
									<%
									end if
									end if
									'GGG Add-on end
									
                                    if pcPrdOrd_BundledDisc>0 then %>
										<tr> 
											<td class="invoice" width="7%">&nbsp;</td>
                                            <td class="invoice"><%=dictLanguage.Item(Session("language")&"_custOrdInvoice_36")%></td>
											<td class="invoice" width="16%">&nbsp;</td>
											<td class="invoice" width="14%"><div align="right">-<%=scCurSign&money(pcPrdOrd_BundledDisc)%> </div></td>
                                        </tr>
                                    <% end if
                					rstemp.moveNext
								loop
								set rstemp=nothing %>
        
                				<% 'RP ADDON-S
								If RewardsActive<>0 Then
									if piRewardValue<>0 then %>
										<tr> 
											<td class="invoice" width="7%">&nbsp;</td>
											<% if RewardsLabel="" then
														RewardsLabel="Rewards Program"
													end if %>
											<td class="invoice"><%=RewardsLabel%></td>
											<td class="invoice" width="16%">&nbsp; 
											</td>
											<td class="invoice" width="14%"><div align="right">-<%=scCurSign&money(piRewardValue) %> </div></td>
										</tr>
				                	<% end if
									End if
								'RP ADDON-E %>
								<%'GGG Add-on start
								if pGWTotal>0 then%>
										<tr>
											<td class="invoice" colspan="3">
											<div align="right"><b><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_7")%></b></div></td>
											<td class="invoice" width="12%"> 
											<div align="right"><%=scCurSign&money(pGWTotal)%></div></td>
										</tr>
								<%
								end if
								'GGG Add-on end%>
								<tr>
									<td class="invoice" colspan="2" rowspan="9">&nbsp;</td> 
									<td class="invoice" width="16%">
										<div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_27")%></div>
									</td>
									<td class="invoice" width="14%"> 
										<div align="right"><%=scCurSign&money(postage)%></div>
									</td>
								</tr>
								
								<% if serviceHandlingFee<>0 then %>
								<tr> 
									<td class="invoice" width="16%"> 
									<div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_28")%></div></td>
									<td class="invoice" width="14%"> 
									<div align="right"><%=scCurSign&money(serviceHandlingFee)%></div></td>
                				</tr>
								<% end if %>
								
								<% if PayCharge>0 then %>
									<tr> 
										<td class="invoice" width="16%"> 
											<div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_29")%></div></td>
										<td class="invoice" width="14%"> 
											<div align="right"><%=scCurSign&money(PayCharge)%></div></td>
									</tr>
								<% end if %>
								<%
								' If the store is using VAT and VAT is > 0, don't show any taxes here, but show VAT after the total
								if NOT (pord_VAT>0) then

										if isNull(ptaxDetails) OR trim(ptaxDetails)="" then %>
											<tr> 
												<td class="invoice" width="16%"> 
													<div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_30")%></div></td>
												<td class="invoice" width="14%"> 
													<div align="right"><%=scCurSign&money(ptaxAmount)%></div></td>
											</tr>
										<% else %>
											<% taxArray=split(ptaxDetails,",")
											tempTaxAmount=0
											for i=0 to (ubound(taxArray)-1)
												taxDesc=split(taxArray(i),"|")
												if taxDesc(0)<>"" then %>
												 <tr> 
													<td class="invoice" width="16%"> 
														<div align="right"><b><%=ucase(taxDesc(0))%></b></div></td>
													<% pDisTax=(money(taxDesc(1))) %>
													<td class="invoice" width="14%"> 
														<div align="right"><%=scCurSign&pDisTax%></div></td>
												</tr>
												<% end if 
											next %>
										<% end if
									end if %>
				                	<% if instr(pdiscountDetails,"- ||") or (pcv_CatDiscounts>"0") then
										'Check if more then one discount code was utilized
										if instr(pdiscountDetails,",") then
											DiscountDetailsArry=split(pdiscountDetails,",")
											intArryCnt=ubound(DiscountDetailsArry)
										else
											intArryCnt=0
										end if
										discount=0
										for k=0 to intArryCnt
											if intArryCnt=0 then
												pTempDiscountDetails=pdiscountDetails
											else
												pTempDiscountDetails=DiscountDetailsArry(k)
											end if
											if instr(pTempDiscountDetails,"- ||") then 
												discounts = split(pTempDiscountDetails,"- ||")
												discountType = discounts(0)
												tdiscount = discounts(1)
											else
												tdiscount=0
											end if
											discount=discount+tdiscount
										Next %>
                						<tr> 
											<td class="invoice" width="16%"> 
												<div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_31")%></div></td>
											<td class="invoice" width="14%"> 
												<div align="right">-<%=scCurSign&money(discount+pcv_CatDiscounts)%></div></td>
										</tr>
									<% end if %>
									<%'GGG Add-on start
									IF GCAmount>"0" THEN%>
									<tr> 
										<td class="invoice" width="16%"> 
											<div align="right"><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_8")%></div></td>
										<td class="invoice" width="14%"> 
											<div align="right">-<%=scCurSign&money(GCAmount)%></div></td>
									</tr>
									<%END IF
									'GGG Add-on end%>
                					<tr> 
										<td class="invoice" width="16%"> 
											<div align="right"><b><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_32")%></b>:</div></td>
										<td class="invoice" width="14%"> 
											<div align="right"><%=scCurSign&money(ptotal)%></div></td>
									</tr>
									<% 
									' If the store is using VAT and VAT > 0, show it here
									if pord_VAT>0 then %>

                                        
										<% if pcv_IsEUMemberState=1 then %>
										<tr> 
											<td class="invoice" align="right" colspan="2"><span class="pcSmallText"><% response.write dictLanguage.Item(Session("language")&"_orderverify_35") & scCurSign&money(pord_VAT)%></span></td>
										</tr>
										<% else %>
										<tr> 
											<td class="invoice" align="right" colspan="2"><span class="pcSmallText"><% response.write dictLanguage.Item(Session("language")&"_orderverify_42") & scCurSign&money(pord_VAT)%></span></td>
										</tr>
										<% end if %> 
                                        
                                        
									<% end if %>
									<% if NOT isNull(prmaCredit) AND prmaCredit<>"" AND prmaCredit>0 then %>
										<tr>
											<td align="left" valign="top" class="invoice">
												<div align="right"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_34")%></div>
											</td>
											<td align="left" valign="top" class="invoice">
												<div align="right">-<%=scCurSign&money(pRmaCredit)%></div>
											</td>
										</tr>
									<% end if %>
             		</table></td>
          	</tr>
        	</table>
          <%rs.MoveNext
			Wend
			Set rs=Nothing
			%>
			
			<%'GGG Add-on start
			IF (GCDetails<>"") then %>
			<br>
			<table class="pcShowContent">
				<tr>
					<th colspan="2"><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_9")%></th>
				</tr>
				<%
				GCArry=split(GCDetails,"|g|")
				intArryCnt=ubound(GCArry)
			
				for k=0 to intArryCnt
				
				if GCArry(k)<>"" then
					GCInfo = split(GCArry(k),"|s|")
					if GCInfo(2)="" OR IsNull(GCInfo(2)) then
					GCInfo(2)=0
					end if
					pGiftCode=GCInfo(0)
					pGiftUsed=GCInfo(2)
				query="select products.IDProduct,products.Description from pcGCOrdered,Products where products.idproduct=pcGCOrdered.pcGO_idproduct and pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
				set rsG=connTemp.execute(query)

				if not rsG.eof then
					pIdproduct=rsG("idproduct")
					pName=rsG("Description")
					pCode=pGiftCode
					%>
				<tr> 
					<td width="18%" nowrap><b><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_10")%></b></td>
					<td width="82%"><b><%=pName%></b></td>
				</tr>
				<tr> 
					<td width="18%" nowrap valign="top">&nbsp;</td>
					<td width="82%" valign="top">
					<%
					query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
					set rs19=connTemp.execute(query)

					if not rs19.eof then%>
						<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_11")%><b><%=rs19("pcGO_GcCode")%></b><br>
						<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_12")%><%=scCurSign & money(pGiftUsed)%><br><br>
						<%
						pGCAmount=rs19("pcGO_Amount")
						if cdbl(pGCAmount)<=0 then%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_13")%>
						<%else%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_14")%><b><%=scCurSign & money(pGCAmount)%></b>
							<br>
							<%pExpDate=rs19("pcGO_ExpDate")
							if year(pExpDate)="1900" then%>
								<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_15")%>
							<%else
								if scDateFrmt="DD/MM/YY" then
									pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
								else
									pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
								end if%>
								<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_16")%><font color=#ff0000><b><%=pExpDate%></b>
							<%end if%>
							<br>
							<%
							pGCStatus=rs19("pcGO_Status")
							if pGCStatus="1" then%>
								<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_17")%>
							<%else%>
								<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_18")%>
							<%end if%>
						<%end if%>
						<br><br>
					<%end if
					set rs19=nothing%>
					</td>
				</tr>
				<%end if
				set rsG=nothing
				end if
				Next%>
			</table>
			<% END IF
			'GGG Add-on end%>
			
			<%'GGG Add-on start
			IF (pGCs<>"") and (pGCs="1") then %>
			<br>
			<table class="pcShowContent">
				<tr>
					<td colspan="2"><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_19")%></td>
				</tr>
				<%
				query="select * from ProductsOrdered WHERE idOrder="& qry_ID
				set rs11=connTemp.execute(query)
				do while not rs11.eof
					query="select products.Description,pcGCOrdered.pcGO_GcCode from Products,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& qry_ID
					set rsG=connTemp.execute(query)

					if not rsG.eof then
						gIdproduct=rs11("idproduct")
						gName=rsG("Description")
						gCode=rsG("pcGO_GcCode")
						%>
				<tr> 
					<td width="18%" nowrap><b><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_9")%></b></td>
					<td width="82%"><b><%=gName%></b></td>
				</tr>
				<tr> 
					<td width="18%" nowrap valign="top">&nbsp;</td>
					<td width="82%" valign="top">
					<%
					query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & qry_ID
					set rs19=connTemp.execute(query)

					do while not rs19.eof%>
						<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_11")%><b><%=rs19("pcGO_GcCode")%></b><br>
						<%pExpDate=rs19("pcGO_ExpDate")
						if year(pExpDate)="1900" then%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_15")%>
						<%else
							if scDateFrmt="DD/MM/YY" then
								pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
							else
								pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
							end if%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_16")%><font color=#ff0000><b><%=pExpDate%></b></font>
						<%end if%>
						<br>
						<%
						pGCAmount=rs19("pcGO_Amount")
						if cdbl(pGCAmount)<=0 then%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_13")%>
						<%else%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_14")%><b><%=scCurSign & money(pGCAmount)%></b>
						<%end if%><br>
						<%
						pGCStatus=rs19("pcGO_Status")
						if pGCStatus="1" then%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_17")%>
						<%else%>
							<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_18")%>
						<%end if%>
						<br><br>
						<%rs19.movenext
					loop
					set rs19=nothing
					%>
					</td>
				</tr>
				<%end if
				set rsG=nothing
				rs11.MoveNext
				loop
				set rs11=nothing
				%>
			</table>
			<% END IF
			'GGG Add-on end%>
			
			<% if pcomments<>"" then %>
        	<p>
					<table width="100%" align="left" cellpadding="5" cellspacing="0" border="1" class="invoice">
					<tr> 
						<td class="invoice"><%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_35")%>
								<br>
								<br>						
								<%=pcomments%>
								
								<br>
						</td>
					</tr>
       				</table>
       		</p>
			<% end if %>
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