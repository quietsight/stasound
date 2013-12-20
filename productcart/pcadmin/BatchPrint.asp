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
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/ShipFromSettings.asp" -->
<!--#include file="sm_inc.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>Order Details - Printer Friendly</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<STYLE TYPE="text/css">
	P.breakhere {page-break-before: always}
	.text {
		font-family: Arial, Helvetica, sans-serif;
		font-size: 12px;
	}
	</STYLE>
	<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin: 10px; background-image: none;">
<%
dim connTemp, query, rs, qry_ID

'///////////////////////////////
'// SHOW VAT ID 
'// Change "0" to "-1" to show
'///////////////////////////////
pcv_strShowVatId=0 ' -1
'///////////////////////////////

'///////////////////////////////
'// SHOW INTL. ID 
'// Change "0" to "1" to show
'///////////////////////////////
pcv_strShowSSN=0 ' -1
'///////////////////////////////

	IF request.Form("GO")="Submit" THEN
   
		Function CreateInvoiceNum(prenum,tindex)
		Dim tmp1, tmp2
		Dim lenindex
		
		tmp1=prenum
		tmp2=""
		m=len(tmp1)
		do while (m>0) and (isnumeric(mid(tmp1,m,1)))
			tmp2=mid(tmp1,m,1) & tmp2
			m=m-1
			if m=0 then
				exit do
			end if
		loop

		tmp1=mid(tmp1,1,m)

		if tmp2<>"" then
			lenindex=len(tmp2)
		else
			lenindex=0
		end if

		if lenindex<5 then
			lenindex=5
		end if

		if tmp2="" then
			tmp2="0"
		end if

		tmp2=cstr(clng(tmp2)+tindex)

		do while len(tmp2)<lenindex
			tmp2="0" & tmp2
		loop

		CreateInvoiceNum=tmp1 & tmp2

	end function

	call openDb()
	A=split(request("id"),"***")
	For k=lbound(A) to ubound(A)
		IF A(k)<>"" then
			if k<>lbound(A) then%>
				<P CLASS="breakhere">&nbsp;</p>
			<%end if%>
			<%qry_ID=A(k)
			query="SELECT orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr,idcustomer, orderdate, Address, city, state, stateCode, zip, CountryCode, paymentDetails, shipmentDetails, shippingAddress, shippingCity, shippingStateCode, shippingState, shippingZip, shippingCountryCode,pcOrd_shippingPhone,idAffiliate, affiliatePay, discountDetails, pcOrd_GCDetails, pcOrd_GCAmount, taxAmount, total, comments, orderStatus, processDate, shipDate, shipvia, trackingNum, returnDate, returnReason, ShippingFullName, ord_DeliveryDate, ord_OrderName, OrdShipType, OrdPackageNum, iRewardPoints, iRewardPointsCustAccrued, iRewardValue,address2, shippingCompany, shippingAddress2, taxDetails, rmaCredit, SRF, ord_VAT, pcOrd_CatDiscounts, gwAuthCode, gwTransId, paymentCode, pcOrd_GCs, pcOrd_GcCode, pcOrd_GcUsed, pcOrd_IDEvent, pcOrd_GWTotal, adminComments, pcOrd_ShipWeight FROM orders WHERE idOrder=" & qry_ID & ";"

			Set rs=Server.CreateObject("ADODB.Recordset")
			Set rs=connTemp.execute(query)
			Dim pidcustomer, porderdate, pAddress, pAddress2, pcity, pstate, pstateCode, pzip, pCountryCode, ppaymentDetails, pshipmentDetails, pshippingCompany, pshippingAddress, pshippingAddress2, pshippingCity, pshippingStateCode, pshippingState, pshippingZip, pshippingCountryCode, pshippingPhone, pidAffiliate, paffiliatePay, pdiscountDetails, ptaxAmount, ptotal, pcomments, porderStatus, pprocessDate, pshipDate, pshipvia, ptrackingNum, preturnDate, preturnReason,ptaxDetails,pSRF, pord_DeliveryDate, pord_OrderName, pOrdShipType, pOrdPackageNum, pOffPayDescription, pOffPayInfo, pAdminComments, pSubtotal, pnotax, intPackingSlip, ptaxcode, pcSCID

			intPackingSlip=request.Form("packingSlip")
			pshippingEmail=rs("pcOrd_ShippingEmail")
			pshippingFax=rs("pcOrd_ShippingFax")
			pcShowShipAddr=rs("pcOrd_ShowShipAddr")
			pidcustomer=rs("idcustomer")
			porderdate=rs("orderdate")
			porderdate=ShowDateFrmt(porderdate)
			pAddress=rs("Address")
			pcity=rs("city")
			pstate=rs("state")
			pstateCode=rs("stateCode")
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
			if not pcv_CatDiscounts<>"" then
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
			if not gIDEvent<>"" then
				gIDEvent="0"
			end if
			pGWTotal=rs("pcOrd_GWTotal")
			if not pGWTotal<>"" then
				pGWTotal="0"
			end if
			'GGG Add-on end
			pAdminComments=rs("adminComments")
			pcOrd_ShipWeight=rs("pcOrd_ShipWeight")
	
			'// Check if the Customer is European Union 
			Dim pcv_IsEUMemberState
			pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)

			query="SELECT [name],lastname,customerCompany,phone,email,customertype, pcCust_VATID, pcCust_SSN,fax FROM customers WHERE idCustomer=" & pidcustomer
			Set rsCustObj=Server.CreateObject("ADODB.Recordset")
			Set rsCustObj=connTemp.execute(query)
			CustomerName=rsCustObj("name")& " " & rsCustObj("lastname")
			CustomerPhone=rsCustObj("phone")
			CustomerEmail=rsCustObj("email")
			CustomerFax=rsCustObj("fax")
			CustomerCompany=rsCustObj("customerCompany")
			CustomerType=rsCustObj("customertype")
			CustomerVATID=rsCustObj("pcCust_VATID")
			CustomerSSN=rsCustObj("pcCust_SSN")
			set rsCustObj=nothing
						
			While Not rs.EOF %>
				<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
					<tr> 
					<td width="18%" height="71" valign="top"><img src="../pc/catalog/<%=scCompanyLogo%>"></td>
					<td width="39%" height="71" class="invoiceNob"><div align="center">
					<b><%=scCompanyName%></b><br>
					<%=scCompanyAddress%><br>
					<%=scCompanyCity %>, <%=scCompanyState %>&nbsp;<%=scCompanyZip %><br>
					<% if scStoreURL<>"" then 
						response.write scStoreURL&"<br>"
					end if
					if scCompanyPhoneNumber<>"" then
						response.write "Phone: "&scCompanyPhoneNumber&"<br>"
					end if
					if scCompanyFaxNumber<>"" then
						response.write "Fax: "&scCompanyFaxNumber&"<br>"
					end if %>
					</td>
					<td width="43%" height="71" valign="bottom"> 
						<table width="50%" align="right" cellpadding="5" cellspacing="0" class="invoice">
							<tr> 
								<td class="invoice" nowrap>ORDER DATE:
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
					<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
					<tr> 
					<td width="50%" valign="top" align="left">
						<table width="95%" cellpadding="5" cellspacing="0" class="invoice">
							<tr> 
								<td class="invoice">
									<strong>BILL TO</strong>:<br>
									<%=CustomerName%><br>
									<% if CustomerCompany<>"" then
										response.write CustomerCompany & "<BR>"
									end if %>
									<%=pAddress%><br>
									<% if pAddress2<>"" then
										response.write pAddress2&"<BR>"
									end if %>  
									<% response.write pcity & ", " & pStateCode &" "& pzip
									if pCountryCode <> scShipFromPostalCountry then
										response.write "<BR>" & pCountryCode
									end if %>
									<%if CustomerPhone<>"" then%>
									<br>Tel: <%=CustomerPhone%>
									<%end if%>
									<%if CustomerEmail<>"" then%>
									<br>E-mail: <%=CustomerEmail%>
									<%end if%>
									<%if CustomerFax<>"" then%>
										<br>Fax: <%=CustomerFax%>
									<%end if%>
									<% 
									'// Vat ID
									If CustomerVATID<>"" AND pcv_strShowVatId=-1 Then
										response.write "<br />" & dictLanguage.Item(Session("language")&"_Custmoda_26") & CustomerVATID
									End IF
								
									'// SSN
									If CustomerSSN<>"" AND pcv_strShowSSN=-1 Then
										response.write "<br />" & dictLanguage.Item(Session("language")&"_Custmoda_24") & CustomerSSN
									End IF 
									%>                                   
								</td>
							</tr>
						</table>
						<br>
					</td>
					<td rowspan="2" width="50%" valign="top"> 
					<table align="right" width="95%" cellpadding="5" cellspacing="0" class="invoice">
						<% dim strInvoiceNum, strAlterInvoiceNum
						intPackingSlip=request.Form("packingSlip")
						strAlterInvoiceNum=request.Form("AlterInvoiceNum")
						if strAlterInvoiceNum="" then
							strInvoiceNum=(scpre+int(qry_ID))
						else
							strInvoiceNum=CreateInvoiceNum(strAlterInvoiceNum,k)
						end if %>
						<tr> 
							<td class="invoice"><b>INVOICE #: <%=strInvoiceNum%></b></td>
						</tr>
						<% ' Calculate customer number using sccustpre constant
						Dim pcCustomerNumber
						pcCustomerNumber = (sccustpre + int(pidcustomer))
						%>
						<tr> 
							<td class="invoice">CUSTOMER ID: <%=pcCustomerNumber%></td>
						</tr>
						<% if scOrderName="1" then
							If trim(pord_OrderName) <> "" Then%>
								<tr><td valign="top" align="left" class="invoice">ORDER NAME: <%= pord_OrderName %></td></tr>
							<% 	End If
						end if %>

						<% 
						If trim(pord_DeliveryDate) <> "1/1/1900" and trim(pord_DeliveryDate) <> "" Then
							if scDateFrmt="DD/MM/YY" then
								pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
							else
								pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
							end if
							pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
							%>
							<tr>
								<td class="invoice">
									DELIVERY DATE/TIME: <%=pord_DeliveryDate & ", " & pord_DeliveryTime%>
								</td>
							</tr>
						<% End If %>
						<%
						'GGG Add-on start
						if gIDEvent<>"0" then	
							query="select pcEvents.pcEv_name, pcEvents.pcEv_Date, customers.name, customers.lastname from pcEvents, Customers where Customers.idcustomer=pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
							set rs1=connTemp.execute(query)
										
							geName=rs1("pcEv_name")
							geDate=rs1("pcEv_Date")
							if year(geDate)="1900" then
							geDate=""
							end if
							if gedate<>"" then
								if pDateFrmt="DD/MM/YY" then
								gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
								else
								gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
								end if
							end if
							gReg=rs1("name") & " " & rs1("lastname")
							%>
							<tr><td class="invoice">EVENT NAME: <%=geName %></td></tr>
							<tr><td class="invoice">EVENT DATE: <%=geDate %></td></tr>
							<tr><td class="invoice">REGISTRANT'S NAME: <%=gReg %></td></tr>
						<% 	End If
						'GGG Add-on end%>
						<tr>                   
							<td class="invoice">SHIPPED VIA:
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
									if NOT isNumeric(serviceHandlingFee) then
										serviceHandlingFee=0
									end if
									if NOT isNumeric(Postage) then
										Postage=0
									end if
								end if
								%>
								</td>
							</tr>            
						<%
							if pOrdShipType=0 then
								pDisShipType="Residential"
							else
								pDisShipType="Commercial" 
							end if
							if varShip<>"0" then
						%>
								<tr> 
									<td class="invoice">SHIPPING TYPE: <%=pDisShipType%></td>
								</tr>
						<%
							end if
							'Clear variable no it does not affect next invoice
							varShip=""
							
							payment = split(ppaymentDetails,"||")
							PaymentType=trim(payment(0))
											
							'Get payment nickname
							query="SELECT paymentDesc,paymentNickName FROM paytypes WHERE paymentDesc = '" & replace(PaymentType,"'","''") & "';"
							Set rsTemp=Server.CreateObject("ADODB.Recordset")
							Set rsTemp=connTemp.execute(query)
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
								end select
							end if
							
							'End get authorization and transaction IDs

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
								if instr(PaymentType,"FREE") AND len(PaymentType)<6 then
								else %>
									<tr> 
										<td class="invoice">&nbsp;</td>
									</tr>
									<tr> 
										<td class="invoice">PAYMENT METHOD: 
													<%
														if PaymentName <> "" and PaymentName <> PaymentType then
															Response.Write PaymentName & " (" & PaymentType & ")"
															else
															Response.Write PaymentType
														end if
													%>
										<% if PayCharge>0 then %>
											<br>ADDITIONAL FEE: 
											<%response.write scCurSign&money(PayCharge)%>
										<% end if %>
										<% if varTransID<>"" then %>
										<br><%=varTransName%>: <%=varTransID%>
										<% end if %>
										<% if varAuthCode<>"" then %>
										<br><%=varAuthName%>: <%=varAuthCode%>
										<% end if %>
								</td>
							</tr>
						<% end if %>
						
						<% 
						' Show offline payment custom field
						' Look for an offline payment record to get custom field and name of payment
						query = "SELECT offlinepayments.AccNum, payTypes.Cprompt FROM offlinepayments LEFT JOIN payTypes ON offlinepayments.idPayment = payTypes.idPayment WHERE offlinepayments.idOrder = " & qry_ID & ";"
						Set rsOP = Server.CreateObject("ADODB.Recordset")
						set rsOP = connTemp.execute(query)
						
						if not rsOP.EOF then
							pOffPayDescription = rsOP("Cprompt")
							pOffPayInfo = rsOP("AccNum")

							'If customer entered a value for the custom field, print it and
							'its associated description.
							if len(pOffPayInfo) > 0 then %>
								<tr>
									<td class="invoice">
										<% if len(pOffPayDescription) > 0 then
											response.write (pOffPayDescription & ": ")
										else
											response.write ("# ")
										end if
										response.write pOffPayInfo %>
														
									</td>
								</tr>
							<% end if
						end if
						
						set rsOP = nothing

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' Print Payment Information
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
						'// Retrieve information on whether card information should be shown
						'// IF yes, show entire card number. IF no, show only last 4 digits
						Dim pcIntShowLast4
						IF trim(uCase(request.Form("showCCInfo")))="NO" THEN
							pcIntShowLast4=1
						ELSE
							pcIntShowLast4=0
						END IF
						
		
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' START Custom Cards 
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		

							myquery="SELECT customCardOrders.idCCOrder, customCardOrders.idOrder, customCardOrders.strFormValue, customCardOrders.strRuleName, customCardOrders.idCustomCardRules FROM customCardOrders WHERE ((customCardOrders.idOrder)=" & qry_ID & ") ORDER BY customCardOrders.idCCOrder;"
							Set rsCC=Server.CreateObject("ADODB.Recordset")
							set rsCC=connTemp.execute(myquery)
							custcardtype=0

							if NOT rsCC.eof then
								intShowBtn=1
								custcardtype=1
								ccCnt=0
								do until rsCC.eof
									pIdCCOrder=rsCC("idCCOrder")
									pStrFormValue=rsCC("strFormValue")
									pStrRuleName=rsCC("strRuleName")
									pTempIdCCRules=rsCC("idCustomCardRules")
									'check length of field
									myquery="SELECT intlengthOfField, intmaxInput FROM customCardRules WHERE idcustomCardRules="&pTempIDCCRules&";"
									Set rsRulObj=Server.CreateObject("ADODB.Recordset")
									Set rsRulObj=connTemp.execute(myquery)
									
									if pcIntShowLast4=1 then
										pStrFormValue= ShowLastFour(pStrFormValue)
									end if
										
									if rsRulObj.eof then
										pLOF="20"
										pMaxInput="999"
									else
										pLOF=rsRulObj("intlengthOfField")
										pMaxInput=rsRulObj("intmaxInput")
									end if
									set rsRulObj=nothing
									'pIdCCR=rsCC("idcustomCardRules")
									if pMaxInput="" or pMaxInput="0" then
										pMaxInput= pLOF
									end if
									ccCnt=ccCnt+1%>
									<tr> 
										<td class="invoice"><%=UCASE(pStrRuleName)&": "&pStrFormValue%></td>
									</tr>
									<% rsCC.moveNext
								loop %>
							<% end if
							rsCC.Close
							set rsCC=nothing
							'End custom payment details

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' END Custom Cards
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' START Offline Credit Cards
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
							query="SELECT cardType,cardNumber,expiration,pcSecurityKeyID FROM creditCards WHERE idOrder=" & qry_ID & ";"
							Set rsCC=Server.CreateObject("ADODB.Recordset")
							set rsCC=connTemp.execute(query)
							if NOT rsCC.eof then
								pcardType=rsCC("cardType")
								pcardNumber=rsCC("cardNumber")
								pexpiration=rsCC("expiration")
								pcv_SecurityKeyID = rsCC("pcSecurityKeyID")
								tmp_pexpiration=pexpiration
								pexpiration=Month(tmp_pexpiration) & "/" & Year(tmp_pexpiration)
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
								end if
								
								pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
								
								Dim VarCCNum
								VarCCNum=pcardNumber
								
								if pcIntShowLast4=1 then
									VarCCNum2=ShowLastFour(enDeCrypt(VarCCNum, pcv_SecurityPass))
								else
									VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)				
								end if

								%>
								<tr>
									<td class="invoice">CARD TYPE :&nbsp;<%=CCType%></td>
								</tr>
								<tr>
									<td class="invoice">CARD NUMBER :&nbsp;<%=VarCCNum2%></td>
								</tr>
								<tr>
									<td class="invoice">EXPIRATION DATE:&nbsp;<%=pexpiration%></td>
								</tr>
							<% end if

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' END Offline Credit Cards
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' START Authorize.net
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  							query="SELECT ccnum,ccexp,pcSecurityKeyID FROM authorders WHERE idOrder=" & qry_ID & ";"
							Set rsCC=Server.CreateObject("ADODB.Recordset")
							set rsCC=connTemp.execute(query)
							if NOT rsCC.eof then
								pcardNumber=rsCC("ccnum")
								pexpiration=rsCC("ccexp")
								pcv_SecurityKeyID = rsCC("pcSecurityKeyID")
								tmp_pexpiration=trim(pexpiration)
								if Len(tmp_pexpiration)=4 then
									pexpiration=Left(tmp_pexpiration,2) & "/" & Right(tmp_pexpiration,2)
									else
									pexpiration=tmp_pexpiration
								end if
								if pcardNumber="*" then %>
									<tr>
										<td class="invoice">The credit card  number has been purged from database.</td>
									</tr>
								<% else
									pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

									VarCCNum=pcardNumber
										VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)				
									VarCCType=ShowCardType(VarCCNum2)
									if pcIntShowLast4=1 then
										VarCCNum2=ShowLastFour(VarCCNum2)
									end if
									%>
									<tr>
                                        <td class="invoice">CARD TYPE :&nbsp;<%=VarCCType%></td>
                                    </tr>
									<tr>
										<td class="invoice">CARD NUMBER:&nbsp;<%=VarCCNum2%></td>
									</tr>
									<tr>
										<td class="invoice">EXPIRATION DATE:&nbsp;<%=pexpiration%></td>
									</tr>
								<% end if
							end if
			
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' END Authorize.net
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' START EIG
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  							query="SELECT ccnum, ccexp, cctype, pcSecurityKeyID FROM pcPay_EIG_Authorize WHERE idOrder=" & qry_ID & ";"
							Set rsCC=Server.CreateObject("ADODB.Recordset")
							set rsCC=connTemp.execute(query)
							if NOT rsCC.eof then
								pcardNumber=rsCC("ccnum")
								pexpiration=rsCC("ccexp")
								VarCCType=rsCC("cctype")
								pcv_SecurityKeyID = rsCC("pcSecurityKeyID")
								tmp_pexpiration=trim(pexpiration)
								if Len(tmp_pexpiration)=4 then
									pexpiration=Left(tmp_pexpiration,2) & "/" & Right(tmp_pexpiration,2)
									else
									pexpiration=tmp_pexpiration
								end if
								if pcardNumber="*" then %>
									<tr>
										<td class="invoice">The credit card  number has been purged from database.</td>
									</tr>
								<% else								
									pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
									VarCCNum=pcardNumber
									VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)
									if pcIntShowLast4=1 then
										VarCCNum2=ShowLastFour(VarCCNum2)
									end if
									Select Case VarCCType
										Case "V": VarCCType="Visa"
										Case "M": VarCCType="MasterCard"
										Case "A": VarCCType="American Express"
										Case "D": VarCCType="Discover"
										Case Else : VarCCType=VarCCType
									End Select
									%>
									<tr>
                                        <td class="invoice">CARD TYPE :&nbsp;<%=VarCCType%></td>
                                    </tr> 
									<tr>
										<td class="invoice">CARD NUMBER:&nbsp;<%=VarCCNum2%></td>
									</tr>
									<tr>
										<td class="invoice">EXPIRATION DATE:&nbsp;<%=pexpiration%></td>
									</tr>
                                    
								<% end if
							end if
			
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' END EIG
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
  
  		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' START PayFlow Pro 
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

							query="SELECT acct,expdate,pcSecurityKeyID FROM pfporders WHERE idOrder=" & qry_ID & ";"
							Set rsCC=Server.CreateObject("ADODB.Recordset")
							set rsCC=connTemp.execute(query)
							if NOT rsCC.eof then
								pcardNumber=rsCC("acct")
								pexpiration=rsCC("expdate")
								pcv_SecurityKeyID = rsCC("pcSecurityKeyID")
								tmp_pexpiration=trim(pexpiration)
								if Len(tmp_pexpiration)=4 then
									pexpiration=Left(tmp_pexpiration,2) & "/" & Right(tmp_pexpiration,2)
									else
									pexpiration=tmp_pexpiration
								end if
								if pcardNumber="*" then %>
									<tr>
										<td class="invoice">The credit card number has been purged from database.</td>
									</tr>
								<% else
									pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

									VarCCNum=pcardNumber
										VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)				
									VarCCType=ShowCardType(VarCCNum2)
									if pcIntShowLast4=1 then
										VarCCNum2=ShowLastFour(VarCCNum2)
									end if
									%>
									<tr>
                                        <td class="invoice">CARD TYPE :&nbsp;<%=VarCCType%></td>
                                    </tr>
									<tr>
										<td class="invoice">CARD NUMBER:&nbsp;<%=VarCCNum2%></td>
									</tr>
									<tr>
										<td class="invoice">EXPIRATION DATE:&nbsp;<%=pexpiration%></td>
									</tr>
								<% end if
							end if
			
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' END PayFlow Pro 
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' START NetBilling
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

							query="SELECT ccnum,ccexp,pcSecurityKeyID FROM netbillorders WHERE idOrder=" & qry_ID & ";"
							Set rsCC=Server.CreateObject("ADODB.Recordset")
							set rsCC=connTemp.execute(query)
							if NOT rsCC.eof then
								pcardNumber=rsCC("ccnum")
								pexpiration=rsCC("ccexp")
								pcv_SecurityKeyID = rsCC("pcSecurityKeyID")
								tmp_pexpiration=trim(pexpiration)
								if Len(tmp_pexpiration)=4 then
									pexpiration=Left(tmp_pexpiration,2) & "/" & Right(tmp_pexpiration,2)
									else
									pexpiration=tmp_pexpiration
								end if
								if pcardNumber="*" then %>
									<tr>
										<td class="invoice">The credit card number has been purged from database.</td>
									</tr>
								<% else

									pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

									VarCCNum=pcardNumber
										VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)				
									VarCCType=ShowCardType(VarCCNum2)
									if pcIntShowLast4=1 then
										VarCCNum2=ShowLastFour(VarCCNum2)
									end if									
									%>
									<tr>
                                        <td class="invoice">CARD TYPE :&nbsp;<%=VarCCType%></td>
                                    </tr>
									<tr>
										<td class="invoice">CARD NUMBER:&nbsp;<%=VarCCNum2%></td>
									</tr>
									<tr>
										<td class="invoice">EXPIRATION DATE:&nbsp;<%=pexpiration%></td>
									</tr>
								<% end if
							end if
							
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' END NetBilling
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' START USAePay
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				query="SELECT ccCard,ccExp,pcSecurityKeyID FROM pcPay_USAePay_Orders WHERE idOrder=" & qry_ID & ";"
				Set rsCC=Server.CreateObject("ADODB.Recordset")
				set rsCC=connTemp.execute(query)
				if NOT rsCC.eof then
					pcardNumber=rsCC("ccCard")
					pexpiration=rsCC("ccExp")
					pcv_SecurityKeyID = rsCC("pcSecurityKeyID")
					tmp_pexpiration=trim(pexpiration)
					if Len(tmp_pexpiration)=4 then
						pexpiration=Left(tmp_pexpiration,2) & "/" & Right(tmp_pexpiration,2)
						else
						pexpiration=tmp_pexpiration
					end if
					if pcardNumber="*" then %>
						<tr>
							<td class="invoice">The credit card number has been purged from database.</td>
						</tr>
					<% else
					
						pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

						VarCCNum=pcardNumber
							VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)				
						VarCCType=ShowCardType(VarCCNum2)
						if pcIntShowLast4=1 then
							VarCCNum2=ShowLastFour(VarCCNum2)
						end if 						
					%>
                        <tr>
                            <td class="invoice">CARD TYPE :&nbsp;<%=VarCCType%></td>
                        </tr>
						<tr>
							<td class="invoice">CARD NUMBER:&nbsp;<%=VarCCNum2%></td>
						</tr>
						<tr>
							<td class="invoice">EXPIRATION DATE:&nbsp;<%=pexpiration%></td>
						</tr>
						<% end if
					end if
			
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' END USAePay
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					
						If intPackingSlip<>1 Then 'Hide this information if this is a packing slip
							If RewardsActive <> 0 And piRewardPoints > 0 Then 
								iDollarValue = piRewardPoints * (RewardsPercent / 100) %>          
								<tr>
									<td class="invoice"><%=ucase(RewardsLabel)%>:<br>The customer used <%=piRewardPoints & " " & RewardsLabel%> on this purchase. for a dollar value of <%=scCurSign&money(iDollarValue)%>.
									</td>
								</tr>
							<% end if %>
							<% If RewardsActive <> 0 And piRewardPointsCustAccrued > 0 Then %>
								<tr> 
									<td class="invoice"><%=ucase(RewardsLabel)%>:<br>
								The customer accrued <%=piRewardPointsCustAccrued & " " & RewardsLabel%> on this purchase.</td>
								</tr>
							<% end if
						end if 'End hide this information if this is a packing slip %>

						<% 'if discount was present, show type here
						if instr(pdiscountDetails,",") then
							DiscountDetailsArry=split(pdiscountDetails,",")
							intArryCnt=ubound(DiscountDetailsArry)
						else
							intArryCnt=0
						end if
					
						dim discounts, discountType 

						for j=0 to intArryCnt
							if intArryCnt=0 then
								pTempDiscountDetails=pdiscountDetails
							else
								pTempDiscountDetails=DiscountDetailsArry(j)
							end if
							if instr(pTempDiscountDetails,"- ||") then
								discounts = split(pTempDiscountDetails,"- ||")
								discountType = discounts(0)
								discount = discounts(1) 
								if discountType<>"" then %>
									<tr>
										<td class="invoice">DISCOUNT/PROMOTION: <%=discountType%></td>
									</tr>
								<% end if
							end if
						Next %>
		
						<%'start of gift certificates
						if GCDetails<>"" then
							GCArry=split(GCDetails,"|g|")
							intArryCnt=ubound(GCArry)
								
							for m=0 to intArryCnt
					
								if GCArry(m)<>"" then
									GCInfo = split(GCArry(m),"|s|")
									if GCInfo(2)="" OR IsNull(GCInfo(2)) then
										GCInfo(2)=0
									end if
									%>
									<tr>
										<td class="invoice">GIFT CERTIFICATE: <%=GCInfo(1) & " (" & GCInfo(0) & ")"%></td>
									</tr>
								<% end if
							Next
						end if
						'end if gift certificates									
						%>

						<% 'if category-based quantity discounts were applied, show them here
							If intPackingSlip<>1 then 'Hide this information if this is a packing slip
								if pcv_CatDiscounts <> "0" then %>
									<tr>
										<td class="invoice">CATEGORY DISCOUNTS: <%=scCurSign&money(pcv_CatDiscounts)%></td>
									</tr>
							<% end if
						end if %>
					</table>
				</td>
			</tr>

			<tr> 
				<td width="50%" valign="top"> 
					<table width="95%" cellpadding="5" cellspacing="0" class="invoice">      
					<tr>              
					<td class="invoice">
					<strong>SHIP TO</strong>:<br>
								<% if pcShowShipAddr="0" then %>
									<% response.write "(Same as billing address)" %>
								<% else %>
									<% if pshippingFullName<>"" then
										response.write pshippingFullName
									else
										response.write CustomerName
									end if %>
									<br>
									<% if pshippingCompany<>"" then 
										response.write pshippingCompany & "<br>"
									else
										if pshippingFullName="" and customerCompany <> "" then
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
									<% if pshippingPhone <> "" then %>
										<br>Tel: <%=pshippingPhone%>
									<% end if %>
								        <% if pshippingEmail <> "" then %>
								            <br>E-mail: <%=pshippingEmail%>
								        <% end if %>
								        <% if pshippingFax <> "" then %>
								            <br>Fax: <%=pshippingFax%>
								        <% end if %>
								<% end if %>											
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<br>
			<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
				<tr> 
					<td>
						<table width="100%" cellpadding="5" cellspacing="0" border="1" class="invoice">
								<tr> 
									<td width="8%" class="invoice"><b>QTY</b></td>
									<td class="invoice"><b>SKU - DESCRIPTION</b></td>
									<td width="16%" class="invoice"><div align="right"><b>
										<% if intPackingSlip<>1 then %>UNIT PRICE<% end if %></b></div></td>
									<td width="12%" class="invoice"><div align="right"><b>
											<% if intPackingSlip<>1 then %>TOTAL<% end if %></b></div></td>
										</tr>
                		<% query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.unitPrice, ProductsOrdered.xfdetails"
										'BTO ADDON-S
										if scBTO=1 then
											query=query&", ProductsOrdered.idconfigSession"
										end if
										'BTO ADDON-E
										query=query&",ProductsOrdered.rmaSubmitted, ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts, pcDropShipper_ID, pcPrdOrd_BackOrder, pcPrdOrd_SentNotice, ProductsOrdered.pcPO_GWOpt,pcPO_GWNote,pcPO_GWPrice, pcPrdOrd_BundledDisc, pcSC_ID FROM ProductsOrdered WHERE ProductsOrdered.idOrder=" & qry_ID & ";"
										Set rsTemp=Server.CreateObject("ADODB.Recordset")
										set rsTemp=connTemp.execute(query)
										dim intTotalWeight
										intTotalWeight=int(0)
										tmpAllPrdSubTotal=0
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

											if scBTO=1 then
												pidConfigSession=rstemp("idConfigSession")
											end if
											prmaSubmitted=rstemp("rmaSubmitted")
											QDiscounts=rstemp("QDiscounts")
											ItemsDiscounts=rstemp("ItemsDiscounts")

											pcv_IntDropShipperId=rstemp("pcDropShipper_ID")
											if IsNull(pcv_IntDropShipperId) or pcv_IntDropShipperId="" then
												pcv_IntDropShipperId=0
											end if
						
											pcv_IntDropNotified=rstemp("pcPrdOrd_SentNotice")
											if IsNull(pcv_IntDropNotified) or pcv_IntDropNotified="" then
												pcv_IntDropNotified=0
											end if
					
							
											pcv_IntBackOrder=rstemp("pcPrdOrd_BackOrder")
											if IsNull(pcv_IntBackOrder) or pcv_IntBackOrder="" then
												pcv_IntBackOrder=0
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
											pcPrdOrd_BundledDisc=rstemp("pcPrdOrd_BundledDisc")
											pcSCID=rstemp("pcSC_ID")
											if pcSCID="" Or (IsNull(pcSCID)) then
												pcSCID=0
											end if
											
											query="SELECT sku,description,weight,pcprod_QtyToPound FROM products WHERE idproduct="& pidProduct
											Set rsTemp2=Server.CreateObject("ADODB.Recordset")
											set rsTemp2=connTemp.execute(query)
											psku=rsTemp2("sku")
											pDescription=rsTemp2("description")
											pWeight=rsTemp2("weight")
											pcv_QtyToPound=rsTemp2("pcprod_QtyToPound")
											'// no extra loop involved to calculate weight, let the calculation happen here
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
					
														query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
														set rsConfigObj=conntemp.execute(query)
														if NOT isNumeric(ArrQuantity(i)) then
															pIntQty=1
														else
															pIntQty=ArrQuantity(i)
														end if
														if not rsConfigObj.eof then
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
												<td width="8%" class="invoice"><%=pquantity%></td>
												<td class="invoice">
													<%=psku%> - <%=pDescription%>
						                            <%
													'// Show sale icon, if applicable
													pcShowSaleIcon
													%>
						                        </td>
												<td width="16%" class="invoice">
												<% If intPackingSlip<>1	Then 'Hide this information if this is a packing slip %>
													<div align="right"><%=scCurSign&money(punitPrice1)%></div>
												<% End if %>
												</td>
												<td width="12%" class="invoice">
												<% If intPackingSlip<>1	Then 'Hide this information if this is a packing slip %>
													<div align="right"><%=scCurSign&money(pRowPrice1)%></div>
												<% End if %>
												</td>
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
													If intPackingSlip <>1 Then 
													%>
													<tr> 
														<td class="invoice">&nbsp;</td>
														<td class="invoice" colspan="3">
													<table width="100%" cellspacing="2" cellpadding="0" bgcolor="#FFFFCC" class="invoiceBto">
															<tr> 
																<td colspan="3" class="invoiceNob"><u>Customizations:</u></td>
															</tr>
																<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								
								query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pidProduct 
								set rsQ=server.CreateObject("ADODB.RecordSet") 
								set rsQ=conntemp.execute(query)

								btDisplayQF=rsQ("displayQF")
								set rsQ=nothing
											
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
																<td width="20%" valign="top" class="invoiceNob"><%=pcategoryDesc%>:</td>
									<td width="70%" valign="top" class="invoiceNob"><%=psku%> - <%=pdescription%>
									<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%> - QTY: <%=ArrQuantity(i)%><%end if%>
									</td>
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
									<td width="10%" valign="top" nowrap class="invoiceNob">
										<div align="right">
											<%if (intPackingSlip<>1) then%>
												<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then%>
													<%=scCurSign & money((ArrValue(i)+UPrice)*pQuantity)%>
												<%else
													if tmpDefault=1 then%>Included<%end if%>
												<%end if%>
											<%end if%>
										</div>
									</td>
															</tr>
															<% 'no extra loop for weight calculation - let it happen here
															intItemWeight=int(pItemWeight)*pIntQTY*pquantity
																 intTotalWeight=intTotalWeight+intItemWeight
																 set rsConfigObj=nothing
																	next
																	set rsConfigObj=nothing %>
														</table>

														</td>
													</tr>
													<% End If '// If intPackingSlip <>1 Then %>
												<% end if %>
											<% end if
											'BTO ADDON-E %>
									

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
			
					<tr> 
						<td width="8%" class="invoice">&nbsp;</td>
						<td class="invoice" style="padding-left:10px;"><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></td>
						<% 
						tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
						if tempPrice="" or tempPrice=0 then
							%><td width="16%" class="invoice"></td><td width="12%" class="invoice"></td><%
						else
							'// Adjust for Quantity Discounts
							tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice)
							%>
							<td width="16%" class="invoice"><% if intPackingSlip <>1 then %><div align="right"><%=scCurSign&money(tempPrice)%></div><% end if %></td>	
							<td width="12%" class="invoice">
								<% if intPackingSlip <>1 then %>
									<div align="right">			 
									<%
									tAprice=(tempPrice*Cdbl(pquantity))
									response.write scCurSign&money(tAprice) 
									%>
									</div>
								<% end if %>
								</td>
								<%
							end if %>
					</tr>	
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
							
											<% if len(pxdetails)>3 then %>
												<tr> 
													<td width="8%" class="invoice">&nbsp;</td>
													<td class="invoice" style="padding-left:10px;"><%=replace(pxdetails,"|","<br>")%></td>
													<td width="16%" class="invoice">&nbsp;</td>
													<td width="12%" class="invoice">&nbsp;</td>
												</tr>
											<% end if %>
                <!-- if RMA -->
				<% if NOT isNull(prmaSubmitted) AND prmaSubmitted<>"" AND prmaSubmitted>0 then %>
				<tr> 
                    <td width="8%" class="invoice"><%=prmaSubmitted%></td>
                    <td class="invoice" style="padding-left:10px;">RETURNED</td>
					<td width="16%" class="invoice">&nbsp;</td>
					<td width="12%" class="invoice">&nbsp;</td>
				</tr>
                <% end if	%>
                <!-- end of RMA -->                
											<%'BTO ADDON-S
											pRowPrice=(punitPrice)*(pquantity)
											pExtRowPrice=pRowPrice
		Charges=0
											If scBTO=1 then
												pidConfigSession=trim(pidConfigSession)
												if pidConfigSession<>"0" then
													ItemsDiscounts=trim(ItemsDiscounts)
													if (ItemsDiscounts<>"") and (CDbl(ItemsDiscounts)<>"0") then
					If intPackingSlip=1	Then 'Hide this information if this is a packing slip %>
															<tr><td colspan="4" class="invoice">&nbsp;</td></tr>
														<% else %>
														<tr> 
														<td width="8%" class="invoice">&nbsp;</td>
														<td class="invoice">&nbsp;</td>
														<td width="16%" class="invoice">ITEM DISCOUNTS:</td>
														<td width="22%" class="invoice"><div align="right"><font color="#FF0000"><%=scCurSign&money(-1*ItemsDiscounts)%></font></div></td>
														</tr>
					<% end if
														pRowPrice=pRowPrice-Cdbl(ItemsDiscounts)
	end if
				pIdConfigSession=trim(pidConfigSession)
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
	' Hide if packing slip
						If intPackingSlip <> 1 Then %>
	<tr bgcolor="#FFFFFF" class="small"> 
	<td width="5%" bgcolor="#FFFFFF" class="invoice">&nbsp;</td>
	<td class="invoice" colspan="3">

	<table width="100%" cellspacing="0" cellpadding="2" bgcolor="#FFFFCC" class="invoiceBto">
	<tr class="small"> 
	<td colspan="3" class="invoiceNob"><u>Additional Charges</u></td>
	</tr>
										<% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
																				query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
																				set rsConfigObj=connTemp.execute(query)
																				pcategoryDesc=rsConfigObj("categoryDesc")
																				pdescription=rsConfigObj("description")
																				psku=rsConfigObj("sku")
																				pItemWeight=rsConfigObj("weight")
																				intTotalWeight=intTotalWeight+int(pItemWeight)
																				if (CDbl(ArrCValue(i))>0)then
																					Charges=Charges+cdbl(ArrCValue(i))
                                            end if %>
																				<tr> 
																				<td width="20%" class="invoiceNob" valign="top"><%=pcategoryDesc%>:</td>
																				<td width="70%" class="invoiceNob" valign="top"><%=psku%> - <%=pdescription%></td>
																				<td width="10%" class="invoiceNob" nowrap valign="top"><%if (ArrCValue(i)>0) and (intPackingSlip<>1) then%><div align="right"><%=scCurSign & money(ArrCValue(i))%></div><%end if%></td>
																				</tr>
																				<% set rsConfigObj=nothing
																				next
																				set rsConfigObj=nothing 
																				pRowPrice=pRowPrice+Cdbl(Charges)%>
																				</table>

																</td>
															</tr>
															<% End If '//If intPackingSlip <> 1 Then %>		
															<% end if
														end if %>
            <% 'BTO Additional Charges

												end if 
											end if 'BTO
											
													QDiscounts=trim(QDiscounts)
											if (QDiscounts<>"") and (CDbl(QDiscounts)<>"0") and intPackingSlip<>1 then %>
                                            <tr bgcolor="#FFFFFF" class="small"> 
                                            <td width="5%" bgcolor="#FFFFFF" class="invoice">&nbsp;</td>
                                            <td class="invoice" colspan="3">

                                            <table width="100%" cellspacing="0" cellpadding="2" bgcolor="#FFFFCC" class="invoiceBto">
													<tr> 
                                            <td width="90%" colspan="2" class="invoiceNob" valign="top">Quantity Discounts:</td>
                                            <td width="10%" class="invoiceNob" nowrap valign="top"><div align="right"><font color="#FF0000"><%=scCurSign&money(-1*QDiscounts)%></font></div></td>
                                            </tr>
                                            </table>
                                            </td>
													</tr>
                                            
                                      		<% pRowPrice=pRowPrice-Cdbl(QDiscounts)
	end if
	
	if pExtRowPrice<>pRowPrice then
		if intPackingSlip<>1 then%>
													<tr> 
													<td width="8%" class="invoice">&nbsp;</td>
													<td class="invoice">&nbsp;</td>
													<td width="16%" class="invoice"><div align="right">PRODUCT SUBTOTAL:</div></td>
													<td width="22%" class="invoice"><div align="right"><%=scCurSign&money(pRowPrice)%></div></td>
													</tr>
											<% end if
											end if 
											
											
											'GGG Add-on start
											if pGWOpt<>"0" then
											query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
											set rsG=connTemp.execute(query)
											if not rsG.eof then%>
													<tr> 
														<td class="invoice" width="8%">&nbsp;</td>
														<td class="invoice" colspan="3">
															Gift Wrapping: <%=rsG("pcGW_OptName")%><%if (pGWPrice>0) and (intPackingSlip<>1) then%> - Price:&nbsp;<%=scCurSign & money(pGWPrice)%><%end if%>
															<%if pGWText<>"" then%>
																<br>
																Gift Notes:<br><%=pGWText%>
															<%end if%>
														</td>
													</tr>
											<%
											end if
											end if
	
	tmpAllPrdSubTotal=tmpAllPrdSubTotal+CDbl(pRowPrice)
	
	'GGG Add-on end
	pcPrdOrd_BundledDisc=trim(pcPrdOrd_BundledDisc)
	if (pcPrdOrd_BundledDisc<>"") and (CDbl(pcPrdOrd_BundledDisc)<>"0") then
	tmpAllPrdSubTotal=tmpAllPrdSubTotal-CDbl(pcPrdOrd_BundledDisc) %>
        <tr> 
            <td width="8%" class="invoice">&nbsp;</td>
            <td class="invoice">Bundle Discount</td>
            <td width="16%" class="invoice">&nbsp;</td>
            <td width="12%" class="invoice"><div align="right"><%=scCurSign&money(-1*pcPrdOrd_BundledDisc)%></div></td>
        </tr>
	<% end if	
	rstemp.moveNext
	loop
	set rstemp=nothing %>
	
	<%if Cdbl(tmpAllPrdSubTotal) <> 0 and intPackingSlip <> 1 then %>
		<tr> 
			<td class="invoice" align="right" colspan="3"><b>ALL PRODUCTS SUBTOTAL</b></td>
			<td width="12%" class="invoice"><div align="right"><%=scCurSign&money(tmpAllPrdSubTotal)%></div></td>
		</tr>
	<%end if%>
                
				<% 'RP ADDON-S
				If RewardsActive <> 0 Then
				if piRewardValue <> 0 and intPackingSlip <> 1 then %>
				<tr> 
					<td width="8%" class="invoice">&nbsp;</td>
					<% if RewardsLabel="" then
								RewardsLabel="Rewards Program"
							end if %>
					<td class="invoice"><%=RewardsLabel%></td>
					<td width="16%" class="invoice">&nbsp;</td>
					<td width="12%" class="invoice"><div align="right">-<%=scCurSign&money(piRewardValue)%></div></td>
				</tr>
				<% end if
				End if
				'RP ADDON-E %>
				
				<%'GGG Add-on start
				if (pGWTotal>0) and (intPackingSlip<>1) then%>
						<tr>
							<td class="invoice" align="right" colspan="3">GIFT WRAPPING:</td>
							<td class="invoice" align="right" width="12%"><%=scCurSign&money(pGWTotal)%></td>
						</tr>
				<%
				end if
				'GGG Add-on end%>
				
				<% 	'Start IF statement for Packing Slip

						if intPackingSlip<>1 then %>
				 <tr>
					<td colspan="2" rowspan="8" class="invoice">&nbsp;</td> 
					<td width="16%" class="invoice"><div align="right">SHIPPING:</div></td>
					<td width="12%" class="invoice"><div align="right"><%=scCurSign&money(postage)%></div></td>
				</tr>

				 <% if serviceHandlingFee<>0 then %>
				 <tr> 
					<td width="16%" class="invoice"><div align="right">SHIPPING &amp;<BR>HANDLING FEES:</div></td>
					<td width="12%" class="invoice"><div align="right"><%=scCurSign&money(serviceHandlingFee)%></div></td>
				</tr>
				<% end if %>
					
					<% if pcOrd_ShipWeight>0 then
						intTotalWeight=pcOrd_ShipWeight
					end if
if intTotalWeight <> 0 then
		intTotalWeight=round(intTotalWeight,0)
%>
				<tr> 
					<td width="16%" class="invoice"><div align="right">ORDER WEIGHT:</div></td>
					<td width="12%" class="invoice">
					<% if scShipFromWeightUnit="KGS" then
					pKilos=Int(intTotalWeight/1000)
					pWeight_g=intTotalWeight-(pKilos*1000)
					Response.Write pKilos&" kg "
					if pWeight_g>0 then 
					response.write pWeight_g&" g"
					end if
					else 
					pPounds=Int(intTotalWeight/16)
					pWeight_oz=intTotalWeight-(pPounds*16)
					Response.Write pPounds&" lbs. "
					if pWeight_oz>0 then 
					response.write pWeight_oz&" oz."
					end if
					end if %>
					</td>
				</tr>
<% end if %>
<% if pOrdPackageNum <> 1 then %>
<tr> 
					<td width="16%" class="invoice"><div align="right">N. OF PACKAGES:</div></td>
					<td width="12%" class="invoice"><%=pOrdPackageNum%></td>
</tr>
<% end if %>

				<% if PayCharge>0 then %>
				<tr> 
					<td width="16%" class="invoice"><div align="right">PROCESSING FEES:</div></td>
					<td width="12%" class="invoice"><div align="right"><%=scCurSign&money(PayCharge)%></div></td>
				</tr>
				<% end if %>
					<% if NOT (pord_VAT>0) then 

						if isNull(ptaxDetails) OR trim(ptaxDetails)="" then %> 
							<tr> 
								<td width="16%" class="invoice"><div align="right">TAXES:</div></td>
								<td width="12%" class="invoice"><div align="right"><%=scCurSign&money(ptaxAmount)%></div></td>
							</tr>
						<% else %>
							<% taxArray=split(ptaxDetails,",")
							tempTaxAmount=0
							for i=0 to (ubound(taxArray)-1)
								taxDesc=split(taxArray(i),"|")
								if taxDesc(0)<>"" then
								'State Taxes|1.27875,Country Taxes|0.34875,%>
                                 <tr> 
                                    <td width="16%" class="invoice"><div align="right"><%=ucase(taxDesc(0))%></div></td>
                                    <% pDisTax=(money(taxDesc(1))) %>
                                    <td width="12%" class="invoice"><div align="right"><%=scCurSign&pDisTax%></div></td>
                                </tr>
                                <% end if
							next %>
					<% end if 
				end if %>
				<% if instr(pdiscountDetails,"- ||") or (pcv_CatDiscounts>"0") then
					if instr(pdiscountDetails,",") then
						DiscountDetailsArry=split(pdiscountDetails,",")
						intArryCnt=ubound(DiscountDetailsArry)
					else
						intArryCnt=0
					end if
					
					discount=0
					
					for m=0 to intArryCnt
						if intArryCnt=0 then
							pTempDiscountDetails=pdiscountDetails
						else
							pTempDiscountDetails=DiscountDetailsArry(m)
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
					<td width="16%" class="invoice"><div align="right">DISCOUNTS:</div></td>
					<td width="12%" class="invoice"><div align="right">-<%=scCurSign&money(discount+pcv_CatDiscounts)%></div></td>
				</tr>
				<% end if %>
					<%if GCAmount>"0" then%>
					<tr> 
						<td width="16%" class="invoice"><div align="right">GIFT CERTIFICATE AMOUNT:</div></td>
						<td width="12%" class="invoice"><div align="right">-<%=scCurSign&money(GCAmount)%></div></td>
					</tr>
					<%end if%>
					<tr> 
					<td width="16%" class="invoice"><div align="right"><b>TOTAL:</b></div></td>
					<td width="12%" class="invoice"><div align="right"><%=scCurSign&money(ptotal)%></div></td>
				</tr>
					<% if pord_VAT>0 then %>
						<% if pcv_IsEUMemberState=1 then %>
                            <tr> 
                                <td class="invoice" colspan="2"><div align="right">Includes <%=scCurSign&money(pord_VAT)%> of VAT</div></td>
                            </tr>
                        <% else %>
                            <tr> 
                                <td class="invoice" colspan="2"><div align="right"><%=scCurSign&money(pord_VAT)%> of VAT Removed</div></td>
                            </tr>
                        <% end if %>                        
					<% end if %>
				<% if NOT isNull(prmaCredit) AND prmaCredit<>"" AND prmaCredit>0 then %>
					<tr>
						<td width="16%" class="invoice"><div align="right">CREDIT ISSUED:</div></td>
						<td width="12%" class="invoice"><div align="right">-<%=scCurSign&money(pRmaCredit)%></div></td>
					</tr>
				<% end if %>
	<% end if 'end IF statement for Packing Slip %>

	</table>

</td>											
</tr>
</table>

<%'GGG Add-on start
	IF (GCDetails<>"") and (intPackingSlip<>1) then %>
	<br>
	<table width="100%" cellpadding="5" cellspacing="0" border="1" class="invoice">
	<tr>
		<td colspan="2" class="invoice"><b>The following Gift Certificate was used for this order:</b></td>
	</tr>
		<%GCArry=split(GCDetails,"|g|")
		intArryCnt=ubound(GCArry)
			
		for m=0 to intArryCnt
				
		if GCArry(m)<>"" then
			GCInfo = split(GCArry(m),"|s|")
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
	<td width="18%" nowrap  class="invoice"><b>Gift Certificate Product Name:</b></td>
	<td width="82%" class="invoice"><b><%=pName%></b></td>
</tr>
<tr> 
	<td width="18%" nowrap valign="top" class="invoice">&nbsp;</td>
	<td width="82%" valign="top" class="invoice">
	<%
	query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
	set rs19=connTemp.execute(query)
				
	if not rs19.eof then%>
	Gift Certificate Code: <b><%=rs19("pcGO_GcCode")%></b><br>
	Used for this order:&nbsp;<%=scCurSign & money(pGiftUsed)%><br><br>
	<%
	pGCAmount=rs19("pcGO_Amount")
	if cdbl(pGCAmount)<=0 then%>
	This Gift Certificate has been completely redeemed.
	<%else%>
	Available Amount: <b><%=scCurSign & money(pGCAmount)%></b>
	<br>
	<%pExpDate=rs19("pcGO_ExpDate")
	if year(pExpDate)="1900" then%>
	This Gift Certificate does not expire.
	<%else
	if pDateFrmt="DD/MM/YY" then
	pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
	else
	pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
	end if%>
	Expiration Date: <font color=#ff0000><b><%=pExpDate%></b></font>
	<%end if%>
	<br>
	<%
	pGCStatus=rs19("pcGO_Status")
	if pGCStatus="1" then%>
	Status: Active
	<%else%>
	Status: Inactive
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
IF (pGCs<>"") and (pGCs="1") and (intPackingSlip<>1) then %>
<br>
<table width="100%" cellpadding="5" cellspacing="0" border="1" class="invoice">
<tr>
	<td colspan="2" class="invoice"><b>GIFT CERTIFICATES</b></td>
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
	<td width="18%" nowrap  class="invoice"><b>Gift Certificate Product Name:</b></td>
	<td width="82%" class="invoice"><b><%=gName%></b></td>
</tr>
<tr> 
	<td width="18%" nowrap valign="top" class="invoice">&nbsp;</td>
	<td width="82%" valign="top" class="invoice">
	<%
	query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & qry_ID
	set rs19=connTemp.execute(query)
				
	do while not rs19.eof%>
	Gift Certificate Code: <b><%=rs19("pcGO_GcCode")%></b><br>
	<%pExpDate=rs19("pcGO_ExpDate")
	if year(pExpDate)="1900" then%>
	This Gift Certificate does not expire.
	<%else
	if pDateFrmt="DD/MM/YY" then
	pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
	else
	pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
	end if%>
	Expiration Date: <font color=#ff0000><b><%=pExpDate%></b></font>
	<%end if%>
	<br>
	<%
	pGCAmount=rs19("pcGO_Amount")
	if cdbl(pGCAmount)<=0 then%>
	This Gift Certificate has been completely redeemed.
	<%else%>
	Available Amount: <b><%=scCurSign & money(pGCAmount)%></b>
	<%end if%><br>
	<%
	pGCStatus=rs19("pcGO_Status")
	if pGCStatus="1" then%>
	Status: Active
	<%else%>
	Status: Inactive
	<%end if%>
	<br><br>
	<%	
	rs19.movenext
	loop
	set rs19=nothing
	%>
	</td>
</tr>
<%
end if
set rsG=nothing

rs11.MoveNext
loop
set rs11=nothing
%>
</table>
<% END IF
'GGG Add-on end%>
				
<% if pcomments<>"" then %>
		<br />
		<table width="100%" align="left" cellpadding="5" cellspacing="0" border="1" class="invoice">
			<tr>
			<td class="invoice">
			COMMENTS:
			<br><br>
			<%=pcomments%>
			<br>
			</td>
			</tr>
		</table>
		<br /><br />
<% end if %>

	<%
		pShowAdminComments = request("showAdminComments")
		if pShowAdminComments = "YES" and pAdminComments<>"" then
	%>
	<br /><br />
	<table width="100%" align="left" cellpadding="5" cellspacing="0" border="1" class="invoice">
	<tr> 
		<td class="invoice">
			ADMIN COMMENTS:
			<br><br>
			<%=pAdminComments%>
			<br>
		</td>
	</tr>
	</table>
	<%
		end if
	%>
	
<%rs.MoveNext
Wend
set rs=nothing
%>
<%end if%>
<%
	Next
	closedb()
%>

<% else
	TmpStr=""
	Count=request("count")
	if (Count="") or (Count="0") then
		response.redirect "menu.asp"
	end if

	For m=1 to Count
		if (request("check" & m)="1") and (request("idord" & m)<>"") then
			TmpStr=TmpStr & request("idord" & m) & "***"
		end if
	Next

	if TmpStr="" then
		response.redirect "menu.asp"
	end if
	%>
	<script language="JavaScript" type="text/javascript">
		function checkCR(evt) {
			var evt  = (evt) ? evt : ((event) ? event : null);
			var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
			if ((evt.keyCode == 13) && (node.type=="text")) {return false;}
		}
		document.onkeypress = checkCR;
	</script>
	<div id="pcCPmain" style="width: 400px; background-image: none;" align="center">
	<form action="BatchPrint.asp" method="post" name="invoice" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
				<tr>
					<th>Generate Invoice or Packing Slip?</th>
				</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
				<tr>
					<td>If you select packing slip, then prices will be hidden.</td>
				</tr>
			<tr>
				<td>
				<input type="radio" name="packingSlip" id="packingSlip" value="1" class="clearBorder">Packing Slip<br>
				<input type="radio" name="packingSlip" id="packingSlip" value="0" checked class="clearBorder">Invoice
				</td>
			</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<tr>
				<td><b>Enter invoice number</b></td>
			</tr>
			<tr>
			<td>If no invoice number is specified, the order number auto-generated by ProductCart will be used as the invoice number.<br>
				<b>Note: </b>You can use both characters and numbers for invoice number. You only need to enter the initial invoice number,  ProductCart will increase it by one automatically (e.g: EI-PC-ORD00011)</td>
			</tr>
			<tr>
				<td>
				#
				<input name="AlterInvoiceNum" type="text" id="AlterInvoiceNum">
				<%qry_ID=tmpStr %>
				<input type="hidden" name="id" value="<%=qry_ID%>">
				</td>
			</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<% 

						myDisplay=0
						call opendb()
						For k=1 to Count
							if (request("check" & k)="1") and (request("idord" & k)<>"") then
								query="SELECT idOrder from creditCards WHERE idOrder=" & request("idord" & k) & ";"
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)
								if NOT rs.eof then
									myDisplay=1
								end if
								query="SELECT idauthorder from authorders WHERE idOrder=" & request("idord" & k) & ";"
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)
								if NOT rs.eof then
									myDisplay=1
								end if
								query="SELECT idpfporder from pfporders WHERE idOrder=" & request("idord" & k) & ";"
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)
								if NOT rs.eof then
									myDisplay=1
								end if
								query="SELECT idOrder from netbillorders WHERE idOrder=" & request("idord" & k) & ";"
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)
								if NOT rs.eof then
									myDisplay=1
								end if
								query="SELECT idOrder from pcPay_USAePay_Orders WHERE idOrder=" & request("idord" & k) & ";"
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)
								if NOT rs.eof then
									myDisplay=1
								end if
								query="SELECT * from customCardOrders WHERE idOrder=" & request("idord" & k) & ";"
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)
								if NOT rs.eof then
									myDisplay=1
								end if
							end if
						Next
			
						if myDisplay=1 then %>
								<tr>
									<td><b>Would you like to display the card information on the invoice?</b></td>
								</tr>
								<tr>
									<td><input type="radio" name="showCCInfo" value="NO" checked class="clearBorder">
										No, I do not want card information printed on the invoice.</td>
								</tr>
								<tr>
									<td><input type="radio" name="showCCInfo" value="YES" class="clearBorder">
										Yes, I want card information printed on the invoice.</td>
								</tr>
								<tr>
									<td class="pcCPspacer"></td>
								</tr>

						<% else %>
								<tr><td><input type="hidden" name="showCCInfo" value="NO"></td></tr>
						<% end if %>

			<tr>
				<td><b>Would you like to include the Store Manager's Comments?</b></td>
			</tr>
			<tr>
				<td><input type="radio" name="showAdminComments" value="NO" checked class="clearBorder">
					No, don't print the administrator's comments on the invoice.</td>
			</tr>
			<tr>
				<td><input type="radio" name="showAdminComments" value="YES" class="clearBorder">
					Yes, print the administrator's comments on the invoice.</td>
			</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<tr>
				<td><input type="submit" name="GO" value="Submit" class="submit2"></td>
			</tr>
		</table>
	</form>
    </div>
	<% 
	call closedb()
	END IF
	%>
</body>
</html>