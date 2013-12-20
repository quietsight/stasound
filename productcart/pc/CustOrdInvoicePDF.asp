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
<!--#include file="../pdf/fpdf.asp"-->
<%Dim pdf%>
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
			query="SELECT orders.pcOrd_OrderKey,orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr, idcustomer, orderdate, Address, city, stateCode,state, zip,CountryCode, paymentDetails, shipmentDetails, shippingAddress, shippingCity, shippingStateCode, shippingState, shippingZip, shippingCountryCode, pcOrd_shippingPhone, idAffiliate, affiliatePay, discountDetails, pcOrd_GCDetails, pcOrd_GCAmount, taxAmount, total, comments, orderStatus, processDate, shipDate, shipvia, trackingNum, returnDate, returnReason, ShippingFullName, ord_DeliveryDate, ord_OrderName, iRewardPoints, iRewardPointsCustAccrued, iRewardValue, address2, shippingCompany, shippingAddress2, taxDetails, rmaCredit, SRF, ord_VAT, pcOrd_CatDiscounts, gwAuthCode, gwTransId, paymentCode, pcOrd_GCs, pcOrd_GcCode, pcOrd_GcUsed, pcOrd_IDEvent, pcOrd_GWTotal FROM orders WHERE idOrder=" & qry_ID & ";"
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
			
			pcv_OrderKey=rs("pcOrd_OrderKey")
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
			ListPDFFiles=""
			While Not rs.EOF %>
			<%Set pdf=CreateJsObject("FPDF")
			pdf.CreatePDF "P", "mm", "A4"
			pdf.SetPath("../pdf")
			pdf.LoadExtension("table")
			pdf.LoadExtension("Rotate")
			pdf.SetSubject ("Order Invoice")
			pdf.SetTitle("Order Invoice ID#" & (scpre+int(qry_ID)))
			pdf.SetCreator("ProductCart software - NetSource Commerce. (www.earlyimpact.com)")
			pdf.SetAuthor(scCompanyName)
			pdf.SetLeftMargin(10)
			pdf.SetRightMargin(10)
			pdf.SetDisplayMode("real")
			pdf.Open()
			
			tmpModData=""
			tmpModData=tmpModData & "this.Header=function Header()" & vbcrlf
			tmpModData=tmpModData & "{" & vbcrlf
			tmpModData=tmpModData & "this.SetFont('Arial','',22);" & vbcrlf
			tmpModData=tmpModData & "this.SetTextColor(204,204,204);" & vbcrlf
			tmpModData=tmpModData & "this.RotatedText(7,47,'Order Invoice',90);" & vbcrlf
			tmpModData=tmpModData & "this.SetTextColor(0,0,0);" & vbcrlf
			tmpModData=tmpModData & "}" & vbcrlf
			tmpModData=tmpModData & "this.Footer=function Footer()" & vbcrlf
			tmpModData=tmpModData & "{" & vbcrlf
			tmpModData=tmpModData & "this.SetY(-10);" & vbcrlf
			tmpModData=tmpModData & "this.SetFont('Arial','',8);" & vbcrlf
			if (Not IsNull(pcv_OrderKey)) AND pcv_OrderKey<>"" then
				tmpPDF=" (Code: " & pcv_OrderKey & ")"
			else
				tmpPDF=""
			end if
			tmpModData=tmpModData & "this.Cell(0,10,'" & replace(scCompanyName,"'","\'") & ": ORDER INVOICE ID#" & (scpre+int(qry_ID)) & tmpPDF & " - Page '+ this.PageNo()+ '/{nb}',0,0,'L');" & vbcrlf
			tmpModData=tmpModData & "}" & vbcrlf
			
			SavedFile = "../pdf/models/HeaderFooter.js"
			findit = Server.MapPath(Savedfile)
			Set fso = server.CreateObject("Scripting.FileSystemObject")
			Err.number=0
			Set f = fso.OpenTextFile(findit, 2)
			f.WriteLine tmpModData
			f.close
			Set f=nothing
			Set fso=nothing
			
			pdf.LoadModels("HeaderFooter")
			pdf.AddPage()
			
			pdf.Table.Border.Color="000000"
			pdf.Table.TextAlign = "L"
			pdf.SetAligns "L", "L", "C"
			pdf.SetBorders 0,0,0
			pdf.SetFont "Arial","",8

			pdf.Table.Border.Width = 0

			'START - Company Information & Order Date
			pdf.SetColumns 100,50,40
			pdf.SetCustomBorders 0,0,1
			pdf.SetCustomStyle "B","",""
			if porderdate <> "" then
				tmpDate=porderdate
			else
				tmpDate="N/A"
			end if
			pdf.Row scCompanyName,"",dictLanguage.Item(Session("language")&"_custOrdInvoice_1") & tmpDate
			pdf.SetCustomBorders 0,0,0
			pdf.SetCustomStyle "","",""
			pdf.Row scCompanyAddress & vbcrlf & scCompanyCity & ", " & scCompanyState & " " & scCompanyZip & vbcrlf & scStoreURL,"",""

			pdf.Ln 5
			'END - Company Information & Order Date
			
			'START - Billing Information
			CurX=pdf.GetX()
			CurY=pdf.GetY()

			pdf.SetColumns 90
			pdf.Table.Border.Width = 0.1
			pdf.SetCustomStyle "B"
			pdf.SetBorders "B"
			pdf.SetFont "Arial","",7
			pdf.SetCellHeight 4
			pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_2")
			pdf.SetFont "Arial","",8
			pdf.SetCustomStyle ""
			pdf.SetBorders "T"
			tmpPDF=""
			tmpPDF=CustomerName & vbcrlf
			if CustomerCompany<>"" then 
				tmpPDF=tmpPDF & CustomerCompany & vbcrlf
			end if
			tmpPDF=tmpPDF & pAddress & vbcrlf
			if pAddress2<>"" then 
				tmpPDF=tmpPDF & pAddress2 & vbcrlf
			end if
			tmpPDF=tmpPDF & pcity&", "&pStateCode&" "&pzip
			if pCountryCode <> scShipFromPostalCountry then
				tmpPDF=tmpPDF & vbcrlf & pCountryCode
			end if
			if CustomerPhone<>"" then
				tmpPDF=tmpPDF & vbcrlf & dictLanguage.Item(Session("language")&"_custOrdInvoice_3") & CustomerPhone
			end if
			if CustomerEmail<>"" then
				tmpPDF=tmpPDF & vbcrlf & dictLanguage.Item(Session("language")&"_custOrdInvoice_4") & CustomerEmail
			end if
			if CustomerFax<>"" then
				tmpPDF=tmpPDF & vbcrlf & "Fax: " & CustomerFax
			end if
			pdf.SetCellHeight 3
			pdf.Row tmpPDF
			NewY=pdf.GetY()
			pdf.SetBorders 0
			'END - Billing Information
			
			'START - Order Information
			pdf.SetXY curX+100,curY
			pdf.SetCellHeight 4
			pdf.SetCustomStyle "B"
			pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_5") & (scpre+int(qry_ID))
			
			'Order Key
			if (Not IsNull(pcv_OrderKey)) AND pcv_OrderKey<>"" then
				tmpPDF=dictLanguage.Item(Session("language")&"_opc_common_1") & " " & pcv_OrderKey
			else
				tmpPDF=""
			end if
			if tmpPDF<>"" then
				CurY=pdf.GetY()
				pdf.SetXY curX+100,curY
				pdf.Row tmpPDF
			end if
			
			pdf.SetFont "Arial","",7
			pdf.SetCustomStyle ""
						
			' Calculate customer number using sccustpre constant
			Dim pcCustomerNumber
			if len(sccustpre)>0 then
				pcCustomerNumber = (sccustpre + int(pidcustomer))
			else
				pcCustomerNumber = (int(pidcustomer))
			end if
			CurY=pdf.GetY()
			pdf.SetXY curX+100,curY
			pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_6") & pcCustomerNumber
			
			if scOrderName="1" then
				if trim(pord_OrderName) <> "" Then
					CurY=pdf.GetY()
					pdf.SetXY curX+100,curY
					pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_7") & pord_OrderName
				end If
			end if
			
			If trim(pord_DeliveryDate) <> "1/1/1900" and trim(pord_DeliveryDate) <> "" Then
				if scDateFrmt="DD/MM/YY" then
					pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
				else
					pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
				end if
				pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
				CurY=pdf.GetY()
				pdf.SetXY curX+100,curY
				pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_8") & pord_DeliveryDate & ", " & pord_DeliveryTime
			End If
			
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
				
				CurY=pdf.GetY()
				pdf.SetXY curX+100,curY
				pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_1") & geName
				CurY=pdf.GetY()
				pdf.SetXY curX+100,curY
				pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_2") & geDate
				CurY=pdf.GetY()
				pdf.SetXY curX+100,curY
				pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_3") & gReg
			else
				geHideAddress=0
			End If
			'GGG Add-on end
			
			tmpPDF=""
			tmpPDF=dictLanguage.Item(Session("language")&"_custOrdInvoice_9") 
			If pSRF="1" then
				tmpPDF=tmpPDF & ship_dictLanguage.Item(Session("language")&"_noShip_b")
			else
				'get shipping details...
				shipping=split(pshipmentDetails,",")
				if ubound(shipping)>1 then
					if NOT isNumeric(trim(shipping(2))) then
						varShip="0"
						tmpPDF=tmpPDF & ship_dictLanguage.Item(Session("language")&"_noShip_a")
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
						tmpPDF=tmpPDF & Service
					End If
				else
					varShip="0"
					tmpPDF=tmpPDF & ship_dictLanguage.Item(Session("language")&"_noShip_a")
				end if
			end if
			CurY=pdf.GetY()
			pdf.SetXY curX+100,curY
			pdf.Row tmpPDF

			payment = split(ppaymentDetails,"||")
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
								else
									CurY=pdf.GetY()
									pdf.SetXY curX+100,curY 
									pdf.Row ""
									tmpPDF=""
									tmpPDF=dictLanguage.Item(Session("language")&"_custOrdInvoice_10")
									if PaymentName <> "" and PaymentName <> PaymentType then
										Dim pcv_strPaymentType
										Select Case PaymentType
											Case "PayPal Website Payments Pro": pcv_strPaymentType=PaymentName
											Case Else: pcv_strPaymentType=PaymentName & " (" & PaymentType & ")"
										End Select
										tmpPDF=tmpPDF & pcv_strPaymentType
									else
										tmpPDF=tmpPDF & PaymentType
									end if
										%>
										<% if PayCharge>0 then
											tmpPDF=tmpPDF & vbcrlf & dictLanguage.Item(Session("language")&"_custOrdInvoice_11") & " " & scCurSign&money(PayCharge)
										end if %>
										<% if varTransID<>"" then
										tmpPDF=tmpPDF & vbcrlf & varTransName & ": " & varTransID
										end if %>
										<% if varAuthCode<>"" then
										tmpPDF=tmpPDF & vbcrlf & varAuthName & ": " & varAuthCode
										end if %>
										<%if varResponseCode <> "" and varISO_Code <> "" Then
										tmpPDF=tmpPDF & vbcrlf & varRespName & " " & varResponseCode & "/" & varISO_Code
										end if %>
									    <% if varIDEBIT_ISSCONF <> ""  and varIDEBIT_ISSNAME <> "" then
										tmpPDF=tmpPDF & vbcrlf & dictLanguage.Item(Session("language")&"_CustviewOrd_48")
										tmpPDF=tmpPDF & vbcrlf & dictLanguage.Item(Session("language")&"_CustviewOrd_49") & " " & varIDEBIT_ISSNAME
										tmpPDF=tmpPDF & vbcrlf & dictLanguage.Item(Session("language")&"_CustviewOrd_50") & " " & varIDEBIT_ISSCONF						
										end if
										tmpPDF=tmpPDF & vbcrlf
										CurY=pdf.GetY()
										pdf.SetXY curX+100,curY
										pdf.Row tmpPDF
								end if %>								
								<% If RewardsActive <> 0 And piRewardPoints > 0 Then 
									iDollarValue = piRewardPoints * (RewardsPercent / 100)
									CurY=pdf.GetY()
									pdf.SetXY curX+100,curY
									pdf.Row ucase(RewardsLabel) & ":" & vbcrlf & dictLanguage.Item(Session("language")&"_custOrdInvoice_12") & piRewardPoints & " " & RewardsLabel & dictLanguage.Item(Session("language")&"_custOrdInvoice_13") & scCurSign&money(iDollarValue)
								end if %>
								<% If RewardsActive <> 0 And piRewardPointsCustAccrued > 0 Then
									CurY=pdf.GetY()
									pdf.SetXY curX+100,curY
									pdf.Row ucase(RewardsLabel) & ":" & vbcrlf & dictLanguage.Item(Session("language")&"_custOrdInvoice_14") & piRewardPointsCustAccrued & " " & RewardsLabel & dictLanguage.Item(Session("language")&"_custOrdInvoice_15")
								end if %>
									<% 'if discount was present, show type here
									'Check if more then one discount code was utilized
									if instr(pdiscountDetails,",") then
										DiscountDetailsArry=split(pdiscountDetails,",")
										intArryCnt=ubound(DiscountDetailsArry)
										for ki=0 to intArryCnt
											if (DiscountDetailsArry(ki)<>"") AND (instr(DiscountDetailsArry(ki),"- ||")=0) then
												DiscountDetailsArry(ki+1)=DiscountDetailsArry(ki)+"," + DiscountDetailsArry(ki+1)
												DiscountDetailsArry(ki)=""
											end if
										next
									else
										intArryCnt=0
									end if

									for ki=0 to intArryCnt
										if intArryCnt=0 then
											pTempDiscountDetails=pdiscountDetails
										else
											pTempDiscountDetails=DiscountDetailsArry(ki)
										end if
										if instr(pTempDiscountDetails,"- ||") then 
											discounts = split(pTempDiscountDetails,"- ||")
											discountType = discounts(0)
											discount = discounts(1)
											if discountType<>"" then
												CurY=pdf.GetY()
												pdf.SetXY curX+100,curY
												pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_16") & discountType
											end if
										end if
									Next %>
									<%'start of gift certificates
									if GCDetails<>"" then
										GCArry=split(GCDetails,"|g|")
										intArryCnt=ubound(GCArry)
				
										for ki=0 to intArryCnt
					
											if GCArry(ki)<>"" then
												GCInfo = split(GCArry(ki),"|s|")
												if GCInfo(2)="" OR IsNull(GCInfo(2)) then
													GCInfo(2)=0
												end if
												CurY=pdf.GetY()
												pdf.SetXY curX+100,curY
												pdf.Row dictLanguage.Item(Session("language")&"_CustviewOrd_15A") & GCInfo(1) & " (" & GCInfo(0) & ")"
											end if
										Next
									end if
									'end if gift certificates									
			'END - Order Information
			
			'START - Shipping Address Information
			if geHideAddress=0 then
				pdf.SetCellHeight 4
				pdf.SetXY curX,NewY+3
				pdf.SetCustomStyle "B"
				pdf.SetBorders "B"
				pdf.SetFont "Arial","",7
				pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_17")
				pdf.SetCellHeight 3
				pdf.SetFont "Arial","",8
				pdf.SetCustomStyle ""
				pdf.SetBorders "T"
				
				tmpPDF=""
											if pcShowShipAddr="0" then
												
												tmpPDF="(Same as billing address)"
	
											ELSE%>
											
												<% 
												if pshippingFullName<>"" then
													tmpPDF=tmpPDF & pshippingFullName
												else
													tmpPDF=tmpPDF & CustomerName
												end if
												tmpPDF=tmpPDF & vbcrlf									
												if pshippingCompany<>"" then 
													tmpPDF=tmpPDF & pshippingCompany & vbcrlf
												else
													if (pshippingFullName = "" or pshippingFullName = CustomerName) and customerCompany <> "" then
														tmpPDF=tmpPDF & customerCompany & vbcrlf
													end if											
												end if
												tmpPDF=tmpPDF & pshippingAddress & vbcrlf
												if pshippingAddress2<>"" then 
													tmpPDF=tmpPDF & pshippingAddress2& vbcrlf
												end if
												tmpPDF=tmpPDF & pshippingcity & ", " & pshippingStateCode & " " & pshippingZip
												if pShippingCountryCode <> scShipFromPostalCountry then
													tmpPDF=tmpPDF & vbcrlf & pShippingCountryCode
												end if
												if pshippingEmail <> "" then
													tmpPDF=tmpPDF & vbcrlf & "E-mail: " & pshippingEmail
												end if
												if pshippingPhone <> "" then
													tmpPDF=tmpPDF & vbcrlf & dictLanguage.Item(Session("language")&"_custOrdInvoice_3") & pshippingPhone
												end if
												if pshippingFax <> "" then
													tmpPDF=tmpPDF & vbcrlf & "Fax: " & pshippingFax
												end if
												%>
                    						<% END IF %>
				<%pdf.Row tmpPDF %>
			<% end if
			'END - Shipping Address Information
			pdf.SetCellHeight 4
			pdf.SetFont "Arial","",8
			pdf.Ln 8
			
			pdf.SetColumns 15,100,40,35
			pdf.SetAligns "R", "L", "R", "R"
			pdf.SetCustomBorders 0,0,0,0
			pdf.SetBorders 0,0,0,0

			'START - Products

			pdf.Table.Border.Width = 0.1
			pdf.SetCustomStyle "B","B","B","B"

			pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_18"),dictLanguage.Item(Session("language")&"_custOrdInvoice_19"),dictLanguage.Item(Session("language")&"_custOrdInvoice_20"),dictLanguage.Item(Session("language")&"_custOrdInvoice_21")

			pdf.SetCustomStyle "","","",""
			
			%>
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

		pdf.SetCustomStyle "","","",""
		pdf.SetAligns "R", "L", "R", "R"
		pdf.SetColumns 15,100,40,35
		pdf.SetBorders 0,0,0,0
		pdf.Row pquantity,psku & " - " & pDescription,scCurSign&money(punitPrice1),scCurSign&money(pRowPrice1)%>

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
											
											pdf.SetCustomStyle "","U","",""
											pdf.SetAligns "L", "L", "L", "R"
											pdf.SetColumns 15,100,40,35
											pdf.SetBorders "B","BRO","LBRO","LBO"
											pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_22") & ":","",""
											pdf.SetCustomStyle "","","",""
											%>
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
									
									pdf.SetColumns 15,40,100,35
									pdf.SetBorders "BT","BTRO","TLBRO","TLBO"
									
									tmpPDFC=rsConfigObj("categoryDesc")
									
									tmpPDF=""
                         			tmpPDF=rsConfigObj("sku") & " - " & rsConfigObj("description")
									if btDisplayQF=True AND clng(ArrQuantity(i))>1 then
										tmpPDF=tmpPDF & " - " & dictLanguage.Item(Session("language")&"_custOrdInvoice_18") & ": " & ArrQuantity(i)
									end if%>
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
										<%tmpPDF1=""
										if pnoprices<2 then%>
											<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then%>
												<%tmpPDF1=scCurSign & money((ArrValue(i)+UPrice)*pQuantity)%>
											<%else
												if tmpDefault=1 then%>
													<%tmpPDF1=dictLanguage.Item(Session("language")&"_defaultnotice_1")%>
												<%end if
											end if%>
										<% end if %>
                      				<% set rsConfigObj=nothing
									pdf.Row "", tmpPDFC & ":",tmpPDF,tmpPDF1
									next
									set rsConfigObj=nothing %>
                		<% end if %>
                	<% end if
					'	BTO ADDON-E
					
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
							tmpFirst=1
							For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
								pdf.SetCustomStyle "","","",""
								pdf.SetColumns 15,100,40,35
								pdf.SetAligns "L", "L", "R", "R"
								if tmpFirst=1 then
									tmpFirst=0
									pdf.SetBorders "B","BRO","LBRO","LBO"
								else
									pdf.SetBorders "BT","BTRO","TLBRO","TLBO"
								end if
								tmpPDF1=pcArray_strOptions(pcv_intOptionLoopCounter)
								
								tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
							
								if tempPrice="" or tempPrice=0 then
									tmpPDF2=""
									tmpPDF3=""
								else
									'// Adjust for Quantity Discounts
									tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice)
									tmpPDF2=scCurSign&money(tempPrice)
									tAprice=(tempPrice*Cdbl(pquantity))
									tmpPDF3=scCurSign&money(tAprice) 
								end if 
								pdf.Row "",tmpPDF1,tmpPDF2,tmpPDF3
							
							Next
							'#####################
							' END LOOP
							'#####################					
			
				end if
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: SHOW PRODUCT OPTIONS
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				%>
				<!-- end options -->

				<% if pxdetails<>"" then
					pdf.SetCustomStyle "","","",""
					pdf.SetColumns 15,100,40,35
					pdf.SetAligns "L", "L", "R", "R"
					pdf.SetBorders 0,0,0,0
					pdf.Row "",replace(pxdetails,"<br>",vbcrlf),"",""
				end if %>

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
												pdf.SetCustomStyle "","","",""
												pdf.SetColumns 15,100,40,35
												pdf.SetAligns "L", "L", "R", "R"
												pdf.SetBorders 0,0,0,0
												pdf.Row "","",dictLanguage.Item(Session("language")&"_custOrdInvoice_23"),scCurSign&money(-1*ItemsDiscounts)
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
													if ArrCProduct(0)<>"na" then 
														pdf.SetCustomStyle "","U","",""
														pdf.SetAligns "L", "L", "L", "R"
														pdf.SetBorders "B","BRO","LBRO","LBO"
														pdf.SetColumns 15,100,40,35
														pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_24"),"",""
														pdf.SetCustomStyle "","","",""
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
																		tmpPDF1=rsConfigObj("categoryDesc") & ":"
																		tmpPDF2=rsConfigObj("sku") & " - " & rsConfigObj("description")
																		tmpPDF3=""
																		if pnoprices<2 then
																			if ArrCValue(i)>0 then
																				tmpPDF3=scCurSign & money(ArrCValue(i))
																			end if
																		end if
																		set rsConfigObj=nothing
																		pdf.SetColumns 15,40,100,35
																		pdf.SetBorders "BT","BTRO","TLBRO","TLBO"
																		pdf.Row "",tmpPDF1,tmpPDF2,tmpPDF3
																	next
																	set rsConfigObj=nothing
																	pRowPrice=pRowPrice+Cdbl(Charges)%>
							<% end if 'Have Additional Charges
						end if
						'BTO Additional Charges %>
                			<% end if
                end if 'BTO %>

                			<% QDiscounts=trim(QDiscounts)
											if QDiscounts="" then
												QDiscounts=0
											end if
				
                if (QDiscounts<>"") and (CDbl(QDiscounts)<>"0") then
					pdf.SetCustomStyle "","",""
					pdf.SetColumns 15,140,35
					pdf.SetAligns "L", "R", "R"
					pdf.SetBorders 0,0,0
					pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_25"),scCurSign&money(-1*QDiscounts)
					pRowPrice=pRowPrice-Cdbl(QDiscounts)
                end if %>
											
                <% if pExtRowPrice<>pRowPrice then
					pdf.SetCustomStyle "","","",""
					pdf.SetColumns 15,100,40,35
					pdf.SetAligns "L", "L", "R", "R"
					pdf.SetBorders 0,0,0,0
					pdf.Row "","",dictLanguage.Item(Session("language")&"_custOrdInvoice_26"),scCurSign&money(pRowPrice)
				end if %>

                					<% 'GGG Add-on start
									if pGWOpt<>"0" then
									query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
									set rsG=connTemp.execute(query)
									if not rsG.eof then
										pdf.SetCustomStyle "",""
										pdf.SetColumns 15,175
										pdf.SetAligns "L", "L"
										pdf.SetBorders 0,0
										tmpPDF=dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_4") & " " & rsG("pcGW_OptName") & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_5") & " " & scCurSign & money(pGWPrice)
										if pGWText<>"" then
											tmpPDF=tmpPDF & vbcrlf
											tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_6") & vbcrlf & replace(pGWText,"<br>",vbcrlf)
										end if
										pdf.Row "",tmpPDF
									end if
									end if
									'GGG Add-on end
									
                                    if pcPrdOrd_BundledDisc>0 then
										pdf.SetCustomStyle "","","",""
										pdf.SetColumns 15,100,40,35
										pdf.SetAligns "L", "L", "R", "R"
										pdf.SetBorders 0,0,0,0
										pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_36"),"","-" & scCurSign&money(pcPrdOrd_BundledDisc)
                                    end if
                					rstemp.moveNext
								loop
								set rstemp=nothing
								'END - Products%>
<%'START - Fees & Order Total%>        
                				<% 'RP ADDON-S
								If RewardsActive<>0 Then
									if piRewardValue<>0 then
										pdf.SetCustomStyle "","","",""
										pdf.SetColumns 15,100,40,35
										pdf.SetAligns "L", "L", "R", "R"
										pdf.SetBorders 0,0,0,0
										if RewardsLabel="" then
											RewardsLabel="Rewards Program"
										end if
										pdf.Row "",RewardsLabel,"","-" & scCurSign&money(piRewardValue)
									end if
								End if
								'RP ADDON-E %>
								<%'GGG Add-on start
								if pGWTotal>0 then
										pdf.SetCustomStyle "B",""
										pdf.SetColumns 155,35
										pdf.SetAligns "R", "R"
										pdf.SetBorders 0,0
										pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_7"),scCurSign&money(pGWTotal)
										pdf.SetCustomStyle "",""
								end if
								'GGG Add-on end
								
								pdf.SetCustomStyle "","",""
								pdf.SetColumns 115,40,35
								pdf.SetAligns "L", "R", "R"
								pdf.SetBorders 0,0,0
								pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_27"),scCurSign&money(postage)
								%>
								
								
								<% if serviceHandlingFee<>0 then
									pdf.SetCustomStyle "","",""
									pdf.SetColumns 115,40,35
									pdf.SetAligns "L", "R", "R"
									pdf.SetBorders "TB",0,0
									pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_28"),scCurSign&money(serviceHandlingFee)
								end if %>
								
								<% if PayCharge>0 then
									pdf.SetCustomStyle "","",""
									pdf.SetColumns 115,40,35
									pdf.SetAligns "L", "R", "R"
									pdf.SetBorders "TB",0,0
									pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_29"),scCurSign&money(PayCharge)
								end if %>
								<%
								
								' If the store is using VAT and VAT is > 0, don't show any taxes here, but show VAT after the total
								if NOT (pord_VAT>0) then

										if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
											pdf.SetCustomStyle "","",""
											pdf.SetColumns 115,40,35
											pdf.SetAligns "L", "R", "R"
											pdf.SetBorders "TB",0,0
											pdf.Row "", dictLanguage.Item(Session("language")&"_custOrdInvoice_30"),scCurSign&money(ptaxAmount)
										else %>
											<% taxArray=split(ptaxDetails,",")
											tempTaxAmount=0
											for i=0 to (ubound(taxArray)-1)
												taxDesc=split(taxArray(i),"|")
												if taxDesc(0)<>"" then
													pDisTax=(money(taxDesc(1)))
													pdf.SetCustomStyle "","",""
													pdf.SetColumns 115,40,35
													pdf.SetAligns "L", "R", "R"
													pdf.SetBorders "TB",0,0
													pdf.Row "",ucase(taxDesc(0)),scCurSign&pDisTax
												end if 
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
										for ki=0 to intArryCnt
											if intArryCnt=0 then
												pTempDiscountDetails=pdiscountDetails
											else
												pTempDiscountDetails=DiscountDetailsArry(ki)
											end if
											if instr(pTempDiscountDetails,"- ||") then 
												discounts = split(pTempDiscountDetails,"- ||")
												discountType = discounts(0)
												tdiscount = discounts(1)
											else
												tdiscount=0
											end if
											discount=discount+tdiscount
										Next
										pdf.SetCustomStyle "","",""
										pdf.SetColumns 115,40,35
										pdf.SetAligns "L", "R", "R"
										pdf.SetBorders "TB",0,0
										pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_31"),"-" & scCurSign&money(discount+pcv_CatDiscounts)
									end if %>
									<%'GGG Add-on start
									IF GCAmount>"0" THEN
										pdf.SetCustomStyle "","",""
										pdf.SetColumns 115,40,35
										pdf.SetAligns "L", "R", "R"
										pdf.SetBorders "TB",0,0
										pdf.Row "",dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_8"),"-" & scCurSign&money(GCAmount)
									END IF
									'GGG Add-on end
									
									pdf.SetCustomStyle "","B","B"
									pdf.SetColumns 115,40,35
									pdf.SetAligns "L", "R", "R"
									pdf.SetBorders "T",0,0
									pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_32") & ":",scCurSign&money(ptotal)
									pdf.SetCustomStyle "","",""
									%>
									<% 
									' If the store is using VAT and VAT > 0, show it here
									if pord_VAT>0 then %>

                                        
										<% if pcv_IsEUMemberState=1 then
											pdf.SetCustomStyle ""
											pdf.SetColumns 190
											pdf.SetAligns "R"
											pdf.SetBorders 0
											pdf.Row dictLanguage.Item(Session("language")&"_orderverify_35") & scCurSign&money(pord_VAT)
										else
											pdf.SetCustomStyle ""
											pdf.SetColumns 190
											pdf.SetAligns "R"
											pdf.SetBorders 0
											pdf.Row dictLanguage.Item(Session("language")&"_orderverify_42") & scCurSign&money(pord_VAT)
										end if %> 
                                        
                                        
									<% end if %>
									<% if NOT isNull(prmaCredit) AND prmaCredit<>"" AND prmaCredit>0 then
										pdf.SetCustomStyle "","",""
										pdf.SetColumns 115,40,35
										pdf.SetAligns "L", "R", "R"
										pdf.SetBorders 0,0,0
										pdf.Row "",dictLanguage.Item(Session("language")&"_custOrdInvoice_34"),"-" & scCurSign&money(pRmaCredit)
									end if %>
          <%rs.MoveNext
			Wend
			Set rs=Nothing
			
			%>
			
			<%'GGG Add-on start
			IF (GCDetails<>"") then
			
				pdf.Ln 8
			
				pdf.Table.Border.Width=0
				pdf.SetCustomStyle "B"
				pdf.SetColumns 190
				pdf.SetAligns "L"
				pdf.SetBorders 0
				pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_9")
				pdf.SetCustomStyle ""
			
				GCArry=split(GCDetails,"|g|")
				intArryCnt=ubound(GCArry)
			
				for ki=0 to intArryCnt
				
				if GCArry(ki)<>"" then
					GCInfo = split(GCArry(ki),"|s|")
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
					
					pdf.SetCustomStyle "",""
					pdf.SetColumns 30,160
					pdf.SetAligns "L","L"
					pdf.SetBorders 0,0
					pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_10"),pName
					tmpPDF=""
					query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
					set rs19=connTemp.execute(query)

					if not rs19.eof then
						tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_11") & rs19("pcGO_GcCode") & vbcrlf
						tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_12") & scCurSign & money(pGiftUsed) & vbcrlf & vbcrlf
						pGCAmount=rs19("pcGO_Amount")
						if cdbl(pGCAmount)<=0 then%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_13")%>
						<%else%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_14") & scCurSign & money(pGCAmount)
							tmpPDF=tmpPDF & vbcrlf
							pExpDate=rs19("pcGO_ExpDate")
							if year(pExpDate)="1900" then%>
								<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_15")%>
							<%else
								if scDateFrmt="DD/MM/YY" then
									pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
								else
									pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
								end if
								tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_16") & pExpDate
							end if
							tmpPDF=tmpPDF & vbcrlf
							pGCStatus=rs19("pcGO_Status")
							if pGCStatus="1" then%>
								<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_17")%>
							<%else%>
								<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_18")%>
							<%end if%>
						<%end if
						tmpPDF=tmpPDF & vbcrlf & vbcrlf
					end if
					set rs19=nothing
					pdf.Row "",tmpPDF
				end if
				set rsG=nothing
				end if
				Next%>
			<% END IF
			'GGG Add-on end%>
			
			<%'GGG Add-on start
			IF (pGCs<>"") and (pGCs="1") then
			
				pdf.Ln 8
				
				pdf.Table.Border.Width=0
				pdf.SetCustomStyle "B"
				pdf.SetColumns 190
				pdf.SetAligns "L"
				pdf.SetBorders 0
				pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_19")
				pdf.SetCustomStyle ""
				
				query="select * from ProductsOrdered WHERE idOrder="& qry_ID
				set rs11=connTemp.execute(query)
				do while not rs11.eof
					query="select products.Description,pcGCOrdered.pcGO_GcCode from Products,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& qry_ID
					set rsG=connTemp.execute(query)

					if not rsG.eof then
						gIdproduct=rs11("idproduct")
						gName=rsG("Description")
						gCode=rsG("pcGO_GcCode")
						
					pdf.SetCustomStyle "",""
					pdf.SetColumns 30,160
					pdf.SetAligns "L","L"
					pdf.SetBorders 0,0
					pdf.Row dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_10"),gName
					tmpPDF=""
					
					query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & qry_ID
					set rs19=connTemp.execute(query)

					do while not rs19.eof%>
						<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_11") & rs19("pcGO_GcCode") & vbcrlf
						pExpDate=rs19("pcGO_ExpDate")
						if year(pExpDate)="1900" then%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_15")%>
						<%else
							if scDateFrmt="DD/MM/YY" then
								pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
							else
								pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
							end if%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_16") & pExpDate
						end if
						tmpPDF=tmpPDF & vbcrlf
						pGCAmount=rs19("pcGO_Amount")
						if cdbl(pGCAmount)<=0 then%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_13")%>
						<%else%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_14") & scCurSign & money(pGCAmount)
						end if
						tmpPDF=tmpPDF & vbcrlf
						pGCStatus=rs19("pcGO_Status")
						if pGCStatus="1" then%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_17")%>
						<%else%>
							<%tmpPDF=tmpPDF & dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_18")%>
						<%end if
						tmpPDF=tmpPDF & vbcrlf & vbcrlf
						rs19.movenext
					loop
					set rs19=nothing
					pdf.Row "",tmpPDF
				end if
				set rsG=nothing
				rs11.MoveNext
				loop
				set rs11=nothing
			END IF
			'GGG Add-on end%>
			
			<% if pcomments<>"" then
				
				pdf.Ln 8
				
				pdf.Table.Border.Width=0
				pdf.SetCustomStyle ""
				pdf.SetColumns 190
				pdf.SetAligns "L"
				pdf.SetBorders 0
				pdf.Row dictLanguage.Item(Session("language")&"_custOrdInvoice_35") & vbcrlf & vbcrlf & replace(pcomments,"<br>",vbcrlf)
				pdf.SetCustomStyle ""
			end if %>
		<%
		
		pdf.Close()
		if (Not IsNull(pcv_OrderKey)) AND pcv_OrderKey<>"" then
			tmpExt="-" & pcv_OrderKey
		else
			tmpExt=""
		end if
		tmpPDFFile="Order-Invoice-ID" & (scpre+int(qry_ID)) & tmpExt & ".pdf"
		ListPDFFiles=ListPDFFiles & "<li>" & "<a href='catalog/" & tmpPDFFile & "'>" & tmpPDFFile & "</a></li>"
		pdf.Output server.mappath("./catalog/" & tmpPDFFile),1
		
		%>
 
		<div class="pcSuccessMessage">
			Order Invoice ID# <%=(scpre+int(qry_ID))%> was generated successfully!<br>
			<ul><%=ListPDFFiles%></ul>
			Please <a href="catalog/<%=tmpPDFFile%>">right click</a> on the link and choose "Save Target As..." to download this invoice.
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