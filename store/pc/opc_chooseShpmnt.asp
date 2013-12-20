<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 'SB S %>
<!--#include file="inc_sb.asp"-->
<% 'SB E %>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languages_ship.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/FedEXconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/FedEXWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/CPconstants.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include file="opc_contentType.asp" -->
<script src="../includes/spry/SpryTabbedPanels.js" type="text/javascript"></script>
<link href="../includes/spry/SpryTabbedPanels-SHIP.css" rel="stylesheet" type="text/css" />
<% On Error Resume Next
dim query, conntemp, rsTemp

dim pcHideEstimateDeliveryTimes
if ( scHideEstimateDeliveryTimes <> "" ) then
	pcHideEstimateDeliveryTimes = scHideEstimateDeliveryTimes
else
	pcHideEstimateDeliveryTimes = "0"
end if


Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

Call SetContentType()

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

Dim pcCartArray
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex" and check to see dbSession was not defined
'*****************************************************************************************************
%>
<!--#include file="pcVerifySession.asp"-->
<%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex" and check to see dbSession was not defined
'*****************************************************************************************************
Sub UpdateNullShipper(tmpvalue)
	Dim query,rs
	call opendb()
	query="UPDATE pcCustomerSessions SET pcCustSession_NullShipper='"& tmpvalue &"' WHERE idDbSession="&session("pcSFIdDbSession")&" AND randomKey="&session("pcSFRandomKey")&" AND idCustomer="&session("idCustomer")&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
End Sub

Sub UpdateNullShipRates(tmpvalue)
	Dim query,rs
	call opendb()
	query="UPDATE pcCustomerSessions SET pcCustSession_NullShipRates='"& tmpvalue &"' WHERE idDbSession="&session("pcSFIdDbSession")&" AND randomKey="&session("pcSFRandomKey")&" AND idCustomer="&session("idCustomer")&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
End Sub

ppcCartIndex=Session("pcCartIndex")

if request("ShippingChargeSubmit")<>"" then
	call opendb()

	pcStrShippingArray=URLDecode(getUserInput(request("Shipping"),0))
	pcIntOrdPackageNumber=URLDecode(getUserInput(request("ordPackageNum"),0))
	'// If there is any shipping at all, then we should have at least one package.
	if (pcIntOrdPackageNumber="" OR pcIntOrdPackageNumber=0) AND len(pcStrShippingArray)>0 then
		pcIntOrdPackageNumber=1
	end if

	query="UPDATE pcCustomerSessions SET pcCustSession_OrdPackageNumber="&pcIntOrdPackageNumber&", pcCustSession_ShippingArray='"&pcStrShippingArray&"' WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&"));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rs=nothing

	call closedb()
	response.clear
	Call SetContentType()
	response.write "OK"
	session("OPCstep")=4
	response.end
	'response.Redirect("tax.asp")
end if

shipmentTotal=Cdbl(0)

call openDb()

'//UPS Variables
query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

UPS_active=rs("active")
UPS_userid=trim(rs("userID"))
UPS_password=trim(rs("password"))
UPS_license_key=trim(rs("AccessLicense"))

'//CPS Variables
query="SELECT active, shipServer, userID FROM ShipmentTypes WHERE idshipment=7;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
CP_active=rs("active")
CP_server=trim(rs("shipserver"))
CP_userid=trim(rs("userID"))

'// FedEX Variables SD
Dim pcv_strAccountNameWS, pcv_strMeterNumberWS, pcv_strCarrierCodeWS
Dim pcv_strMethodNameWS, pcv_strMethodReplyWS, fedex_postdataWS, objFedExWSClass, objOutputXMLDocWS, srvFEDEXWSXmlHttp, FEDEXWS_result, FEDEXWS_URL, pcv_strErrorMsgWS

query="SELECT active, shipServer, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
FedEX_server=trim(rs("shipserver"))
FedEX_active=rs("active")
FedEX_AccountNumber=trim(rs("userID"))
FedEX_MeterNumber=trim(rs("password"))
FEDEX_Environment=rs("AccessLicense")

'// FedEX Variables WS
query="SELECT active, shipServer, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=9;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
If NOT rs.EOF Then
	FedEXWS_server=trim(rs("shipserver"))
	FedEXWS_active=rs("active")
	FedEXWS_AccountNumber=trim(rs("userID"))
	FedEXWS_MeterNumber=trim(rs("password"))
	FEDEXWS_Environment=rs("AccessLicense")
End If

ErrPageName = "login.asp"

'//USPS Variables
query="SELECT active, shipServer, userID FROM ShipmentTypes WHERE idshipment=4;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
usps_userid=trim(rs("userID"))
usps_server=trim(rs("shipserver"))
usps_active=rs("active")

set rs=nothing

err.number=0

'Retreive the saved shipping information from the customer sessions table
query="SELECT  pcCustomerSessions.idDbSession, pcCustomerSessions.randomKey, pcCustomerSessions.idCustomer, pcCustomerSessions.pcCustSession_BillingStateCode, pcCustomerSessions.pcCustSession_BillingCity, pcCustomerSessions.pcCustSession_BillingProvince, pcCustomerSessions.pcCustSession_BillingPostalCode, pcCustomerSessions.pcCustSession_BillingCountryCode, pcCustomerSessions.pcCustSession_CustomerEmail, pcCustomerSessions.pcCustSession_ShippingResidential, pcCustomerSessions.pcCustSession_ShippingCity, pcCustomerSessions.pcCustSession_ShippingStateCode, pcCustomerSessions.pcCustSession_ShippingProvince, pcCustomerSessions.pcCustSession_ShippingPostalCode, pcCustomerSessions.pcCustSession_ShippingCountryCode FROM pcCustomerSessions WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&")) ORDER BY idDbSession DESC;"

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
pcBillingStateCode=rs("pcCustSession_BillingStateCode")
pcBillingCity=rs("pcCustSession_BillingCity")
pcBillingProvince=rs("pcCustSession_BillingProvince")
pcBillingPostalCode=rs("pcCustSession_BillingPostalCode")
if NOT isNull(pcBillingPostalCode) then
	pcBillingPostalCode=pcf_PostCodes(pcBillingPostalCode)
end if
pcBillingCountryCode=rs("pcCustSession_BillingCountryCode")
pcCustomerEmail=rs("pcCustSession_CustomerEmail")
pResidentialShipping=rs("pcCustSession_ShippingResidential")
pcShippingCity=rs("pcCustSession_ShippingCity")
pcShippingStateCode=rs("pcCustSession_ShippingStateCode")
pcShippingProvince=rs("pcCustSession_ShippingProvince")
pcShippingPostalCode=rs("pcCustSession_ShippingPostalCode")
if NOT isNull(pcShippingPostalCode) then
	pcShippingPostalCode=pcf_PostCodes(pcShippingPostalCode)
end if
pcShippingCountryCode=rs("pcCustSession_ShippingCountryCode")

If Not len(pcShippingCity)>0 Then
	pcShippingCity = pcBillingCity
End If
If Not len(pcShippingStateCode)>0 Then
	pcShippingStateCode = pcBillingStateCode
End If
If Not len(pcShippingProvince)>0 Then
	pcShippingProvince = pcBillingProvince
End If
If Not len(pcShippingPostalCode)>0 Then
	pcShippingPostalCode = pcBillingPostalCode
End If
If Not len(pcShippingCountryCode)>0 Then
	pcShippingCountryCode = pcBillingCountryCode
End If

set rs=nothing


'// Do you want the cart shippable total to include discounts?
Dim pcv_intIncludeDiscounts
pcv_intIncludeDiscounts = 1

'// Calculate total price of the order, total weight and product total quantities

'////////////////////////////////////////////////////////////////////////////////
'// START  - Check IF recalculating shipping charges
'////////////////////////////////////////////////////////////////////////////////
' // Scenario: discount has been applied on orderVerify and order total
' // has changed and free shipping no longer applies.
pShipSubTotal=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
pShipCDSubTotal=Cdbl(calculateCategoryDiscounts(pcCartArray, ppcCartIndex))
if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
	'// Calculate Promo Price
	TotalPromotions=pcf_GetPromoTotal(Session("pcPromoSession"),Session("pcPromoIndex"))
end if
pSubTotal=trim(URLDecode(getUserInput(request.QueryString("pSubTotalCheckFreeShipping"),20)))
if pSubTotal = "" or isNull(pSubTotal) then
	' Not coming from orderVerify.asp, so calculate normally
	if session("SF_DiscountTotal")="" then session("SF_DiscountTotal")=0
	if session("SF_RewardPointTotal")="" then session("SF_RewardPointTotal")=0
	pSubTotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
	pSubTotal=pSubTotal-pShipCDSubTotal-TotalPromotions-session("SF_DiscountTotal")-session("SF_RewardPointTotal")
	If pcv_intIncludeDiscounts=1 Then
		pShipSubTotal=pSubTotal
	End If
else
	' Coming from orderVerify.asp, so overwrite pShipSubTotal
	' The sub total is updated on orderVerify.asp so we do not need to subtract anything
	pShipSubTotal=pSubTotal
end if
'////////////////////////////////////////////////////////////////////////////////
'// END  - Check IF recalculating shipping charges
'////////////////////////////////////////////////////////////////////////////////

pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
intUniversalWeight=pShipWeight
pCartQuantity=Int(calculateCartQuantity(pcCartArray, ppcCartIndex))
pCartShipQuantity=Int(calculateCartShipQuantity(pcCartArray, ppcCartIndex))
pCartSurcharge=Cdbl(calculateTotalProductSurcharge(pcCartArray, ppcCartIndex))

call closedb()

' check if state was entered for Shipping Address (only if Canada/US)
if pcShippingCountryCode<>"" then
	' use shipping codes
	If len(pcShippingStateCode)>0 Then
		pcShippingProvince=pcShippingStateCode
	end if
	Universal_destination_provOrState=pcShippingProvince
	Universal_destination_country=pcShippingCountryCode
	Universal_destination_postal=pcShippingPostalCode
	Universal_destination_city=pcShippingCity
else
	' use billing
	if pcBillingProvince="" then
		pcBillingProvince=pcBillingStateCode
	end if
	Universal_destination_provOrState=pcBillingProvince
	Universal_destination_country=pcBillingCountryCode
	Universal_destination_postal=pcBillingPostalCode
	Universal_destination_city=pcBillingCity
end if

' if customer use anotherState, insert a dummy state code to simplify SQL sentence
if Universal_destination_provOrState="" then
   Universal_destination_provOrState="**"
end if

shipcompany=scShipService

If pShipWeight="0" Then
	call opendb()

	query="SELECT active FROM ShipmentTypes WHERE active<>0"
	set rs=connTemp.execute(query)
	if rs.eof then '// There are NO active dynamic shipping services
		pcv_NoDynamicShipping="1"
	end if

	query="SELECT idFlatShiptype,WQP FROM FlatShipTypes"
	set rsShpObj=server.CreateObject("ADODB.RecordSet")
	set rsShpObj=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsShpObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if rsShpObj.eof then
		call UpdateNullShipper("Yes")
		call closeDb()
		set rsShpObj=nothing
		response.Clear()
		Call SetContentType()
		If pcv_NoDynamicShipping="1" Then
			response.write "OK|*|<div class='pcSuccessMessage'>" & dictLanguage.Item(Session("language")&"_opc_ship_1") & "</div>"
			session("OPCstep")=4
			response.end
		End If
		'response.redirect "tax.asp?idDbSession="& pIdDbSession &"&randomKey="& pRandomKey
	else
		dim flagShp
		flagShp=0
		do until rsShpObj.eof
			intIdFlatShipType=rsShpObj("idFlatShiptype")
			pShpObjType=rsShpObj("WQP")
			select case pShpObjType
			case "Q"
				flagShp=1
			case "P"
				flagShp=1
			case "O"
				flagShp=1
			case "I"
				flagShp=1
			case "W"
				'do nothing
			end select
			rsShpObj.movenext
		loop
		set rsShpObj=nothing

		if flagShp=0 then
			call UpdateNullShipper("Yes")
		else
			call UpdateNullShipper("No")
		End if
	end if
	call closedb()
Else
	call UpdateNullShipper("No")
End If

If pCartShipQuantity=0 then
	call UpdateNullShipper("Yes")
	response.Clear()
	Call SetContentType()
	response.write "OK|*|<div class='pcSuccessMessage'>" & dictLanguage.Item(Session("language")&"_opc_ship_1") & "</div>"
	session("OPCstep")=4
	response.end
	'response.redirect "tax.asp?idDbSession="& pIdDbSession &"&randomKey="& pRandomKey
end if

iShipmentTypeCnt=0
%>
<!--#include file="ShipRates.asp"-->
<%err.number=0
err.description=""
call openDB()
query="SELECT shipService.serviceCode, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation FROM shipService WHERE (((shipService.serviceActive)=-1)) ORDER BY shipService.servicePriority;"

set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if rs.eof then
	call UpdateNullShipper("Yes")
	call closedb()
	response.Clear()
	Call SetContentType()
	response.write "OK|*|<div class='pcSuccessMessage'>" & dictLanguage.Item(Session("language")&"_opc_ship_1") & "</div>"
	session("OPCstep")=4
	response.end
	'response.redirect "tax.asp?idDbSession="& pIdDbSession &"&randomKey="& pRandomKey
else %>
<script language="JavaScript"><!--
function newWindow(file,window) {
		PackageWindow=open(file,window,'resizable=no,width=500,height=600,scrollbars=1');
		if (PackageWindow.opener == null) PackageWindow.opener = self;
}
//--></script>
	<form name="ShipChargeForm" id="ShipChargeForm">
	<table class="pcShowContent">
		<%
		'=============================================================================
		' START optional shipping-related message
		'  - if the feature is on
		'  - if it's setup to show the message at the top
		'=============================================================================
		if PC_SECTIONSHOW="TOP" AND PC_RATESONLY="NO" then %>
			<tr>
				<td>
					<table class="pcShowContent">
						<tr>
							<td class="pcSectionTitle">
								<p><%=PC_SHIP_DETAIL_TITLE%></p>
							</td>
						</tr>
						<tr>
							<td>
								<p><%=PC_SHIP_DETAILS%></p>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="pcSpacer"></td>
			</tr>
		<% end if
		'=============================================================================
		' END optional shipping-related message
		'=============================================================================

		'=============================================================================
		' START show package information
		'  - if the feature is on
		'  - if there is more than 1 package
		'=============================================================================
		if scHideProductPackage <> "-1" then
			if pcv_intTotPackageNum>1 then
			%>
				<tr>
					<td>
						<table class="pcShowContent">
							<tr>
								<td class="pcSectionTitle">
									<p><%=dictLanguage.Item(Session("language")&"_CustviewOrd_38")%></p>
								</td>
							</tr>
							<tr>
								<td>
									<p><%=ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_h")&pcv_intTotPackageNum&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_i") %></p>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="pcSpacer"></td>
				</tr>
			<%
			end if
		end if
		'=============================================================================
		' END show package information
		'=============================================================================
		%>
		<%' load previous entered fields in hidden HTML tags %>
		<tr>
			<td>
			 <input type="hidden" name="ordPackageNum" value="<%=pcv_intTotPackageNum%>">
			<div id="TabbedPanelsShipping" class="TabbedPanels">
			<%
			'=============================================================================
			' START shipping provide selection
			' If more than 1 provider is active, ask customer to choose which one to display
			' This feature was introduced to remain compliant with UPS requirements
			'=============================================================================
			dim strTabOrder
			strTabOrder=""

			if iShipmentTypeCnt=>1 then %>
			  <ul class="TabbedPanelsTabGroup">
				<% if instr(strTabShipmentType,"[/TAB][TAB]") then
					strTabShipmentTypeArry=split(strTabShipmentType,"[/TAB]")
					strFirstTab=""
					strTabs=""
					for itab=0 to ubound(strTabShipmentTypeArry)
						strTabProvider=replace(strTabShipmentTypeArry(itab),"[TAB]","")
						strTabProviderArry=split(strTabProvider,",")
						if strTabProviderArry(0)=strDefaultProvider then
							strFirstTab="<li class=""TabbedPanelsTab"" tabindex=""0"">"&strTabProviderArry(1)&"</li>"
							strTabFirst=strTabProviderArry(0)&","
						else
							strTabs=strTabs&"<li class=""TabbedPanelsTab"" tabindex=""0"">"&strTabProviderArry(1)&"</li>"
							strTabOrder=strTabOrder&strTabProviderArry(0)&","
						end if
					Next
					strTabOrder=strTabFirst&strTabOrder
				else
					if instr(strTabShipmentType,"[/TAB]") then
						strTabProvider=replace(strTabShipmentType,"[TAB]","")
						strTabProvider=replace(strTabProvider,"[/TAB]","")
						strTabProviderArry=split(strTabProvider,",")
						strFirstTab="<li class=""TabbedPanelsTab"" tabindex=""0"">"&strTabProviderArry(1)&"</li>"
						strTabFirst=strTabProviderArry(0)&","
						strTabOrder=strTabFirst
					end if
				end if
				response.write strFirstTab
				response.write strTabs
				%>
			  </ul>
			<% end if
			'=============================================================================
			' END shipping provider selection
			'=============================================================================

			'=============================================================================
			' START display shipping options
			'=============================================================================
			Dim CntFree, DCnt, serviceFree, serviceFreeOverAmt, serviceCode, OrderTotal, shipArray, i, shipDetailsArray, tempRate, tempRateDisplay
			CntFree=0
			DCnt=0
			FedExCnt=0
			FedExWSCnt=0 '// WS
			USPSCnt=0
			UPSCnt=0
			CPCnt=0
			CUSTOMCnt=0
			pcv_Default=0
			do until rs.eof
				serviceCode=rs("serviceCode")
				serviceFree=rs("serviceFree")
				serviceFreeOverAmt=rs("serviceFreeOverAmt")
				serviceHandlingFee=rs("serviceHandlingFee")
				serviceHandlingIntFee=rs("serviceHandlingIntFee")
				serviceShowHandlingFee=rs("serviceShowHandlingFee")
				serviceLimitation=rs("serviceLimitation")
				customerLimitation=0
				if serviceLimitation<>0 then
					if serviceLimitation=1 then
						if Universal_destination_country=scShipFromPostalCountry then
							customerLimitation=1
						end if
					end if
					if serviceLimitation=2 then
						if Universal_destination_country<>scShipFromPostalCountry then
							customerLimitation=1
						end if
					end if
					if serviceLimitation=3 then
						if ucase(trim(Universal_destination_country))<>"US" then
							customerLimitation=1
						else
							if ucase(trim(Universal_destination_provOrState))="AK" OR ucase(trim(Universal_destination_provOrState))="HI" then
								customerLimitation=1
							end if
						end if
					end if
					if serviceLimitation=4 then
						if ucase(trim(Universal_destination_country))<>"US" then
							customerLimitation=1
						else
							if ucase(trim(Universal_destination_provOrState))<>"AK" AND ucase(trim(Universal_destination_provOrState))<>"HI" then
								customerLimitation=1
							end if
						end if
					end if
				end if

				if customerLimitation=0 then
					shipArray=split(availableShipStr,"|?|")
					for i=lbound(shipArray) to (Ubound(shipArray))
						shipDetailsArray=split(shipArray(i),"|")

						if ubound(shipDetailsArray)>0 then
							if shipDetailsArray(1)=serviceCode then
								tempRate=shipDetailsArray(3)
								if ubound(shipDetailsArray)>4 then
									pcvNegRate=shipDetailsArray(5)
									if isNumeric(pcvNegRate) AND pcvNegRate<>0 then
									if ucase(shipDetailsArray(0))="UPS" then
										if pcv_UseNegotiatedRates=1 AND pcvNegRate<>"NONE"  then
											tempRate=pcvNegRate
											end if
										end if
									end if
								end if
								tempRate=(cDbl(tempRate)+cDbl(pCartSurcharge))

								tempRateDisplay=scCurSign&money(tempRate)
								If serviceShowHandlingFee="0" then
									tempRate=(cDbl(tempRate)+cDbl(serviceHandlingFee))
									tempRateDisplay=scCurSign&money(tempRate)
									serviceHandlingFee="0"
								End If
								'if (ucase(shipDetailsArray(0))=ucase(session("provider"))) then
									If serviceFree="-1" and Cdbl(pSubTotal)>Cdbl(serviceFreeOverAmt) then
										tempRate="0"
										tempRateDisplay= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_f")
										CntFree=CntFree+1
									End If
									pshipDetailsArray2= shipDetailsArray(2)
									pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&reg;</sup>","")
									pshipDetailsArray2= replace(pshipDetailsArray2,"&reg;","")
									pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>SM</sup>","")

									if pcv_Default=0 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
										pcv_Default=1
									else
										x_checked="XCHECK"
									end if

									select case ucase(shipDetailsArray(0))
									case "FEDEX", "FedEX", "FedEx", "FEDEXWS", "FedExWS"
										DCnt=DCnt+1
										FedExCnt=FedExCnt+1
										if FedExCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
											x_checked=" checked"
										else
											if FedExCnt=1 AND pcv_Default=0 then
												x_checked="FCHECK"
											end if
										end if
										strFEDEX=strFEDEX&"<tr class='opcShippingSelect'><td><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</td><td>"
										if shipDetailsArray(4)<>"NA" AND pcHideEstimateDeliveryTimes <> "-1" then
											strFEDEX=strFEDEX&shipDetailsArray(4)
										end if

										strFEDEX=strFEDEX&"</td><td>"&tempRateDisplay&"</td></tr>"
									case "USPS"
										DCnt=DCnt+1
										USPSCnt=USPSCnt+1
										if USPSCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
											x_checked=" checked"
										else
											if USPSCnt=1 AND pcv_Default=0 then
												x_checked="FCHECK"
											end if
										end if
										strUSPS=strUSPS&"<tr class='opcShippingSelect'><td><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</td><td>"
										if shipDetailsArray(4)<>"NA" AND pcHideEstimateDeliveryTimes <> "-1"  then
											strUSPS=strUSPS&shipDetailsArray(4)
										end if
										strUSPS=strUSPS&"</td><td>"&tempRateDisplay&"</td></tr>"
									case "UPS"
										DCnt=DCnt+1
										UPSCnt=UPSCnt+1
										if UPSCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
											x_checked=" checked"
										else
											if UPSCnt=1 AND pcv_Default=0 then
												x_checked="FCHECK"
											end if
										end if
										strUPS=strUPS&"<tr class='opcShippingSelect'><td><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</td><td>"
										if shipDetailsArray(4)<>"NA" AND pcHideEstimateDeliveryTimes <> "-1"  then
											strUPS=strUPS&shipDetailsArray(4)
										end if
										strUPS=strUPS&"</td><td>"&tempRateDisplay&"</td></tr>"
									case "CP"
										DCnt=DCnt+1
										CPCnt=CPCnt+1
										if CPCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
											x_checked=" checked"
										else
											if CPCnt=1 AND pcv_Default=0 then
												x_checked="FCHECK"
											end if
										end if
										strCP=strCP&"<tr class='opcShippingSelect'><td><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</td><td>"
										if shipDetailsArray(4)<>"NA" AND pcHideEstimateDeliveryTimes <> "-1"  then
											strCP=strCP&shipDetailsArray(4)
										end if
										strCP=strCP&"</td><td>"&tempRateDisplay&"</td></tr>"
									case "CUSTOM"
										DCnt=DCnt+1
										CUSTOMCnt=CUSTOMCnt+1
										if CUSTOMCnt=1 AND pcv_Default=1 AND ucase(scDefaultProvider)=ucase(shipDetailsArray(0))then
											x_checked=" checked"
										else
											if CUSTOMCnt=1 AND pcv_Default=0 then
												x_checked="FCHECK"
											end if
										end if
										strCUSTOM=strCUSTOM&"<tr class='opcShippingSelect'><td><input type='radio' name='Shipping' value='"&shipDetailsArray(0)&","&pshipDetailsArray2&","&tempRate&","&serviceHandlingFee&","&serviceShowHandlingFee&","&shipDetailsArray(1)&"' class='clearBorder'"&x_checked&">&nbsp;"&shipDetailsArray(2)&"</td><td>"
										if shipDetailsArray(4)<>"NA" AND pcHideEstimateDeliveryTimes <> "-1"  then
											strCUSTOM=strCUSTOM&shipDetailsArray(4)
										end if
										strCUSTOM=strCUSTOM&"</td><td>"&tempRateDisplay&"</td></tr>"
									end select
									%>

								<%
								'end if
							end if
						end if
					next
					tempRate=""
					tempRateDisplay=""
				end if
				rs.movenext
			loop
			set rs=nothing
			call closeDb()

			'//Replace last instance of each provider string with an alternate class
			strUPSLastInStr = split(strUPS, "<tr class='opcShippingSelect'>")
			strReplace = "<tr class='opcShippingSelect'>"&strUPSLastInStr(Ubound(strUPSLastInStr))
			strNew = "<tr>"&strUPSLastInStr(Ubound(strUPSLastInStr))
			strUPS = replace(strUPS, strReplace, strNew)

			strUSPSLastInStr = split(strUSPS, "<tr class='opcShippingSelect'>")
			strReplace = "<tr class='opcShippingSelect'>"&strUSPSLastInStr(Ubound(strUSPSLastInStr))
			strNew = "<tr>"&strUSPSLastInStr(Ubound(strUSPSLastInStr))
			strUSPS = replace(strUSPS, strReplace, strNew)

			strFedExLastInStr = split(strFedEx, "<tr class='opcShippingSelect'>")
			strReplace = "<tr class='opcShippingSelect'>"&strFedExLastInStr(Ubound(strFedExLastInStr))
			strNew = "<tr>"&strFedExLastInStr(Ubound(strFedExLastInStr))
			strFedEx = replace(strFedEx, strReplace, strNew)

			strCPLastInStr = split(strCP, "<tr class='opcShippingSelect'>")
			strReplace = "<tr class='opcShippingSelect'>"&strCPLastInStr(Ubound(strCPLastInStr))
			strNew = "<tr>"&strCPLastInStr(Ubound(strCPLastInStr))
			strCP = replace(strCP, strReplace, strNew)

			strCUSTOMLastInStr = split(strCUSTOM, "<tr class='opcShippingSelect'>")
			strReplace = "<tr class='opcShippingSelect'>"&strCUSTOMLastInStr(Ubound(strCUSTOMLastInStr))
			strNew = "<tr>"&strCUSTOMLastInStr(Ubound(strCUSTOMLastInStr))
			strCUSTOM = replace(strCUSTOM, strReplace, strNew)

			'//ENSURE THERE IS AT LEAST ONE OPTION CHECKED - Can happen if no rates are returned by the default provider that is set in the CP
			if pcv_Default=0 then
				strFEDEX=replace(strFEDEX,"XCHECK","")
				strUSPS=replace(strUSPS,"XCHECK","")
				strUPS=replace(strUPS,"XCHECK","")
				strCP=replace(strCP,"XCHECK","")
				strCUSTOM=replace(strCUSTOM,"XCHECK","")
			else
				strFEDEX=replace(strFEDEX,"XCHECK","")
				strUSPS=replace(strUSPS,"XCHECK","")
				strUPS=replace(strUPS,"XCHECK","")
				strCP=replace(strCP,"XCHECK","")
				strCUSTOM=replace(strCUSTOM,"XCHECK","")
				strFEDEX=replace(strFEDEX,"FCHECK","")
				strUSPS=replace(strUSPS,"FCHECK","")
				strUPS=replace(strUPS,"FCHECK","")
				strCP=replace(strCP,"FCHECK","")
				strCUSTOM=replace(strCUSTOM,"FCHECK","")
			end if
			inttotalUPSWeight=0
			for uCnt=1 to pcv_intPackageNum
				intTotalUPSWeight=intTotalUPSWeight+session("UPSPackWeight"&uCnt)
			next
			%>

			<%
			strContent=""

			if (pcHideEstimateDeliveryTimes <> "-1") then
				strProviderHeader="<table class='pcShowContent'><tr><th width='40%'>&nbsp;"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_a")&"</th><th width='40%'>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_b")&"</th><th width='20%'>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_c")&"</th></tr><tr><td colspan='3' class='pcSpacer'></td></tr>"
			else
				strProviderHeader="<table class='pcShowContent'><tr><th width='60%'>&nbsp;"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_a")&"</th><th width='10%'>&nbsp;</th><th width='30%'>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_c")&"</th></tr><tr><td colspan='3' class='pcSpacer'></td></tr>"
			end if

			strUPSFooter="<tr bgcolor='#FFFFFF'><td colspan='3'><table width='100%' border='0' cellspacing='0' cellpadding='3'><tr><td><a href=""javascript:;"" onclick=""newWindow('pcUPSTimeInTransit.asp?sResidential="&pResidentialShipping&"&sPackageCnt="&pcv_intTotPackageNum&"&sWeight="&pShipWeight&"&sState="&universal_destination_provOrstate&"&sCity="&universal_destination_city&"&sPC="&universal_destination_postal&"&sCountry="&universal_destination_country&"','ProductWindow')"">Time In Transit</a>: Calculate estimated transit time for the various UPS services.</td></tr><tr><td colspan='3'><hr></td></tr></table><table width='100%' border='0' cellspacing='0' cellpadding='3'><tr> <td width='9%' valign='top'><img src='../UPSLicense/LOGO_S2.jpg' width='45' height='50'></td><td width='91%' rowspan='2' valign='top'><p><b>UPS OnLine&reg; Tools Rates & Service Selection</b></p><p>Notice: UPS fees do not necessarily represent UPS published rates and may include charges levied by the store owner.</p><p class='pcSmallText'>UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OFUNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</p></td></tr><tr> <td class='pcSpacer'></td></tr></table></td></tr>"
%>
			<div class="TabbedPanelsContentGroup">
				<% strTabOrderArry=split(strTabOrder,",")
				for iContent=0 to ubound(strTabOrderArry)-1
					select case strTabOrderArry(iContent)
					case "USPS"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK"," checked")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div class=""TabbedPanelsContent"">"&strProviderHeader&strUSPS&"</Table></div>"
					case "CP"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK"," checked")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div class=""TabbedPanelsContent"">"&strProviderHeader&strCP&"</Table></div>"
					case "FEDEX", "FedEX", "FedEx", "FEDEXWS", "FedExWS"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK"," checked")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div class=""TabbedPanelsContent"">"&strProviderHeader&strFEDEX&"<tr><td colspan='3'><p class='pcSmallText'>FedEx service marks are owned by Federal Express Corporation and used with permission.</p></td></tr></Table></div>"
					case "UPS"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK"," checked")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK","")
						end if
						strContent=strContent&"<div class=""TabbedPanelsContent"">"&strProviderHeader&strUPS&strUPSFooter&"</Table></div>"
					case "CUSTOM"
						if iContent=0 AND pcv_Default=0 then
							strFEDEX=replace(strFEDEX,"FCHECK","")
							strUSPS=replace(strUSPS,"FCHECK","")
							strUPS=replace(strUPS,"FCHECK","")
							strCP=replace(strCP,"FCHECK","")
							strCUSTOM=replace(strCUSTOM,"FCHECK"," checked")
						end if
						strContent=strContent&"<div class=""TabbedPanelsContent"">"&strProviderHeader&strCUSTOM&"</Table></div>"
					end select
				next %>
				<% response.write strContent %>
			</div>
			</div>
			<script type="text/javascript">
			<!--
			try {
				var TabbedPanelsShipping = new Spry.Widget.TabbedPanels("TabbedPanelsShipping");
			} catch(err) { }
			//-->
			</script>

			</td></tr>

			<% call UpdateNullShipRates("No")
			dim intCRates
			intCRates=0
			If DCnt=0 then
				if scAlwNoShipRates="-1" then
					call UpdateNullShipRates("Yes")
					ShowJSSubmitValidation="1"
					'=============================================================================
					' START show messages about no shipping options available
					' No shipping rates and checkout allowed
					'=============================================================================
					%>
					<tr>
						<td colspan="4">
							<p>
							<% intCRates=1 %>
							<% = ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_g")%>
							</p>
						</td>
					</tr>
					<tr>
						<td colspan="4"><div id="ShippingChargesLoader"></div></td>
					</tr>
					<tr>
						<td colspan="4">
							<input type="image" name="ShippingChargeSubmit" id="ShippingChargeSubmit" src="<%=RSlayout("pcLO_Update")%>" border="0">
						</td>
					</tr>
				<% else
					'=============================================================================
					' No shipping rates and checkout NOT allowed
					'=============================================================================
					response.Clear()
					Call SetContentType()
					response.write "STOP|*|" & dictLanguage.Item(Session("language")&"_opc_ship_2")
					session("OPCstep")=0
					response.end
				end if
				'=============================================================================
				' END show messages about no shipping options available
				'=============================================================================
			else
			ShowJSSubmitValidation="1" %>
				<tr>
					<td colspan="4"><div id="ShippingChargesLoader"></div></td>
				</tr>
				<tr>
					<td colspan="4">
						<input type="image" name="ShippingChargeSubmit" id="ShippingChargeSubmit" src="<%=RSlayout("pcLO_Update")%>" border="0">
					</td>
				</tr>
			<% end if %>

			<%if ShowJSSubmitValidation="1" then%>
			<tr>
				<td colspan="4">
					<script>
						//*Submit Shipping Charges Form
						$('#ShippingChargeSubmit').click(function(){
						{
							$.ajax({
								type: "POST",
								url: "opc_chooseShpmnt.asp",
								data: $('#ShipChargeForm').formSerialize() + "&ShippingChargeSubmit=yes",
								timeout: 45000,
								error: function (XMLHttpRequest, textStatus, errorThrown) {
									if (textStatus=='timeout') {
										$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_72"))%>');
										$("#GlobalErrorDialog").dialog('open');
										return false;
									} else {
										$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_73"))%>');
										$("#GlobalErrorDialog").dialog('open');
										return false;
									}
								},
								global: false,
								success: function(data, textStatus){
								if (data=="SECURITY")
								{
									// Session Expired
									window.location="msg.asp?message=1";
								}
								else
								{
									if (data=="OK")
									{
										$("#ShippingChargesLoader").hide();
										//$("#ShippingChargesLoader").html('<img src="images/pcv4_st_icon_success_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_ship_4"))%>');

										getShipMethod();
										getTaxContents();
										btnShow1("OK","Ship");
									}
									else
									{
										$("#ShippingChargesLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> ' + data);
										$("#ShippingChargesLoader").show();
										btnShow1("Error","SC");
									}
									}
								}
							});
							return(false);
						}
						return(false);
						});
					</script>
				</td>
			</tr>
			<%end if%>


			<%'=============================================================================
			' END display shipping options
			'=============================================================================

			'=============================================================================
			' START optional shipping-related message
			'  - if the feature is on
			'  - if it's setup to show the message at the bottom
			'=============================================================================
			dim intDisplay
			intDisplay=0
			if PC_SECTIONSHOW="BTM" then
				intDisplay=1
				if PC_RATESONLY="YES" then
					if intCRates=1 then
						intDisplay=0
					end if
				end if
			end if
			if intDisplay=1 then
				%>
				<tr>
					<td>
						<table class="pcShowContent">
							<tr class="pcSectionTitle">
								<td><%=PC_SHIP_DETAIL_TITLE%></td>
							</tr>
							<tr>
								<td>
									<p><%=PC_SHIP_DETAILS%></p>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<%
			end if
			'=============================================================================
			' END optional shipping-related message
			'=============================================================================
			%>
	</table>
	</form>
<% end if %>
<% call closeDb()
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing %>
