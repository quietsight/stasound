<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<%
'// Set to Residential
pResidentialShipping=1

'// Procede if the Country Code exists
IF Session("pcSFCountryCode")<>"" then

	shipmentTotal=Cdbl(0)

	'// UPS Variables
	query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	ups_license_key=trim(rstemp("AccessLicense"))
	ups_userid=trim(rstemp("userID"))
	ups_password=trim(rstemp("password"))
	ups_active=rstemp("active")

	'// CPS Variables
	query="SELECT active, shipServer, userID FROM ShipmentTypes WHERE idshipment=7"
	set rstemp=conntemp.execute(query)
	CP_userid=trim(rstemp("userID"))
	CP_server=trim(rstemp("shipserver"))
	CP_active=rstemp("active")

	'// FedEX Variables SD
	query="SELECT active, shipServer, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=1;"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	FedEX_server=trim(rstemp("shipserver"))
	FedEX_active=rstemp("active")
	FedEX_AccountNumber=trim(rstemp("userID"))
	FedEX_MeterNumber=trim(rstemp("password"))
	FEDEX_Environment=rstemp("AccessLicense")

	'// FedEX Variables WS
	query="SELECT active, shipServer, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=9;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	FedEXWS_server=trim(rs("shipserver"))
	FedEXWS_active=rs("active")
	FedEXWS_AccountNumber=trim(rs("userID"))
	FedEXWS_MeterNumber=trim(rs("password"))
	FEDEXWS_Environment=rs("AccessLicense")

	'// USPS Variables
	query="SELECT active, shipServer, userID, [password] FROM ShipmentTypes WHERE idshipment=4"
	set rstemp=conntemp.execute(query)
	usps_userid=trim(rstemp("userID"))
	usps_server=trim(rstemp("shipserver"))
	usps_active=rstemp("active")

	err.number=0

	'// page name
	pcStrPageName = "pcPay_GoogleCheckout_Shipping.asp"

	'// set error to zero
	pcv_intErr=0

	'// Set the Sessions Generated from Google
	CountryCode=Session("pcSFCountryCode")
	StateCode=Session("pcSFStateCode")
	Province=Session("pcSFProvince")
	city=Session("pcSFcity")
	zip=Session("pcSFzip")

	If pcv_intErr=0 Then

		ppcCartIndex=Session("pcCartIndex")

		'// Calculate total price of the order, total weight and product total quantities
		pSubTotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
		pShipSubTotal=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
		pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
		intUniversalWeight=pShipWeight
		pCartQuantity=Int(calculateCartQuantity(pcCartArray, ppcCartIndex))
		pCartShipQuantity=Int(calculateCartShipQuantity(pcCartArray, ppcCartIndex))
		pCartTotalWeight=Int(calculateCartWeight(pcCartArray, ppcCartIndex))

		Session("pSubTotal")=pSubTotal
		Session("pShipSubTotal")=pShipSubTotal
		Session("pShipWeight")=pShipWeight
		Session("intUniversalWeight")=intUniversalWeight
		Session("pCartQuantity")=pCartQuantity
		Session("pCartShipQuantity")=pCartShipQuantity
		Session("pCartTotalWeight")=pCartTotalWeight

		if Province="" then
			Province=StateCode
		end if

		Universal_destination_provOrState=Province
		Universal_destination_country=CountryCode
		if CountryCode<>"" then
			session("DestinationCountry")=CountryCode
		end if
		if Universal_destination_country="" then
			Universal_destination_country=session("DestinationCountry")
		end if
		Universal_destination_postal=zip
		Universal_destination_city=city

		'// If customer use anotherState, insert a dummy state code to simplify SQL sentence
		if Universal_destination_provOrState="" then
			 Universal_destination_provOrState="**"
		end if

		shipcompany=scShipService

		If pShipWeight="0" Then
			query="SELECT idFlatShiptype,WQP FROM FlatShipTypes"
			set rsShpObj=server.CreateObject("ADODB.RecordSet")
			set rsShpObj=conntemp.execute(query)
			if rsShpObj.eof then
				Session("nullShipper")="Yes"
				set rsShpObj=nothing
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
						'// Do Nothing
					end select
					rsShpObj.movenext
				loop
				set rsShpObj=nothing

				if flagShp=0 then
					Session("nullShipper")="Yes"
				else
					Session("nullShipper")="No"
				End if
			end if
		Else
			Session("nullShipper")="No"
		End If

		If pCartShipQuantity=0 then
			Session("nullShipper")="Yes"
		end if

		iShipmentTypeCnt=0
		%>
		<!--#include file="pcPay_GoogleCheckout_ShipRates.asp"-->
		<%
		session("strDefaultProvider")=strDefaultProvider
		session("iShipmentTypeCnt")=iShipmentTypeCnt
		session("strOptionShipmentType")=strOptionShipmentType
		session("availableShipStr")=availableShipStr
		session("iUPSFlag")=iUPSFlag

		availableShipStr=session("availableShipStr")

		query="SELECT shipService.idshipservice, shipService.serviceCode, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation FROM shipService WHERE (((shipService.serviceActive)=-1)) ORDER BY shipService.servicePriority;"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if rs.eof then
			Session("nullShipper")="Yes"
		else
			'// AVAILABLE RATES ARRAY
			Dim CntFree, DCnt, serviceFree, serviceFreeOverAmt, serviceCode, OrderTotal, shipArray, i, shipDetailsArray, tempRate, tempRateDisplay
			CntFree=0
			DCnt=0
			do until rs.eof
				pidshipservice=rs("idshipservice")
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
					'///////////////////////////////////////////////////////////////////////////////////
					'// START: SET TOTALS
					'///////////////////////////////////////////////////////////////////////////////////
					ptaxLoc=taxLoc
					pcv_ArrayGoogleShippingMethods = split(Session("attrPriceList"), "||||")
					shipArray=split(availableShipStr,"|?|")
					for i=lbound(shipArray) to (Ubound(shipArray))
						shipDetailsArray=split(shipArray(i),"|")
						if ubound(shipDetailsArray)>0 then

							if shipDetailsArray(1)=serviceCode then
								tempRate=shipDetailsArray(3)
								tempRateDisplay=money(tempRate)

								If serviceHandlingFee>0 then
									tempRate=(cDbl(shipDetailsArray(3))+cDbl(serviceHandlingFee))
									tempRateDisplay=money(tempRate)
								End If

								If serviceFree="-1" and Int(pSubTotal)>Int(serviceFreeOverAmt) then
									tempRateDisplay="0.00"
									CntFree=CntFree+1
								End If

								DCnt=DCnt+1
								Dim pshipDetailsArray2
								pshipDetailsArray2= shipDetailsArray(2)
								pshipDetailsArray2= replace(pshipDetailsArray2,"&reg;","")
								pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>SM</sup>","")
								pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>","")
								pshipDetailsArray2= replace(pshipDetailsArray2,"</sup>","")
								pshipDetailsArray2=trim(pshipDetailsArray2)
								pcv_strGoogleShippingPrices = pcv_strGoogleShippingPrices & pshipDetailsArray2 & "|?|" & pcf_CurrencyField(tempRateDisplay) & "|?|" & serviceCode & "||||"


								Session(pshipDetailsArray2)=tempRateDisplay
								Session(pshipDetailsArray2 & "_handling") = pcf_CurrencyField(serviceHandlingFee)
								Session(pshipDetailsArray2 & "_id") = pidshipservice
								%>
								<!--#include file="pcPay_GoogleCheckout_Tax.asp" -->
								<%
							end if
						end if
					next
					tempRate=""
					tempRateDisplay=""
					'///////////////////////////////////////////////////////////////////////////////////
					'// END: SET TOTALS
					'///////////////////////////////////////////////////////////////////////////////////
				end if
				rs.movenext
			loop
			set rs=nothing

		end if
	End If '// If pcv_intErr>0 Then
END IF '// IF request("SubmitShip")="" AND request("ddjumpflag")="" then
%>