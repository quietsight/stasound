<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx Web Services Shipping Configuration - Select Shipping Services" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

		<% Dim query, rs, connTemp
		call openDb()


		if request.querystring("mode")="InAct" then
			' inactivate
			set rs=Server.CreateObject("ADODB.Recordset")

			query="UPDATE ShipmentTypes SET active=0, international=0 WHERE idShipment=9;"
			set rs=connTemp.execute(query)

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEXWS"
		end if


		if request.querystring("mode")="Act" then
			' activate
			set rs=Server.CreateObject("ADODB.Recordset")
			query = "SELECT * FROM ShipmentTypes WHERE idShipment=9;"
			set rs=connTemp.execute(query)
			IF RS.EOF THEN
				CALL CLOSEDB()
				RESPONSE.REDIRECT "upddb_v46.asp"
				response.end
			END IF
			query="UPDATE ShipmentTypes SET active=-1, international=0 WHERE idShipment=9;"
			set rs=connTemp.execute(query)

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEXWS"
		end if


		if request.querystring("mode")="del" then
			'remove
			set rs=Server.CreateObject("ADODB.Recordset")
			'clear all informatin out of shipService for service

			query="UPDATE ShipmentTypes SET shipServer='', active=0, international=0 WHERE idShipment=9;"
			set rs=connTemp.execute(query)

			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='PRIORITY_OVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FIRST_OVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='STANDARD_OVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_2_DAY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_EXPRESS_SAVER';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_PRIORITY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_ECONOMY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_FIRST';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_1_DAY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_2_DAY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_3_DAY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_GROUND';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='GROUND_HOME_DELIVERY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_PRIORITY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_ECONOMY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_GROUND';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_NATIONAL_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='SMART_POST';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='EUROPE_FIRST_INTERNATIONAL_PRIORITY';"
			set rs=connTemp.execute(query)
			set rs=nothing

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEXWS"
		end if

		'check for real integers
		Function validNum2(strInput)
			DIM iposition		' Current position of the character or cursor
			validNum2 =  true
			if isNULL(strInput) OR trim(strInput)="" then
				validNum2 = false
			else
				'loop through each character in the string and validate that it is a number or integer
				For iposition=1 To Len(trim(strInput))
					if InStr(1, "12345676890", mid(strInput,iposition,1), 1) = 0 then
						validNum2 =  false
						Exit For
					end if
				Next
			end if
		end Function

		if request.form("submit")<>"" then

			pcStrService=request.form("FEDEXWS_SERVICE")
			if pcStrService="" then
				response.redirect "FedEXWS_EditShipOptions.asp?msg="&Server.URLEncode("Select at least one service.")
				response.end
			end if
			pcStrFreeShip=""
			pcStrHandling=""
			servicePriorityStr=""

			'PRIORITY_OVERNIGHT
			If request.form("FREE-PRIORITY_OVERNIGHT")="YES" then
				pcFreeAmount=request.form("AMT-PRIORITY_OVERNIGHT")
				pcStrFreeShip=pcStrFreeShip&"PRIORITY_OVERNIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-PRIORITY_OVERNIGHT")<>"0" AND request.form("HAND-PRIORITY_OVERNIGHT")<>"" then
				If isNumeric(request.form("HAND-PRIORITY_OVERNIGHT"))=true then
					pcStrHandling=pcStrHandling&"PRIORITY_OVERNIGHT|"&replacecomma(request.form("HAND-PRIORITY_OVERNIGHT"))&"|"&request.form("SHFEE-PRIORITY_OVERNIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-PRIORITY_OVERNIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"PRIORITY_OVERNIGHT|"&servicePriority&","

			'FIRST_OVERNIGHT
			If request.form("FREE-FIRST_OVERNIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FIRST_OVERNIGHT")
				pcStrFreeShip=pcStrFreeShip&"FIRST_OVERNIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FIRST_OVERNIGHT")<>"0" AND request.form("HAND-FIRST_OVERNIGHT")<>"" then
				If isNumeric(request.form("HAND-FIRST_OVERNIGHT"))=true then
					pcStrHandling=pcStrHandling&"FIRST_OVERNIGHT|"&replacecomma(request.form("HAND-FIRST_OVERNIGHT"))&"|"&request.form("SHFEE-FIRST_OVERNIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FIRST_OVERNIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FIRST_OVERNIGHT|"&servicePriority&","

			'STANDARD_OVERNIGHT
			If request.form("FREE-STANDARD_OVERNIGHT")="YES" then
				pcFreeAmount=request.form("AMT-STANDARD_OVERNIGHT")
				pcStrFreeShip=pcStrFreeShip&"STANDARD_OVERNIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-STANDARD_OVERNIGHT")<>"0" AND request.form("HAND-STANDARD_OVERNIGHT")<>"" then
				If isNumeric(request.form("HAND-STANDARD_OVERNIGHT"))=true then
					pcStrHandling=pcStrHandling&"STANDARD_OVERNIGHT|"&replacecomma(request.form("HAND-STANDARD_OVERNIGHT"))&"|"&request.form("SHFEE-STANDARD_OVERNIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-STANDARD_OVERNIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"STANDARD_OVERNIGHT|"&servicePriority&","

			'FEDEX_2_DAY
			If request.form("FREE-FEDEX_2_DAY")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX_2_DAY")
				pcStrFreeShip=pcStrFreeShip&"FEDEX_2_DAY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX_2_DAY")<>"0" AND request.form("HAND-FEDEX_2_DAY")<>"" then
				If isNumeric(request.form("HAND-FEDEX_2_DAY"))=true then
					pcStrHandling=pcStrHandling&"FEDEX_2_DAY|"&replacecomma(request.form("HAND-FEDEX_2_DAY"))&"|"&request.form("SHFEE-FEDEX_2_DAY")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX_2_DAY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX_2_DAY|"&servicePriority&","

			'FEDEX_EXPRESS_SAVER
			If request.form("FREE-FEDEX_EXPRESS_SAVER")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX_EXPRESS_SAVER")
				pcStrFreeShip=pcStrFreeShip&"FEDEX_EXPRESS_SAVER|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX_EXPRESS_SAVER")<>"0" AND request.form("HAND-FEDEX_EXPRESS_SAVER")<>"" then
				If isNumeric(request.form("HAND-FEDEX_EXPRESS_SAVER"))=true then
					pcStrHandling=pcStrHandling&"FEDEX_EXPRESS_SAVER|"&replacecomma(request.form("HAND-FEDEX_EXPRESS_SAVER"))&"|"&request.form("SHFEE-FEDEX_EXPRESS_SAVER")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX_EXPRESS_SAVER")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX_EXPRESS_SAVER|"&servicePriority&","

			'INTERNATIONAL_PRIORITY
			If request.form("FREE-INTERNATIONAL_PRIORITY")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONAL_PRIORITY")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONAL_PRIORITY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONAL_PRIORITY")<>"0" AND request.form("HAND-INTERNATIONAL_PRIORITY")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONAL_PRIORITY"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONAL_PRIORITY|"&replacecomma(request.form("HAND-INTERNATIONAL_PRIORITY"))&"|"&request.form("SHFEE-INTERNATIONAL_PRIORITY")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONAL_PRIORITY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONAL_PRIORITY|"&servicePriority&","

			'INTERNATIONAL_ECONOMY
			If request.form("FREE-INTERNATIONAL_ECONOMY")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONAL_ECONOMY")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONAL_ECONOMY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONAL_ECONOMY")<>"0" AND request.form("HAND-INTERNATIONAL_ECONOMY")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONAL_ECONOMY"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONAL_ECONOMY|"&replacecomma(request.form("HAND-INTERNATIONAL_ECONOMY"))&"|"&request.form("SHFEE-INTERNATIONAL_ECONOMY")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONAL_ECONOMY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONAL_ECONOMY|"&servicePriority&","

			'INTERNATIONAL_FIRST
			If request.form("FREE-INTERNATIONAL_FIRST")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONAL_FIRST")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONAL_FIRST|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONAL_FIRST")<>"0" AND request.form("HAND-INTERNATIONAL_FIRST")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONAL_FIRST"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONAL_FIRST|"&replacecomma(request.form("HAND-INTERNATIONAL_FIRST"))&"|"&request.form("SHFEE-INTERNATIONAL_FIRST")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONAL_FIRST")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONAL_FIRST|"&servicePriority&","

			'FEDEX_1_DAY_FREIGHT
			If request.form("FREE-FEDEX_1_DAY_FREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX_1_DAY_FREIGHT")
				pcStrFreeShip=pcStrFreeShip&"FEDEX_1_DAY_FREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX_1_DAY_FREIGHT")<>"0" AND request.form("HAND-FEDEX_1_DAY_FREIGHT")<>"" then
				If isNumeric(request.form("HAND-FEDEX_1_DAY_FREIGHT"))=true then
					pcStrHandling=pcStrHandling&"FEDEX_1_DAY_FREIGHT|"&replacecomma(request.form("HAND-FEDEX_1_DAY_FREIGHT"))&"|"&request.form("SHFEE-FEDEX_1_DAY_FREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX_1_DAY_FREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX_1_DAY_FREIGHT|"&servicePriority&","

			'FEDEX_2_DAY_FREIGHT
			If request.form("FREE-FEDEX_2_DAY_FREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX_2_DAY_FREIGHT")
				pcStrFreeShip=pcStrFreeShip&"FEDEX_2_DAY_FREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX_2_DAY_FREIGHT")<>"0" AND request.form("HAND-FEDEX_2_DAY_FREIGHT")<>"" then
				If isNumeric(request.form("HAND-FEDEX_2_DAY_FREIGHT"))=true then
					pcStrHandling=pcStrHandling&"FEDEX_2_DAY_FREIGHT|"&replacecomma(request.form("HAND-FEDEX_2_DAY_FREIGHT"))&"|"&request.form("SHFEE-FEDEX_2_DAY_FREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX_2_DAY_FREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX_2_DAY_FREIGHT|"&servicePriority&","

			'FEDEX_3_DAY_FREIGHT
			If request.form("FREE-FEDEX_3_DAY_FREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX_3_DAY_FREIGHT")
				pcStrFreeShip=pcStrFreeShip&"FEDEX_3_DAY_FREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX_3_DAY_FREIGHT")<>"0" AND request.form("HAND-FEDEX_3_DAY_FREIGHT")<>"" then
				If isNumeric(request.form("HAND-FEDEX_3_DAY_FREIGHT"))=true then
					pcStrHandling=pcStrHandling&"FEDEX_3_DAY_FREIGHT|"&replacecomma(request.form("HAND-FEDEX_3_DAY_FREIGHT"))&"|"&request.form("SHFEE-FEDEX_3_DAY_FREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX_3_DAY_FREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX_3_DAY_FREIGHT|"&servicePriority&","

			'FEDEX_GROUND
			If request.form("FREE-FEDEX_GROUND")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX_GROUND")
				pcStrFreeShip=pcStrFreeShip&"FEDEX_GROUND|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX_GROUND")<>"0" AND request.form("HAND-FEDEX_GROUND")<>"" then
				If isNumeric(request.form("HAND-FEDEX_GROUND"))=true then
					pcStrHandling=pcStrHandling&"FEDEX_GROUND|"&replacecomma(request.form("HAND-FEDEX_GROUND"))&"|"&request.form("SHFEE-FEDEX_GROUND")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX_GROUND")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX_GROUND|"&servicePriority&","

			'GROUND_HOME_DELIVERY
			If request.form("FREE-GROUND_HOME_DELIVERY")="YES" then
				pcFreeAmount=request.form("AMT-GROUND_HOME_DELIVERY")
				pcStrFreeShip=pcStrFreeShip&"GROUND_HOME_DELIVERY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-GROUND_HOME_DELIVERY")<>"0" AND request.form("HAND-GROUND_HOME_DELIVERY")<>"" then
				If isNumeric(request.form("HAND-GROUND_HOME_DELIVERY"))=true then
					pcStrHandling=pcStrHandling&"GROUND_HOME_DELIVERY|"&replacecomma(request.form("HAND-GROUND_HOME_DELIVERY"))&"|"&request.form("SHFEE-GROUND_HOME_DELIVERY")&","
				End If
			End if
			servicePriority=request.form("SP-GROUND_HOME_DELIVERY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"GROUND_HOME_DELIVERY|"&servicePriority&","

			'INTERNATIONAL_PRIORITY_FREIGHT
			If request.form("FREE-INTERNATIONAL_PRIORITY_FREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONAL_PRIORITY_FREIGHT")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONAL_PRIORITY_FREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONAL_PRIORITY_FREIGHT")<>"0" AND request.form("HAND-INTERNATIONAL_PRIORITY_FREIGHT")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONAL_PRIORITY_FREIGHT"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONAL_PRIORITY_FREIGHT|"&replacecomma(request.form("HAND-INTERNATIONAL_PRIORITY_FREIGHT"))&"|"&request.form("SHFEE-INTERNATIONAL_PRIORITY_FREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONAL_PRIORITY_FREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONAL_PRIORITY_FREIGHT|"&servicePriority&","

			'INTERNATIONAL_ECONOMY_FREIGHT
			If request.form("FREE-INTERNATIONAL_ECONOMY_FREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONAL_ECONOMY_FREIGHT")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONAL_ECONOMY_FREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONAL_ECONOMY_FREIGHT")<>"0" AND request.form("HAND-INTERNATIONAL_ECONOMY_FREIGHT")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONAL_ECONOMY_FREIGHT"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONAL_ECONOMY_FREIGHT|"&replacecomma(request.form("HAND-INTERNATIONAL_ECONOMY_FREIGHT"))&"|"&request.form("SHFEE-INTERNATIONAL_ECONOMY_FREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONAL_ECONOMY_FREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONAL_ECONOMY_FREIGHT|"&servicePriority&","



			'SMART_POST
			If request.form("FREE-SMART_POST")="YES" then
				pcFreeAmount=request.form("AMT-SMART_POST")
				pcStrFreeShip=pcStrFreeShip&"SMART_POST|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-SMART_POST")<>"0" AND request.form("HAND-SMART_POST")<>"" then
				If isNumeric(request.form("HAND-SMART_POST"))=true then
					pcStrHandling=pcStrHandling&"SMART_POST|"&replacecomma(request.form("HAND-SMART_POST"))&"|"&request.form("SHFEE-SMART_POST")&","
				End If
			End if
			servicePriority=request.form("SP-SMART_POST")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"SMART_POST|"&servicePriority&","




			set rs=Server.CreateObject("ADODB.Recordset")

			query="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=9;"
			set rs=connTemp.execute(query)

			'clear all informatin out of shipService for FedEX
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='PRIORITY_OVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FIRST_OVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='STANDARD_OVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_2_DAY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_EXPRESS_SAVER';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_PRIORITY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_ECONOMY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_FIRST';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_1_DAY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_2_DAY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_3_DAY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX_GROUND';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='GROUND_HOME_DELIVERY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_PRIORITY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONAL_ECONOMY_FREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='SMART_POST';"
			set rs=connTemp.execute(query)


			Dim i
			shipServiceArray=split(pcStrService,", ")

			for i=0 to ubound(shipServiceArray)
				query="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
			next

			pcStrFreeShipArray=split(pcStrFreeShip,",")
			for i=0 to (ubound(pcStrFreeShipArray)-1)
				pcFreeOverAmt=split(pcStrFreeShipArray(i),"|")
				if pcFreeOverAmt(1)>0 then
					pcServiceFree=-1
				else
					pcServiceFree=0
				end if
				query="UPDATE shipService SET serviceFree="&pcServiceFree&",serviceFreeOverAmt="&pcFreeOverAmt(1)&" WHERE serviceCode='"&pcFreeOverAmt(0)&"';"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
			next

			pcStrHandlingArray=split(pcStrHandling,",")
			for i=0 to (ubound(pcStrHandlingArray)-1)
				pcShipHandAmt=split(pcStrHandlingArray(i),"|")
				query="UPDATE shipService SET serviceHandlingFee="&pcShipHandAmt(1)&", serviceShowHandlingFee="&pcShipHandAmt(2)&" WHERE serviceCode='"&pcShipHandAmt(0)&"';"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
			next

				servicePriorityStrArray=split(servicePriorityStr,",")
				for i=0 to (ubound(servicePriorityStrArray)-1)
					SetServicePriority=split(servicePriorityStrArray(i),"|")
					query="UPDATE shipService SET servicePriority="&SetServicePriority(1)&" WHERE serviceCode='"&SetServicePriority(0)&"';"
					set rs=connTemp.execute(query)
				next

			set rs=nothing
			call closedb()
			if session("FedExWSSetUP")="YES" then
				response.redirect "FEDEXWS_EditSettings.asp"
			else
				response.redirect "viewshippingoptions.asp#FedEXWS"
			end if
		else %>


		<% ' START show message, if any %>
			<!--#include file="pcv4_showMessage.asp"-->
		<% 	' END show message %>

			<form name="form1" method="post" action="FedEXWS_EditShipOptions.asp" class="pcForms">
				<table class="pcCPcontent">
					<% query="SELECT serviceCode, serviceActive, servicePriority, serviceDescription,serviceFree,serviceFreeOverAmt,serviceHandlingFee,serviceShowHandlingFee FROM shipService ORDER BY servicePriority, idShipService ASC;"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					pcv_FormString=""
					do until rs.eof
						pServiceCode=rs("serviceCode")
						pServiceActive=rs("serviceActive")
						pServicePriority=rs("servicePriority")
						pServiceDescription=rs("serviceDescription")
						pServiceFree=rs("serviceFree")
						pServiceFreeOverAmt=rs("serviceFreeOverAmt")
						pServiceHandlingFee=rs("serviceHandlingFee")
						pServiceShowHandlingFee =rs("serviceShowHandlingFee")
						if pServiceActive="-1" then
							pServiceCheck="checked"
						else
							pServiceCheck=""
						end if
						if pServiceShowHandlingFee="0" then
							pServiceHandlingFeeChecked="checked"
						else
							pServiceHandlingFeeChecked=""
						end if
						if pServiceFree="-1" then
							pServiceFreeChecked="checked"
						else
							pServiceFreeChecked=""
						end if
						pTempString="<tr bgcolor='#DDEEFF'><td width='4%'><input type='checkbox' name='FEDEXWS_SERVICE' value='XXXX' "&pServiceCheck&"></td><td width='77%'><font color='#000000'><b>"&pServiceDescription&"</b></font></td><td width='19%' align='right'><strong>Order:&nbsp;</strong><input name='SP-XXXX' type='text' id='SP-XXXX' size='2' maxlength='10' value='"&pServicePriority&"'></td></tr>||||||||||<tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='FREE-XXXX' type='checkbox' id='FREE-XXXX' value='YES' "&pServiceFreeChecked&">Offer free shipping for orders over "&scCurSign&" <input name='AMT-XXXX' type='text' id='AMT-XXXX' size='6' maxlength='10' value='"&money(pServiceFreeOverAmt)&"'></td></tr><tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee "&scCurSign&" <input name='HAND-XXXX' type='text' id='HAND-XXXX' size='6' maxlength='10' value='"&money(pServiceHandlingFee)&"'></td></tr><tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='SHFEE-XXXX' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='SHFEE-XXXX' value='0' "&pServiceHandlingFeeChecked&">Integrate into shipping rate.<br><br></td></tr>"

						select case pServiceCode
							case "FIRST_OVERNIGHT"
								pTempString=replace(pTempString,"XXXX","FIRST_OVERNIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "PRIORITY_OVERNIGHT"
								pTempString=replace(pTempString,"XXXX","PRIORITY_OVERNIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "STANDARD_OVERNIGHT"
								pTempString=replace(pTempString,"XXXX","STANDARD_OVERNIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX_2_DAY"
								pTempString=replace(pTempString,"XXXX","FEDEX_2_DAY")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX_EXPRESS_SAVER"
								pTempString=replace(pTempString,"XXXX","FEDEX_EXPRESS_SAVER")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONAL_PRIORITY"
								pTempString=replace(pTempString,"XXXX","INTERNATIONAL_PRIORITY")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONAL_ECONOMY"
								pTempString=replace(pTempString,"XXXX","INTERNATIONAL_ECONOMY")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONAL_FIRST"
								pTempString=replace(pTempString,"XXXX","INTERNATIONAL_FIRST")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX_1_DAY_FREIGHT"
								pTempString=replace(pTempString,"XXXX","FEDEX_1_DAY_FREIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX_2_DAY_FREIGHT"
								pTempString=replace(pTempString,"XXXX","FEDEX_2_DAY_FREIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX_3_DAY_FREIGHT"
								pTempString=replace(pTempString,"XXXX","FEDEX_3_DAY_FREIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX_GROUND"
								pTempString=replace(pTempString,"XXXX","FEDEX_GROUND")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "GROUND_HOME_DELIVERY"
								pTempString=replace(pTempString,"XXXX","GROUND_HOME_DELIVERY")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONAL_PRIORITY_FREIGHT"
								pTempString=replace(pTempString,"XXXX","INTERNATIONAL_PRIORITY_FREIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONAL_ECONOMY_FREIGHT"
								pTempString=replace(pTempString,"XXXX","INTERNATIONAL_ECONOMY_FREIGHT")
								pTempString=replace(pTempString,"||||||||||","")
								pcv_FormString=pcv_FormString&pTempString
							case "SMART_POST"
								pTempString=replace(pTempString,"XXXX","SMART_POST")
								pTempString=replace(pTempString,"||||||||||","<tr><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><font color='#ff0000'>To use SmartPost you must have SmartPost enabled for your FedEx account. Please contact your FedEx Account Representative for more information.</font></td></tr>")
								pcv_FormString=pcv_FormString&pTempString
						end select
						rs.moveNext
					loop
					response.write pcv_FormString
					set rs=nothing
					call closedb()
					%>

					<tr>
						<td colspan="3"><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">FedEx service marks are owned by Federal Express Corporation and used with permission.</div>
	</td>
					</tr>
					<tr>
						<td colspan="3" align="center"><input type="submit" name="Submit" value="Submit" class="submit2"></td>
					</tr>
				</table>
			</form>
			<% end if
			call closedb() %>
<!--#include file="AdminFooter.asp"-->