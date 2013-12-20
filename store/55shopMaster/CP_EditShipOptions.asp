<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit Canada Post Shipping Services" %>
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
<table class="pcCPcontent">
	<tr>
		<td>
			<% 
			Dim query, rs, connTemp
			call openDb()


			if trim(lcase(request.querystring("mode")))="inact" then
				'set inactive
				query="UPDATE ShipmentTypes SET active=0 WHERE idShipment=7;"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
				response.redirect "viewshippingoptions.asp#CP"
			end if

			if trim(lcase(request.querystring("mode")))="act" then
				'set active
				query="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=7;"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
				response.redirect "viewshippingoptions.asp#CP"
			end if


			if trim(lcase(request.querystring("mode")))="del" then
				'remove
				set rs=Server.CreateObject("ADODB.Recordset")
				'clear all informatin out of shipService for UPS
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1010';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1020';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1030';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1040';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1120';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1130';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1220';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1230';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2005';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2010';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2015';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2020';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2025';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2030';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2040';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2050';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3005';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3010';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3015';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3020';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3025';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3040';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3050';"
				set rs=connTemp.execute(query)
				'set inactive
				query="UPDATE ShipmentTypes SET active=0, international=0 WHERE idShipment=7;"
				set rs=connTemp.execute(query)
				
				set rs=nothing

				call closedb()
				response.redirect "viewshippingoptions.asp#CP"
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
				ServiceStr=request.form("CP_Service")
				if ServiceStr="" then
					response.redirect "CP_EditShipOptions.asp?msg="&Server.URLEncode("Select at least one service.")
					response.end
				end if
				freeshipStr=""
				handlingStr=""
				servicePriorityStr=""
				
				If request.form("free1010")="YES" then
					freeamt=request.form("amt1010")
					freeshipStr=freeshipStr&"1010|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1010")<>"0" AND request.form("handling1010")<>"" then
					If isNumeric(request.form("handling1010"))=true then
						handlingStr=handlingStr&"1010|"&replacecomma(request.form("handling1010"))&"|"&request.form("shfee1010")&","
					End If
				End if
				servicePriority=request.form("servicePriority1010")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1010|"&servicePriority&","
				
				If request.form("free1020")="YES" then
					freeamt=request.form("amt1020")
					freeshipStr=freeshipStr&"1020|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1020")<>"0" AND request.form("handling1020")<>"" then
					If isNumeric(request.form("handling1020"))=true then
						handlingStr=handlingStr&"1020|"&replacecomma(request.form("handling1020"))&"|"&request.form("shfee1020")&","
					End If
				End if
				servicePriority=request.form("servicePriority1020")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1020|"&servicePriority&","

				If request.form("free1030")="YES" then
					freeamt=request.form("amt1030")
					freeshipStr=freeshipStr&"1030|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1030")<>"0" AND request.form("handling1030")<>"" then
					If isNumeric(request.form("handling1030"))=true then
						handlingStr=handlingStr&"1030|"&replacecomma(request.form("handling1030"))&"|"&request.form("shfee1030")&","
					End If
				End if
				servicePriority=request.form("servicePriority1030")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1030|"&servicePriority&","

				If request.form("free1040")="YES" then
					freeamt=request.form("amt1040")
					freeshipStr=freeshipStr&"1040|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1040")<>"0" AND request.form("handling1040")<>"" then
					If isNumeric(request.form("handling1040"))=true then
						handlingStr=handlingStr&"1040|"&replacecomma(request.form("handling1040"))&"|"&request.form("shfee1040")&","
					End If
				End if
				servicePriority=request.form("servicePriority1040")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1040|"&servicePriority&","

				If request.form("free1120")="YES" then
					freeamt=request.form("amt1120")
					freeshipStr=freeshipStr&"1120|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1120")<>"0" AND request.form("handling1120")<>"" then
					If isNumeric(request.form("handling1120"))=true then
						handlingStr=handlingStr&"1120|"&replacecomma(request.form("handling1120"))&"|"&request.form("shfee1120")&","
					End If
				End if
				servicePriority=request.form("servicePriority1120")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1120|"&servicePriority&","

				If request.form("free1130")="YES" then
					freeamt=request.form("amt1130")
					freeshipStr=freeshipStr&"1130|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1130")<>"0" AND request.form("handling1130")<>"" then
					If isNumeric(request.form("handling1130"))=true then
						handlingStr=handlingStr&"1130|"&replacecomma(request.form("handling1130"))&"|"&request.form("shfee1130")&","
					End If
				End if
				servicePriority=request.form("servicePriority1130")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1130|"&servicePriority&","

				If request.form("free1220")="YES" then
					freeamt=request.form("amt1220")
					freeshipStr=freeshipStr&"1220|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1220")<>"0" AND request.form("handling1220")<>"" then
					If isNumeric(request.form("handling1220"))=true then
						handlingStr=handlingStr&"1220|"&replacecomma(request.form("handling1220"))&"|"&request.form("shfee1220")&","
					End If
				End if
				servicePriority=request.form("servicePriority1220")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1220|"&servicePriority&","

				If request.form("free1230")="YES" then
					freeamt=request.form("amt1230")
					freeshipStr=freeshipStr&"1230|"&replacecomma(freeamt)&","
				End if
				If request.form("handling1230")<>"0" AND request.form("handling1230")<>"" then
					If isNumeric(request.form("handling1230"))=true then
						handlingStr=handlingStr&"1230|"&replacecomma(request.form("handling1230"))&"|"&request.form("shfee1230")&","
					End If
				End if
				servicePriority=request.form("servicePriority1230")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"1230|"&servicePriority&","

				If request.form("free2005")="YES" then
					freeamt=request.form("amt2005")
					freeshipStr=freeshipStr&"2005|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2005")<>"0" AND request.form("handling2005")<>"" then
					If isNumeric(request.form("handling2005"))=true then
						handlingStr=handlingStr&"2005|"&replacecomma(request.form("handling2005"))&"|"&request.form("shfee2005")&","
					End If
				End if
				servicePriority=request.form("servicePriority2005")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2005|"&servicePriority&","

				If request.form("free2010")="YES" then
					freeamt=request.form("amt2010")
					freeshipStr=freeshipStr&"2010|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2010")<>"0" AND request.form("handling2010")<>"" then
					If isNumeric(request.form("handling2010"))=true then
						handlingStr=handlingStr&"2010|"&replacecomma(request.form("handling2010"))&"|"&request.form("shfee2010")&","
					End If
				End if
				servicePriority=request.form("servicePriority2010")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2010|"&servicePriority&","

				If request.form("free2015")="YES" then
					freeamt=request.form("amt2015")
					freeshipStr=freeshipStr&"2015|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2015")<>"0" AND request.form("handling2015")<>"" then
					If isNumeric(request.form("handling2015"))=true then
						handlingStr=handlingStr&"2015|"&replacecomma(request.form("handling2015"))&"|"&request.form("shfee2015")&","
					End If
				End if
				servicePriority=request.form("servicePriority2015")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2015|"&servicePriority&","

				If request.form("free2020")="YES" then
					freeamt=request.form("amt2020")
					freeshipStr=freeshipStr&"2020|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2020")<>"0" AND request.form("handling2020")<>"" then
					If isNumeric(request.form("handling2020"))=true then
						handlingStr=handlingStr&"2020|"&replacecomma(request.form("handling2020"))&"|"&request.form("shfee2020")&","
					End If
				End if
				servicePriority=request.form("servicePriority2020")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2020|"&servicePriority&","
			
				If request.form("free2025")="YES" then
					freeamt=request.form("amt2025")
					freeshipStr=freeshipStr&"2025|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2025")<>"0" AND request.form("handling2025")<>"" then
					If isNumeric(request.form("handling2025"))=true then
						handlingStr=handlingStr&"2025|"&replacecomma(request.form("handling2025"))&"|"&request.form("shfee2025")&","
					End If
				End if
				servicePriority=request.form("servicePriority2025")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2025|"&servicePriority&","
			
				If request.form("free2030")="YES" then
					freeamt=request.form("amt2030")
					freeshipStr=freeshipStr&"2030|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2030")<>"0" AND request.form("handling2030")<>"" then
					If isNumeric(request.form("handling2030"))=true then
						handlingStr=handlingStr&"2030|"&replacecomma(request.form("handling2030"))&"|"&request.form("shfee2030")&","
					End If
				End if
				servicePriority=request.form("servicePriority2030")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2030|"&servicePriority&","
			
				If request.form("free2040")="YES" then
					freeamt=request.form("amt2040")
					freeshipStr=freeshipStr&"2040|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2040")<>"0" AND request.form("handling2040")<>"" then
					If isNumeric(request.form("handling2040"))=true then
						handlingStr=handlingStr&"2040|"&replacecomma(request.form("handling2040"))&"|"&request.form("shfee2040")&","
					End If
				End if
				servicePriority=request.form("servicePriority2040")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2040|"&servicePriority&","
			
				If request.form("free2050")="YES" then
					freeamt=request.form("amt2050")
					freeshipStr=freeshipStr&"2050|"&replacecomma(freeamt)&","
				End if
				If request.form("handling2050")<>"0" AND request.form("handling2050")<>"" then
					If isNumeric(request.form("handling2050"))=true then
						handlingStr=handlingStr&"2050|"&replacecomma(request.form("handling2050"))&"|"&request.form("shfee2050")&","
					End If
				End if
				servicePriority=request.form("servicePriority2050")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"2050|"&servicePriority&","

				If request.form("free3005")="YES" then
					freeamt=request.form("amt3005")
					freeshipStr=freeshipStr&"3005|"&replacecomma(freeamt)&","
				End if
				If request.form("handling3005")<>"0" AND request.form("handling3005")<>"" then
					If isNumeric(request.form("handling3005"))=true then
						handlingStr=handlingStr&"3005|"&replacecomma(request.form("handling3005"))&"|"&request.form("shfee3005")&","
					End If
				End if
				servicePriority=request.form("servicePriority3005")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"3005|"&servicePriority&","

				If request.form("free3010")="YES" then
					freeamt=request.form("amt3010")
					freeshipStr=freeshipStr&"3010|"&replacecomma(freeamt)&","
				End if
				If request.form("handling3010")<>"0" AND request.form("handling3010")<>"" then
					If isNumeric(request.form("handling3010"))=true then
						handlingStr=handlingStr&"3010|"&replacecomma(request.form("handling3010"))&"|"&request.form("shfee3010")&","
					End If
				End if
				servicePriority=request.form("servicePriority3010")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"3010|"&servicePriority&","

				If request.form("free3015")="YES" then
					freeamt=request.form("amt3015")
					freeshipStr=freeshipStr&"3015|"&replacecomma(freeamt)&","
				End if
				If request.form("handling3015")<>"0" AND request.form("handling3015")<>"" then
					If isNumeric(request.form("handling3015"))=true then
						handlingStr=handlingStr&"3015|"&replacecomma(request.form("handling3015"))&"|"&request.form("shfee3015")&","
					End If
				End if
				servicePriority=request.form("servicePriority3015")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"3015|"&servicePriority&","

				If request.form("free3020")="YES" then
					freeamt=request.form("amt3020")
					freeshipStr=freeshipStr&"3020|"&replacecomma(freeamt)&","
				End if
				If request.form("handling3020")<>"0" AND request.form("handling3020")<>"" then
					If isNumeric(request.form("handling3020"))=true then
						handlingStr=handlingStr&"3020|"&replacecomma(request.form("handling3020"))&"|"&request.form("shfee3020")&","
					End If
				End if
				servicePriority=request.form("servicePriority3020")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"3020|"&servicePriority&","

				If request.form("free3025")="YES" then
					freeamt=request.form("amt3025")
					freeshipStr=freeshipStr&"3025|"&replacecomma(freeamt)&","
				End if
				If request.form("handling3025")<>"0" AND request.form("handling3025")<>"" then
					If isNumeric(request.form("handling3025"))=true then
						handlingStr=handlingStr&"3025|"&replacecomma(request.form("handling3025"))&"|"&request.form("shfee3025")&","
					End If
				End if
				servicePriority=request.form("servicePriority3025")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"3025|"&servicePriority&","

				If request.form("free3040")="YES" then
					freeamt=request.form("amt3040")
					freeshipStr=freeshipStr&"3040|"&replacecomma(freeamt)&","
				End if
				If request.form("handling3040")<>"0" AND request.form("handling3040")<>"" then
					If isNumeric(request.form("handling3040"))=true then
						handlingStr=handlingStr&"3040|"&replacecomma(request.form("handling3040"))&"|"&request.form("shfee3040")&","
					End If
				End if
				servicePriority=request.form("servicePriority3040")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"3040|"&servicePriority&","

				If request.form("free3050")="YES" then
					freeamt=request.form("amt3050")
					freeshipStr=freeshipStr&"3050|"&replacecomma(freeamt)&","
				End if
				If request.form("handling3050")<>"0" AND request.form("handling3050")<>"" then
					If isNumeric(request.form("handling3050"))=true then
						handlingStr=handlingStr&"3050|"&replacecomma(request.form("handling3050"))&"|"&request.form("shfee3050")&","
					End If
				End if
				servicePriority=request.form("servicePriority3050")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"3050|"&servicePriority&","


				set rs=Server.CreateObject("ADODB.Recordset")
				'clear all informatin out of shipService for USPS
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1010';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1020';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1030';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1040';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1120';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1130';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1220';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='1230';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2005';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2010';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2015';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2020';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2025';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2030';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2040';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='2050';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3005';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3010';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3015';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3020';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3025';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3040';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='3050';"
				set rs=connTemp.execute(query)
				Dim i
				shipServiceArray=split(ServiceStr,", ")
				for i=0 to ubound(shipServiceArray)
					query="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
				set rs=Server.CreateObject("ADODB.Recordset")
					set rs=connTemp.execute(query)
				next
			
				freeshipStrArray=split(freeshipStr,",")
				for i=0 to (ubound(freeshipStrArray)-1)
					freeoveramt=split(freeshipStrArray(i),"|")
					if freeoveramt(1)>0 then
						serviceFree=-1
					else
						serviceFree=0
					end if
					query="UPDATE shipService SET serviceFree="&serviceFree&",serviceFreeOverAmt="&freeoveramt(1)&" WHERE serviceCode='"&freeoveramt(0)&"';"
					response.write query
					set rs=connTemp.execute(query)
				next
			
				handlingStrArray=split(handlingStr,",")
				for i=0 to (ubound(handlingStrArray)-1)
					shiphandamt=split(handlingStrArray(i),"|")
					query="UPDATE shipService SET serviceHandlingFee="&shiphandamt(1)&", serviceShowHandlingFee="&shiphandamt(2)&" WHERE serviceCode='"&shiphandamt(0)&"';"
					'response.write query
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
				response.redirect "viewshippingoptions.asp#CP"			
			else %>
				<form name="form1" method="post" action="CP_EditShipOptions.asp" class="pcForms">
                    <table class="pcCPcontent">
                        <tr>
                            <td colspan="2" class="pcCPspacer">
                                <% ' START show message, if any %>
                                    <!--#include file="pcv4_showMessage.asp"-->
                                <% 	' END show message %>
                            </td>
                        </tr>
                        
						<% query="SELECT serviceCode, serviceActive, servicePriority, serviceDescription,serviceFree,serviceFreeOverAmt,serviceHandlingFee,serviceShowHandlingFee FROM shipService ORDER BY serviceActive, servicePriority;"
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
							pTempString="<tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='CP_Service' value='XXXX' "&pServiceCheck&"></td><td width='77%'><font color='#000000'><b>"&pServiceDescription&"</b></font></td><td width='19%' align='right'><strong>Order:&nbsp;</strong><input name='servicePriorityXXXX' type='text' id='servicePriorityXXXX' size='2' maxlength='10' value='"&pServicePriority&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='freeXXXX' type='checkbox' id='freeXXXX' value='YES' "&pServiceFreeChecked&">Offer free shipping for orders over "&scCurSign&" <input name='amtXXXX' type='text' id='amtXXXX' size='6' maxlength='10' value='"&money(pServiceFreeOverAmt)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee "&scCurSign&" <input name='handlingXXXX' type='text' id='handlingXXXX' size='6' maxlength='10' value='"&money(pServiceHandlingFee)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfeeXXXX' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfeeXXXX' value='0' "&pServiceHandlingFeeChecked&">Integrate into shipping rate.</td></tr>"

							select case pServiceCode
								case "1010"
									pTempString=replace(pTempString,"XXXX","1010")
									pcv_FormString=pcv_FormString&pTempString
								case "1020"
									pTempString=replace(pTempString,"XXXX","1020")
									pcv_FormString=pcv_FormString&pTempString
								case "1030"
									pTempString=replace(pTempString,"XXXX","1030")
									pcv_FormString=pcv_FormString&pTempString
								case "1040"
									pTempString=replace(pTempString,"XXXX","1040")
									pcv_FormString=pcv_FormString&pTempString
								case "1120"
									pTempString=replace(pTempString,"XXXX","1120")
									pcv_FormString=pcv_FormString&pTempString
								case "1130"
									pTempString=replace(pTempString,"XXXX","1130")
									pcv_FormString=pcv_FormString&pTempString
								case "1220"
									pTempString=replace(pTempString,"XXXX","1220")
									pcv_FormString=pcv_FormString&pTempString
								case "1230"
									pTempString=replace(pTempString,"XXXX","1230")
									pcv_FormString=pcv_FormString&pTempString
								case "2005"
									pTempString=replace(pTempString,"XXXX","2005")
									pcv_FormString=pcv_FormString&pTempString
								case "2010"
									pTempString=replace(pTempString,"XXXX","2010")
									pcv_FormString=pcv_FormString&pTempString
								case "2015"
									pTempString=replace(pTempString,"XXXX","2015")
									pcv_FormString=pcv_FormString&pTempString
								case "2020"
									pTempString=replace(pTempString,"XXXX","2020")
									pcv_FormString=pcv_FormString&pTempString
								case "2025"
									pTempString=replace(pTempString,"XXXX","2025")
									pcv_FormString=pcv_FormString&pTempString
								case "2030"
									pTempString=replace(pTempString,"XXXX","2030")
									pcv_FormString=pcv_FormString&pTempString
								case "2040"
									pTempString=replace(pTempString,"XXXX","2040")
									pcv_FormString=pcv_FormString&pTempString
								case "2050"
									pTempString=replace(pTempString,"XXXX","2050")
									pcv_FormString=pcv_FormString&pTempString
								case "3005"
									pTempString=replace(pTempString,"XXXX","3005")
									pcv_FormString=pcv_FormString&pTempString
								case "3010"
									pTempString=replace(pTempString,"XXXX","3010")
									pcv_FormString=pcv_FormString&pTempString
								case "3015"
									pTempString=replace(pTempString,"XXXX","3015")
									pcv_FormString=pcv_FormString&pTempString
								case "3020"
									pTempString=replace(pTempString,"XXXX","3020")
									pcv_FormString=pcv_FormString&pTempString
								case "3025"
									pTempString=replace(pTempString,"XXXX","3025")
									pcv_FormString=pcv_FormString&pTempString
								case "3040"
									pTempString=replace(pTempString,"XXXX","3040")
									pcv_FormString=pcv_FormString&pTempString
								case "3050"
									pTempString=replace(pTempString,"XXXX","3050")
									pcv_FormString=pcv_FormString&pTempString
							end select
							rs.moveNext
						loop 
						response.write pcv_FormString      
						set rs=nothing
						call closedb()
						%>
                        <tr>
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
					
						<tr> 
							<td colspan="2" align="center">
							<input type="submit" name="Submit" value="Submit" class="ibtnGrey"></td>
						</tr>
					</table>
  		  		</form>
			<% end if 
			call closedb() %>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->