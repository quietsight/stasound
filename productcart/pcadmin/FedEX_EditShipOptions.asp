<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEX<sup>&reg;</sup> Shipping Configuration - Edit Services" %>
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
<table width="94%" border="0" align="center" cellpadding="4" cellspacing="0">
	<tr>
		<td>
		<% Dim query, rs, connTemp
		call openDb()
		

		if request.querystring("mode")="InAct" then
			' inactivate
			set rs=Server.CreateObject("ADODB.Recordset")
		
			query="UPDATE ShipmentTypes SET active=0 WHERE idShipment=1;"
			set rs=connTemp.execute(query)

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEX"
		end if	
		

		if request.querystring("mode")="Act" then
			' activate
			set rs=Server.CreateObject("ADODB.Recordset")
		
			query="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=1;"
			set rs=connTemp.execute(query)

			call closedb()
			response.redirect "viewshippingoptions.asp#FedEX"
		end if				
		
		
		if request.querystring("mode")="del" then
			'remove
			set rs=Server.CreateObject("ADODB.Recordset")
			'clear all informatin out of shipService for service
		
			query="UPDATE ShipmentTypes SET shipServer='', active=0, international=0 WHERE idShipment=1;"
			set rs=connTemp.execute(query)

			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='PRIORITYOVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FIRSTOVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='STANDARDOVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX2DAY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEXEXPRESSSAVER';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALPRIORITY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALECONOMY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALFIRST';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX1DAYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX2DAYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX3DAYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEXGROUND';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='GROUNDHOMEDELIVERY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALPRIORITYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALECONOMYFREIGHT';"
			set rs=connTemp.execute(query)
			set rs=nothing
			
			call closedb()
			response.redirect "viewshippingoptions.asp#FedEX"
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

			pcStrService=request.form("FEDEX_SERVICE")
			if pcStrService="" then
				response.redirect "FedEX_EditShipOptions.asp?msg="&Server.URLEncode("Select at least one service.")
				response.end
			end if
			pcStrFreeShip=""
			pcStrHandling=""
			servicePriorityStr=""
			
			'PRIORITYOVERNIGHT
			If request.form("FREE-PRIORITYOVERNIGHT")="YES" then
				pcFreeAmount=request.form("AMT-PRIORITYOVERNIGHT")
				pcStrFreeShip=pcStrFreeShip&"PRIORITYOVERNIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-PRIORITYOVERNIGHT")<>"0" AND request.form("HAND-PRIORITYOVERNIGHT")<>"" then
				If isNumeric(request.form("HAND-PRIORITYOVERNIGHT"))=true then
					pcStrHandling=pcStrHandling&"PRIORITYOVERNIGHT|"&replacecomma(request.form("HAND-PRIORITYOVERNIGHT"))&"|"&request.form("SHFEE-PRIORITYOVERNIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-PRIORITYOVERNIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"PRIORITYOVERNIGHT|"&servicePriority&","
				
			'FIRSTOVERNIGHT
			If request.form("FREE-FIRSTOVERNIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FIRSTOVERNIGHT")
				pcStrFreeShip=pcStrFreeShip&"FIRSTOVERNIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FIRSTOVERNIGHT")<>"0" AND request.form("HAND-FIRSTOVERNIGHT")<>"" then
				If isNumeric(request.form("HAND-FIRSTOVERNIGHT"))=true then
					pcStrHandling=pcStrHandling&"FIRSTOVERNIGHT|"&replacecomma(request.form("HAND-FIRSTOVERNIGHT"))&"|"&request.form("SHFEE-FIRSTOVERNIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FIRSTOVERNIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FIRSTOVERNIGHT|"&servicePriority&","
				
			'STANDARDOVERNIGHT
			If request.form("FREE-STANDARDOVERNIGHT")="YES" then
				pcFreeAmount=request.form("AMT-STANDARDOVERNIGHT")
				pcStrFreeShip=pcStrFreeShip&"STANDARDOVERNIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-STANDARDOVERNIGHT")<>"0" AND request.form("HAND-STANDARDOVERNIGHT")<>"" then
				If isNumeric(request.form("HAND-STANDARDOVERNIGHT"))=true then
					pcStrHandling=pcStrHandling&"STANDARDOVERNIGHT|"&replacecomma(request.form("HAND-STANDARDOVERNIGHT"))&"|"&request.form("SHFEE-STANDARDOVERNIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-STANDARDOVERNIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"STANDARDOVERNIGHT|"&servicePriority&","
				
			'FEDEX2DAY
			If request.form("FREE-FEDEX2DAY")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX2DAY")
				pcStrFreeShip=pcStrFreeShip&"FEDEX2DAY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX2DAY")<>"0" AND request.form("HAND-FEDEX2DAY")<>"" then
				If isNumeric(request.form("HAND-FEDEX2DAY"))=true then
					pcStrHandling=pcStrHandling&"FEDEX2DAY|"&replacecomma(request.form("HAND-FEDEX2DAY"))&"|"&request.form("SHFEE-FEDEX2DAY")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX2DAY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX2DAY|"&servicePriority&","
				
			'FEDEXEXPRESSSAVER
			If request.form("FREE-FEDEXEXPRESSSAVER")="YES" then
				pcFreeAmount=request.form("AMT-FEDEXEXPRESSSAVER")
				pcStrFreeShip=pcStrFreeShip&"FEDEXEXPRESSSAVER|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEXEXPRESSSAVER")<>"0" AND request.form("HAND-FEDEXEXPRESSSAVER")<>"" then
				If isNumeric(request.form("HAND-FEDEXEXPRESSSAVER"))=true then
					pcStrHandling=pcStrHandling&"FEDEXEXPRESSSAVER|"&replacecomma(request.form("HAND-FEDEXEXPRESSSAVER"))&"|"&request.form("SHFEE-FEDEXEXPRESSSAVER")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEXEXPRESSSAVER")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEXEXPRESSSAVER|"&servicePriority&","
				
			'INTERNATIONALPRIORITY
			If request.form("FREE-INTERNATIONALPRIORITY")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONALPRIORITY")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONALPRIORITY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONALPRIORITY")<>"0" AND request.form("HAND-INTERNATIONALPRIORITY")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONALPRIORITY"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONALPRIORITY|"&replacecomma(request.form("HAND-INTERNATIONALPRIORITY"))&"|"&request.form("SHFEE-INTERNATIONALPRIORITY")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONALPRIORITY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONALPRIORITY|"&servicePriority&","
				
			'INTERNATIONALECONOMY
			If request.form("FREE-INTERNATIONALECONOMY")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONALECONOMY")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONALECONOMY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONALECONOMY")<>"0" AND request.form("HAND-INTERNATIONALECONOMY")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONALECONOMY"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONALECONOMY|"&replacecomma(request.form("HAND-INTERNATIONALECONOMY"))&"|"&request.form("SHFEE-INTERNATIONALECONOMY")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONALECONOMY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONALECONOMY|"&servicePriority&","
				
			'INTERNATIONALFIRST
			If request.form("FREE-INTERNATIONALFIRST")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONALFIRST")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONALFIRST|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONALFIRST")<>"0" AND request.form("HAND-INTERNATIONALFIRST")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONALFIRST"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONALFIRST|"&replacecomma(request.form("HAND-INTERNATIONALFIRST"))&"|"&request.form("SHFEE-INTERNATIONALFIRST")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONALFIRST")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONALFIRST|"&servicePriority&","
				
			'FEDEX1DAYFREIGHT
			If request.form("FREE-FEDEX1DAYFREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX1DAYFREIGHT")
				pcStrFreeShip=pcStrFreeShip&"FEDEX1DAYFREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX1DAYFREIGHT")<>"0" AND request.form("HAND-FEDEX1DAYFREIGHT")<>"" then
				If isNumeric(request.form("HAND-FEDEX1DAYFREIGHT"))=true then
					pcStrHandling=pcStrHandling&"FEDEX1DAYFREIGHT|"&replacecomma(request.form("HAND-FEDEX1DAYFREIGHT"))&"|"&request.form("SHFEE-FEDEX1DAYFREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX1DAYFREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX1DAYFREIGHT|"&servicePriority&","
				
			'FEDEX2DAYFREIGHT
			If request.form("FREE-FEDEX2DAYFREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX2DAYFREIGHT")
				pcStrFreeShip=pcStrFreeShip&"FEDEX2DAYFREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX2DAYFREIGHT")<>"0" AND request.form("HAND-FEDEX2DAYFREIGHT")<>"" then
				If isNumeric(request.form("HAND-FEDEX2DAYFREIGHT"))=true then
					pcStrHandling=pcStrHandling&"FEDEX2DAYFREIGHT|"&replacecomma(request.form("HAND-FEDEX2DAYFREIGHT"))&"|"&request.form("SHFEE-FEDEX2DAYFREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX2DAYFREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX2DAYFREIGHT|"&servicePriority&","
				
			'FEDEX3DAYFREIGHT
			If request.form("FREE-FEDEX3DAYFREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-FEDEX3DAYFREIGHT")
				pcStrFreeShip=pcStrFreeShip&"FEDEX3DAYFREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEX3DAYFREIGHT")<>"0" AND request.form("HAND-FEDEX3DAYFREIGHT")<>"" then
				If isNumeric(request.form("HAND-FEDEX3DAYFREIGHT"))=true then
					pcStrHandling=pcStrHandling&"FEDEX3DAYFREIGHT|"&replacecomma(request.form("HAND-FEDEX3DAYFREIGHT"))&"|"&request.form("SHFEE-FEDEX3DAYFREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEX3DAYFREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEX3DAYFREIGHT|"&servicePriority&","
				
			'FEDEXGROUND
			If request.form("FREE-FEDEXGROUND")="YES" then
				pcFreeAmount=request.form("AMT-FEDEXGROUND")
				pcStrFreeShip=pcStrFreeShip&"FEDEXGROUND|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-FEDEXGROUND")<>"0" AND request.form("HAND-FEDEXGROUND")<>"" then
				If isNumeric(request.form("HAND-FEDEXGROUND"))=true then
					pcStrHandling=pcStrHandling&"FEDEXGROUND|"&replacecomma(request.form("HAND-FEDEXGROUND"))&"|"&request.form("SHFEE-FEDEXGROUND")&","
				End If
			End if
			servicePriority=request.form("SP-FEDEXGROUND")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"FEDEXGROUND|"&servicePriority&","
				
			'GROUNDHOMEDELIVERY
			If request.form("FREE-GROUNDHOMEDELIVERY")="YES" then
				pcFreeAmount=request.form("AMT-GROUNDHOMEDELIVERY")
				pcStrFreeShip=pcStrFreeShip&"GROUNDHOMEDELIVERY|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-GROUNDHOMEDELIVERY")<>"0" AND request.form("HAND-GROUNDHOMEDELIVERY")<>"" then
				If isNumeric(request.form("HAND-GROUNDHOMEDELIVERY"))=true then
					pcStrHandling=pcStrHandling&"GROUNDHOMEDELIVERY|"&replacecomma(request.form("HAND-GROUNDHOMEDELIVERY"))&"|"&request.form("SHFEE-GROUNDHOMEDELIVERY")&","
				End If
			End if
			servicePriority=request.form("SP-GROUNDHOMEDELIVERY")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"GROUNDHOMEDELIVERY|"&servicePriority&","
				
			'INTERNATIONALPRIORITYFREIGHT
			If request.form("FREE-INTERNATIONALPRIORITYFREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONALPRIORITYFREIGHT")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONALPRIORITYFREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONALPRIORITYFREIGHT")<>"0" AND request.form("HAND-INTERNATIONALPRIORITYFREIGHT")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONALPRIORITYFREIGHT"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONALPRIORITYFREIGHT|"&replacecomma(request.form("HAND-INTERNATIONALPRIORITYFREIGHT"))&"|"&request.form("SHFEE-INTERNATIONALPRIORITYFREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONALPRIORITYFREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONALPRIORITYFREIGHT|"&servicePriority&","
				
			'INTERNATIONALECONOMYFREIGHT
			If request.form("FREE-INTERNATIONALECONOMYFREIGHT")="YES" then
				pcFreeAmount=request.form("AMT-INTERNATIONALECONOMYFREIGHT")
				pcStrFreeShip=pcStrFreeShip&"INTERNATIONALECONOMYFREIGHT|"&replacecomma(pcFreeAmount)&","
			End if
			If request.form("HAND-INTERNATIONALECONOMYFREIGHT")<>"0" AND request.form("HAND-INTERNATIONALECONOMYFREIGHT")<>"" then
				If isNumeric(request.form("HAND-INTERNATIONALECONOMYFREIGHT"))=true then
					pcStrHandling=pcStrHandling&"INTERNATIONALECONOMYFREIGHT|"&replacecomma(request.form("HAND-INTERNATIONALECONOMYFREIGHT"))&"|"&request.form("SHFEE-INTERNATIONALECONOMYFREIGHT")&","
				End If
			End if
			servicePriority=request.form("SP-INTERNATIONALECONOMYFREIGHT")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"INTERNATIONALECONOMYFREIGHT|"&servicePriority&","
				
			set rs=Server.CreateObject("ADODB.Recordset")
			
			query="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=1;"
			set rs=connTemp.execute(query)

			'clear all informatin out of shipService for FedEX
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='PRIORITYOVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FIRSTOVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='STANDARDOVERNIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX2DAY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEXEXPRESSSAVER';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALPRIORITY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALECONOMY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALFIRST';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX1DAYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX2DAYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEX3DAYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='FEDEXGROUND';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='GROUNDHOMEDELIVERY';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALPRIORITYFREIGHT';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='INTERNATIONALECONOMYFREIGHT';"
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
			if session("FedExSetUP")="YES" then
				response.redirect "FEDEX_EditSettings.asp"
			else
				response.redirect "viewshippingoptions.asp#FedEX"
			end if			
		else %>
            <form name="form1" method="post" action="FedEX_EditShipOptions.asp" class="pcForms">
                <table class="pcCPcontent">
                    <% if request.querystring("msg")<>"" then %>
                        <tr class="normal"> 
                            <td colspan="2"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="4">
                                    <tr> 
                                        <td width="4%"><img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"></td>
                                        <td width="96%"><font color="#FF9900"><b><%=request.querystring("msg")%></b></font></td>
                                    </tr>
                                 </table>
                             </td>
                        </tr>
                    <% end if %>
			
                    <tr class="normal"> 
                        <td colspan="2"><span style="font-weight: bold"><br>
                          U.S. Express and Ground Package Services </span></td>
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
						pTempString="<tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='FEDEX_SERVICE' value='XXXX' "&pServiceCheck&"></td><td width='77%'><font color='#000000'><b>"&pServiceDescription&"</b></font></td><td width='19%' align='right'><strong>Order:&nbsp;</strong><input name='SP-XXXX' type='text' id='SP-XXXX' size='2' maxlength='10' value='"&pServicePriority&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='FREE-XXXX' type='checkbox' id='FREE-XXXX' value='YES' "&pServiceFreeChecked&">Offer free shipping for orders over "&scCurSign&" <input name='AMT-XXXX' type='text' id='AMT-XXXX' size='6' maxlength='10' value='"&money(pServiceFreeOverAmt)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee "&scCurSign&" <input name='HAND-XXXX' type='text' id='HAND-XXXX' size='6' maxlength='10' value='"&money(pServiceHandlingFee)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='SHFEE-XXXX' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='SHFEE-XXXX' value='0' "&pServiceHandlingFeeChecked&">Integrate into shipping rate.</td></tr>"
	
						select case pServiceCode
							case "FIRSTOVERNIGHT"
								pTempString=replace(pTempString,"XXXX","FIRSTOVERNIGHT")
								pcv_FormString=pcv_FormString&pTempString
							case "PRIORITYOVERNIGHT"
								pTempString=replace(pTempString,"XXXX","PRIORITYOVERNIGHT")
								pcv_FormString=pcv_FormString&pTempString
							case "STANDARDOVERNIGHT"
								pTempString=replace(pTempString,"XXXX","STANDARDOVERNIGHT")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX2DAY"
								pTempString=replace(pTempString,"XXXX","FEDEX2DAY")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEXEXPRESSSAVER"
								pTempString=replace(pTempString,"XXXX","FEDEXEXPRESSSAVER")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONALPRIORITY"
								pTempString=replace(pTempString,"XXXX","INTERNATIONALPRIORITY")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONALECONOMY"
								pTempString=replace(pTempString,"XXXX","INTERNATIONALECONOMY")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONALFIRST"
								pTempString=replace(pTempString,"XXXX","INTERNATIONALFIRST")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX1DAYFREIGHT"
								pTempString=replace(pTempString,"XXXX","FEDEX1DAYFREIGHT")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX2DAYFREIGHT"
								pTempString=replace(pTempString,"XXXX","FEDEX2DAYFREIGHT")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEX3DAYFREIGHT"
								pTempString=replace(pTempString,"XXXX","FEDEX3DAYFREIGHT")
								pcv_FormString=pcv_FormString&pTempString
							case "FEDEXGROUND"
								pTempString=replace(pTempString,"XXXX","FEDEXGROUND")
								pcv_FormString=pcv_FormString&pTempString
							case "GROUNDHOMEDELIVERY"
								pTempString=replace(pTempString,"XXXX","GROUNDHOMEDELIVERY")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONALPRIORITYFREIGHT"
								pTempString=replace(pTempString,"XXXX","INTERNATIONALPRIORITYFREIGHT")
								pcv_FormString=pcv_FormString&pTempString
							case "INTERNATIONALECONOMYFREIGHT"
								pTempString=replace(pTempString,"XXXX","INTERNATIONALECONOMYFREIGHT")
								pcv_FormString=pcv_FormString&pTempString
						end select
						rs.moveNext
					loop 
					response.write pcv_FormString      
					set rs=nothing
					call closedb()
					%>
			
                    <tr class="normal"> 
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr class="normal"> 
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