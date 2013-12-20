<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit USPS Shipping Services" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="AdminHeader.asp"-->
<table width="94%" border="0" align="center" cellpadding="4" cellspacing="0">
	<tr>
		<td>
			<% Dim query, rs, connTemp
			call openDb()
			
			if request.querystring("mode")="InAct" then
				'set inactive
				query="UPDATE ShipmentTypes SET active=0 WHERE idShipment=4;"
				set rs=connTemp.execute(query)
			set rs=nothing
			
			call closedb()
				response.redirect "viewshippingoptions.asp#USPS"
			end if			
		
					
			if request.querystring("mode")="Act" then
				'set active
				query="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=4;"
				set rs=connTemp.execute(query)
			set rs=nothing
			
			call closedb()
				response.redirect "viewshippingoptions.asp#USPS"
			end if						
				
					
			if request.querystring("mode")="del" then
				'remove
				set rs=Server.CreateObject("ADODB.Recordset")
				'clear all informatin out of shipService for UPS
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9901';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9902';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9903';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9904';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9905';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9906';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9907';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9908';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9909';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9910';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9911';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9912';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9913';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9914';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9915';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9916';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9917';"
				set rs=connTemp.execute(query)
				'set inactive
				query="UPDATE ShipmentTypes SET active=0, international=0 WHERE idShipment=4;"
				set rs=connTemp.execute(query)
			set rs=nothing
			
			call closedb()
				response.redirect "viewshippingoptions.asp#USPS"
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
				ServiceStr=request.form("USPS_Service")
				if ServiceStr="" then
					response.redirect "USPS_EditShipOptions.asp?msg="&Server.URLEncode("Select at least one service.")
					response.end
				end if
				freeshipStr=""
				handlingStr=""
				servicePriorityStr=""
				
				If request.form("free9901")="YES" then
					freeamt=request.form("amt9901")
					freeshipStr=freeshipStr&"9901|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9901")<>"0" AND request.form("handling9901")<>"" then
					If isNumeric(request.form("handling9901"))=true then
						handlingStr=handlingStr&"9901|"&replacecomma(request.form("handling9901"))&"|"&request.form("shfee9901")&","
					End If
				End if
				servicePriority=request.form("servicePriority9901")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9901|"&servicePriority&","
				
				If request.form("free9902")="YES" then
					freeamt=request.form("amt9902")
					freeshipStr=freeshipStr&"9902|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9902")<>"0" AND request.form("handling9902")<>"" then
					If isNumeric(request.form("handling9902"))=true then
						handlingStr=handlingStr&"9902|"&replacecomma(request.form("handling9902"))&"|"&request.form("shfee9902")&","
					End If
				End if
				servicePriority=request.form("servicePriority9902")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9902|"&servicePriority&","

				If request.form("free9903")="YES" then
					freeamt=request.form("amt9903")
					freeshipStr=freeshipStr&"9903|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9903")<>"0" AND request.form("handling9903")<>"" then
					If isNumeric(request.form("handling9903"))=true then
						handlingStr=handlingStr&"9903|"&replacecomma(request.form("handling9903"))&"|"&request.form("shfee9903")&","
					End If
				End if
				servicePriority=request.form("servicePriority9903")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9903|"&servicePriority&","

				If request.form("free9904")="YES" then
					freeamt=request.form("amt9904")
					freeshipStr=freeshipStr&"9904|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9904")<>"0" AND request.form("handling9904")<>"" then
					If isNumeric(request.form("handling9904"))=true then
						handlingStr=handlingStr&"9904|"&replacecomma(request.form("handling9904"))&"|"&request.form("shfee9904")&","
					End If
				End if
				servicePriority=request.form("servicePriority9904")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9904|"&servicePriority&","

				If request.form("free9905")="YES" then
					freeamt=request.form("amt9905")
					freeshipStr=freeshipStr&"9905|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9905")<>"0" AND request.form("handling9905")<>"" then
					If isNumeric(request.form("handling9905"))=true then
						handlingStr=handlingStr&"9905|"&replacecomma(request.form("handling9905"))&"|"&request.form("shfee9905")&","
					End If
				End if
				servicePriority=request.form("servicePriority9905")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9905|"&servicePriority&","

				If request.form("free9906")="YES" then
					freeamt=request.form("amt9906")
					freeshipStr=freeshipStr&"9906|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9906")<>"0" AND request.form("handling9906")<>"" then
					If isNumeric(request.form("handling9906"))=true then
						handlingStr=handlingStr&"9906|"&replacecomma(request.form("handling9906"))&"|"&request.form("shfee9906")&","
					End If
				End if
				servicePriority=request.form("servicePriority9906")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9906|"&servicePriority&","

				If request.form("free9907")="YES" then
					freeamt=request.form("amt9907")
					freeshipStr=freeshipStr&"9907|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9907")<>"0" AND request.form("handling9907")<>"" then
					If isNumeric(request.form("handling9907"))=true then
						handlingStr=handlingStr&"9907|"&replacecomma(request.form("handling9907"))&"|"&request.form("shfee9907")&","
					End If
				End if
				servicePriority=request.form("servicePriority9907")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9907|"&servicePriority&","

				If request.form("free9908")="YES" then
					freeamt=request.form("amt9908")
					freeshipStr=freeshipStr&"9908|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9908")<>"0" AND request.form("handling9908")<>"" then
					If isNumeric(request.form("handling9908"))=true then
						handlingStr=handlingStr&"9908|"&replacecomma(request.form("handling9908"))&"|"&request.form("shfee9908")&","
					End If
				End if
				servicePriority=request.form("servicePriority9908")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9908|"&servicePriority&","

				If request.form("free9909")="YES" then
					freeamt=request.form("amt9909")
					freeshipStr=freeshipStr&"9909|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9909")<>"0" AND request.form("handling9909")<>"" then
					If isNumeric(request.form("handling9909"))=true then
						handlingStr=handlingStr&"9909|"&replacecomma(request.form("handling9909"))&"|"&request.form("shfee9909")&","
					End If
				End if
				servicePriority=request.form("servicePriority9909")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9909|"&servicePriority&","

				If request.form("free9910")="YES" then
					freeamt=request.form("amt9910")
					freeshipStr=freeshipStr&"9910|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9910")<>"0" AND request.form("handling9910")<>"" then
					If isNumeric(request.form("handling9910"))=true then
						handlingStr=handlingStr&"9910|"&replacecomma(request.form("handling9910"))&"|"&request.form("shfee9910")&","
					End If
				End if
				servicePriority=request.form("servicePriority9910")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9910|"&servicePriority&","

				If request.form("free9911")="YES" then
					freeamt=request.form("amt9911")
					freeshipStr=freeshipStr&"9911|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9911")<>"0" AND request.form("handling9911")<>"" then
					If isNumeric(request.form("handling9911"))=true then
						handlingStr=handlingStr&"9911|"&replacecomma(request.form("handling9911"))&"|"&request.form("shfee9911")&","
					End If
				End if
				servicePriority=request.form("servicePriority9911")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9911|"&servicePriority&","

				If request.form("free9912")="YES" then
					freeamt=request.form("amt9912")
					freeshipStr=freeshipStr&"9912|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9912")<>"0" AND request.form("handling9912")<>"" then
					If isNumeric(request.form("handling9912"))=true then
						handlingStr=handlingStr&"9912|"&replacecomma(request.form("handling9912"))&"|"&request.form("shfee9912")&","
					End If
				End if
				servicePriority=request.form("servicePriority9912")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9912|"&servicePriority&","
			
				If request.form("free9913")="YES" then
					freeamt=request.form("amt9913")
					freeshipStr=freeshipStr&"9913|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9913")<>"0" AND request.form("handling9913")<>"" then
					If isNumeric(request.form("handling9913"))=true then
						handlingStr=handlingStr&"9913|"&replacecomma(request.form("handling9913"))&"|"&request.form("shfee9913")&","
					End If
				End if
				servicePriority=request.form("servicePriority9913")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9913|"&servicePriority&","
			
				If request.form("free9914")="YES" then
					freeamt=request.form("amt9914")
					freeshipStr=freeshipStr&"9914|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9914")<>"0" AND request.form("handling9914")<>"" then
					If isNumeric(request.form("handling9914"))=true then
						handlingStr=handlingStr&"9914|"&replacecomma(request.form("handling9914"))&"|"&request.form("shfee9914")&","
					End If
				End if
				servicePriority=request.form("servicePriority9914")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9914|"&servicePriority&","
			
				If request.form("free9915")="YES" then
					freeamt=request.form("amt9915")
					freeshipStr=freeshipStr&"9915|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9915")<>"0" AND request.form("handling9915")<>"" then
					If isNumeric(request.form("handling9915"))=true then
						handlingStr=handlingStr&"9915|"&replacecomma(request.form("handling9915"))&"|"&request.form("shfee9915")&","
					End If
				End if
				servicePriority=request.form("servicePriority9915")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9915|"&servicePriority&","
			
				If request.form("free9916")="YES" then
					freeamt=request.form("amt9916")
					freeshipStr=freeshipStr&"9916|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9916")<>"0" AND request.form("handling9916")<>"" then
					If isNumeric(request.form("handling9916"))=true then
						handlingStr=handlingStr&"9916|"&replacecomma(request.form("handling9916"))&"|"&request.form("shfee9916")&","
					End If
				End if
				servicePriority=request.form("servicePriority9916")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9916|"&servicePriority&","

				If request.form("free9917")="YES" then
					freeamt=request.form("amt9917")
					freeshipStr=freeshipStr&"9917|"&replacecomma(freeamt)&","
				End if
				If request.form("handling9917")<>"0" AND request.form("handling9917")<>"" then
					If isNumeric(request.form("handling9917"))=true then
						handlingStr=handlingStr&"9917|"&replacecomma(request.form("handling9917"))&"|"&request.form("shfee9917")&","
					End If
				End if
				servicePriority=request.form("servicePriority9917")
				If NOT validNum2(servicePriority) then
					servicePriority="0"
				End if
				servicePriorityStr=servicePriorityStr&"9917|"&servicePriority&","

				set rs=Server.CreateObject("ADODB.Recordset")
				'clear all informatin out of shipService for USPS
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9901';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9902';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9903';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9904';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9905';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9906';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9907';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9908';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9909';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9910';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9911';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9912';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9913';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9914';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9915';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9916';"
				set rs=connTemp.execute(query)
				query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9917';"
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
				response.redirect "viewshippingoptions.asp#USPS"			
			else %>
				<form name="form1" method="post" action="USPS_EditShipOptions.asp">
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
                        <tr bgcolor="#FFFFFF" class="normal"> 
                            <td colspan="2">&nbsp;</td>
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
							pTempString="<tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='XXXX' "&pServiceCheck&"></td><td width='77%'><font color='#000000'><b>"&pServiceDescription&"</b></font></td><td width='19%' align='right'><strong>Order:&nbsp;</strong><input name='servicePriorityXXXX' type='text' id='servicePriorityXXXX' size='2' maxlength='10' value='"&pServicePriority&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='freeXXXX' type='checkbox' id='freeXXXX' value='YES' "&pServiceFreeChecked&">Offer free shipping for orders over "&scCurSign&" <input name='amtXXXX' type='text' id='amtXXXX' size='6' maxlength='10' value='"&money(pServiceFreeOverAmt)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee "&scCurSign&" <input name='handlingXXXX' type='text' id='handlingXXXX' size='6' maxlength='10' value='"&money(pServiceHandlingFee)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfeeXXXX' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfeeXXXX' value='0' "&pServiceHandlingFeeChecked&">Integrate into shipping rate.</td></tr>"

							select case pServiceCode
								case "9901"
									pTempString=replace(pTempString,"XXXX","9901")
									pcv_FormString=pcv_FormString&pTempString
								case "9902"
									pTempString=replace(pTempString,"XXXX","9902")
									pcv_FormString=pcv_FormString&pTempString
								case "9903"
									pTempString=replace(pTempString,"XXXX","9903")
									pcv_FormString=pcv_FormString&pTempString
								case "9904"
									pTempString=replace(pTempString,"XXXX","9904")
									pcv_FormString=pcv_FormString&pTempString
								case "9905"
									pTempString=replace(pTempString,"XXXX","9905")
									pcv_FormString=pcv_FormString&pTempString
								case "9906"
									pTempString=replace(pTempString,"XXXX","9906")
									pcv_FormString=pcv_FormString&pTempString
								case "9907"
									pTempString=replace(pTempString,"XXXX","9907")
									pcv_FormString=pcv_FormString&pTempString
								case "9908"
									pTempString=replace(pTempString,"XXXX","9908")
									pcv_FormString=pcv_FormString&pTempString
								case "9909"
									pTempString=replace(pTempString,"XXXX","9909")
									pcv_FormString=pcv_FormString&pTempString
								case "9910"
									pTempString=replace(pTempString,"XXXX","9910")
									pcv_FormString=pcv_FormString&pTempString
								case "9911"
									pTempString=replace(pTempString,"XXXX","9911")
									pcv_FormString=pcv_FormString&pTempString
								case "9912"
									pTempString=replace(pTempString,"XXXX","9912")
									pcv_FormString=pcv_FormString&pTempString
								case "9913"
									pTempString=replace(pTempString,"XXXX","9913")
									pcv_FormString=pcv_FormString&pTempString
								case "9914"
									pTempString=replace(pTempString,"XXXX","9914")
									pcv_FormString=pcv_FormString&pTempString
								case "9915"
									pTempString=replace(pTempString,"XXXX","9915")
									pcv_FormString=pcv_FormString&pTempString
								case "9916"
									pTempString=replace(pTempString,"XXXX","9916")
									pcv_FormString=pcv_FormString&pTempString
								case "9917"
									pTempString=replace(pTempString,"XXXX","9917")
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