<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="UPS OnLine&reg; Tools Shipping Configuration - Edit Services" %>
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
		<% Dim query, rs, connTemp
		call openDb()
		
		if request.querystring("mode")="InAct" then
            'set inactive	
			set rs=Server.CreateObject("ADODB.Recordset")
			query="UPDATE ShipmentTypes SET active=0 WHERE idShipment=3;"
			set rs=connTemp.execute(query)
			response.redirect "viewshippingoptions.asp#UPS"
		end if		
		
		if request.querystring("mode")="Act" then
            'set active	
			set rs=Server.CreateObject("ADODB.Recordset")
			query="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=3;"
			set rs=connTemp.execute(query)
			response.redirect "viewshippingoptions.asp#UPS"
		end if			
		
		if request.querystring("mode")="del" then
			'remove
			set rs=Server.CreateObject("ADODB.Recordset")
			'clear all informatin out of shipService for UPS
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='01';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='02';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='03';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='07';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='08';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='11';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='12';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='13';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='14';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='54';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='59';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='65';"
			set rs=connTemp.execute(query)
			'set inactive
			query="UPDATE ShipmentTypes SET active=0, international=0 WHERE idShipment=3;"
			set rs=connTemp.execute(query)
			response.redirect "viewshippingoptions.asp#UPS"
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
			ServiceStr=request.form("UPS_Service")
			if ServiceStr="" then
				response.redirect "UPS_EditShipOptions.asp?msg="&Server.URLEncode("Select at least one service.")
				response.end
			end if
			freeshipStr=""
			handlingStr=""
			servicePriorityStr=""
				
			If request.form("free01")="YES" then
				freeamt=request.form("amt01")
				freeshipStr=freeshipStr&"01|"&replacecomma(freeamt)&","
			End if
			If request.form("handling01")<>"0" AND request.form("handling01")<>"" then
				If isNumeric(request.form("handling01"))=true then
					handlingStr=handlingStr&"01|"&replacecomma(request.form("handling01"))&"|"&request.form("shfee01")&","
				End If
			End if
			servicePriority=request.form("servicePriority01")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"01|"&servicePriority&","
				
			If request.form("free02")="YES" then
				freeamt=request.form("amt02")
				freeshipStr=freeshipStr&"02|"&replacecomma(freeamt)&","
			End if
			If request.form("handling02")<>"0" AND request.form("handling02")<>"" then
				If isNumeric(request.form("handling02"))=true then
					handlingStr=handlingStr&"02|"&replacecomma(request.form("handling02"))&"|"&request.form("shfee02")&","
				End If
			End if
			servicePriority=request.form("servicePriority02")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"02|"&servicePriority&","
			'//NA FOR CANADA ORIGIN
			If request.form("free03")="YES" then
				freeamt=request.form("amt03")
				freeshipStr=freeshipStr&"03|"&replacecomma(freeamt)&","
			End if
			If request.form("handling03")<>"0" AND request.form("handling03")<>"" then
				If isNumeric(request.form("handling03"))=true then
					handlingStr=handlingStr&"03|"&replacecomma(request.form("handling03"))&"|"&request.form("shfee03")&","
				End If
			End if
			servicePriority=request.form("servicePriority03")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"03|"&servicePriority&","
			'// -------------------
			If request.form("free07")="YES" then
				freeamt=request.form("amt07")
				freeshipStr=freeshipStr&"07|"&replacecomma(freeamt)&","
			End if
			If request.form("handling07")<>"0" AND request.form("handling07")<>"" then
				If isNumeric(request.form("handling07"))=true then
					handlingStr=handlingStr&"07|"&replacecomma(request.form("handling07"))&"|"&request.form("shfee07")&","
				End If
			End if
			servicePriority=request.form("servicePriority07")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"07|"&servicePriority&","
			
			If request.form("free08")="YES" then
				freeamt=request.form("amt08")
				freeshipStr=freeshipStr&"08|"&replacecomma(freeamt)&","
			End if
			If request.form("handling08")<>"0" AND request.form("handling08")<>"" then
				If isNumeric(request.form("handling08"))=true then
					handlingStr=handlingStr&"08|"&replacecomma(request.form("handling08"))&"|"&request.form("shfee08")&","
				End If
			End if
			servicePriority=request.form("servicePriority08")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"08|"&servicePriority&","

			If request.form("free11")="YES" then
				freeamt=request.form("amt11")
				freeshipStr=freeshipStr&"11|"&replacecomma(freeamt)&","
			End if
			If request.form("handling11")<>"0" AND request.form("handling11")<>"" then
				If isNumeric(request.form("handling11"))=true then
					handlingStr=handlingStr&"11|"&replacecomma(request.form("handling11"))&"|"&request.form("shfee11")&","
				End If
			End if
			servicePriority=request.form("servicePriority11")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"11|"&servicePriority&","

			If request.form("free12")="YES" then
				freeamt=request.form("amt12")
				freeshipStr=freeshipStr&"12|"&replacecomma(freeamt)&","
			End if
			If request.form("handling12")<>"0" AND request.form("handling12")<>"" then
				If isNumeric(request.form("handling12"))=true then
					handlingStr=handlingStr&"12|"&replacecomma(request.form("handling12"))&"|"&request.form("shfee12")&","
				End If
			End if
			servicePriority=request.form("servicePriority12")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"12|"&servicePriority&","

			If request.form("free13")="YES" then
				freeamt=request.form("amt13")
				freeshipStr=freeshipStr&"13|"&replacecomma(freeamt)&","
			End if
			If request.form("handling13")<>"0" AND request.form("handling13")<>"" then
				If isNumeric(request.form("handling13"))=true then
					handlingStr=handlingStr&"13|"&replacecomma(request.form("handling13"))&"|"&request.form("shfee13")&","
				End If
			End if
			servicePriority=request.form("servicePriority13")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"13|"&servicePriority&","
			
			If request.form("free14")="YES" then
				freeamt=request.form("amt14")
				freeshipStr=freeshipStr&"14|"&replacecomma(freeamt)&","
			End if
			If request.form("handling14")<>"0" AND request.form("handling14")<>"" then
				If isNumeric(request.form("handling14"))=true then
					handlingStr=handlingStr&"14|"&replacecomma(request.form("handling14"))&"|"&request.form("shfee14")&","
				End If
			End if
			servicePriority=request.form("servicePriority14")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"14|"&servicePriority&","
			
			If request.form("free54")="YES" then
				freeamt=request.form("amt54")
				freeshipStr=freeshipStr&"54|"&replacecomma(freeamt)&","
			End if
			If request.form("handling54")<>"0" AND request.form("handling54")<>"" then
				If isNumeric(request.form("handling54"))=true then
					handlingStr=handlingStr&"54|"&replacecomma(request.form("handling54"))&"|"&request.form("shfee54")&","
				End If
			End if
			servicePriority=request.form("servicePriority54")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"54|"&servicePriority&","
			
			'//NA FOR CANADA ORIGIN
			If request.form("free59")="YES" then
				freeamt=request.form("amt59")
				freeshipStr=freeshipStr&"59|"&replacecomma(freeamt)&","
			End if
			If request.form("handling59")<>"0" AND request.form("handling59")<>"" then
				If isNumeric(request.form("handling59"))=true then
					handlingStr=handlingStr&"59|"&replacecomma(request.form("handling59"))&"|"&request.form("shfee59")&","
				End If
			End if
			servicePriority=request.form("servicePriority59")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"59|"&servicePriority&","
			
			'// -------------------
			'//NA FOR CANADA ORIGIN
			If request.form("free65")="YES" then
				freeamt=request.form("amt65")
				freeshipStr=freeshipStr&"65|"&replacecomma(freeamt)&","
			End if
			If request.form("handling65")<>"0" AND request.form("handling65")<>"" then
				If isNumeric(request.form("handling65"))=true then
					handlingStr=handlingStr&"65|"&replacecomma(request.form("handling65"))&"|"&request.form("shfee65")&","
				End If
			End if
			servicePriority=request.form("servicePriority65")
			If NOT validNum2(servicePriority) then
				servicePriority="0"
			End if
			servicePriorityStr=servicePriorityStr&"65|"&servicePriority&","
			
			'// -------------------
			set rs=Server.CreateObject("ADODB.Recordset")
			'clear all informatin out of shipService for UPS
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='01';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='02';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='03';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='07';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='08';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='11';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='12';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='13';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='14';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='54';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='59';"
			set rs=connTemp.execute(query)
			query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='65';"
			set rs=connTemp.execute(query)
			
			Dim i
			shipServiceArray=split(ServiceStr,", ")
			for i=0 to ubound(shipServiceArray)
				query="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
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
			response.redirect "viewshippingoptions.asp#UPS"			
		else %>
        
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>

            <form name="form1" method="post" action="UPS_EditShipOptions.asp" class="pcForms">
				<table class="pcCPcontent">                            
                    <tr> 
                        <td colspan="2">&nbsp;</td>
                    </tr>
                    <% query="SELECT serviceCode,serviceActive, servicePriority, serviceDescription,serviceFree,serviceFreeOverAmt,serviceHandlingFee,serviceShowHandlingFee FROM shipService ORDER BY serviceActive, servicePriority;"
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
						pTempString="<tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='UPS_Service' value='XXXX' "&pServiceCheck&"></td><td width='77%'><font color='#000000'><b>"&pServiceDescription&"</b></font></td><td width='19%' align='right'><strong>Order:&nbsp;</strong><input name='servicePriorityXXXX' type='text' id='servicePriorityXXXX' size='2' maxlength='10' value='"&pServicePriority&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='freeXXXX' type='checkbox' id='freeXXXX' value='YES' "&pServiceFreeChecked&">Offer free shipping for orders over "&scCurSign&" <input name='amtXXXX' type='text' id='amtXXXX' size='6' maxlength='10' value='"&money(pServiceFreeOverAmt)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee "&scCurSign&" <input name='handlingXXXX' type='text' id='handlingXXXX' size='6' maxlength='10' value='"&money(pServiceHandlingFee)&"'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfeeXXXX' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfeeXXXX' value='0' "&pServiceHandlingFeeChecked&">Integrate into shipping rate.</td></tr><tr><td colspan=2 class=pcCPspacer></td></tr>"

						select case pServiceCode
							case "01"
								pTempString=replace(pTempString,"XXXX","01")
								pcv_FormString=pcv_FormString&pTempString
							case "13"
								pTempString=replace(pTempString,"XXXX","13")
								pcv_FormString=pcv_FormString&pTempString
							case "14"
								pTempString=replace(pTempString,"XXXX","14")
								pcv_FormString=pcv_FormString&pTempString
							case "02"
								pTempString=replace(pTempString,"XXXX","02")
								pcv_FormString=pcv_FormString&pTempString
							case "59" 'Not for Canada Origin
								pTempString=replace(pTempString,"XXXX","59")
								pcv_FormString=pcv_FormString&pTempString
							case "12"
								pTempString=replace(pTempString,"XXXX","12")
								pcv_FormString=pcv_FormString&pTempString
							case "65" 'Not for Canada Origin
								pTempString=replace(pTempString,"XXXX","65")
								pcv_FormString=pcv_FormString&pTempString
							case "11"
								pTempString=replace(pTempString,"XXXX","11")
								pcv_FormString=pcv_FormString&pTempString
							case "03" 'Not for Canada Origin
								pTempString=replace(pTempString,"XXXX","03")
								pcv_FormString=pcv_FormString&pTempString
							case "07"
								pTempString=replace(pTempString,"XXXX","07")
								pcv_FormString=pcv_FormString&pTempString
							case "08"
								pTempString=replace(pTempString,"XXXX","08")
								pcv_FormString=pcv_FormString&pTempString
							case "54"
								pTempString=replace(pTempString,"XXXX","54")
								pcv_FormString=pcv_FormString&pTempString
						end select
						rs.moveNext
					loop 
					response.write pcv_FormString      
					set rs=nothing
					call closedb()
					%>
                    <tr class="normal"> 
                        <td colspan="2" align="left">
                        <input type="submit" name="Submit" value="Update UPS Shipping Services" class="submit2">
                        &nbsp;
                        <input type="button" name="back" value="Back" onClick="document.location.href='viewShippingOptions.asp'"></td>
                    </tr>
                </table>
            </form>
		<% end if %>
        </td>
    </tr>
    <tr align="center">
        <td>
        <hr>
        	<table>
                <tr>
                    <td width="58" valign="top" bgcolor="#FFFFFF"><div align="right"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50" /></div></td>
                    <td width="457" valign="top" bgcolor="#FFFFFF"><div align="center">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br>
                    THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br> 
                    UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</div></td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->