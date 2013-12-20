<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
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
<%
Dim query, conntemp, rstemp, rs

pcv_EOSC="Y"

shipmentTotal=Cdbl(0)
call openDb()

'//UPS Variables
query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
ups_license_key=trim(rstemp("AccessLicense"))
ups_userid=trim(rstemp("userID"))
ups_password=trim(rstemp("password"))
ups_active=rstemp("active")

'//CPS Variables
query="SELECT active, shipServer, userID FROM ShipmentTypes WHERE idshipment=7"
set rstemp=conntemp.execute(query)
CP_userid=trim(rstemp("userID"))
CP_server=trim(rstemp("shipserver"))
CP_active=rstemp("active")

'//FedEX Variables SD
Dim pcv_strAccountNameWS, pcv_strMeterNumberWS, pcv_strCarrierCodeWS
Dim pcv_strMethodNameWS, pcv_strMethodReplyWS, fedex_postdataWS, objFedExWSClass, objOutputXMLDocWS, srvFEDEXWSXmlHttp, FEDEXWS_result, FEDEXWS_URL, pcv_strErrorMsgWS

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

'//USPS Variables
query="SELECT active, shipServer, userID, [password] FROM ShipmentTypes WHERE idshipment=4"
set rstemp=conntemp.execute(query)
usps_userid=trim(rstemp("userID"))
usps_server=trim(rstemp("shipserver"))
usps_active=rstemp("active")

err.number=0 %>
<% query="SELECT shippingStateCode, shippingState, shippingCity, shippingCountryCode, shippingZip, ordShiptype FROM orders WHERE idOrder="&request.Querystring("idorder")&";"
set rstemp=conntemp.execute(query)
pStateCode=rstemp("shippingStateCode")
pState=rstemp("shippingState")
pCity=rstemp("shippingCity")
pCountryCode=rstemp("shippingCountryCode")
pZip=rstemp("shippingZip")
pordShiptype=rstemp("ordShiptype")
if pordShiptype = 0 then
	pResidentialShipping="-1"
else
	pResidentialShipping="0"
end if

' calculate total price of the order, total weight and product total quantities
pSubTotal=request.QueryString("subtotal")
pShipWeight=request.QueryString("weight")
intUniversalWeight=pShipWeight
pCartShipQuantity=request.QueryString("cartQTY")
pShipSubTotal=pSubTotal

if pState="" then
	pState=pStateCode
end if
Universal_destination_provOrState=pState
Universal_destination_country=pCountryCode
Universal_destination_postal=pZip
Universal_destination_city=pCity

' if customer use anotherState, insert a dummy state code to simplify SQL sentence
if Universal_destination_provOrState="" then
	Universal_destination_provOrState="**"
end if

shipcompany=scShipService

If pShipWeight="0" Then
	query="SELECT idFlatShiptype,WQP FROM FlatShipTypes"
	set rsShpObj=conntemp.execute(query)
	if rsShpObj.eof then
		call closeDb()
	else
		dim flagShp
		flagShp=0
		do until rsShpObj.eof
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
		if flagShp=0 then
		else
			Session("nullShipper")="No"
		End if
	end if
Else
	Session("nullShipper")="No"
End If

If pCartShipQuantity=0 then
end if
%>
<!--#include file="../pc/ShipRates.asp"-->
<% call openDB()
set rs=Server.CreateObject("ADODB.RecordSet")
query="SELECT serviceCode, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
set rs=connTemp.execute(query)
if rs.eof then
else %>
<html>
<head>
<title>Shipping Rates</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">
	<table class="pcCPMainTable">
		<tr>
			<td>
		<table class="pcCPcontent">
		<tr>
			<td height="174" valign="top">
				<table width="100%" border="0" cellspacing="0" cellpadding="4">
					<tr>
						<td>
							<table width="437" border="0" cellpadding="1" cellspacing="0">
								<tr>
									<td width="435">  <table width="100%" border="0" cellspacing="1" cellpadding="3">
										<tr>
											<th width="52%" align="left"><%response.write ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_a")%></th>
											<th colspan="2" align="left">
											  <%response.write ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_c")%></th>
											</tr>
										<% Dim CntFree, DCnt, serviceFree, serviceFreeOverAmt, serviceCode, OrderTotal, shipArray, i, shipDetailsArray, tempRate, tempRateDisplay
										CntFree=0
										DCnt=0
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
																if ucase(shipDetailsArray(0))="UPS" then
																	if pcv_UseNegotiatedRates=1 AND pcvNegRate<>"NONE"  then
																		tempRate=pcvNegRate
																	end if
																end if
															end if
															tempRateDisplay=scCurSign&money(tempRate)
															If serviceShowHandlingFee="0" then
																tempRate=(cDbl(tempRate)+cDbl(serviceHandlingFee))
																tempRateDisplay=scCurSign&money(tempRate)
																serviceHandlingFee="0"
															End If
															If serviceFree="-1" and Cdbl(pSubTotal)>Cdbl(serviceFreeOverAmt) then
																tempRate="0"
																tempRateDisplay= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_f")
																CntFree=CntFree+1
															End If
															DCnt=DCnt+1
															%>

															<% Dim pshipDetailsArray2
															pshipDetailsArray2= shipDetailsArray(2)
															pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>&reg;</sup>","")
															pshipDetailsArray2= replace(pshipDetailsArray2,"&reg;", "")
															pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>SM</sup>","")
															pshipDetailsArray0= shipDetailsArray(0)
	 %>
															<tr bgcolor="#FFFFFF"><FORM NAME="inputForm<%=DCnt%>" onSubmit="return setForm<%=DCnt%>();" class="pcForms">
															<td width="52%"><%=shipDetailsArray(2)%></td>
															<td width="27%"><%=tempRateDisplay%>&nbsp;<INPUT NAME="inputField<%=DCnt%>" TYPE="hidden" VALUE="<%=money(tempRate)%>"><INPUT NAME="inputProvider<%=DCnt%>" TYPE="hidden" VALUE="<%=pshipDetailsArray0%>"><INPUT NAME="inputService<%=DCnt%>" TYPE="hidden" VALUE="<%=pshipDetailsArray2%>"><INPUT NAME="inputHandlingFee<%=DCnt%>" TYPE="hidden" VALUE="<%=money(serviceHandlingFee)%>">	<INPUT NAME="inputServiceCode<%=DCnt%>" TYPE="hidden" VALUE="<%=serviceCode%>">														</td>
															<td width="21%"><INPUT TYPE="submit" name="UPD" VALUE="Select Rate" onSubmit="return setForm();"></td>
															</form></tr>
															<%
														end if
													end if
												next
												tempRate=""
												tempRateDisplay=""
											end if
											rs.movenext
										loop

										response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf
										for i=1 to DCnt
											response.write "function setForm"&i&"() {"&vbCrlf
											response.write "opener.document.EditOrder.Shipping.value = document.inputForm"&i&".inputField"&i&".value;"&vbCrlf
											response.write "opener.document.EditOrder.shippingProvider.value = document.inputForm"&i&".inputProvider"&i&".value;"&vbCrlf
											response.write "opener.document.EditOrder.handling.value = document.inputForm"&i&".inputHandlingFee"&i&".value;"&vbCrlf
											response.write "opener.document.EditOrder.shippingService.value = document.inputForm"&i&".inputService"&i&".value;"&vbCrlf
											response.write "opener.document.EditOrder.shippingServiceCode.value = document.inputForm"&i&".inputServiceCode"&i&".value;"&vbCrlf
											response.write "self.close();"&vbCrlf
											response.write "return false;"&vbCrlf
											response.write "}"&vbCrlf
										next
										response.write "//--></SCRIPT>"&vbCrlf
										set rs=nothing
										call closeDb() %>
										<% if CntFree>0 then %>
											<tr>
											<td colspan="3">
											<%response.write ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_e")%>											</td>
											</tr>
										<% end if %>
										<% if iUPSFlag=1 then %>
											<tr bgcolor="#FFFFFF">
											<td colspan="3">

											<table width="427" border="0" cellspacing="0" cellpadding="2">
											<tr>
											<td width="45" valign="top"><img src="../UPSLicense/LOGO_S2.gif" width="45" height="50"></td>
											<td width="374" rowspan="2" valign="top"><p>
											<b>UPS OnLine&reg; Tools Rates
											& Service Selection</b><br>
											Notice: UPS fees do not necessarily
											represent UPS published rates
											and may include charges levied
											by the store owner.</p>
											<p> UPS<sup>®</sup>, UPS Brandmark and COLOR BROWN<sup>®</sup>
										 <br>are trademarks of United Parcel Service of America, Inc. All Rights Reserved</p></td>
											</tr>
											<tr>
											<td>&nbsp;</td>
											</tr>
											</table>											</td>
											</tr>
										<% end if %>
											<% If DCnt=0 then %>
											<tr bgcolor="#FFFFFF">
												<td colspan="3"><br>
													&nbsp;
													<%response.write ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_d")%>												</td>
											</tr>
											<% end if %>
										</table></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
				  <td>&nbsp;            </td>
					</tr>
				</table>
			<% end if %>
		</td>
	</tr>
</table>
</td>
</tr>
</table>
</body>
</html>