<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
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
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%
Dim query, conntemp, rstemp, rs

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

call openDb()
%>

<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<html>
<head>
<title><%response.write ship_dictLanguage.Item(Session("language")&"_viewCart_b")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 5px;">
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<%
				pcv_isZipRequired = true
				pcv_isCityRequired = true
				pcv_isShipCountryCodeRequired=true

				'// Use the Request object to toggle State (based of Country selection)
				pcv_isShipStateCodeRequired=True
				pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
				if  len(pcv_strStateCodeRequired)>0 then
					pcv_isShipStateCodeRequired=pcv_strStateCodeRequired
				end if

				'// Use the Request object to toggle Province (based of Country selection)
				pcv_isShipProvinceCodeRequired=False
				pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
				if  len(pcv_strProvinceCodeRequired)>0 then
					pcv_isShipProvinceCodeRequired=pcv_strProvinceCodeRequired
				end if

				IF request("SubmitShip")="" AND request("ddjumpflag")="" then %>
					<form action="estimateShipCost.asp" method="post" name="shipCost" class="pcForms">
						<table class="pcShowContent">
							<% msg=getUserInput(request.querystring("msg"),0)
							If msg<>"" then %>
								<tr>
									<td colspan="2"><div class="pcErrorMessage"><%=msg%></div></td>
								</tr>
							<% end if %>
							<tr>
								<td colspan="2"><h2><%=ship_dictLanguage.Item(Session("language")&"_viewCart_b")%></h2></td>
							</tr>
							<%
							'///////////////////////////////////////////////////////////
							'// START: COUNTRY AND STATE/ PROVINCE CONFIG
							'///////////////////////////////////////////////////////////

							'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
							pcv_isStateCodeRequired = pcv_isShipStateCodeRequired '// determines if validation is performed (true or false)
							pcv_isProvinceCodeRequired = pcv_isShipProvinceCodeRequired '// determines if validation is performed (true or false)
							pcv_isCountryCodeRequired = pcv_isShipCountryCodeRequired '// determines if validation is performed (true or false)

							'// Required Info
							pcv_strTargetForm = "shipCost" '// Name of Form
							pcv_strCountryBox = "CountryCode" '// Name of Country Dropdown
							pcv_strTargetBox = "StateCode" '// Name of State Dropdown
							pcv_strProvinceBox =  "Province" '// Name of Province Field

							'// Set local Country to Session
							if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strCountryBox) = CountryCode
							end if

							'// Set local State to Session
							if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strTargetBox) = StateCode
							end if

							'// Set local Province to Session
							if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Province
							end if
							%>
							<!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
							<%
							'///////////////////////////////////////////////////////////
							'// END: COUNTRY AND STATE/ PROVINCE CONFIG
							'///////////////////////////////////////////////////////////
							%>
							<%
							'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
							pcs_CountryDropdown
							%>
							<tr>
								<td width="25%"><p><%response.write dictLanguage.Item(Session("language")&"_vShipAdd_3")%></p></td>
								<td width="75%"><p><input name="city" type="text" size="30"
								value="<%=pcf_FillFormField ("city", pcv_isCityRequired) %>">
								<% pcs_RequiredImageTag "city", pcv_isCityRequired %>
								</p></td>
							</tr>
							<%
							'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
							pcs_StateProvince
							%>
							<tr>
								<td>
									<p><%response.write dictLanguage.Item(Session("language")&"_vShipAdd_5")%></p>
								</td>
								<td>
									<p><input name="zip" type="text" size="10"
									value="<%=pcf_FillFormField ("zip", pcv_isZipRequired) %>">
									<% pcs_RequiredImageTag "zip", pcv_isZipRequired %>
									</p>
								</td>
							</tr>

							<%
							'// ProductCart v4.5 - Commercial vs. Residential
							Dim pcComResShipAddress
							if scComResShipAddress = "0" then
							%>
							<tr>
								<td align="right"><p><input type="radio" name="residentialShipping" value="-1" checked class="clearBorder"></p></td>
								<td>
									<p><%response.write ship_dictLanguage.Item(Session("language")&"_login_c")%></p>
								</td>
							</tr>
							<tr>
								<td align="right"><p><input type="radio" name="residentialShipping" value="0" class="clearBorder"></p></td>
								<td>
									<p><%response.write ship_dictLanguage.Item(Session("language")&"_login_d")%></p>
								</td>
							</tr>
							<tr>
								<td class="pcSpacer"></td>
							</tr>
							<%
							else
								Select Case scComResShipAddress
								Case "1"
									pcComResShipAddress="-1"
								Case "2"
									pcComResShipAddress="0"
								Case "3"
									if session("customerType")="1" then
										pcComResShipAddress="0"
									else
										pcComResShipAddress="-1"
									end if
								End Select
							%>
							<tr>
								<td class="pcSpacer"><input type="hidden" name="residentialShipping" value="<%=pcComResShipAddress%>"></td>
							</tr>
							<%
							end if
							%>
							<tr>
								<td><input type="submit" name="SubmitShip" value="Submit" id="submit" class="submit2">
								</td>
								<td>&nbsp;</td>
							</tr>
						</table>
					</form>

				<% ELSE

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

					'// FedEX Variables SD
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
					<%
					'// page name
					pcStrPageName = "estimateShipCost.asp"

					'//set error to zero
					pcv_intErr=0

					'//generic error for page
					pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_Custmoda_18"))

					pResidentialShipping=request("residentialShipping")

					pcs_ValidateTextField	"CountryCode",  pcv_isShipCountryCodeRequired , 4
					pcs_ValidateTextField	"zip", pcv_isZipRequired, 10
					if request("ddjumpflag")="YES"then
						pcs_ValidateTextField	"city", false, 30
						pcs_ValidateStateProvField	"StateCode", false, 4
						pcs_ValidateStateProvField	"Province", false, 50
					else
						pcs_ValidateTextField	"city", pcv_isCityRequired, 30
						pcs_ValidateStateProvField	"StateCode", pcv_isShipStateCodeRequired, 4
						pcs_ValidateStateProvField	"Province", pcv_isShipProvinceCodeRequired, 50
					end if
					CountryCode=Session("pcSFCountryCode")
					StateCode=Session("pcSFStateCode")
					Province=Session("pcSFProvince")
					city=Session("pcSFcity")
					zip=Session("pcSFzip")
					If pcv_intErr>0 Then
						response.redirect pcStrPageName&"?reID="&reID&"&msg=" & pcv_strGenericPageError
					Else
						pcCartArray=Session("pcCartSession")
						ppcCartIndex=Session("pcCartIndex")

						' calculate total price of the order, total weight and product total quantities
						pSubTotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
						pShipCDSubTotal=Cdbl(calculateCategoryDiscounts(pcCartArray, ppcCartIndex))
						pSubTotal=pSubTotal-pShipCDSubTotal
						pShipSubTotal=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
						pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
						intUniversalWeight=pShipWeight
						pCartQuantity=Int(calculateCartQuantity(pcCartArray, ppcCartIndex))
						pCartShipQuantity=Int(calculateCartShipQuantity(pcCartArray, ppcCartIndex))
						pCartSurcharge=Cdbl(calculateTotalProductSurcharge(pcCartArray, ppcCartIndex))

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

						' if customer use anotherState, insert a dummy state code to simplify SQL sentence
						if Universal_destination_provOrState="" then
							 Universal_destination_provOrState="**"
						end if

						shipcompany=scShipService

						If pShipWeight="0" Then
							call opendb()
							query="SELECT idFlatShiptype,WQP FROM FlatShipTypes"
							set rsShpObj=server.CreateObject("ADODB.RecordSet")
							set rsShpObj=conntemp.execute(query)
							if rsShpObj.eof then
								Session("nullShipper")="Yes"
								call closeDb()
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
										'do nothing
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
							call closedb()
						Else
							Session("nullShipper")="No"
						End If

						If pCartShipQuantity=0 then
							Session("nullShipper")="Yes"
						end if

						iShipmentTypeCnt=0

						if session("provider")="" OR request("provider")<>"" then
							session("provider")=request("provider")
						end if

						if (session("availableShipStr")="" or session("provider")="") OR request("ddjumpflag")="" then %>
							<!--#include file="ShipRates.asp"-->
							<% session("strDefaultProvider")=strDefaultProvider
							session("iShipmentTypeCnt")=iShipmentTypeCnt
							session("strOptionShipmentType")=strOptionShipmentType
							session("availableShipStr")=availableShipStr
							session("iUPSFlag")=iUPSFlag
						end if

						availableShipStr=session("availableShipStr")
						call openDB()
						query="SELECT shipService.serviceCode, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation FROM shipService WHERE (((shipService.serviceActive)=-1)) ORDER BY shipService.servicePriority;"
						set rs=Server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if rs.eof then
							Session("nullShipper")="Yes"
						else %>
							<table class="pcShowContent">
								<tr>
									<td colspan="2">
									<h2><%response.write ship_dictLanguage.Item(Session("language")&"_viewCart_b")%></h2>
									</td>
								</tr>
								<%
								'if count shows that more then 1 shipmentType is active, show customer choice
								if session("iShipmentTypeCnt")>1 then %>
									<tr>
										<td width="25%">
											<p><%=ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_10")%></p>
										</td>
										<td width="75%">
											<form action="estimateShipCost.asp" method="post" class="pcForms">
												<select name="provider" onChange="javascript:if (this.value != '') {this.form.submit();}">
												<% strTempOptionShipmentType=session("strOptionShipmentType")
												if session("provider")<>"" then
													strTempOptionShipmentType=replace(strTempOptionShipmentType,"value="&session("provider")&"","value="&session("provider")&" selected")
												else
													strTempOptionShipmentType=replace(strTempOptionShipmentType,"value="&session("strDefaultProvider")&"","value="&session("strDefaultProvider")&" selected")
													session("provider")=session("strDefaultProvider")
												end if %>
												<%=strTempOptionShipmentType%>
												</select>
												<input type="hidden" name="ddjumpflag" value="YES">
												<input type="hidden" name="CountryCode" value="<%=Session("pcSFCountryCode")%>">
												<input type="hidden" name="StateCode" value="<%=Session("pcSFStateCode")%>">
												<input type="hidden" name="Province" value="<%=Session("pcSFProvince")%>">
												<input type="hidden" name="city" value="<%=Session("pcSFcity")%>">
												<input type="hidden" name="zip" value="<%=Session("pcSFzip")%>">
												<input type="hidden" name="residentialShipping" value="<%=request("residentialShipping")%>">
											</form>
										</td>
									</tr>
								<% else
									session("provider")=session("strDefaultProvider")%>
								<% end if %>
								<tr>
									<td colspan="2">
										<table class="pcShowContent">
											<tr>
												<th width="43%"><%= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_a")%></th>
												<th width="38%"><%if pcHideEstimateDeliveryTimes <> "-1" then %><%= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_b")%><%end if%></th>
												<th width="19%"><%= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_c")%></th>
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
																		if UPS_USENEGOTIATEDRATES=1 AND pcvNegRate<>"NONE"  then
																			tempRate=pcvNegRate
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
																if (ucase(shipDetailsArray(0))=ucase(session("provider"))) then
																	If serviceFree="-1" and Cdbl(pSubTotal)>Cdbl(serviceFreeOverAmt) then
																		tempRate="0"
																		tempRateDisplay= ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_f")
																		CntFree=CntFree+1
																	End If
																	DCnt=DCnt+1%>
																	<% Dim pshipDetailsArray2
																	pshipDetailsArray2= shipDetailsArray(2)
																	pshipDetailsArray2= replace(pshipDetailsArray2,"&reg;","")
																	pshipDetailsArray2= replace(pshipDetailsArray2,"<sup>SM</sup>","") %>
																	<tr>
																		<td><p><%=shipDetailsArray(2)%></p></td>
																		<td><p><% if shipDetailsArray(4)<>"NA" AND pcHideEstimateDeliveryTimes <> "-1" then %><%=shipDetailsArray(4)%><% end if %></p></td>
																		<td width="19%"><p><%=tempRateDisplay%></p></td>
																	</tr>
																<% end if
															end if
														end if
													next
													tempRate=""
													tempRateDisplay=""
												end if
												rs.movenext
											loop
											set rs=nothing
											call closeDb() %>
											<% if CntFree>0 then %>
												<tr>
													<td colspan="3">
														<p><%=ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_e")%></p>
													</td>
												</tr>
											<% end if %>
											<% if session("iUPSFlag")=1 AND ucase(session("provider"))="UPS"then %>
												<tr>
													<td colspan="3">
														<table width="427" border="0" cellspacing="0" cellpadding="2">
															<tr>
																<td width="45" valign="top"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50"></td>
																<td width="374" rowspan="2" valign="top" class="pcSmallText">
																<p><b>UPS OnLine&reg; Tools Rates &amp; Service Selection</b></p>
																<p>Notice: UPS fees do not necessarily represent UPS published rates	and may include charges levied by the store owner.</p>
																	<p>UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br />
THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br />
UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</p>
																</td>
															</tr>
															<tr>
																<td>&nbsp;</td>
															</tr>
														</table>
													</td>
												</tr>
											<% end if %>
											<% If DCnt=0 then %>
												<tr>
													<td colspan="3">
														<p><%response.write ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_d")%></p>
													</td>
												</tr>
											<% end if %>
										</table>
									</td>
								</tr>
								<tr>
									<td colspan="2" align="right">
										<form name="viewn" class="pcForms">
											<input type="image" src="images/close.gif" align="right" value="Close Window" onClick="self.close()" id="submit">
										</form>
									</td>
								</tr>
							</table>
						<% end if
					End If '// If pcv_intErr>0 Then
				END IF '// IF request("SubmitShip")="" AND request("ddjumpflag")="" then
				%>
			</td>
		</tr>
	</table>
	</div>
<% call closeDb()
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing %>
</body>
</html>