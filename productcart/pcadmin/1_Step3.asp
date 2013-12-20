<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="UPS OnLine&reg; Tools Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<% pcPageName="UPS_EditSettings.asp"

'/////////////////////////////////////////////////////
'// Get Local Variables for Setting
'/////////////////////////////////////////////////////
pcStrPickupType = UPS_PICKUP_TYPE
pcStrPackageType = UPS_PACKAGE_TYPE
pcStrPackageHeight = UPS_HEIGHT
pcStrPackageWidth = UPS_WIDTH
pcStrPackageLength = UPS_LENGTH
pcStrPackageDimUnit = UPS_DIM_UNIT
pcStrShipperCompanyName = UPS_COMPANYNAME
pcStrShipperAttentionName = UPS_ATTENTION
pcStrShipperAddress1 = UPS_ADDRESS1
pcStrShipperAddress2 = UPS_ADDRESS2
pcStrShipperAddress3 = UPS_ADDRESS3
pcStrShipperCity = UPS_CITY
pcStrShipperState = UPS_STATE
pcStrShipperPostalCode = UPS_POSTALCODE
pcStrShipperCountryCode = UPS_COUNTRY
pcStrShipperPhone = UPS_PHONE
pcStrShipperFax = UPS_FAX
pcCurInsuredValue = UPS_INSUREDVALUE
pcStrDynamicInsuredValue = UPS_DYNAMICINSUREDVALUE
pcStrUseNegotiatedRates = UPS_USENEGOTIATEDRATES
pcStrShipperNumber = UPS_SHIPPERNUM

pcv_isPickupTypeRequired=false
pcv_isPackageTypeRequired=false
pcv_isPackageHeightRequired=false
pcv_isPackageWidthRequired=false
pcv_isPackageLengthRequired=false
pcv_isPackageDimUnitRequired=false
pcv_isShipperCompanyNameRequired=true 
pcv_isShipperAttentionNameRequired=false 
pcv_isShipperAddress1Required=true
pcv_isShipperAddress2Required=false
pcv_isShipperAddress3Required=false
pcv_isShipperCityRequired=true
pcv_isShipperStateRequired=false
pcv_isShipperPostalCodeRequired=false
pcv_isShipperCountryCodeRequired=true
pcv_isShipperPhoneRequired=false
pcv_isShipperFaxRequired=false
pcv_isInsuredValueRequired=false
pcv_isDynamicInsuredValueRequired=false
pcv_isUseNegotiatedRatesRequired=false
pcv_isShipperNumberRequired=false
%>
<table class="pcCPcontent">
	<tr>
		<td>
			<% if request.form("submit")<>"" then
									
				'/////////////////////////////////////////////////////
				'// Validate Fields and Set Sessions	
				'/////////////////////////////////////////////////////
				
				'// set errors to none
				pcv_intErr=0
				
				'// generic error for page
				pcv_strGenericPageError = "One of more fields were not filled in correctly."
					
				'// Clear error string
				pcv_strErrorMsg = ""
				pcs_ValidateTextField	"PickupType", pcv_isPickupTypeRequired, 0
				pcs_ValidateTextField	"PackageType", pcv_isPackageTypeRequired, 0
				pcs_ValidateTextField	"PackageHeight", pcv_isPackageHeightRequired, 0
				pcs_ValidateTextField	"PackageWidth", pcv_isPackageWidthRequired, 0
				pcs_ValidateTextField	"PackageLength", pcv_isPackageLengthRequired, 0
				pcs_ValidateTextField	"PackageDimUnit", pcv_isPackageDimUnitRequired, 0
				pcs_ValidateTextField	"ShipperCompanyName", pcv_isShipperCompanyNameRequired, 35
				pcs_ValidateTextField	"ShipperAttentionName", pcv_isShipperAttentionNameRequired, 35
				pcs_ValidateTextField	"ShipperAddress1", pcv_isShipperAddress1Required, 35
				pcs_ValidateTextField	"ShipperAddress2", pcv_isShipperAddress2Required, 35
				pcs_ValidateTextField	"ShipperAddress3", pcv_isShipperAddress3Required, 35
				pcs_ValidateTextField	"ShipperCity", pcv_isShipperCityRequired, 30
				pcs_ValidateTextField	"ShipperState", pcv_isShipperStateRequired, 5
				pcs_ValidateTextField	"ShipperPostalCode", pcv_isShipperPostalCodeRequired, 10
				pcs_ValidateTextField	"ShipperCountryCode", pcv_isShipperCountryCodeRequired, 2
				pcs_ValidateTextField	"ShipperPhone", pcv_isShipperPhoneRequired, 15
				pcs_ValidateTextField	"ShipperFax", pcv_isShipperFaxRequired, 14
				pcs_ValidateTextField	"InsuredValue", pcv_isInsuredValueRequired, 0
				pcs_ValidateTextField	"DynamicInsuredValue", pcv_isDynamicInsuredValueRequired, 0
				pcs_ValidateTextField	"UseNegotiatedRates", pcv_isUseNegotiatedRatesRequired, 0
				pcs_ValidateTextField	"ShipperNumber", pcv_isShipperNumberRequired, 0

				'/////////////////////////////////////////////////////
				'// Check for Validation Errors
				'/////////////////////////////////////////////////////
				If pcv_intErr>0 Then
					response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
				End If
						
				'/////////////////////////////////////////////////////
				'// Set Local Variables for Setting
				'/////////////////////////////////////////////////////

				pcStrPickupType = removeSQ(Session("pcAdminPickupType"))
				pcStrPackageType = removeSQ(Session("pcAdminPackageType"))
				'pcStrClassificationType = removeSQ(Session("pcAdminClassificationType"))
				select case pcStrPackageType
					case "01"
						pcStrClassificationType="01"
					case "03"
						pcStrClassificationType="03"
					case "11"
						pcStrClassificationType="04"
				end select
				pcStrPackageHeight = removeSQ(Session("pcAdminPackageHeight"))
				pcStrPackageWidth = removeSQ(Session("pcAdminPackageWidth"))
				pcStrPackageLength = removeSQ(Session("pcAdminPackageLength"))
				pcStrPackageDimUnit = removeSQ(Session("pcAdminPackageDimUnit"))
				pcStrShipperCompanyName = removeSQ(Session("pcAdminShipperCompanyName"))
				pcStrShipperAttentionName = removeSQ(Session("pcAdminShipperAttentionName"))
				pcStrShipperAddress1 = removeSQ(Session("pcAdminShipperAddress1"))
				pcStrShipperAddress2 = removeSQ(Session("pcAdminShipperAddress2"))
				pcStrShipperAddress3 = removeSQ(Session("pcAdminShipperAddress3"))
				pcStrShipperCity = removeSQ(Session("pcAdminShipperCity"))
				pcStrShipperState = removeSQ(Session("pcAdminShipperState"))
				pcStrShipperPostalCode = removeSQ(Session("pcAdminShipperPostalCode"))
				pcStrShipperCountryCode = removeSQ(Session("pcAdminShipperCountryCode"))
				pcStrShipperPhone = removeSQ(Session("pcAdminShipperPhone"))
				pcStrShipperFax = removeSQ(Session("pcAdminShipperFax"))
				pcCurInsuredValue = removeSQ(Session("pcAdminInsuredValue"))
				pcStrDynamicInsuredValue = removeSQ(Session("pcAdminDynamicInsuredValue"))
				pcStrUseNegotiatedRates = removeSQ(Session("pcAdminUseNegotiatedRates"))
				pcStrShipperNumber = removeSQ(Session("pcAdminShipperNumber"))
				'/////////////////////////////////////////////////////
				'// Update database with new Settings
				'/////////////////////////////////////////////////////
				%>
				<!--#include file="pcAdminSaveUPSConstants.asp"-->
				<%
				response.redirect "1_Step4.asp"
				response.end
			else
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Config Client-Side Validation
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				response.write "<script language=""JavaScript"">"&vbcrlf
				response.write "<!--"&vbcrlf	
				response.write "function Form1_Validator(theForm)"&vbcrlf
				response.write "{"&vbcrlf

				StrGenericJSError="One or more fields were not filled in correctly"
				
				pcs_JavaDropDownList "PickupType", pcv_isPickupTypeRequired, StrGenericJSError
				pcs_JavaDropDownList "PackageType", pcv_isPackageTypeRequired, StrGenericJSError
				pcs_JavaTextField	"PackageHeight", pcv_isPackageHeightRequired, StrGenericJSError
				pcs_JavaTextField	"PackageWidth", pcv_isPackageWidthRequired, StrGenericJSError
				pcs_JavaTextField	"PackageLength", pcv_isPackageLengthRequired, StrGenericJSError
				pcs_JavaTextField	"PackageDimUnit", pcv_isPackageDimUnitRequired, StrGenericJSError
				pcs_JavaTextField	"ShipperCompanyName", pcv_isShipperCompanyNameRequired, StrGenericJSError
				pcs_JavaTextField	"ShipperAttentionName", pcv_isShipperAttentionNameRequired, StrGenericJSError
				pcs_JavaTextField	"ShipperAddress1", pcv_isShipperAddress1Required, StrGenericJSError
				pcs_JavaTextField	"ShipperAddress2", pcv_isShipperAddress2Required, StrGenericJSError
				pcs_JavaTextField	"ShipperAddress3", pcv_isShipperAddress3Required, StrGenericJSError
				pcs_JavaTextField	"ShipperCity", pcv_isShipperCityRequired, StrGenericJSError
				pcs_JavaDropDownList "ShipperState", pcv_isShipperStateRequired, StrGenericJSError
				pcs_JavaTextField	"ShipperPostalCode", pcv_isShipperPostalCodeRequired, StrGenericJSError
				pcs_JavaDropDownList "ShipperCountryCode", pcv_isShipperCountryCodeRequired, StrGenericJSError
				pcs_JavaTextField	"ShipperPhone", pcv_isShipperPhoneRequired, StrGenericJSError
				pcs_JavaTextField	"ShipperFax", pcv_isShipperFaxRequired, StrGenericJSError
				pcs_JavaTextField	"InsuredValue", pcv_isInsuredValueRequired, StrGenericJSError
				pcs_JavaTextField	"DynamicInsuredValue", pcv_isDynamicInsuredValueRequired, StrGenericJSError
				pcs_JavaTextField	"UseNegotiatedRates", pcv_isUseNegotiatedRatesRequired, StrGenericJSError
				pcs_JavaTextField	"ShipperNumber", pcv_isShipperNumberRequired, StrGenericJSError
				response.write "return (true);"&vbcrlf
				response.write "}"&vbcrlf
				response.write "//-->"&vbcrlf
				response.write "</script>"&vbcrlf
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' End Config Client-Side Validation
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%>
            
				<% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>

            
                <form name="form1" method="post" action="1_Step3.asp" class="pcForms">
                    <table class="pcCPcontent">
                        <script type="text/javascript" language="JavaScript1.1">
                            <!--
                            function setOptions(chosen) {
                                var selbox = document.form1.PickupType;
                                
                                selbox.options.length = 0;
                                if (chosen == " ") {
                                    selbox.options[selbox.options.length] = new Option('Please select one of the classification options above.',' ');
                                }
                                if (chosen == "01") {
                                    selbox.options[selbox.options.length] = new Option('Daily Pickup','01');
                                }
                                if (chosen == "03") {
                                    selbox.options[selbox.options.length] = new Option('Daily Pickup','01');
                                    selbox.options[selbox.options.length] = new Option('UPS Customer Counter','03');
                                    selbox.options[selbox.options.length] = new Option('One Time Pickup','06');
                                    selbox.options[selbox.options.length] = new Option('On Call Air','07');
                                    selbox.options[selbox.options.length] = new Option('Letter Center/UPS Drop Box','19');
                                    selbox.options[selbox.options.length] = new Option('Air Service Center','20');
                                }
                                if (chosen == "04") {
                                    selbox.options[selbox.options.length] = new Option('UPS Customer Counter','03');
                                    selbox.options[selbox.options.length] = new Option('Suggested Retail Rates (UPS Store)','11');
                                }
                            }
                            // -->
                        </script>
                        <tr>
                            <th colspan="3">Shipper Address </th>
                        </tr>
                        <tr>
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td width="14%"><p>Company Name: </p></td>
                          <td width="86%"><p><input name="ShipperCompanyName" type="text" value="<%=pcStrShipperCompanyName%>" size="35" maxlength="35" >
                            <% pcs_RequiredImageTag "ShipperCompanyName", pcv_isShipperCompanyNameRequired %>
    *Use	only	alphanumeric	Characters	- &quot;&amp;&quot; symbols are not allowed.</p></td>
                        </tr>
                        <tr>
                            <td width="14%"><p>Attention Name: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperAttentionName" value="<%=pcStrShipperAttentionName%>" size="35" maxlength="35" >
                            <% pcs_RequiredImageTag "ShipperAttentionName", pcv_isShipperAttentionNameRequired %></p></td>
                        <tr>
                            <td width="14%"><p>Address Line 1: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperAddress1" value="<%=pcStrShipperAddress1%>" size="35" maxlength="35" >
                            <% pcs_RequiredImageTag "ShipperAddress1", pcv_isShipperAddress1Required %></p></td>
                        </tr>
                        <tr>
                            <td width="14%"><p>Address Line 2: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperAddress2" value="<%=pcStrShipperAddress2%>" size="35" maxlength="35" ><% pcs_RequiredImageTag "ShipperAddress2", pcv_isShipperAddress2Required %></p></td>
                        </tr>
                        <tr>
                            <td width="14%"><p>Address Line 3: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperAddress3" value="<%=pcStrShipperAddress3%>" size="35" maxlength="35" ><% pcs_RequiredImageTag "ShipperAddress3", pcv_isShipperAddress3Required %></p></td>
                        </tr>
                        <tr>
                            <td width="14%"><p>City: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperCity" value="<%=pcStrShipperCity%>" size="25" maxlength="30" ><% pcs_RequiredImageTag "ShipperCity", pcv_isShipperCityRequired %></p></td>
                        </tr>
                        <% dim conntemp, rs, query
                            call opendb()
                            '///////////////////////////////////////////////////////////
                            '// START: COUNTRY AND STATE/ PROVINCE CONFIG
                            '///////////////////////////////////////////////////////////
                            ' 
                            ' 1) Place this section ABOVE the Country field
                            ' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
                            ' 3) Additional Required Info
                            
                            '// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
                            pcv_isStateCodeRequired =  False '// determines if validation is performed (true or false)
                            pcv_isProvinceCodeRequired =  False '// determines if validation is performed (true or false)
                            pcv_isCountryCodeRequired =  False '// determines if validation is performed (true or false)
                            
                            '// #3 Additional Required Info
                            pcv_strTargetForm = "form1" '// Name of Form
                            pcv_strCountryBox = "ShipperCountryCode" '// Name of Country Dropdown
                            pcv_strTargetBox = "ShipperState" '// Name of State Dropdown
                            pcv_strProvinceBox =  "ShipperProvince" '// Name of Province Field
                            
                            '// Set local Country to Session
                            if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                                Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrShipperCountryCode
                            end if
                            
                            '// Set local State to Session
                            if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                                Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrShipperState
                            end if
                            
                            '// Set local Province to Session
                            if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                                Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  pcStrShipperState
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
                            <%
                            '// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
                            pcs_StateProvince
                            %>
                            <tr>
                            <td width="14%"><p>Postal Code: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperPostalCode" value="<%=pcStrShipperPostalCode%>" size="10" maxlength="10" ><% pcs_RequiredImageTag "ShipperPostalCode", pcv_isShipperStateRequired %></p></td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td width="14%"><p>Phone Number: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperPhone" value="<%=pcStrShipperPhone%>" size="15" maxlength="15" ><% pcs_RequiredImageTag "ShipperPhone", pcv_isShipperPhoneRequired %></p></td>
                        </tr>
                        <tr>
                            <td width="14%"><p>Fax Number: </p></td>
                            <td width="86%"><p><input type="text" name="ShipperFax" value="<%=pcStrShipperFax%>" size="15" maxlength="14" ><% pcs_RequiredImageTag "ShipperFax", pcv_isShipperFaxRequired %></p></td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <th colspan="3">Settings</th>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="2" bgcolor="#FFFFFF">
                            <h2>UPS Insurance Settings</h2> 
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">To use the value of the cart as the insurance rate value, choose to <span style="font-weight: bold; font-style: italic">Use Dynamic Insurance Rate</span>. If you are not using Dynamic Insurance Rate, you can set a <span style="font-weight: bold; font-style: italic">Flat Rate</span> that will be used for every UPS rate calculation in the store front. The default flat rate will be set to $100.00 if one is not set and you have not selected to use dynamic insurance rates. </td>
                        </tr>
                        <tr>
                            <td align="right"><p><input name="DynamicInsuredValue" type="radio" value="1" <%if pcStrDynamicInsuredValue="1" then%>checked<% end if %>></p></td>
                            <td>Use Dynamic Insurance Rate</td>
                        </tr>
                        <tr>
                            <td align="right"><p><input type="radio" name="DynamicInsuredValue" value="0" <%if pcStrDynamicInsuredValue="0" then%>checked<% end if %>></p></td>
                            <td>Use Flat Rate</td>
                        </tr>
                        <tr>
                            <td width="14%"><p>Flat  Rate Value: </p></td>
                            <td width="86%"><p><input type="text" name="InsuredValue" value="<%=pcCurInsuredValue%>" size="15" maxlength="14" ><% pcs_RequiredImageTag "InsuredValue", pcv_isInsuredValueRequired %></p></td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="2" bgcolor="#FFFFFF">
                            <h2>UPS Account Type</h2> 
                            <p><select name="PickupType" id="PickupType">
                            <option value="01" <%if pcStrPickupType="01" then%>selected<%end if%>>Daily Pickup</option>
                            <option value="03" <%if pcStrPickupType="03" then%>selected<%end if%>>Occasional Pickup</option>
                            <!-- Newly added 06/24/04 -->
                            <option value="11" <%if pcStrPickupType="11" then%>selected<%end if%>>Suggested Retail Rates (UPS Store)</option>
                            </select><% pcs_RequiredImageTag "PickupType", pcv_isPickupTypeRequired %>					
                            </p></td>
                        </tr>
                        <tr>
                            <td colspan="2" bgcolor="#FFFFFF"><p>With a <span style="font-weight: bold; font-style: italic">Daily Pickup Account</span>, a UPS driver will make a regular stop at your location each day, Monday through Friday, to pick up all package types, including:</p>
                                <ul>
                                    <li>Ground shipments</li>
                                    <li>Air shipments</li>
                                    <li>International Shipments</li>
                                </ul>
                             
                                <p>NOTE: A weekly service charge applies, and Daily Rates will be billed to your UPS Account (see the UPS Rates section under Terms and Conditions of Service in the online UPS Service Guide). </p>
                                </p>
                                <p>&nbsp;</p>
                                <p> With an <span style="font-weight: bold; font-style: italic">Occasional Account</span>, you decide if and when you need to schedule a UPS driver to pickup your shipments. Once you have prepared your shipments on UPS.com, you have the following options:&nbsp; </p>
                                <ul>
                                    <li>Schedule an On Call Pickup to have your shipment picked up by a UPS driver for a nominal fee.</li>
                                    <li>Hand you shipments to any UPS driver in your area.</li>
                                    <li>Take your shipments to The UPS Store, a UPS Customer Center, or an Authorized Shipping Outlet.</li>
                                </ul>
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td colspan="2" bgcolor="#FFFFFF"><h2>Account Based Rates</h2> </td>
                        </tr>
                        <tr> 
                            <td colspan="2"><p>Once you have completed the above step, enable ABR in ProductCart along with your UPS Account Number. If you do not have ABR enabled with UPS, enabling this setting will not allow rates to be returned in the store front.</p><br>
                            <p>
                            <label>
                                <input name="UseNegotiatedRates" type="checkbox" id="checkbox" value="1" <%if pcStrUseNegotiatedRates="1" then%>checked<% end if %>>
                            </label>
                            Enable Account Based Rates
                            </p>
                            <p>
                            UPS Account Number: 
                            <label>
                                <input name="ShipperNumber" type="text" id="ShipperNumber" value="<%=pcStrShipperNumber%>" size="6" maxlength="50"><% pcs_RequiredImageTag "ShipperNumber", pcv_isShipperNumberRequired %>	
                            </label>
                            </p>
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="2" bgcolor="#FFFFFF">
                            <h2>Package Type</h2>
                            <select name="PackageType" id="PackageType">
                            <option value="00" selected>Unknown</option>
                            <option value="02" <%if pcStrPackageType="02" then%>selected<%end if%>>Your Packaging</option>
                            <option value="01" <%if pcStrPackageType="01" then%>selected<%end if%>>UPS letter</option>
                            <option value="03" <%if pcStrPackageType="03" then%>selected<%end if%>>UPS Tube</option>
                            <option value="04" <%if pcStrPackageType="04" then%>selected<%end if%>>UPS Pak</option>
                            <option value="21" <%if pcStrPackageType="21" then%>selected<%end if%>>UPS Express Box</option>
                            </select><% pcs_RequiredImageTag "PackageType", pcv_isPackageTypeRequired %>					
                            </p></td>
                        </tr>
                        <tr> 
                            <td colspan="2"><p>If you select "Your Packaging", enter your default package type below. You can override these settings on a product by product basis by using the "Oversized" option.</p></td>
                        </tr>
                        <tr> 
                            <td width="14%"><p>Height: </p></td>
                            <td width="86%"> 
                            <input name="PackageHeight" type="text" id="PackageHeight" value="<%=pcStrPackageHeight%>" size="4" maxlength="4"><% pcs_RequiredImageTag "PackageHeight", pcv_isPackageHeightRequired %>						</td>
                        </tr>
                        <tr> 
                            <td width="14%"><p>Width: </p></td>
                            <td width="86%"><input name="PackageWidth" type="text" id="PackageWidth" value="<%=pcStrPackageWidth%>" size="4" maxlength="4"><% pcs_RequiredImageTag "PackageWidth", pcv_isPackageWidthRequired %>						</td>
                        </tr>
                        <tr> 
                            <td width="14%"><p>Length:</p></td>
                            <td width="86%"><input name="PackageLength" type="text" id="PackageLength" value="<%=pcStrPackageLength%>" size="4" maxlength="4"> <% pcs_RequiredImageTag "PackageLength", pcv_isPackageLengthRequired %>					
                            
                            <font color="#FF0000"> *This is the measurement of the longest side</font></td>
                        </tr>
                        <tr> 
                            <td colspan="2"><p>Measurement Unit: 
                            <% if pcStrPackageDimUnit="cm" then%> <input type="radio" name="PackageDimUnit" value="in"class="clearBorder">
                            Inches 
                            <input type="radio" name="PackageDimUnit" value="cm" checked class="clearBorder">
                            Centimeters 
                            <% else %> <input type="radio" name="PackageDimUnit" value="in" checked class="clearBorder">
                            Inches 
                            <input type="radio" name="PackageDimUnit" value="cm" class="clearBorder">
                            Centimeters 
                            <% end if %><% pcs_RequiredImageTag "PackageDimUnit", pcv_isPackageDimUnitRequired %>					
                            </p> </td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="3">
                            <h2>Notes about shipping <strong>oversized</strong> packages via UPS</h2>
                            You can add oversized information on a <u>product by product</u> basis. Therefore, you don't need to enter oversized package dimensions here. If you do, remember that:
                            <ul>
                            <li>&gt; &quot;Length&quot; should always be the longest side</li>
                            <li>&gt; (length + girth) cannot exceed 130 inches </li>
                            <li>&gt; &quot;Girth&quot; is defined as: (width*2) + (height*2) </li>
                            <li>&gt; For more information, <a href="http://www.ups.com/content/us/en/resources/prepare/oversize.html" target="_blank">click here</a>.</li>
                            </li>
                            </ul>
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"><hr></td>
                        </tr>
                        <tr> 
                            <td colspan="2" align="center"><input type="submit" name="Submit" value="Update Settings" class="submit2"></td>
                        </tr>
                        <tr> 
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr align="center">
                          <td colspan="3">
                            <table>
                                <tr>
                                    <td width="58" valign="top" bgcolor="#FFFFFF"><div align="right"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50" /></div></td>
                                    <td width="457" valign="top" bgcolor="#FFFFFF"><div align="center">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br />THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF <br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</div></td>
                                </tr>
                            </table></td>
                          </tr>
                    </table>
                </form>
			<% end if %>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->