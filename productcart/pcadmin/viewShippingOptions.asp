<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Options Summary" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/FedEXconstants.asp"-->
<!--#include file="../includes/FedEXWSconstants.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/CPconstants.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/pcShipTestModes.asp" -->
<!--#include file="AdminHeader.asp"-->
<link href="../includes/spry/SpryAccordionCPS.css" rel="stylesheet" type="text/css" />
<script src="../includes/spry/SpryAccordion.js" type="text/javascript"></script>
<script src="../includes/spry/SpryURLUtils.js" type="text/javascript"></script>
<script type="text/javascript">
var params = Spry.Utils.getLocationParamsAsObject();
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<%

Dim customServiceActive : customServiceActive = "0"

If request("switch")<>"" then
	pcv_Switch=request("switch")
	pcv_Service=request("service")
	pcv_USPSTM=USPS_TESTMODE
	pcv_UPSTM=UPS_TESTMODE
	if pcv_Service="UPS" then
		if pcv_Switch="TEST" then
			pcv_UPSTM="1"
		else
			pcv_UPSTM="0"
		end if
	end if
	if pcv_Service="USPS" then
		if pcv_Switch="TEST" then
			pcv_USPSTM="1"
		else
			pcv_USPSTM="0"
		end if
	end if
	Dim objFS
	Dim objFile

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")

	'//Get File
	if PPD="1" then
		pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/pcShipTestModes.asp")
	else
		pcStrFileName=Server.Mappath ("../includes/pcShipTestModes.asp")
	end if

	'//Write File
	Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
	objFile.WriteLine CHR(60)&CHR(37)& vbCrLf
	objFile.WriteLine "UPS_TESTMODE = """&pcv_UPSTM&"""" & vbCrLf
	objFile.WriteLine "USPS_TESTMODE = """&pcv_USPSTM&"""" & vbCrLf
	objFile.WriteLine CHR(37)&CHR(62)& vbCrLf
	objFile.Close
	set objFS=nothing
	set objFile=nothing
	'//Redirect
	response.redirect "viewShippingOptions.asp"
end if

Dim connTemp, query, rs
set rs=server.CreateObject("ADODB.RecordSet")
call opendb()

query="SELECT active FROM ShipmentTypes WHERE idShipment=3"
set rs=connTemp.execute(query)
UPSActive=rs("active")
if UPSActive=True or UPSActive<>0 then
	UPSActive="YES"
end if

' check if UPS is disabled
Dim UPSDisabled : UPSDisabled = "NO"
if UPSActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE serviceCode IN ('01','02','03','07','08','11','12','13','14','54','59','65') AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		UPSDisabled = "YES"
	end if
end if


query="SELECT active FROM ShipmentTypes WHERE idShipment=4"
set rs=connTemp.execute(query)
USPSActive=rs("active")
if USPSActive=True or USPSActive<>0 then
	USPSActive="YES"
end if

' check if USPS is disabled
Dim USPSDisabled : USPSDisabled = "NO"
if USPSActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE serviceCode IN ('9901','9902','9903','9904','9905','9906','9907','9908','9909','9910','9911','9912','9913','9914','9915','9916','9917') AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		USPSDisabled = "YES"
	end if
end if



query="SELECT active FROM ShipmentTypes WHERE idShipment=1"
set rs=connTemp.execute(query)
FEDEXActive=rs("active")
if FEDEXActive=True or FEDEXActive<>0 then
	FEDEXActive="YES"
end if


query="SELECT active FROM ShipmentTypes WHERE idShipment=9"
set rs=connTemp.execute(query)
If NOT rs.EOF Then
	FEDEXWSActive=rs("active")
End If
if FEDEXWSActive=True or FEDEXWSActive<>0 then
	FEDEXWSActive="YES"
end if
set rs = nothing

' check if FEDEX is disabled
Dim FEDEXDisabled : FEDEXDisabled = "NO"
if FEDEXActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE serviceCode IN ('PRIORITYOVERNIGHT','STANDARDOVERNIGHT','FIRSTOVERNIGHT','FEDEX2DAY','FEDEXEXPRESSSAVER','INTERNATIONALPRIORITY','INTERNATIONALECONOMY','INTERNATIONALFIRST','FEDEX1DAYFREIGHT','FEDEX2DAYFREIGHT','FEDEX3DAYFREIGHT','FEDEXGROUND','GROUNDHOMEDELIVERY','INTERNATIONALPRIORITY FREIGHT','INTERNATIONALECONOMY FREIGHT') AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		FEDEXDisabled = "YES"
	end if
end if

' check if FEDEX WebServices is disabled
Dim FEDEXWSDisabled : FEDEXWSDisabled = "NO"
if FEDEXWSActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE serviceCode IN ('PRIORITY_OVERNIGHT','STANDARD_OVERNIGHT','FIRST_OVERNIGHT','FEDEX_2_DAY','FEDEX_EXPRESS_SAVER','INTERNATIONAL_PRIORITY','INTERNATIONAL_ECONOMY','INTERNATIONAL_FIRST','FEDEX_1_DAY_FREIGHT','FEDEX_2_DAY_FREIGHT','FEDEX_3_DAY_FREIGHT','FEDEX_GROUND','GROUND_HOME_DELIVERY','INTERNATIONAL_PRIORITY_FREIGHT','INTERNATIONAL_ECONOMY_FREIGHT','INTERNATIONAL_GROUND','FEDEX_FREIGHT','FEDEX_NATIONAL_FREIGHT','SMART_POST','EUROPE_FIRST_INTERNATIONAL_PRIORITY') AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		FEDEXWSDisabled = "YES"
	end if
end if


query="SELECT active FROM ShipmentTypes WHERE idShipment=7"
set rs=connTemp.execute(query)
CPActive=rs("active")
if CPActive=True or CPActive<>0 then
	CPActive="YES"
end if


' check if CP is disabled
Dim CPDisabled : CPDisabled = "NO"
if CPActive<>"YES" then
	query = "SELECT COUNT(*) AS NumServicesActive FROM shipService WHERE serviceCode IN ('1010','1020','1130','1030','1040','1120','1220','1230','2010','2020','2030','2040','2050','3010','3020','3040', '2005', '2015', '2025', '3005', '3015', '3025', '3050') AND serviceActive = -1"
	set rs=connTemp.execute(query)
	if rs("NumServicesActive") <> "0" then
		CPDisabled = "YES"
	end if
end if





%>

<div id="acc1" class="Accordion">

	<div class="AccordionPanel">
		<div class="AccordionPanelTab"><a name="UPS"></a>UPS OnLine&reg; Tools</div>
		<div class="AccordionPanelContent">
		<table class="pcCPcontent">

			<%
			IF UPSActive="YES" THEN

				if UPS_TESTMODE="1" then %>
				<tr class="pcShowProductsMheader">
					<td><p>You are currently runing UPS in <strong>&quot;TEST&quot; mode</strong>. <a href="viewShippingOptions.asp?switch=LIVE&service=UPS">Switch to &quot;LIVE&quot; mode</a>.</p></td>
				</tr>
				<tr class="pcShowProductsMheader">
					<td><p><u>"TEST" mode only affect the printing of UPS shipping labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
					</td>
				</tr>
				<% else %>
				<tr class="pcShowProductsMheader">
					<td><p>You are currently runing UPS in <strong>&quot;LIVE&quot; mode</strong>. <a href="viewShippingOptions.asp?switch=TEST&service=UPS">Switch to &quot;TEST&quot; mode</a>.</p></td>
				</tr>
				<tr class="pcShowProductsMheader">
					<td><p><u>This setting only affect the priting of UPS shippings labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
					</td>
				</tr>
				<% end if %>
				<tr>
					<td class="pcCPspacer"></td>
				</tr>
				<tr>
					<td>
					<ul>
						<li><img src="../pc/images/ups_pri_scr_lbg_sm.jpg" alt="UPS OnLine Tools" align="right">Services: - <a href="UPS_EditShipOptions.asp">Edit</a>
							<ul>
							<% query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
							set rs=connTemp.execute(query)
							do until rs.eof
								VarServiceCode=rs("serviceCode")
								select case VarServiceCode
								case "01","02","03","07","08","11","12","13","14","54","59","65"
									response.write "<li>"&rs("serviceDescription")&"</li>"
								end select
								rs.movenext
							loop %>
							</ul>
						 </li>
					<% 'if ups_license contains info, do not show this link
					query="SELECT ups_UserId FROM ups_license WHERE idUPS=1;"
					set rs=connTemp.execute(query)
					if len(rs("ups_UserId")&"A")=1 then %>
						<li>UPS OnLine&reg; Tools License Information - <a href="UPS_EditLicense.asp">Edit</a></li>
					<% end if %>
					<li>UPS Online Tools: User Preferences - <a href="UPS_Preferences.asp">Edit</a></li>
					<li>UPS Online Tools: Shipping Settings - <a href="UPS_EditSettings.asp">Edit</a></li>
					<li>Current Settings:
						<ul>
							<li>Default packaging:
							<% select case UPS_PACKAGE_TYPE
								case "00"
									response.write "Unknown"
								case "01"
									response.write "UPS Letter"
								case "02"
									response.write "Package"
								case "03"
									response.write "UPS Tube"
								case "04"
									response.write "UPS Pak"
								case "21"
									response.write "UPS Express Box"
								case "24"
									response.write "UPS 25KG Box&reg;"
							end select %>
							</li>
							<li>Default Account Type:
							  <% select case UPS_PICKUP_TYPE
								case "01"
									response.write "Daily Pickup"
								case "03"
									response.write "Occasional Pickup"
								case "11"
									response.write "Suggested Retail Rates (UPS Store)"
							end select %>
							</li>
							<li>Default Package Dimensions:
								<ul>
									<li>Height: <%=UPS_HEIGHT%>&nbsp;<%=UPS_DIM_UNIT%></li>
									<li>Width: <%=UPS_WIDTH%>&nbsp;<%=UPS_DIM_UNIT%></li>
									<li>Length: <%=UPS_LENGTH%>&nbsp;<%=UPS_DIM_UNIT%></li>
								</ul>
							</li>
						</ul>
					</li>
					<li><a href="OrderShippingOptions.asp?Provider=ups">Set Display Order</a></li>
					<li><a href="javascript:if (confirm('You are about to permanantly delete all current UPS settings. You will have to register again with the UPS Online Tools from your ProductCart Control Panel in order to reactivate UPS. This feature should only be used when UPS rates cannot be retrieved and no explanation other than a misconfigured account can be found. Are you sure you want to complete this action?')) location='pcResetUPS.asp'">Reset UPS OnLine&reg; Tools registration.</a></li>
					<li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='UPS_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
					<li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='UPS_EditShipOptions.asp?mode=del'">Remove</a></li>
					</ul>
				  </td>
				</tr>

			<% ELSEIF UPSDisabled = "YES" then %>
				<tr>
					<td>
						<img src="../pc/images/ups_pri_scr_lbg_sm2.jpg" alt="Enable (Reactivate) UPS" hspace="10"><strong>UPS</strong>&nbsp; is disabled - <a href="UPS_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
					</td>
				</tr>
			<% ELSE %>
				<tr>
					<td>
						<img src="../pc/images/ups_pri_scr_lbg_sm2.jpg" alt="Activate UPS" hspace="10"><strong>UPS OnLine&reg; Tools</strong> is not active - <a href="ConfigureOption1.asp">Activate</a>.
					</td>
				</tr>
			<% END IF %>

			<tr align="center">
				<td><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</div></td>
			</tr>
		 </table>
	  </div>
	</div>

	<div class="AccordionPanel">
		<div class="AccordionPanelTab"><a name="USPS"></a>United States Postal Service (USPS) &amp; Endicia</div>
		<div class="AccordionPanelContent">
		<table class="pcCPcontent">

			<% if USPSActive="YES" then %>
				<%call opendb()
				query="SELECT pcES_UserID,pcES_PassP,pcES_AutoRefill,pcES_TriggerAmount,pcES_LogTrans,pcES_Reg,pcES_TestMode FROM pcEDCSettings WHERE pcES_Reg=1;"
				set rsQ=connTemp.execute(query)

				tmpEDCUserID=0
				if not rsQ.eof then
					EndiciaReg=1
					tmpEDCUserID=rsQ("pcES_UserID")
				else
					EndiciaReg=0
				end if
				set rsQ=nothing%>
				<%if EndiciaReg=0 then
				if USPS_TESTMODE="1" then %>
				<tr class="pcShowProductsMheader">
					<td><p>You are currently runing USPS in <strong>&quot;TEST&quot; mode</strong>. <a href="viewShippingOptions.asp?switch=LIVE&service=USPS">Switch to &quot;LIVE&quot; mode</a></p></td>
				</tr>
				<tr class="pcShowProductsMheader">
					<td><p><u>"TEST" mode only affects the printing of USPS shipping labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
					</td>
				</tr>
				<% else %>
				<tr class="pcShowProductsMheader">
					<td><p>You are currently runing USPS in <strong>&quot;LIVE&quot; mode.</strong>. <a href="viewShippingOptions.asp?switch=TEST&service=USPS">Switch to &quot;TEST&quot; mode</a>.</p></td>
				</tr>
				<tr class="pcShowProductsMheader">
					<td><p><u>This setting only affect the printing of USPS shipping labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</p>
					</td>
				</tr>
				<% end if
				end if %>
				<tr>
					<td class="pcCPspacer"></td>
				</tr>
				<tr>
					<td>
						<ul>
							<li>Active Shipping Services - <a href="USPS_EditShipOptions.asp">Edit</a>
								<ul>
								<% dim USPSServiceFound
								USPSServiceFound=0
								query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
								set rs=connTemp.execute(query)
								do until rs.eof
									VarServiceCode=rs("serviceCode")
									select case VarServiceCode
									case "9901","9902","9903","9904","9905","9906","9907","9908","9909","9910","9911","9912","9913","9914","9915","9916","9917"
										response.write "<li>"&rs("serviceDescription")&"</li>"
										USPSServiceFound=1
									end select
									rs.movenext
								loop
								if USPSServiceFound=0 then
									response.write "<font color=#FF0000>NO services are active, please choose at least one service for the USPS Provider.</font>&nbsp; - <a href='USPS_EditShipOptions.asp'>Add</a>"
								end if
								set rs=nothing %>
								</ul>
							</li>
							<li>USPS License Information - <a href="USPS_EditLicense.asp">Edit</a></li>
							<li>USPS Shipping Settings - <a href="USPS_EditSettings.asp">Edit</a></li>
							<li>Current Shipping Settings:
								<ul>
									<li>Default Express Mail packaging: <%=USPS_EM_PACKAGE%></li>
									<% pcv_PMPackage=""
									select case USPS_PM_PACKAGE
									case "Flat Rate Envelope"
										pcv_PMPackage="Priority Mail Flat Rate Envelope, 12.5&quot; x 9.5&quot;"
									case "Flat Rate Box"
										pcv_PMPackage="Priority Mail Box, 12.25&quot; x 15.5&quot; x 3&quot;"
									case "Flat Rate Box1"
										pcv_PMPackage="Priority Mail Flat Rate Box, 14&quot; x 12&quot; x 3.5&quot;"
									case "Flat Rate Box2"
										pcv_PMPackage="Priority Mail Flat Rate Box, 11.25&quot; x 8.75&quot; x 6&quot;"
									end select %>
									<li>Default Priority Mail packaging: <%=pcv_PMPackage%></li>
								</ul>
							</li>
							<li><a href="OrderShippingOptions.asp?Provider=usps">Set Display Order</a></li>
							<li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='USPS_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
							<li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='USPS_EditShipOptions.asp?mode=del'">Remove</a></li>
						</ul>
					</td>
				</tr>
			</table>

			<h2>Endicia's Postage Label Services for USPS</h2>

			<table class="pcCPcontent">
			<tr valign="top">
				<td colspan="2">
					<img src="images/PoweredByEndicia_small.jpg" border="0" align="right" hspace="20">
					<%if EndiciaReg=0 then%>
						You can choose Endicia's Postage Label Services to print USPS postage.<br>
						<a href="EDC_manage.asp">Click here</a> to sign-up for an Endicia account.
					<%else%>
						<%if tmpEDCUserID="0" OR tmpEDCUserID="" then%>
							You signed up to use Endicia's service to print USPS postage.<br>
							Please <a href="EDC_manage.asp">click here</a> to complete the sign up process and activate your account
						<%else%>
							You are using Endicia's Postage Label Services.<br>
							<a href="EDC_manage.asp">Click here</a> to manage your Endicia account.
							<br><br>
							<a href="javascript:if (confirm('You are about to remove Endicia and all of its settings. Would you like to continue?')) location='ECD_remove.asp'">Remove</a> Endicia
						<%end if%>
				  <%end if%>
				</td>
			</tr>
			<tr>
				<td cospan="2" class="pcCPspacer"></td>
			</tr>
		<% elseif USPSDisabled = "YES" then %>
			<tr>
				<td>
					<strong>USPS</strong>&nbsp; is disabled - <a href="USPS_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
				</td>
			</tr>

		<% else %>
			<tr>
				<td><strong>USPS</strong>&nbsp; is not active - <a href="ConfigureOption2.asp">Activate</a> <br><span class="pcSmallText">Note: to take advantage of Endicia's Postage Printing Service, you must first activate USPS.</span></td>
			</tr>
		<% end if %>
		</table>
	  </div>
	</div>

	<div class="AccordionPanel">
		<div class="AccordionPanelTab"><a name="FedExWS"></a>FedEx Web Services</div>
		<div class="AccordionPanelContent">
		<table class="pcCPcontent">

			<% if FedExWSActive="YES" then %>

				<tr>
					<td>
						<ul>
						<li>FedEx Shipping Services - <a href="FedExWS_EditShipOptions.asp">Edit</a>
							<ul>
								<% query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
								set rs=connTemp.execute(query)
								do until rs.eof
									VarServiceCode=rs("serviceCode")
									select case VarServiceCode
										case "PRIORITY_OVERNIGHT","STANDARD_OVERNIGHT","FIRST_OVERNIGHT","FEDEX_2_DAY","FEDEX_EXPRESS_SAVER","INTERNATIONAL_PRIORITY","INTERNATIONAL_ECONOMY","INTERNATIONAL_FIRST","FEDEX_1_DAY_FREIGHT","FEDEX_2_DAY_FREIGHT","FEDEX_3_DAY_FREIGHT","FEDEX_GROUND","GROUND_HOME_DELIVERY","INTERNATIONAL_PRIORITY_FREIGHT","INTERNATIONAL_ECONOMY_FREIGHT","INTERNATIONAL_GROUND","FEDEX_FREIGHT","FEDEX_NATIONAL_FREIGHT","SMART_POST","EUROPE_FIRST_INTERNATIONAL_PRIORITY"
											response.write "<li>"&rs("serviceDescription")&"</li>"
											FEDEXWSServiceFound=1
									end select
									rs.movenext
								loop
								if FEDEXWSServiceFound=0 then
									response.write "<li>NO services are active, please choose at least one service for the FedEx Provider. - <a href='FEDEXWS_EditShipOptions.asp'>Add</a></li>"
								end if
								set rs=nothing %>
							</ul>
						</li>
						<li>FedEx Shipping Settings - <a href="FEDEXWS_EditSettings.asp">Edit</a></li>
						<li>Current Settings:
							<ul>
								<li>Default Package Type:
									<% select case FEDEXWS_FEDEX_PACKAGE
										case "YOUR_PACKAGING"
											response.write "Your Packaging"
										case "FEDEX_TUBE"
											response.write "FedEx&reg; Tube"
										case "FEDEX_PAK"
											response.write "FedEx&reg; Pak"
										case "FEDEX_ENVELOPE"
											response.write "FedEx&reg; Envelope"
										case "FEDEX_BOX"
											response.write "FedEx&reg; Box"
										case "FEDEX_10KG_BOX"
											response.write "FedEx&reg; 10KG Box"
										case "FEDEX_25KG_BOX"
											response.write "FedEx&reg; 25KG Box"
									end select %>
								</li>
								<li>Default Drop-off Type:
								<% select case FEDEXWS_DROPOFF_TYPE
									case "REGULAR_PICKUP"
										response.write "Regular Pickup"
									case "REQUEST_COURIER"
										response.write "Request Courier"
									case "DROP_BOX"
										response.write "Dropbox"
									case "BUSINESS_SERVICE_CENTER"
										response.write "Business Service Center"
									case "STATION"
										response.write "FedEx&reg; Station"
								end select %>
								</li>
								<li>Default Package Dimensions:
									<ul>
									<li>Height:
									<%=FEDEXWS_HEIGHT%>&nbsp;<%=FEDEXWS_DIM_UNIT%></li>
									<li>Width:
									<%=FEDEXWS_WIDTH%>&nbsp;<%=FEDEXWS_DIM_UNIT%></li>
									<li>Length:
									<%=FEDEXWS_LENGTH%>&nbsp;<%=FEDEXWS_DIM_UNIT%></li>
									</ul>
								</li>
							</ul>
						</li>
						<li><a href="OrderShippingOptions.asp?Provider=fedexWS">Set Display Order</a></li>
						<li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='FedEXWS_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
						<li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='FedEXWS_EditShipOptions.asp?mode=del'">Remove</a></li>
						</ul>
					</td>
				</tr>
			<%elseif FEDEXWSDisabled = "YES" then %>
			<tr>
				<td>
					<img src="../pc/images/fedex_corp_logo_sm.gif" alt="Enable (Reactivate)" hspace="10"><strong>FedEx</strong>&nbsp; is disabled - <a href="FedEXWS_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
				</td>
			</tr>
			<tr align="center">
				<td><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">FedEx service marks are owned by Federal Express Corporation and used with permission.</div></td>
			</tr>

			<% ELSE %>
				<tr>
					<td><img src="../pc/images/fedex_corp_logo_sm.gif" alt="Activate FedEx" hspace="10"><strong>FedEx Web Services</strong></strong>&nbsp;is not active - <a href="ConfigureOption5.asp">Activate</a> | <a href="http://wiki.earlyimpact.com/productcart/shipping-federal_express_ws" target="_blank">Help</a></td>
				</tr>
				<tr align="center">
					<td><div style="border: 1px dashed #CCC; margin: 10px; padding: 10px;">FedEx service marks are owned by Federal Express Corporation and used with permission.</div></td>
				</tr>

			<% end if %>

		</table>
	  </div>
	</div>

	<%
	' Only show the old FedEx integration if the store is using it
	if FEDEXActive="YES" OR FEDEXDisabled = "YES" then
	%>

	<div class="AccordionPanel">
		<div class="AccordionPanelTab"><a name="FedEX"></a>FedEx</div>
		<div class="AccordionPanelContent">
		<table class="pcCPcontent">

			<% if FEDEXActive="YES" then %>

				<tr>
					<td>
						<ul>
						<li>FedEx Shipping Services - <a href="FEDEX_EditShipOptions.asp">Edit</a>
							<ul>
								<% query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
								set rs=connTemp.execute(query)
								do until rs.eof
									VarServiceCode=rs("serviceCode")
									select case VarServiceCode
										case "PRIORITYOVERNIGHT","STANDARDOVERNIGHT","FIRSTOVERNIGHT","FEDEX2DAY","FEDEXEXPRESSSAVER","INTERNATIONALPRIORITY","INTERNATIONALECONOMY","INTERNATIONALFIRST","FEDEX1DAYFREIGHT","FEDEX2DAYFREIGHT","FEDEX3DAYFREIGHT","FEDEXGROUND","GROUNDHOMEDELIVERY","INTERNATIONALPRIORITY FREIGHT","INTERNATIONALECONOMY FREIGHT"
											response.write "<li>"&rs("serviceDescription")&"</li>"
											FEDEXServiceFound=1
									end select
									rs.movenext
								loop
								if FEDEXServiceFound=0 then
									response.write "<li>NO services are active, please choose at least one service for the FedEx Provider. - <a href='FEDEX_EditShipOptions.asp'>Add</a></li>"
								end if
								set rs=nothing %>
							</ul>
						</li>
						<li>FedEx Shipping Settings - <a href="FEDEX_EditSettings.asp">Edit</a></li>
						<li>Current Settings:
							<ul>
								<li>Default Package Type:
									<% select case FEDEX_FEDEX_PACKAGE
										case "YOURPACKAGING"
											response.write "Your Packaging"
										case "FEDEXTUBE"
											response.write "FedEx Tube"
										case "FEDEXPAK"
											response.write "FedEx Pak"
										case "FEDEXENVELOPE"
											response.write "FedEx Envelope"
										case "FEDEXBOX"
											response.write "FedEx Box"
										case "FEDEX10KGBOX"
											response.write "FedEx 10KG Box"
										case "FEDEX25KGBOX"
											response.write "FedEx 25KG Box"
									end select %>
								</li>
								<li>Default Drop-off Type:
								<% select case FEDEX_DROPOFF_TYPE
									case "REGULARPICKUP"
										response.write "Regular Pickup"
									case "REQUESTCOURIER"
										response.write "Request Courier"
									case "DROPBOX"
										response.write "Dropbox"
									case "BUSINESSSERVICE CENTER"
										response.write "Business Service Center"
									case "STATION"
										response.write "FedEx Station"
								end select %>
								</li>
								<li>Default Package Dimensions:
									<ul>
									<li>Height:
									<%=FEDEX_HEIGHT%>&nbsp;<%=FEDEX_DIM_UNIT%></li>
									<li>Width:
									<%=FEDEX_WIDTH%>&nbsp;<%=FEDEX_DIM_UNIT%></li>
									<li>Length:
									<%=FEDEX_LENGTH%>&nbsp;<%=FEDEX_DIM_UNIT%></li>
									</ul>
								</li>
							</ul>
						</li>
						<li><a href="OrderShippingOptions.asp?Provider=fedex">Set Display Order</a></li>
						<li><a href="javascript:if (confirm('You are about to disable (inactivate) the FedEx API integration. Would you like to continue?')) location='FedEX_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
						<li><a href="javascript:if (confirm('You are about to remove the FedEx API integration. Would you like to continue?')) location='FedEX_EditShipOptions.asp?mode=del'">Remove</a></li>
						</ul>
					</td>
				</tr>
			<%elseif FEDEXDisabled = "YES" then %>
			<tr>
				<td>
					<img src="../pc/images/fedex_corp_logo_sm.gif" alt="Enable (Reactivate)" hspace="10"><strong>FedEX</strong></strong>&nbsp; is disabled - <a href="FedEX_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
				</td>
			</tr>

			<% ELSE %>
				<tr>
					<td><img src="../pc/images/fedex_corp_logo_sm.gif" alt="Activate FedEx" hspace="10"><strong>FedEX</strong></strong>&nbsp;is not active - <a href="ConfigureOption3.asp">Activate</a></td>
				</tr>
			<% end if %>

		</table>
	  </div>
	</div>

	<% end if %>

	<div class="AccordionPanel">
		<div class="AccordionPanelTab"><a name="CP"></a>Canada Post</div>
		<div class="AccordionPanelContent">
		<table class="pcCPcontent">

			<% if CPActive="YES" then %>

				<tr>
					<td>
						<ul>
						<li>Canada Post Shipping Services - <a href="CP_EditShipOptions.asp">Edit</a>
							<ul>
							<% Dim CPServiceFound
							CPServiceFound=0
							query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1 Order by servicePriority;"
							set rs=connTemp.execute(query)
							do until rs.eof
								VarServiceCode=rs("serviceCode")
								select case VarServiceCode
									case "1010","1020","1130","1030","1040","1120","1220","1230","2010","2020","2030","2040","2050","3010","3020","3040", "2005", "2015", "2025", "3005", "3015", "3025", "3050"
										response.write "<li>"&rs("serviceDescription")&"</li>"
										CPServiceFound=1
								end select
								rs.movenext
							loop
							if CPServiceFound=0 then
								response.write "<li>NO services are active, please choose at least one service for Canada Post. - <a href='CP_EditShipOptions.asp'>Add</a></li>"
							end if
							set rs=nothing %>
							</ul>
						</li>
						<li>Canada Post User License - <a href="CP_EditLicense.asp">Edit</a></li>
						<li>Canada Post Shipping Settings - <a href="CP_EditSettings.asp">Edit</a></li>
						<li>Current Shipping Settings:
							<ul>
								<li>Default Package Dimensions:
									<ul>
										<li>Height:
										<%=CP_Height%>&nbsp;<%=CP_dimUnit%></li>
										<li>Width:
										<%=CP_Width%>&nbsp;<%=CP_dimUnit%></li>
										<li>Length:
										<%=CP_Length%>&nbsp;<%=CP_dimUnit%></li>
									</ul>
								</li>
							</ul>
						<li><a href="OrderShippingOptions.asp?Provider=cp">Set Display Order</a></li>
						<li><a href="javascript:if (confirm('You are about to disable (inactivate) this shipping provider. Would you like to continue?')) location='CP_EditShipOptions.asp?mode=InAct'">Disable (Inactivate)</a></li>
						<li><a href="javascript:if (confirm('You are about to remove this shipping provider and all of its shipping options. This action cannot be undone. Would you like to continue?')) location='CP_EditShipOptions.asp?mode=del'">Remove</a></li>
						</ul>
					</td>
				</tr>
			<%elseif CPDisabled = "YES" then %>
			<tr>
				<td>
					<strong>Canada Post</strong>&nbsp; is disabled - <a href="CP_EditShipOptions.asp?mode=Act">Enable (Reactivate)</a>
				</td>
			</tr>
			<% else %>
				<tr>
					<td><strong>Canada Post</strong>&nbsp;is not active - <a href="ConfigureOption4.asp">Add</a></td>
				</tr>
			<% end if %>

		</table>
	  </div>
	</div>

	<div class="AccordionPanel">
		<div class="AccordionPanelTab"><a name="Custom"></a>Custom Shipping Options</div>
		<div class="AccordionPanelContent">
		<table class="pcCPcontent">
			<% dim iCustomCnt
			iCustCnt=0
			'''query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1;"
			query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceCode LIKE 'C%';"
			set rs=connTemp.execute(query)
			if rs.eof then
			%>
			<tr>
				<td colspan="2">No Custom Shipping Options have been added - <a href="AddCustomShipping.asp">Add</a></td>
			</tr>
			<%
			else
			%>
			<tr>
				<td colspan="2" style="border-bottom: 1px dashed #CCC;" class="cpLinksList"><a href="AddCustomShipping.asp">Add New</a> : <a href="OrderShippingOptions.asp?Provider=x">Set Display Order</a></td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<%
				do until rs.eof
					VarServiceCode=rs("serviceCode")
					customServiceActive = rs("serviceActive")
					if left(VarServiceCode,1)="C" then
						iCustCnt=iCustCnt+1
						VaridFlatShipType=replace(VarServiceCode,"C","")
						query="SELECT FlatShipTypeDesc FROM FlatShipTypes WHERE idFlatShipType="&VaridFlatShipType&";"
						set rsCustObj=Server.CreateObject("ADODB.RecordSet")
						set rsCustObj=connTemp.execute(query)
						%>
						<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist">
						<td width="55%">
							<a href="modFlatShippingRates.asp?refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>"><strong><% =rsCustObj("FlatShipTypeDesc") %></strong></a>
						</td>
						<td width="45%" align="right">
							<span class="cpLinksList">
								<%if customServiceActive = "-1" then%>
									<a href="modFlatShippingRates.asp?refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>">Edit</a> : <a href="javascript:if (confirm('You are about to permanantly disable (inactivate) this shipping type from the database. Are you sure you want to complete this action?')) location='modFlatShippingRates.asp?mode=InAct&refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>'">Disable (Inactivate)</a> : <a href="javascript:if (confirm('You are about to permanantly delete this shipping type from the database. Are you sure you want to complete this action?')) location='modFlatShippingRates.asp?mode=DEL&refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>'">Remove</a>
								<%else%>
									<a href="modFlatShippingRates.asp?mode=Act&refer=viewShippingOptions.asp&idFlatShipType=<%=VaridFlatShipType%>">Enable (Reactivate)</a>
								<%end if%>
							</span>
						 </td>
						</tr>
						<% set rsCustObj=nothing
					end if
					rs.movenext
				loop
				set rs=nothing
				call closedb()
			end if
			%>
		</table>
	  </div>
	</div>
 </div>

<table class="pcCPcontent">
	<tr>
		<td align="center">
			<form class="pcForms" style="margin: 20px;">
				<input type="button" value="Edit Shipping Settings" onClick="location.href='modFromShipper.asp'">
				&nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</form>
		 </td>
	</tr>
</table>

<script type="text/javascript">
	var acc1 = new Spry.Widget.Accordion("acc1", { useFixedPanelHeights: false, defaultPanel: params.panel ? params.panel: -1 });
</script>

<!--#include file="AdminFooter.asp"-->