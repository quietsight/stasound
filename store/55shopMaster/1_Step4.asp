<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="UPS Shipping Preferences" %>
<% Section="Shipping" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<% pcPageName="1_Step4.asp"
'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
dim query, rs, conntemp 

call opendb()

query="SELECT pcUPSPref_Service, pcUPSPref_PackageType, pcUPSPref_PaymentMethod, pcUPSPref_AccountNumber, pcUPSPref_ReadyHours, pcUPSPref_ReadyMinutes, pcUPSPref_ReadyAMPM, pcUPSPref_PUHours, pcUPSPref_PUMinutes, pcUPSPref_RefNumber1, pcUPSPref_RefNumber2, pcUPSPref_RefData1, pcUPSPref_RefData2, pcUPSPref_CODPackage, pcUPSPref_CODAmount, pcUPSPref_CODCurrency, pcUPSPref_CODFunds, pcUPSPref_ShipmentNotification, pcUPSPref_NotifiCode1, pcUPSPref_NotifiCode2, pcUPSPref_NotifiCode3, pcUPSPref_NotifiCode4, pcUPSPref_NotifiCode5, pcUPSPref_NotifiEmail1, pcUPSPref_NotifiEmail2, pcUPSPref_NotifiEmail3, pcUPSPref_NotifiEmail4, pcUPSPref_NotifiEmail5, pcUPSPref_SaturdayDelivery, pcUPSPref_InsuredValue, pcUPSPref_VerbalConfirmation FROM pcUPSPreferences WHERE pcUPSPref_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then
	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	Session("pcAdminUPSService")=rs("pcUPSPref_Service")
	Session("pcAdminUPSPackageType")=rs("pcUPSPref_PackageType")
	Session("pcAdminUPSPaymentMethod")=rs("pcUPSPref_PaymentMethod")
	Session("pcAdminUPSAccountNumber")=rs("pcUPSPref_AccountNumber")
	Session("pcAdminUPSReadyHours")=rs("pcUPSPref_ReadyHours")
	Session("pcAdminUPSReadyMinutes")=rs("pcUPSPref_ReadyMinutes")
	Session("pcAdminUPSReadyAMPM")=rs("pcUPSPref_ReadyAMPM")
	Session("pcAdminUPSPUHours")=rs("pcUPSPref_PUHours")
	Session("pcAdminUPSPUMinutes")=rs("pcUPSPref_PUMinutes")
	Session("pcAdminUPSRefNumber1")=rs("pcUPSPref_RefNumber1")
	Session("pcAdminUPSRefNumber2")=rs("pcUPSPref_RefNumber2")
	Session("pcAdminUPSRefData1")=rs("pcUPSPref_RefData1")
	Session("pcAdminUPSRefData2")=rs("pcUPSPref_RefData2")
	Session("pcAdminUPSCODPackage")=rs("pcUPSPref_CODPackage")
	Session("pcAdminUPSCODAmount")=rs("pcUPSPref_CODAmount")
	Session("pcAdminUPSCODCurrency")=rs("pcUPSPref_CODCurrency")
	Session("pcAdminUPSCODFunds")=rs("pcUPSPref_CODFunds")
	Session("pcAdminUPSShipmentNotification")=rs("pcUPSPref_ShipmentNotification")
	Session("pcAdminUPSNotifiCode1")=rs("pcUPSPref_NotifiCode1")
	Session("pcAdminUPSNotifiCode2")=rs("pcUPSPref_NotifiCode2")
	Session("pcAdminUPSNotifiCode3")=rs("pcUPSPref_NotifiCode3")
	Session("pcAdminUPSNotifiCode4")=rs("pcUPSPref_NotifiCode4")
	Session("pcAdminUPSNotifiCode5")=rs("pcUPSPref_NotifiCode5")
	Session("pcAdminUPSNotifiEmail1")=rs("pcUPSPref_NotifiEmail1")
	Session("pcAdminUPSNotifiEmail2")=rs("pcUPSPref_NotifiEmail2")
	Session("pcAdminUPSNotifiEmail3")=rs("pcUPSPref_NotifiEmail3")
	Session("pcAdminUPSNotifiEmail4")=rs("pcUPSPref_NotifiEmail4")
	Session("pcAdminUPSNotifiEmail5")=rs("pcUPSPref_NotifiEmail5")
	Session("pcAdminUPSSaturdayDelivery")=rs("pcUPSPref_SaturdayDelivery")
	Session("pcAdminUPSInsuredValue")=rs("pcUPSPref_InsuredValue")
	Session("pcAdminUPSVerbalConfirmation")=rs("pcUPSPref_VerbalConfirmation")
end if

pcv_isUPSServiceRequired=false
pcv_isUPSPackageTypeRequired=false
pcv_isUPSPaymentMethodRequired=false
pcv_isUPSAccountNumberRequired=true
pcv_isUPSReadyHoursRequired=false
pcv_isUPSReadyMinutesRequired=false
pcv_isUPSReadyAMPMRequired=false
pcv_isUPSPUHoursRequired=false
pcv_isUPSPUMinutesRequired=false
pcv_isUPSRefNumber1Required=false
pcv_isUPSRefNumber2Required=false
pcv_isUPSRefData1Required=false
pcv_isUPSRefData2Required=false
pcv_isUPSCODPackageRequired=false
pcv_isUPSCODAmountRequired=false
pcv_isUPSCODCurrencyRequired=false
pcv_isUPSCODFundsRequired=false
pcv_isUPSShipmentNotificationRequired=false
pcv_isUPSNotifiCode1Required=false
pcv_isUPSNotifiCode2Required=false
pcv_isUPSNotifiCode3Required=false
pcv_isUPSNotifiCode4Required=false
pcv_isUPSNotifiCode5Required=false
pcv_isUPSNotifiEmail1Required=false
pcv_isUPSNotifiEmail2Required=false
pcv_isUPSNotifiEmail3Required=false
pcv_isUPSNotifiEmail4Required=false
pcv_isUPSNotifiEmail5Required=false
pcv_isUPSSaturdayDeliveryRequired=false
pcv_isUPSInsuredValueRequired=false
pcv_isUPSVerbalConfirmationRequired=false

if request("Submit1")="Update" then
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = "One of more fields were not filled in correctly."
	
	'// validate all fields
	pcs_ValidateTextField	"UPSService", pcv_isUPSServiceRequired, 250
	pcs_ValidateTextField	"UPSPackageType", pcv_isUPSPackageTypeRequired, 250
	pcs_ValidateTextField	"UPSPaymentMethod", pcv_isUPSPaymentMethodRequired, 250
	pcs_ValidateTextField	"UPSAccountNumber", pcv_isUPSAccountNumberRequired, 250
	pcs_ValidateTextField	"UPSReadyHours", pcv_isUPSReadyHoursRequired, 250
	pcs_ValidateTextField	"UPSReadyMinutes", pcv_isUPSReadyMinutesRequired, 250
	pcs_ValidateTextField	"UPSReadyAMPM", pcv_isUPSReadyAMPMRequired, 250
	pcs_ValidateTextField	"UPSPUHours", pcv_isUPSPUHoursRequired, 250
	pcs_ValidateTextField	"UPSPUMinutes", pcv_isUPSPUMinutesRequired, 250
	pcs_ValidateTextField	"UPSRefNumber1", pcv_isUPSRefNumber1Required, 250
	pcs_ValidateTextField	"UPSRefNumber2", pcv_isUPSRefNumber2Required, 250
	pcs_ValidateTextField	"UPSRefData1", pcv_isUPSRefData1Required, 250
	pcs_ValidateTextField	"UPSRefData2", pcv_isUPSRefData2Required, 250
	pcs_ValidateTextField	"UPSCODPackage", pcv_isUPSCODPackageRequired, 250
	pcs_ValidateTextField	"UPSCODAmount", pcv_isUPSCODAmountRequired, 250
	pcs_ValidateTextField	"UPSCODCurrency", pcv_isUPSCODCurrencyRequired, 250
	pcs_ValidateTextField	"UPSCODFunds", pcv_isUPSCODFundsRequired, 250
	pcs_ValidateTextField	"UPSShipmentNotification", pcv_isUPSShipmentNotificationRequired, 250
	pcs_ValidateTextField	"UPSNotifiCode1", pcv_isUPSNotifiCode1Required, 250
	pcs_ValidateTextField	"UPSNotifiCode2", pcv_isUPSNotifiCode2Required, 250
	pcs_ValidateTextField	"UPSNotifiCode3", pcv_isUPSNotifiCode3Required, 250
	pcs_ValidateTextField	"UPSNotifiCode4", pcv_isUPSNotifiCode4Required, 250
	pcs_ValidateTextField	"UPSNotifiCode5", pcv_isUPSNotifiCode5Required, 250
	pcs_ValidateTextField	"UPSNotifiEmail1", pcv_isUPSNotifiEmail1Required, 250
	pcs_ValidateTextField	"UPSNotifiEmail2", pcv_isUPSNotifiEmail2Required, 250
	pcs_ValidateTextField	"UPSNotifiEmail3", pcv_isUPSNotifiEmail3Required, 250
	pcs_ValidateTextField	"UPSNotifiEmail4", pcv_isUPSNotifiEmail4Required, 250
	pcs_ValidateTextField	"UPSNotifiEmail5", pcv_isUPSNotifiEmail5Required, 250
	pcs_ValidateTextField	"UPSSaturdayDelivery", pcv_isUPSSaturdayDeliveryRequired, 250
	pcs_ValidateTextField	"UPSInsuredValue", pcv_isUPSInsuredValueRequired, 250
	pcs_ValidateTextField	"UPSVerbalConfirmation", pcv_isUPSVerbalConfirmationRequired, 250

	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		
		response.write pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
		response.end
	End If
	
	pcStrUPSService = Session("pcAdminUPSService")
	pcStrUPSPackageType = Session("pcAdminUPSPackageType")
	pcStrUPSPaymentMethod = Session("pcAdminUPSPaymentMethod")
	pcStrUPSAccountNumber = Session("pcAdminUPSAccountNumber")
	pcStrUPSReadyHours = Session("pcAdminUPSReadyHours")
	pcStrUPSReadyMinutes = Session("pcAdminUPSReadyMinutes")
	pcStrUPSReadyAMPM = Session("pcAdminUPSReadyAMPM")
	pcStrUPSPUHours = Session("pcAdminUPSPUHours")
	pcStrUPSPUMinutes = Session("pcAdminUPSPUMinutes")
	pcStrUPSRefNumber1 = Session("pcAdminUPSRefNumber1")
	pcStrUPSRefNumber2 = Session("pcAdminUPSRefNumber2")
	pcStrUPSRefData1 = Session("pcAdminUPSRefData1")
	pcStrUPSRefData2 = Session("pcAdminUPSRefData2")
	pcStrUPSCODPackage = Session("pcAdminUPSCODPackage")
	pcStrUPSCODAmount = Session("pcAdminUPSCODAmount")
	pcStrUPSCODCurrency = Session("pcAdminUPSCODCurrency")
	pcStrUPSCODFunds = Session("pcAdminUPSCODFunds")
	pcStrUPSShipmentNotification = Session("pcAdminUPSShipmentNotification")
	pcStrUPSNotifiCode1 = Session("pcAdminUPSNotifiCode1")
	pcStrUPSNotifiCode2 = Session("pcAdminUPSNotifiCode2")
	pcStrUPSNotifiCode3 = Session("pcAdminUPSNotifiCode3")
	pcStrUPSNotifiCode4 = Session("pcAdminUPSNotifiCode4")
	pcStrUPSNotifiCode5 = Session("pcAdminUPSNotifiCode5")
	pcStrUPSNotifiEmail1 = Session("pcAdminUPSNotifiEmail1")
	pcStrUPSNotifiEmail2 = Session("pcAdminUPSNotifiEmail2")
	pcStrUPSNotifiEmail3 = Session("pcAdminUPSNotifiEmail3")
	pcStrUPSNotifiEmail4 = Session("pcAdminUPSNotifiEmail4")
	pcStrUPSNotifiEmail5 = Session("pcAdminUPSNotifiEmail5")
	pcStrUPSSaturdayDelivery = Session("pcAdminUPSSaturdayDelivery")
	pcStrUPSInsuredValue = Session("pcAdminUPSInsuredValue")
	pcStrUPSVerbalConfirmation = Session("pcAdminUPSVerbalConfirmation")

	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
	query="SELECT pcUPSPref_ID FROM pcUPSPreferences;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	if NOT rs.eof then
		query="UPDATE pcUPSPreferences SET pcUPSPref_Service='"&pcStrUPSService&"', pcUPSPref_PackageType='"&pcStrUPSPackageType&"', pcUPSPref_PaymentMethod='"&pcStrUPSPaymentMethod&"', pcUPSPref_AccountNumber='"&pcStrUPSAccountNumber&"', pcUPSPref_ReadyHours='"&pcStrUPSReadyHours&"', pcUPSPref_ReadyMinutes='"&pcStrUPSReadyMinutes&"', pcUPSPref_ReadyAMPM='"&pcStrUPSReadyAMPM&"', pcUPSPref_PUHours='"&pcStrUPSPUHours&"', pcUPSPref_PUMinutes='"&pcStrUPSPUMinutes&"', pcUPSPref_RefNumber1='"&pcStrUPSRefNumber1&"', pcUPSPref_RefNumber2='"&pcStrUPSRefNumber2&"', pcUPSPref_RefData1='"&pcStrUPSRefData1&"', pcUPSPref_RefData2='"&pcStrUPSRefData2&"', pcUPSPref_CODPackage='"&pcStrUPSCODPackage&"', pcUPSPref_CODAmount='"&pcStrUPSCODAmount&"', pcUPSPref_CODCurrency='"&pcStrUPSCODCurrency&"', pcUPSPref_CODFunds='"&pcStrUPSCODFunds&"', pcUPSPref_ShipmentNotification='"&pcStrUPSShipmentNotification&"', pcUPSPref_NotifiCode1='"&pcStrUPSNotifiCode1&"', pcUPSPref_NotifiCode2='"&pcStrUPSNotifiCode2&"', pcUPSPref_NotifiCode3='"&pcStrUPSNotifiCode3&"', pcUPSPref_NotifiCode4='"&pcStrUPSNotifiCode4&"', pcUPSPref_NotifiCode5='"&pcStrUPSNotifiCode5&"', pcUPSPref_NotifiEmail1='"&pcStrUPSNotifiEmail1&"', pcUPSPref_NotifiEmail2='"&pcStrUPSNotifiEmail2&"', pcUPSPref_NotifiEmail3='"&pcStrUPSNotifiEmail3&"', pcUPSPref_NotifiEmail4='"&pcStrUPSNotifiEmail4&"', pcUPSPref_NotifiEmail5='"&pcStrUPSNotifiEmail5&"', pcUPSPref_SaturdayDelivery='"&pcStrUPSSaturdayDelivery&"', pcUPSPref_InsuredValue='"&pcStrUPSInsuredValue&"', pcUPSPref_VerbalConfirmation='"&pcStrUPSVerbalConfirmation&"' WHERE pcUPSPref_ID=1;"
	else
		query="INSERT INTO pcUPSPreferences (pcUPSPref_Service, pcUPSPref_PackageType, pcUPSPref_PaymentMethod, pcUPSPref_AccountNumber, pcUPSPref_ReadyHours, pcUPSPref_ReadyMinutes, pcUPSPref_ReadyAMPM, pcUPSPref_PUHours, pcUPSPref_PUMinutes, pcUPSPref_RefNumber1, pcUPSPref_RefNumber2, pcUPSPref_RefData1, pcUPSPref_RefData2, pcUPSPref_CODPackage, pcUPSPref_CODAmount, pcUPSPref_CODCurrency, pcUPSPref_CODFunds, pcUPSPref_ShipmentNotification, pcUPSPref_NotifiCode1, pcUPSPref_NotifiCode2, pcUPSPref_NotifiCode3, pcUPSPref_NotifiCode4, pcUPSPref_NotifiCode5, pcUPSPref_NotifiEmail1, pcUPSPref_NotifiEmail2, pcUPSPref_NotifiEmail3, pcUPSPref_NotifiEmail4, pcUPSPref_NotifiEmail5, pcUPSPref_SaturdayDelivery, pcUPSPref_InsuredValue, pcUPSPref_VerbalConfirmation) VALUES ('"&pcStrUPSService&"', '"&pcStrUPSPackageType&"', '"&pcStrUPSPaymentMethod&"', '"&pcStrUPSAccountNumber&"', '"&pcStrUPSReadyHours&"', '"&pcStrUPSReadyMinutes&"', '"&pcStrUPSReadyAMPM&"', '"&pcStrUPSPUHours&"', '"&pcStrUPSPUMinutes&"', '"&pcStrUPSRefNumber1&"', '"&pcStrUPSRefNumber2&"', '"&pcStrUPSRefData1&"', '"&pcStrUPSRefData2&"', '"&pcStrUPSCODPackage&"', '"&pcStrUPSCODAmount&"', '"&pcStrUPSCODCurrency&"', '"&pcStrUPSCODFunds&"', '"&pcStrUPSShipmentNotification&"', '"&pcStrUPSNotifiCode1&"', '"&pcStrUPSNotifiCode2&"', '"&pcStrUPSNotifiCode3&"', '"&pcStrUPSNotifiCode4&"', '"&pcStrUPSNotifiCode5&"', '"&pcStrUPSNotifiEmail1&"', '"&pcStrUPSNotifiEmail2&"', '"&pcStrUPSNotifiEmail3&"', '"&pcStrUPSNotifiEmail4&"', '"&pcStrUPSNotifiEmail5&"', '"&pcStrUPSSaturdayDelivery&"', '"&pcStrUPSInsuredValue&"', '"&pcStrUPSVerbalConfirmation&"');"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	set rs=nothing
	call closedb() 
	
	response.Redirect "1_Step5.asp"
end if %>

<form name="form2" method="post" action="<%=pcPageName%>" class="pcForms">
	<table class="pcCPcontent">
		<tr> 
			<th colspan="2">UPS Shipping User Preferences </th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2"><p>Customizing your Preferences will save you time by remembering your most frequently used shipping options. The options you select will appear as defaults on your shipping pages. Please Note that only &quot;<span style="font-style: italic">Account Number</span>&quot; is require, you are not required to make a selection in every category.</p>			  </td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Service Type & Package Type</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
      <td align="right" valign="top">Type of service:</td>
		  <td align="left"><%
					'// Set Carrier Code to local
					pcv_strDropOptions = Session("pcAdminUPSService")
					%>
          <select name="UPSService">
            <option value="01" <%=pcf_SelectOption("UPSService","01")%>>UPS Next Day Air&reg;</option>
            <option value="02" <%=pcf_SelectOption("UPSService","02")%>>UPS 2nd Day Air&reg;</option>
            <option value="03" <%=pcf_SelectOption("UPSService","03")%>>UPS Ground</option>
            <option value="07" <%=pcf_SelectOption("UPSService","07")%>>UPS Worldwide Express<sup>SM</sup></option>
            <option value="08" <%=pcf_SelectOption("UPSService","08")%>>UPS Worldwide Expedited<sup>SM</sup></option>
            <option value="11" <%=pcf_SelectOption("UPSService","11")%>>UPS Standard</option>
            <option value="12" <%=pcf_SelectOption("UPSService","12")%>>UPS 3-Day Select&reg;</option>
            <option value="13" <%=pcf_SelectOption("UPSService","13")%>>UPS Next Day Air Saver&reg;</option>
            <option value="14" <%=pcf_SelectOption("UPSService","14")%>>UPS Next Day Air&reg; Early A.M.&reg;</option>
            <option value="54" <%=pcf_SelectOption("UPSService","54")%>>UPS Worldwide Express Plus<sup>SM</sup></option>
            <option value="59" <%=pcf_SelectOption("UPSService","59")%>>UPS 2nd Day Air A.M.&reg;</option>
          </select>
          <%pcs_RequiredImageTag "UPSService", false %>      </td>
		  </tr>
		<tr>
      <td align="right">Package Type :</td>
		  <td width="75%" align="left"><select name="UPSPackageType" id="UPSPackageType">
          <option value="01" <%=pcf_SelectOption("UPSPackageType","01")%>>UPS Letter</option>
          <option value="02" <%=pcf_SelectOption("UPSPackageType","02")%>>Your Packaging</option>
          <option value="03" <%=pcf_SelectOption("UPSPackageType","03")%>>UPS Tube</option>
          <option value="04" <%=pcf_SelectOption("UPSPackageType","04")%>>UPS PAK</option>
          <option value="21" <%=pcf_SelectOption("UPSPackageType","21")%>>UPS 25KG Box&reg;</option>
          <option value="24" <%=pcf_SelectOption("UPSPackageType","24")%>>UPS 10KG Box&reg;</option>
        </select>
        <%pcs_RequiredImageTag "UPSPackageType", false%></td>
		  </tr>
		<tr>
		  <td colspan="2" class="pcCPspacer"></td>
		  </tr>
		<tr>
		  <td align="right">Insured Value: </td>
		  <td><input name="UPSInsuredValue" type="text" id="UPSInsuredValue" value="<%=pcf_FillFormField("UPSInsuredValue", true)%>">
        <%pcs_RequiredImageTag "UPSInsuredValue", false%></td>
		  </tr>
		<tr>
		  <td colspan="2" class="pcCPspacer"></td>
		  </tr>
		<tr>
      <td align="right"><INPUT tabIndex="25" type="checkbox" value="1" name="UPSSaturdayDelivery" class="clearBorder" <%=pcf_CheckOption("UPSSaturdayDelivery", "1")%>></td>
		  <td><strong>Saturday Delivery </strong> </td>
		  </tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Payment Method </th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
      <td align="right">Payment Method :</td>
		  <td align="left"><select name="UPSPaymentMethod" id="UPSPaymentMethod">
        <option value="PrePaid" <%=pcf_SelectOption("UPSPaymentMethod","PrePaid")%>>PrePaid</option>
        <option value="BillThirdParty" <%=pcf_SelectOption("UPSPaymentMethod","BillThirdParty")%>>Bill 3rd Party</option>
        <option value="ConsigneeBilled" <%=pcf_SelectOption("UPSPaymentMethod","ConsigneeBilled")%>>Consignee Billing</option>
        <option value="FreightCollect" <%=pcf_SelectOption("UPSPaymentMethod","FreightCollect")%>>Freight Collect</option>
      </select>
        <%pcs_RequiredImageTag "UPSPaymentMethod", false%></td>
		  </tr>
		<tr>
      <td align="right"> UPS Account Number:</td>
		  <td align="left">
			<input name="UPSAccountNumber" type="text" id="UPSAccountNumber" value="<%=pcf_FillFormField("UPSAccountNumber", true)%>">
        <%pcs_RequiredImageTag "UPSAccountNumber", true%></td>
		  </tr>
		
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">UPS On Call Pickup&reg; Times </th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
      <td align="right">Shipment Ready At  :</td>
		  <td align="left"><select name="UPSReadyHours" id="UPSReadyHours">
        <option value="01" <%=pcf_SelectOption("UPSReadyHours","01")%>>01</option>
        <option value="02" <%=pcf_SelectOption("UPSReadyHours","02")%>>02</option>
        <option value="03" <%=pcf_SelectOption("UPSReadyHours","03")%>>03</option>
        <option value="04" <%=pcf_SelectOption("UPSReadyHours","04")%>>04</option>
        <option value="05" <%=pcf_SelectOption("UPSReadyHours","05")%>>05</option>
        <option value="06" <%=pcf_SelectOption("UPSReadyHours","06")%>>06</option>
        <option value="07" <%=pcf_SelectOption("UPSReadyHours","07")%>>07</option>
        <option value="08" <%=pcf_SelectOption("UPSReadyHours","08")%>>08</option>
        <option value="09" <%=pcf_SelectOption("UPSReadyHours","09")%>>09</option>
        <option value="10" <%=pcf_SelectOption("UPSReadyHours","10")%>>10</option>
        <option value="11" <%=pcf_SelectOption("UPSReadyHours","11")%>>11</option>
        <option value="12" <%=pcf_SelectOption("UPSReadyHours","12")%>>12</option>
      </select>
			:
			<select name="UPSReadyMinutes" id="UPSReadyMinutes">
				<option value="00" <%=pcf_SelectOption("UPSReadyMinutes","00")%>>00</option>
        <option value="01" <%=pcf_SelectOption("UPSReadyMinutes","01")%>>01</option>
        <option value="02" <%=pcf_SelectOption("UPSReadyMinutes","02")%>>02</option>
        <option value="03" <%=pcf_SelectOption("UPSReadyMinutes","03")%>>03</option>
        <option value="04" <%=pcf_SelectOption("UPSReadyMinutes","04")%>>04</option>
        <option value="05" <%=pcf_SelectOption("UPSReadyMinutes","05")%>>05</option>
        <option value="06" <%=pcf_SelectOption("UPSReadyMinutes","06")%>>06</option>
        <option value="07" <%=pcf_SelectOption("UPSReadyMinutes","07")%>>07</option>
        <option value="08" <%=pcf_SelectOption("UPSReadyMinutes","08")%>>08</option>
        <option value="09" <%=pcf_SelectOption("UPSReadyMinutes","09")%>>09</option>
				<% for iHHCnt=10 to 59 
					response.write "<option value="""&iHHCnt&""" "&pcf_SelectOption("UPSReadyMinutes",""&iHHCnt&"")&">"&iHHCnt&"</option>"
				next %>
			</select>
			&nbsp;<input name="UPSReadyAMPM" type="radio" value="AM" <%=pcf_CheckOption("UPSReadyAMPM","AM")%>>
			A.M. 
			&nbsp;<input name="UPSReadyAMPM" type="radio" value="PM" <%=pcf_CheckOption("UPSReadyAMPM","PM")%>>
			P.M. </td>
		</tr>
		<tr>
      <td align="right">Pick Up by :</td>
		  <td align="left">
        <select name="UPSPUHours" id="UPSPUHours">
        <option value="12" <%=pcf_SelectOption("UPSPUHours","12")%>>12</option>
        <option value="01" <%=pcf_SelectOption("UPSPUHours","01")%>>01</option>
        <option value="02" <%=pcf_SelectOption("UPSPUHours","02")%>>02</option>
        <option value="03" <%=pcf_SelectOption("UPSPUHours","03")%>>03</option>
        <option value="04" <%=pcf_SelectOption("UPSPUHours","04")%>>04</option>
        <option value="05" <%=pcf_SelectOption("UPSPUHours","05")%>>05</option>
        <option value="06" <%=pcf_SelectOption("UPSPUHours","06")%>>06</option>
        <option value="07" <%=pcf_SelectOption("UPSPUHours","07")%>>07</option>
        <option value="08" <%=pcf_SelectOption("UPSPUHours","08")%>>08</option>
        <option value="09" <%=pcf_SelectOption("UPSPUHours","09")%>>09</option>
        <option value="10" <%=pcf_SelectOption("UPSPUHours","10")%>>10</option>
        <option value="11" <%=pcf_SelectOption("UPSPUHours","11")%>>11</option>
      </select>
			:
			<select name="UPSPUMinutes" id="UPSPUMinutes">
				<option value="00" <%=pcf_SelectOption("UPSPUMinutes","00")%>>00</option>
        <option value="01" <%=pcf_SelectOption("UPSPUMinutes","01")%>>01</option>
        <option value="02" <%=pcf_SelectOption("UPSPUMinutes","02")%>>02</option>
        <option value="03" <%=pcf_SelectOption("UPSPUMinutes","03")%>>03</option>
        <option value="04" <%=pcf_SelectOption("UPSPUMinutes","04")%>>04</option>
        <option value="05" <%=pcf_SelectOption("UPSPUMinutes","05")%>>05</option>
        <option value="06" <%=pcf_SelectOption("UPSPUMinutes","06")%>>06</option>
        <option value="07" <%=pcf_SelectOption("UPSPUMinutes","07")%>>07</option>
        <option value="08" <%=pcf_SelectOption("UPSPUMinutes","08")%>>08</option>
        <option value="09" <%=pcf_SelectOption("UPSPUMinutes","09")%>>09</option>
				<% for iHHCnt=10 to 59 
					response.write "<option value="""&iHHCnt&""" "&pcf_SelectOption("UPSPUMinutes",""&iHHCnt&"")&">"&iHHCnt&"</option>"
				next %>
			</select>
P.M. </td>
		  </tr>
		
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Reference Numbers  </th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
      <td align="right">Reference #1 Type  :</td>
		  <td align="left"><select name="UPSRefNumber1" id="UPSRefNumber1">
				<option value="NONE">None Selected</option>
        <option value="AJ" <%=pcf_SelectOption("UPSRefNumber1","AJ")%>>Acct. Rec. Customer Acct.</option>
        <option value="AT" <%=pcf_SelectOption("UPSRefNumber1","AT")%>>Appropriation Number</option>
        <option value="BM" <%=pcf_SelectOption("UPSRefNumber1","BM")%>>Bill of Lading Number</option>
        <option value="9V" <%=pcf_SelectOption("UPSRefNumber1","9V")%>>COD Number</option>
				<option value="ON" <%=pcf_SelectOption("UPSRefNumber1","ON")%>>Dealer Order Number</option>
				<option value="DP" <%=pcf_SelectOption("UPSRefNumber1","DP")%>>Department Number</option>
				<option value="EI" <%=pcf_SelectOption("UPSRefNumber1","EI")%>>Employer's ID Number</option>
				<option value="3Q" <%=pcf_SelectOption("UPSRefNumber1","3Q")%>>FDA Product Code																																																																<option value="TJ" <%=pcf_SelectOption("UPSRefNumber1","TJ")%>>Federal Taxpayer ID No.</option>
				<option value="IK" <%=pcf_SelectOption("UPSRefNumber1","IK")%>>Invoice Number</option>
				<option value="MK" <%=pcf_SelectOption("UPSRefNumber1","MK")%>>Manifest Key Number</option>
				<option value="MJ" <%=pcf_SelectOption("UPSRefNumber1","MJ")%>>Model Number</option>
				<option value="PM" <%=pcf_SelectOption("UPSRefNumber1","PM")%>>Part Number</option>
				<option value="PC" <%=pcf_SelectOption("UPSRefNumber1","PC")%>>Production Code</option>
				<option value="PO" <%=pcf_SelectOption("UPSRefNumber1","PO")%>>Purchase Order Number</option>
				<option value="RQ" <%=pcf_SelectOption("UPSRefNumber1","RQ")%>>Purchase Req. Number</option>
				<option value="RZ" <%=pcf_SelectOption("UPSRefNumber1","RZ")%>>Return Authorization No.</option>
				<option value="SA" <%=pcf_SelectOption("UPSRefNumber1","SA")%>>Salesperson Number</option>
				<option value="SE" <%=pcf_SelectOption("UPSRefNumber1","SE")%>>Serial Number</option>
				<option value="SY" <%=pcf_SelectOption("UPSRefNumber1","SY")%>>Social Security Number</option>
				<option value="ST" <%=pcf_SelectOption("UPSRefNumber1","ST")%>>Store Number</option>
				<option value="TN" <%=pcf_SelectOption("UPSRefNumber1","TN")%>>Transaction Ref. No.</option>
				</option>
      </select></td>
		  </tr>
		<tr>
      <td align="right">Reference #1  :</td>
		  <td align="left"><input name="UPSRefData1" type="text" value="<%=pcf_FillFormField("UPSRefData1", false)%>" size="35"></td>
		  </tr>
		<tr>
      <td align="right">Reference #2 Type  :</td>
		  <td align="left"><select name="UPSRefNumber2" id="UPSRefNumber2">
				<option value="NONE">None Selected</option>
        <option value="AJ" <%=pcf_SelectOption("UPSRefNumber2","AJ")%>>Acct. Rec. Customer Acct.</option>
        <option value="AT" <%=pcf_SelectOption("UPSRefNumber2","AT")%>>Appropriation Number</option>
        <option value="BM" <%=pcf_SelectOption("UPSRefNumber2","BM")%>>Bill of Lading Number</option>
        <option value="9V" <%=pcf_SelectOption("UPSRefNumber2","9V")%>>COD Number</option>
        <option value="ON" <%=pcf_SelectOption("UPSRefNumber2","ON")%>>Dealer Order Number</option>
        <option value="DP" <%=pcf_SelectOption("UPSRefNumber2","DP")%>>Department Number</option>
        <option value="EI" <%=pcf_SelectOption("UPSRefNumber2","EI")%>>Employer's ID Number</option>
        <option value="3Q" <%=pcf_SelectOption("UPSRefNumber2","3Q")%>>FDA Product Code																																																																
        <option value="TJ" <%=pcf_SelectOption("UPSRefNumber2","TJ")%>>Federal Taxpayer ID No.</option>
        <option value="IK" <%=pcf_SelectOption("UPSRefNumber2","IK")%>>Invoice Number</option>
        <option value="MK" <%=pcf_SelectOption("UPSRefNumber2","MK")%>>Manifest Key Number</option>
        <option value="MJ" <%=pcf_SelectOption("UPSRefNumber2","MJ")%>>Model Number</option>
        <option value="PM" <%=pcf_SelectOption("UPSRefNumber2","PM")%>>Part Number</option>
        <option value="PC" <%=pcf_SelectOption("UPSRefNumber2","PC")%>>Production Code</option>
        <option value="PO" <%=pcf_SelectOption("UPSRefNumber2","PO")%>>Purchase Order Number</option>
        <option value="RQ" <%=pcf_SelectOption("UPSRefNumber2","RQ")%>>Purchase Req. Number</option>
        <option value="RZ" <%=pcf_SelectOption("UPSRefNumber2","RZ")%>>Return Authorization No.</option>
        <option value="SA" <%=pcf_SelectOption("UPSRefNumber2","SA")%>>Salesperson Number</option>
        <option value="SE" <%=pcf_SelectOption("UPSRefNumber2","SE")%>>Serial Number</option>
        <option value="SY" <%=pcf_SelectOption("UPSRefNumber2","EY")%>>Social Security Number</option>
        <option value="ST" <%=pcf_SelectOption("UPSRefNumber2","ST")%>>Store Number</option>
        <option value="TN" <%=pcf_SelectOption("UPSRefNumber2","TN")%>>Transaction Ref. No.</option>
        </option>
      </select></td>
		  </tr>
		<tr>
      <td align="right">Reference #2  :</td>
		  <td align="left"><input name="UPSRefData2" type="text" value="<%=pcf_FillFormField("UPSRefData2", false)%>" size="35"></td>
		  </tr>
		<tr>
		  <td colspan="2" class="pcCPspacer"></td>
		  </tr>
		<tr>
      <th colspan="2">C.O.D. Preferences </th>
		  </tr>
		<tr>
      <td align="right"><input type="checkbox" name="UPSCODPackage" value="1" class="clearBorder" <%=pcf_CheckOption("UPSCODPackage", "1")%>></td>
		  <td><span style="font-weight: bold">Set C.O.D. required by default </span></td>
		  </tr>
		<tr>
      <td align="right"><b>Collection Amount:</b></td>
		  <td align="left"><input name="UPSCODAmount" type="text" id="UPSCODAmount" value="<%=pcf_FillFormField("UPSCODAmount", false)%>">
          <%pcs_RequiredImageTag "UPSCODAmount", false%>      </td>
		  </tr>
		<tr>
      <td align="right"><b>Collection Currency:</b></td>
		  <td align="left"><select name="UPSCODCurrency" id="UPSCODCurrency">
          <option value="USD" <%=pcf_SelectOption("UPSCODCurrency","USD")%>>USD</option>
          <option value="AUD" <%=pcf_SelectOption("UPSCODCurrency","AUD")%>>AUD</option>
          <option value="CAD" <%=pcf_SelectOption("UPSCODCurrency","CAD")%>>CAD</option>
          <option value="CHF" <%=pcf_SelectOption("UPSCODCurrency","CHF")%>>CHF</option>
          <option value="CZK" <%=pcf_SelectOption("UPSCODCurrency","CZK")%>>CZK</option>
          <option value="DKK" <%=pcf_SelectOption("UPSCODCurrency","DKK")%>>DKK</option>
          <option value="EUR" <%=pcf_SelectOption("UPSCODCurrency","EUR")%>>EUR</option>
          <option value="GBP" <%=pcf_SelectOption("UPSCODCurrency","GBP")%>>GBP</option>
          <option value="GRD" <%=pcf_SelectOption("UPSCODCurrency","GRD")%>>GRD</option>
          <option value="HKD" <%=pcf_SelectOption("UPSCODCurrency","HKD")%>>HKD</option>
          <option value="HUF" <%=pcf_SelectOption("UPSCODCurrency","HUF")%>>HUF</option>
          <option value="INR" <%=pcf_SelectOption("UPSCODCurrency","INR")%>>INR</option>
          <option value="MXN" <%=pcf_SelectOption("UPSCODCurrency","MXN")%>>MXN</option>
          <option value="MYR" <%=pcf_SelectOption("UPSCODCurrency","MYR")%>>MYR</option>
          <option value="NOK" <%=pcf_SelectOption("UPSCODCurrency","NOK")%>>NOK</option>
          <option value="NZD" <%=pcf_SelectOption("UPSCODCurrency","NZD")%>>NZD</option>
          <option value="PLN" <%=pcf_SelectOption("UPSCODCurrency","PLN")%>>PLN</option>
          <option value="SEK" <%=pcf_SelectOption("UPSCODCurrency","SEK")%>>SEK</option>
          <option value="SGD" <%=pcf_SelectOption("UPSCODCurrency","SGD")%>>SGD</option>
          <option value="THB" <%=pcf_SelectOption("UPSCODCurrency","THB")%>>THB</option>
          <option value="TWD" <%=pcf_SelectOption("UPSCODCurrency","TWD")%>>TWD</option>
        </select>
          <%pcs_RequiredImageTag "UPSCODCurrency", false%>      </td>
		  </tr>
		<tr>
      <td align="right"><b>Collection Fund Type:</b></td>
		  <td align="left"><select name="UPSCODFunds" id="UPSCODFunds">
          <option value="0" <%=pcf_SelectOption("UPSCODFunds","0")%>>Cash</option>
          <option value="8" <%=pcf_SelectOption("UPSCODFunds","8")%>>Check, Cashier&rsquo;s Check or
            Money Order</option>
        </select>
          <%pcs_RequiredImageTag "UPSCODFunds", false%>      </td>
		  </tr>
		
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
				</tr>
			<tr>
				<th colspan="2">Shipment Notification</th>
				</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td align="right">
					<input type="checkbox" name="UPSShipmentNotification" value="1" class="clearBorder" <%=pcf_CheckOption("UPSShipmentNotification", "1")%>>				</td>
				<td><strong>Turn Shipment Notification on by default</strong> </td>
			</tr>
		<tr>
      <td align="right">Notification Type: </td>
		  <td align="left">
				<select name="UPSNotifiCode1" id="UPSNotifiCode1">
					<option value="0" <%=pcf_SelectOption("UPSNotifiCode1","0")%>>None Selected</option>
					<option value="6" <%=pcf_SelectOption("UPSNotifiCode1","6")%>>QVN Ship Notification</option>
					<option value="7" <%=pcf_SelectOption("UPSNotifiCode1","7")%>>QVN Exception Notification</option>
					<option value="8" <%=pcf_SelectOption("UPSNotifiCode1","8")%>>QVN Delivery Notification</option>
				</select>
				<%pcs_RequiredImageTag "UPSNotifiCode1", false%>
				&nbsp;E-Mail Address:
				<input name="UPSNotifiEmail1" type="text" id="UPSNotifiEmail1" value="<%=pcf_FillFormField("UPSNotifiEmail1", false)%>">
		    <%pcs_RequiredImageTag "UPSNotifiEmail1", false%></td>
		</tr>
		<tr>
      <td align="right">Notification Type:</td>
		  <td align="left">
				<select name="UPSNotifiCode2" id="UPSNotifiCode2">
					<option value="0" <%=pcf_SelectOption("UPSNotifiCode2","0")%>>None Selected</option>
					<option value="6" <%=pcf_SelectOption("UPSNotifiCode2","6")%>>QVN Ship Notification</option>
					<option value="7" <%=pcf_SelectOption("UPSNotifiCode2","7")%>>QVN Exception Notification</option>
					<option value="8" <%=pcf_SelectOption("UPSNotifiCode2","8")%>>QVN Delivery Notification</option>
        </select>
        <%pcs_RequiredImageTag "UPSNotifiCode2", false %>
		    &nbsp;E-Mail Address:
		    <input name="UPSNotifiEmail2" type="text" id="UPSNotifiEmail2" value="<%=pcf_FillFormField("UPSNotifiEmail2", false)%>">
		    <%pcs_RequiredImageTag "UPSNotifiEmail2", false%>			</td>
		</tr>
		<tr>
      <td align="right">Notification Type:</td>
		  <td align="left">
				<select name="UPSNotifiCode3" id="UPSNotifiCode3">
					<option value="0" <%=pcf_SelectOption("UPSNotifiCode3","0")%>>None Selected</option>
					<option value="6" <%=pcf_SelectOption("UPSNotifiCode3","6")%>>QVN Ship Notification</option>
					<option value="7" <%=pcf_SelectOption("UPSNotifiCode3","7")%>>QVN Exception Notification</option>
					<option value="8" <%=pcf_SelectOption("UPSNotifiCode3","8")%>>QVN Delivery Notification</option>
        </select>
				<%pcs_RequiredImageTag "UPSNotifiCode3", false %>
				&nbsp;E-Mail Address:
				<input name="UPSNotifiEmail3" type="text" id="UPSNotifiEmail3" value="<%=pcf_FillFormField("UPSNotifiEmail3", false)%>">
				<%pcs_RequiredImageTag "UPSNotifiEmail3", false %>			</td>
		</tr>
		<tr>
      <td align="right">Notification Type:</td>
		  <td align="left">
				<select name="UPSNotifiCode4" id="UPSNotifiCode4">
					<option value="0" <%=pcf_SelectOption("UPSNotifiCode4","0")%>>None Selected</option>
					<option value="6" <%=pcf_SelectOption("UPSNotifiCode4","6")%>>QVN Ship Notification</option>
					<option value="7" <%=pcf_SelectOption("UPSNotifiCode4","7")%>>QVN Exception Notification</option>
					<option value="8" <%=pcf_SelectOption("UPSNotifiCode4","8")%>>QVN Delivery Notification</option>
        </select>
				<%pcs_RequiredImageTag "UPSNotifiCode4", false %>
				&nbsp;E-Mail Address:
				<input name="UPSNotifiEmail4" type="text" id="UPSNotifiEmail4" value="<%=pcf_FillFormField("UPSNotifiEmail4", false)%>">
		    <%pcs_RequiredImageTag "UPSNotifiEmail4", false %>			</td>
		</tr>
		<tr>
      <td align="right">Notification Type:</td>
		  <td align="left">
				<select name="UPSNotifiCode5" id="UPSNotifiCode5">
					<option value="0" <%=pcf_SelectOption("UPSNotifiCode5","0")%>>None Selected</option>
					<option value="6" <%=pcf_SelectOption("UPSNotifiCode5","6")%>>QVN Ship Notification</option>
					<option value="7" <%=pcf_SelectOption("UPSNotifiCode5","7")%>>QVN Exception Notification</option>
					<option value="8" <%=pcf_SelectOption("UPSNotifiCode5","8")%>>QVN Delivery Notification</option>
        </select>
				<%pcs_RequiredImageTag "UPSNotifiCode5", false %>
				&nbsp;E-Mail Address:
				<input name="UPSNotifiEmail5" type="text" id="UPSNotifiEmail5" value="<%=pcf_FillFormField("UPSNotifiEmail5", false)%>">
		    <%pcs_RequiredImageTag "UPSNotifiEmail5", false %>			</td>
		</tr>
		<tr>
      <td align="right">&nbsp;</td>
		  <td align="left">&nbsp;</td>
		  </tr>
		<tr>
      <td align="right"><input type="checkbox" name="UPSVerbalConfirmation" value="1" class="clearBorder" <%=pcf_CheckOption("UPSVerbalConfirmation", "1")%>></td>
		  <td><strong>Verbal Confirmation </strong> </td>
		</tr>
		
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center">
				<input name="Submit1" type="submit" value="Update" class="submit2"> 
				&nbsp;
				<input name="back" type="button" onClick="javascript:history.back()" value="Back">			</td>
		</tr>
		<tr>
		  <td colspan="2" align="center"><br>
		  <table>
		    <tr>
		      <td width="58" valign="top" bgcolor="#FFFFFF"><div align="right"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50" /></div></td>
            <td width="457" valign="top" bgcolor="#FFFFFF"><div align="center">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br />              THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br />              UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</div></td>
          </tr>
		    </table></td></tr>
		</table>
	</form>
<!--#include file="AdminFooter.asp"-->