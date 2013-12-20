<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pageTitle="Shipping Wizard" %>
<% response.Buffer=true %>
<% section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="AdminHeader.asp"-->

<%
Dim connTemp,rs,query
call opendb()

pcv_IdOrder=request("orderID")
if pcv_IdOrder="" then
	pcv_IdOrder=0
end if
pcv_PackID=request("packID")
if pcv_PackID="" then
	pcv_PackID=0
end if
%>
<table class="pcCPcontent">
	<tr>
		<td valign="top">
            <table  border="0" cellpadding="0" cellspacing="0" width="60%">
                <tr>
                    <td width="100%">Order ID#: <b><%=(scpre+int(pcv_IdOrder))%></b></td>
                </tr>
            </table>
		</td>
	</tr>
</table>
	
<%	' Look up shipping method
'Get Tracking info
query="SELECT pcPackageInfo_UPSServiceCode, pcPackageInfo_ShipMethod, pcPackageInfo_TrackingNumber FROM pcPackageInfo WHERE idorder=" & pcv_IdOrder & " AND pcPackageInfo_ID="&pcv_PackID&";"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)
pcPackageInfo_UPSServiceCode=rs("pcPackageInfo_UPSServiceCode")
pcPackageInfo_ShipMethod=rs("pcPackageInfo_ShipMethod")
pcPackageInfo_TrackingNumber=rs("pcPackageInfo_TrackingNumber")
set rs=nothing
select case pcPackageInfo_UPSServiceCode
	case "E"
		pcv_serviceCode="Express Mail Label"
	case "D"
		pcv_serviceCode="Delivery Confirmation Mail Label"
	case "S"
		pcv_serviceCode="Signature Confirmation Mail Label"
end select

Dim pshipmentDetails, pSRF, pShippingMethod
query="SELECT shipmentDetails, SRF FROM orders WHERE idOrder=" & pcv_IdOrder & ";"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)
pshipmentDetails=rs("shipmentDetails")
pSRF=rs("SRF")
set rs=nothing
		
If pSRF="1" then
	pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_b")
else
	shipping=split(pshipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
			varShip="0"
			pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
		else
			Shipper=shipping(0)
			Service=shipping(1)
			Postage=trim(shipping(2))
			if ubound(shipping)=>3 then
				serviceHandlingFee=trim(shipping(3))
				if NOT isNumeric(serviceHandlingFee) then
					serviceHandlingFee=0
				end if
			else
				serviceHandlingFee=0
			end if
		end if
	else
		varShip="0"
		pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	end if 
end if
	
if pSRF="1" then
	pShippingMethod=pshipmentDetails
else
	if varShip<>"0"  then
		pShippingMethod=Service
	else
		pShippingMethod=pshipmentDetails
	end if 
end if

if request("smode")="E" then
	pShippingMethod="Express Mail"
else
	if pcPackageInfo_ShipMethod<>"" then
		pShippingMethod=pcPackageInfo_ShipMethod
	end if
end if
		
' Look up today's date
Dim varMonth, varDay, varYear
varMonth=Month(Date)
varDay=Day(Date)
varYear=Year(Date) 
dim dtInputStr
dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
if scDateFrmt="DD/MM/YY" then
	dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
end if

			
' Setup default Order Shipped message
		
' Get customer information 
query="SELECT idcustomer,orderDate FROM orders WHERE idOrder="& pcv_IdOrder
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=conntemp.execute(query)
pIdCustomer=rs("idcustomer")
pcv_orderDate=rs("orderDate")
set rs=nothing

query="SELECT name,lastname FROM customers WHERE idcustomer="& pIdCustomer
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=conntemp.execute(query)
pcv_CustomerName = rs("name")&" "&rs("lastname")

' Prepare message
customerShippedEmail=""
personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
personalmessage=replace(personalmessage,"<TODAY_DATE>",dtInputStr)
personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",pcv_CustomerName)
personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(pcv_IdOrder)))
personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pcv_orderDate))
If scShippedEmail<>"" Then
	customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf & vbCrLf
end if
CustomerShippedEmail=replace(CustomerShippedEmail,"//","/")
CustomerShippedEmail=replace(CustomerShippedEmail,"http:/","http://")
CustomerShippedEmail=replace(CustomerShippedEmail,"https:/","https://")
CustomerShippedEmail=replace(CustomerShippedEmail,"''",chr(39))

%>
	
<Form name="form1" method="post" action="sds_ShipOrderWizard3.asp?action=add&usps=1" class="pcForms">
	<input type="hidden" name="USPSLabelOnly" value="1">
    <input type="hidden" name="PackID" value="<%=pcv_PackID%>">
    <table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <th colspan="2">Specify Shipment Details</th>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td width="18%">Shipment Method:</td>
            <td width="82%"><input type="text" name="pcv_method" value="<%=pShippingMethod%>" size="40"></td>
        </tr>
		<tr>
		  <td colspan="2"><b>The &quot;Label Tracking Number&quot;  is associated with the  <%=pcv_serviceCode%> that has been generated for this shipment. If you have a Tracking Number to add for this shipment, such as a Priority Mail Tracking Number, you can enter it in the &quot;Tracking Number&quot; field.</b></td>
          </tr>
		<tr>
			<td>Tracking Number:</td>
			<td><input type="text" name="pcv_tracking" value="<%=pcPackageInfo_TrackingNumber%>" size="40"></td>
		</tr>
		<tr>
			<td>Shipped Date:</td>
			<td><input type="text" name="pcv_shippedDate" value="<%=dtInputStr%>" size="40"> <span class="pcCPnotes">Date Format: <%=scDateFrmt%></span></td>
		</tr>
		<tr>
			<td valign="top">Comments:</td>
			<td valign="top">
			<textarea name="pcv_AdmComments" size="40" rows="10" cols="65"><%=CustomerShippedEmail%></textarea>
			<div style="margin: 10px 15px 15px 0;" class="pcCPnotes">Please note that additional text will appear in the message that is emailed to the customer depending on whether this is a partial or final shipment, and depending on which shipping provider was used for the shipment, if any. The additional text can be edited by editing the file &quot;includes/languages_ship.asp". We recommend that you ship a few test orders in different scenarios to become familiar with the way the final message appears.</div>			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>
			<input type="submit" name="submit1" value="Finalize Shipment" class="submit2">
			&nbsp;<input type=button name="Back" value="Back" onclick="javascript:history.back();">
			<input type=hidden name="PrdList" value="<%=pcv_PrdList%>">
			<input type=hidden name="idorder" value="<%=pcv_IdOrder%>">			</td>
		</tr>
    </table>
</Form>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->