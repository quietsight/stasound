<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="ClearUPSSessions.asp"-->
<!--#include file="ClearFedExSessions.asp"-->
<!--#include file="ClearEndiciaSessions.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/pcUPSClass.asp"-->
<!--#include file="../includes/pcUSPSClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ShipFromSettings.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="pcPayPal_GetLiveStatus.asp"-->
<!--#include file="sm_inc.asp"-->
<%
'// Clear shipping-related session variables
'// in preparation for new order information
Session("pcAdminOrderID")=""
Session("pcAdminPackageCount")=""
Session("pcGlobalArray")=""
session("pcEDCMailClass")=""
session("pcEDCLabelType")=""
session("pcEDCShape1")=""
session("pcEDCPakValue1")=""
session("pcEDCIM1")=""
session("pcEDCInValue1")=""
%>
<!--#include file="../includes/EndiciaFunctions.asp"-->
<%
Dim pageTitle, Section
pageTitle="View &amp; Process Order"
pageIcon="pcv4_icon_orders.gif"
Section="orders"
%>
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<style>

	#pcCPmain ul {
		margin: 0px;
		padding: 0;
	}

	#pcCPmain ul li {
		margin: 0px;
	}

	div.TabbedMenu ul {
	text-align:left;
	margin:0 0 0 60px;
	padding:0;
	cursor:pointer;
	}

	div.TabbedMenu ul li {
	display:inline;
	list-style:none;
	margin:0 0.3em;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu ul li a {
	position:relative;
	z-index:0;
	font-weight:bold;
	border:solid 2px #e1e1e1;
	border-bottom-width:0;
	padding:0.3em;
	background-color:#ffffcc;
	color:black;
	text-decoration:none;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu ul li a.current {
	background-color:#F5F5F5;
	border:solid 2px #CCCCCC;
	border-bottom-width:0;
	position:relative;z-index:2;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu ul li a.current:hover {
	background-color:#F5F5F5;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu ul li a:hover {
	z-index:2;
	background-color:#F5F5F5;
	border-bottom:0;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu a span {display:none;}

	div.TabbedMenu a:hover span {
		display:block;
		position:absolute;
		top:2.3em;
		background-color:#F5F5F5;
		border-bottom:thin dotted gray;
		border-top:thin dotted gray;
		font-weight:normal;
		left:0;
		padding:1px 2px;
		cursor:pointer;
		font-size:12px;
	}

	div.TabbedPanes {
		padding: 1em;
		border: dashed 2px #CCCCCC;
		background-color: #F5F5F5;
		display: none;
		text-align:left;
		position:relative;z-index:1;
		margin-top:0.15em;
	}

</style>
<%
'// Define Tab Count
dim k, pcTabCount, strTabCnt
pcTabCount=7
strTabCnt=""
for k=1 to pcTabCount
	if k=1 then
		strTabCnt=strTabCnt&"""tab"&k&""""
	else
		strTabCnt=strTabCnt&",""tab"&k&""""
	end if
next
%>
<!--#include file="../includes/javascripts/pcCPTabs.asp"-->

<% Dim connTemp, qry_ID, query, rs

qry_ID=request.querystring("id")

call openDb()

'// FedEx Config
set objFedExClass = New pcFedExClass
pcv_strFedExEnabled = objFedExClass.pcf_FedExEnabled
pcv_strFedExPackagesExist = objFedExClass.pcf_FedExPackages(qry_ID)

'// UPS Config
set objUPSClass = New pcUPSClass
pcv_strUPSEnabled = objUPSClass.pcf_UPSEnabled
pcv_strUPSPackagesExist = objUPSClass.pcf_UPSPackages(qry_ID)

'// USPS Config
set objUSPSClass = New pcUSPSClass
pcv_strUSPSEnabled = objUSPSClass.pcf_USPSEnabled
pcv_strUSPSURLActive = objUSPSClass.pcf_USPSURLActive
pcv_strUSPSPackagesExist = objUSPSClass.pcf_USPSPackages(qry_ID)

'// PayPal Config
set objPayPalClass = New pcPayPalClass

If Not isNumeric(qry_ID) then
	response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

query="SELECT pcOrd_Archived,idcustomer, orderdate, Address, city, state, stateCode, zip, CountryCode, paymentDetails, shipmentDetails, shippingAddress, shippingCity, shippingStateCode, shippingState, shippingZip, pcOrd_shippingPhone, pcOrd_ShippingEmail, shippingCountryCode, idAffiliate, affiliatePay, discountDetails, pcOrd_GCDetails,pcOrd_GCAmount, taxAmount,  total, comments, orderStatus, processDate, shipDate, shipvia, trackingNum, returnDate, returnReason, ShippingFullName, ord_DeliveryDate, ord_OrderName, iRewardPoints, iRewardPointsCustAccrued, iRewardValue, address2, shippingCompany, shippingAddress2, taxDetails, adminComments, rmaCredit, DPs, gwAuthCode, gwTransId, gwTransParentId, paymentCode, SRF, ordShipType, ordPackageNum, ord_VAT, pcOrd_CatDiscounts, pcOrd_Payer, pcOrd_PaymentStatus, pcOrd_CustAllowSeparate, pcOrd_CustRequestStr, pcOrd_GCs, pcOrd_GcCode, pcOrd_GcUsed, pcOrd_IDEvent, pcOrd_GWTotal, pcOrd_Time, pcOrd_ShipWeight, pcOrd_GoogleIDOrder, pcOrd_CustomerIP, pcOrd_EligibleForProtection, pcOrd_AVSRespond, pcOrd_CVNResponse, pcOrd_PartialCCNumber, pcOrd_BuyerAccountAge, pcOrd_OrderKey FROM orders WHERE idOrder=" & qry_ID & ";"

Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)

Dim pidcustomer, porderdate, pAddress, pAddress2, pcity, pState, pstateCode, pzip, pCountryCode, ppaymentDetails, pshipmentDetails, pshippingCompany, pshippingAddress, pshippingAddress2, pshippingCity, pOrdshippingStateCode,pshippingState, pshippingZip, pshippingPhone, pshippingEmail, pshippingCountryCode, pidAffiliate, paffiliatePay, pdiscountDetails, ptaxAmount, ptotal, pcomments, porderStatus, pprocessDate, pshipDate, pshipvia, ptrackingNum, preturnDate, preturnReason, ptaxDetails,padminComments,prmaCredit,pSRF, pOrdShipType, pOrdPackageNum, gwTransParentId, pTotalAdj, paffiliatePayActual

pidorder=qry_ID
pOrdArc=rs("pcOrd_Archived")
if pOrdArc="" OR IsNULL(pOrdArc) then
	pOrdArc=0
end if
pidcustomer=rs("idcustomer")
porderdate=rs("orderdate")
pPayPalOriginalAuthorizedDate=porderdate
pPayPalAuthorizedDate=porderdate
porderdate=ShowDateFrmt(porderdate)
pAddress=rs("Address")
pcity=rs("city")
pstate=rs("state")
pstateCode=rs("stateCode")
pzip=rs("zip")
pCountryCode=rs("CountryCode")
ppaymentDetails=trim(rs("paymentDetails"))
pshipmentDetails=rs("shipmentDetails")
pshippingAddress=rs("shippingAddress")
pshippingCity=rs("shippingCity")
pOrdshippingStateCode=rs("shippingStateCode")
pshippingState=rs("shippingState")
pshippingZip=rs("shippingZip")
pshippingPhone=rs("pcOrd_shippingPhone")
pshippingEmail=rs("pcOrd_ShippingEmail")
pshippingCountryCode=rs("shippingCountryCode")
pidAffiliate=rs("idaffiliate")
paffiliatePay=rs("affiliatePay")
pdiscountDetails=rs("discountDetails")
GCDetails=rs("pcOrd_GCDetails")
GCAmount=rs("pcOrd_GCAmount")
if GCAmount="" OR IsNull(GCAmount) then
	GCAmount=0
end if
ptaxAmount=rs("taxAmount")
ptotal=rs("total")
pcomments=rs("comments")
porderStatus=rs("orderStatus")
	'// If the order is incomplete, a different page must be loaded
	if porderStatus=1 then
		set rs=nothing
		call closedb()
		response.redirect("ordDetailsIncomplete.asp?id=" & pidorder)
	end if
pprocessDate=rs("processDate")
pprocessDate=ShowDateFrmt(pprocessDate)
pshipDate=rs("shipDate")
pshipDate=ShowDateFrmt(pshipDate)
pshipvia=rs("shipvia")
ptrackingNum=rs("trackingNum")
preturnDate=rs("returnDate")
preturnDate=ShowDateFrmt(preturnDate)
preturnReason=rs("returnReason")
pshippingFullName=rs("ShippingFullName")
pord_DeliveryDate=rs("ord_DeliveryDate")
pord_OrderName=rs("ord_OrderName")
if isNULL(pord_OrderName) OR pord_OrderName="" then
	pord_OrderName="No Name"
end if
piRewardPoints=rs("iRewardPoints")
piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
piRewardValue=rs("iRewardValue")
pAddress2=rs("address2")
pshippingCompany=rs("shippingCompany")
pshippingAddress2=rs("shippingAddress2")
ptaxDetails=rs("taxDetails")
padminComments=rs("adminComments")
prmaCredit=rs("rmaCredit")
pcDPs=rs("DPs")
pcgwAuthCode=rs("gwAuthCode")
pcgwTransId=rs("gwTransId")
pcgwTransParentId=rs("gwTransParentId")
pcpaymentCode=rs("paymentCode")
pSRF=rs("SRF")
pOrdShipType=rs("ordShipType")
pOrdPackageNum=rs("ordPackageNum")
pord_VAT=rs("ord_VAT")
pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
	pcv_CatDiscounts="0"
end if
pcOrd_Payer=rs("pcOrd_Payer")
pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
if isNULL(pcv_PaymentStatus) OR pcv_PaymentStatus="" then
	pcv_PaymentStatus="0"
end if

'------------------------------
' Start Back-ordering
'------------------------------
' Is the customer allowed to request multiple shipments?
' This applies to orders that contain back-ordered items
' 0 = waiting to hear from customer
' 1 = customer wants separate shipments
' 2 = customer wants one shipment
pcv_CustAllow=rs("pcOrd_CustAllowSeparate")
	if isNULL(pcv_CustAllow) or pcv_CustAllow="" then
		pcv_CustAllow="0"
	end if

pcv_CustRequestStr=rs("pcOrd_CustRequestStr")
	if isNULL(pcv_CustRequestStr) or pcv_CustRequestStr="" then
		pcv_CustRequestStr="NA"
	end if
'------------------------------
' End Back-ordering
'------------------------------

'GGG Add-on start
pGCs=rs("pcOrd_GCs")
pGiftCode=rs("pcOrd_GcCode")
pGiftUsed=rs("pcOrd_GcUsed")
gIDEvent=rs("pcOrd_IDEvent")
if gIDEvent<>"" then
else
	gIDEvent="0"
end if
pGWTotal=rs("pcOrd_GWTotal")
if pGWTotal<>"" then
else
	pGWTotal="0"
end if
'GGG Add-on end

'------------------------------
' Order time: retrieve and format
pcv_OrderTime=rs("pcOrd_Time")
if pcv_OrderTime<>"" and not isNull(pcv_OrderTime) then
	if scDateFrmt="DD/MM/YY" then
		pcv_OrderTime = FormatDateTime(pcv_OrderTime, 4)
	else
		pcv_OrderTime = FormatDateTime(pcv_OrderTime, 3)
	end if
else
	pcv_OrderTime=""
end if
'------------------------------
pcOrd_ShipWeight=rs("pcOrd_ShipWeight")
pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder")
pcv_strCustomerIP = rs("pcOrd_CustomerIP")
pcv_strEligibleForProtection = rs("pcOrd_EligibleForProtection")
pcv_strAVSRespond = rs("pcOrd_AVSRespond")
pcv_strCVNResponse = rs("pcOrd_CVNResponse")
pcv_strPartialCCNumber = rs("pcOrd_PartialCCNumber")
pcv_strBuyerAccountAge = rs("pcOrd_BuyerAccountAge")
pcOrderKey=rs("pcOrd_OrderKey")

'// Calculate total adjusted for credits
if trim(prmaCredit)="" or IsNull(prmaCredit) then
	prmaCredit=0
end if
pTotalAdj=pTotal-prmaCredit

'// Check if the Customer is European Union
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)


set rs=nothing

'If it is the pending PayPal payment, get the Authorized Date
query="SELECT AuthorizedDate FROM pcPay_PayPal_Authorize WHERE idOrder=" & qry_ID & ";"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)
if not rs.eof then
	pPayPalAuthorizedDate=rs("AuthorizedDate")
end if
set rs=nothing

'---------------------------------------------------
' START SDBA
' Find out if the order contains drop-shipping
' and/or back-ordered products
'---------------------------------------------------

DIM pcv_haveDropPrds, pcv_haveBOPrds
pcv_haveDropPrds=0
pcv_haveBOPrds=0

' Loop through the products looking for drop-shipping products where a drop-shipper is specified
query="SELECT pcDropShipper_ID FROM ProductsOrdered WHERE idorder=" & qry_ID & " AND pcDropShipper_ID>0;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if not rs.eof then
	pcv_haveDropPrds=1 ' There are drop-shipping products
end if
set rs=nothing

' Loop through the products looking for back-ordered products
query="SELECT Products.idproduct,ProductsOrdered.pcPrdOrd_BackOrder FROM Products INNER JOIN ProductsOrdered ON (products.idproduct=ProductsOrdered.idproduct AND products.pcProd_IsDropShipped=0) WHERE productsOrdered.idorder=" & qry_ID & " AND productsOrdered.pcPrdOrd_BackOrder=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if not rs.eof then
	pcv_haveBOPrds=1 ' There are back-ordered products
end if
set rs=nothing

' Loop through the products looking for drop-shipping products where a drop-shipper is not specified
query="SELECT products.pcProd_IsDropShipped FROM Products INNER JOIN ProductsOrdered ON (products.idproduct=ProductsOrdered.idproduct AND products.pcProd_IsDropShipped=1) WHERE productsordered.idorder=" & qry_ID & " AND productsOrdered.pcDropShipper_ID=0;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if not rs.eof then
	pcv_haveDropPrds=1 ' There are drop-shipping products, but the drop-shipper is not known
end if
set rs=nothing

'---------------------------------------------------
' END SDBA
'---------------------------------------------------

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
varShip="1"
%>
<script>
	function winSale(fileName)
	{
		myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=650,height=300')
		myFloater.location.href=fileName;
	}

	function openshipwin(fileName)
	{
		myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=500,height=500')
		myFloater.location.href=fileName;
		checkwin();
	}
	function openwin(fileName)
	{
		myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=600,height=480')
		myFloater.location.href=fileName;
	}
	function checkwin()
	{
		if (myFloater.closed)
		{
			location="Orddetails.asp?id=<%=qry_ID%>&ActiveTab=2";
		}
		else
		{
			setTimeout('checkwin()',500);
		}
	}

	function CalPop(sInputName)
	{
		window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
	}

	function isDigit(s)
	{
	var test=""+s;
	var OK2reset;
	if(test=="."||test==","||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
			{
			return(true) ;
			}
			return(false);
	}

	function allDigit(s)
	{
			var test=""+s ;
			for (var k=0; k <test.length; k++)
			{
				var c=test.substring(k,k+1);
				if (isDigit(c)==false)
				{
					return (false);
				}
			}
			return (true);
	}

</script>
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf
pcs_JavaTextField	"name", true, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"email", true, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"address", true, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"city", true, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"zip", true, dictLanguage.Item(Session("language")&"_NewCust_3")
response.write "// If the user wants to change the status of the order..."&vbcrlf
response.write "if (theForm.checkreset.value == 'true')"&vbcrlf
response.write "{"&vbcrlf
response.write "	//reset the click flag"&vbcrlf
response.write "	theForm.checkreset.value = '';"&vbcrlf

response.write "	//proceed to reset the status if there's a new status value"&vbcrlf
response.write "	if (theForm.resetstat.value != theForm.oldstat.value)"&vbcrlf
response.write "	{"&vbcrlf
response.write "		return(true);"&vbcrlf
response.write "	}"&vbcrlf
response.write "	else"&vbcrlf
response.write "	{"&vbcrlf
response.write "		theForm.resetstat.focus();"&vbcrlf
response.write "		return(false);"&vbcrlf
response.write "	}"&vbcrlf
response.write "}"&vbcrlf

response.write "//Otherwise, check if the user clicked 'Cancel' or didn't use the reset button"&vbcrlf
response.write "else"&vbcrlf
response.write "{"&vbcrlf
response.write "	//If the user Cancelled (rather than clicking a different button)..."&vbcrlf
response.write "	if (theForm.checkreset.value == 'false')"&vbcrlf
response.write "	{"&vbcrlf
response.write "		//cancel the form"&vbcrlf
response.write "		theForm.checkreset.value = '';"&vbcrlf
response.write "		theForm.resetstat.focus();"&vbcrlf
response.write "		return(false);"&vbcrlf
response.write "	}"&vbcrlf
response.write "}"&vbcrlf
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>

<%if request("msg1")<>"" then%>
	<table class="pcCPcontent">
		<tr>
			<td>
				<div class="pcCPmessage">
					PayPal Gateway Transaction Error<br>
					<%=replace(replace(request("msg1"),"</div>",""),"<div align=""left"">","")%>
				</div>
			</td>
		</tr>
	</table>
<%end if%>

<%if request("msg2")<>"" then%>
	<table class="pcCPcontent">
		<tr>
			<td>
				<div class="pcCPmessage">
					NetSource Commerce Gateway Transaction Error<br>
					<%=replace(replace(request("msg2"),"</div>",""),"<div align=""left"">","")%>
				</div>
			</td>
		</tr>
	</table>
<%end if%>

<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer">
			<%
			msg=getUserInput(request.querystring("msg"),0)
			if msg<>"" then
				if msg="1" then %>
					<div class="pcCPmessage">
						We have sent your new order status to Google Merchant Center.
						Refresh this page in 30 seconds to view your updated order status.
						<br />
						<a href="Orddetails.asp?id=<%=qry_ID%>">Click Here to refresh this page &gt;&gt;</a>
					</div>
				<%
				else
					msg=replace(msg, "&gt;", ">")
					msg=replace(msg, "&lt;", "<")
					pcvMessageType=request.querystring("s")
					if not validNum(pcvMessageType) then pcvMessageType=0
					if pcvMessageType=1 then %>
						<div class="pcCPmessageSuccess"><%=msg%></div>
					<%
						else
					%>
						<div class="pcCPmessage"><%=msg%></div>
					<%
					end if
				end if
			end if %>
			<h2>Order #: <%=(scpre+int(qry_ID))%>&nbsp;|&nbsp;Date: <%=porderdate%><% if pcv_OrderTime<>"" then %>&nbsp;|&nbsp;Time: <%=pcv_OrderTime%><%end if%>&nbsp;|&nbsp;Total: <%=scCurSign&money(ptotalAdj)%>
			&nbsp;&nbsp;<a href="OrdInvoice.asp?id=<%=qry_ID%>" target="_blank"><img src="images/print_small.gif" alt="View Printer-Friendly Version"></a></h2>
		</td>
	</tr>
</table>
<% dim intActiveTab, pcTab1Style, pcTab2Style, pcTab3Style, pcTab4Style, pcTab5Style

intActiveTab=request("ActiveTab")
if intActiveTab="" then
	intActiveTab=1
end if

pcTab1Style="display:none"
pcTab2Style="display:none"
pcTab3Style="display:none"
pcTab4Style="display:none"
pcTab5Style="display:none"
pcTab1Class=""
pcTab2Class=""
pcTab3Class=""
pcTab4Class=""
pcTab5Class=""

select case intActiveTab
	case "1"
		pcTab1Style="display:block"
		pcTab1Class="current"
	case "2"
		pcTab2Style="display:block"
		pcTab2Class="current"
	case "3"
		pcTab3Style="display:block"
		pcTab3Class="current"
	case "4"
		pcTab4Style="display:block"
		pcTab4Class="current"
	case "5"
		pcTab5Style="display:block"
		pcTab5Class="current"
	case "6"
		pcTab6Style="display:block"
		pcTab6Class="current"
	case "7"
		pcTab7Style="display:block"
		pcTab7Class="current"
	case else
		pcTab1Style="display:block"
		pcTab1Class="current"
end select
%>

<form id="form2" name="form2" method="get" action="processOrder.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" name="ActiveTab" value="<%=intActiveTab%>">
	<input type="hidden" name="qry_ID" value="<%=request.querystring("id")%>">
	<input type="hidden" name="idcustomer" value="<%=pidcustomer%>">
	<table class="pcCPcontent">
		<tr>
			<td valign="top">
				<div class="TabbedMenu">
					<ul>
						<li><a id="tabs1" class="<%=pcTab1Class%>" onclick="change('tabs1', 'current');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab1');form2.ActiveTab.value = 1">General</a></li>
						<li><a id="tabs2" class="<%=pcTab2Class%>" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab2');form2.ActiveTab.value = 2">Product Details</a></li>
						<li><a id="tabs3" class="<%=pcTab3Class%>" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', 'current');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab3');form2.ActiveTab.value = 3">Process/Update</a></li>
						<li><a id="tabs4" class="<%=pcTab4Class%>" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Shipping Center</a></li>
						<li><a id="tabs5" class="<%=pcTab5Class%>" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', 'current');change('tabs6', '');change('tabs7', '');showTab('tab5');form2.ActiveTab.value = 5">Payment Status</a></li>
						<li><a id="tabs6" class="<%=pcTab6Class%>" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', 'current');change('tabs7', '');showTab('tab6');form2.ActiveTab.value = 6">Billing &amp; Shipping</a></li>
						<li><a id="tabs7" class="<%=pcTab7Class%>" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', 'current');showTab('tab7');form2.ActiveTab.value = 7">Other Details</a></li>
					</ul>
				</div>

				<%
				'// Set Flag for Google Checkout Order
				pcArrayPayment = split(ppaymentDetails,"||")
				PaymentType=pcArrayPayment(0)

				pcv_strDeactivateStatus=0
				if trim(PaymentType)="Google Checkout" then
					pcv_strDeactivateStatus=1
				end if

				'// Set Flag for PayPal Special Features
				pcv_WPPSpecialFeatures=0
				if trim(PaymentType)="PayPal Express Checkout" OR trim(PaymentType)="PayPal WebPayments Pro" OR trim(PaymentType)="PayPal Website Payments Pro"then
					'// Determine which API to use (US or UK)
					query="SELECT pcPay_PayPal.pcPay_PayPal_Partner, pcPay_PayPal.pcPay_PayPal_Vendor FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
					set rsPayPalType=Server.CreateObject("ADODB.Recordset")
					set rsPayPalType=conntemp.execute(query)
					pcPay_PayPal_Partner=rsPayPalType("pcPay_PayPal_Partner")
					pcPay_PayPal_Vendor=rsPayPalType("pcPay_PayPal_Vendor")
					if isNULL(pcPay_PayPal_Partner)=True then pcPay_PayPal_Partner=""
					if isNULL(pcPay_PayPal_Vendor)=True then pcPay_PayPal_Vendor=""
					if pcPay_PayPal_Partner<>"" AND pcPay_PayPal_Vendor<>"" then
						pcPay_PayPal_Version = "UK"
					else
						pcPay_PayPal_Version = "US"
					end if
					set rsPayPalType=nothing
					if pcPay_PayPal_Version = "US" then
						pcv_WPPSpecialFeatures=1
					end if
				end if
				'--------------
				' START TAB 1
				'--------------
				%>
				<div id="tab1" class="TabbedPanes" style="<%=pcTab1Style%>">
					<table class="pcCPcontent">
						<% if pcv_strDeactivateStatus=1 then %>
							<tr>
								<th colspan="2">Google Checkout Order – General Information<a name="top"></a></th>
							</tr>
						<% else %>
							<tr>
								<th colspan="2">General Information<a name="top"></a></th>
							</tr>
						<% end if %>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<% query="SELECT [name],lastName,pcCust_Guest,customerCompany,phone,email,customerType,idrefer FROM customers WHERE idcustomer="& pidcustomer
						Set rs=Server.CreateObject("ADODB.Recordset")
						Set rs=connTemp.execute(query)
						pname=rs("name")
						plastName=rs("lastName")
						pcvGuest=rs("pcCust_Guest")
						if pcvGuest="" OR IsNull(pcvGuest) then
							pcvGuest=0
						end if
						pcustomerCompany=rs("customerCompany")
						pphone=rs("phone")
						pemail=rs("email")
						pcustomerType=rs("customerType")
						pidrefer=rs("idrefer")
						set rs=nothing

						'see if an RMA request has been made for this order
						dim RMAVar,RMAStatus
						RMAVar=0
						RMAStatus=0
						query="SELECT idRMA,rmaNumber,rmaDateTime,rmaReturnStatus FROM PCReturns WHERE idOrder="& qry_ID
						Set rs=Server.CreateObject("ADODB.Recordset")
						Set rs=connTemp.execute(query)
						If NOT rs.eof then
							pIdRMA=rs("idRMA")
							prmaNumber=rs("rmaNumber")
							prmaDate=rs("rmaDateTime")
							prmaStatus=rs("rmaReturnStatus")
							if NOT isNull(prmaStatus) OR prmaStatus<>"" then
								RMAStatus=1
							end if
							RMAVar=1
						end if

						' Check to see if Product Reviews are active
						query = "SELECT pcRS_Active, pcRS_SendReviewReminder FROM pcRevSettings;"
						set rs=connTemp.execute(query)
						pcv_Active=rs("pcRS_Active")
						if isNull(pcv_Active) or pcv_Active="" then
							pcv_Active="0"
						end if
						pcv_ActiveReminder=rs("pcRS_SendReviewReminder")
						if isNull(pcv_ActiveReminder) or pcv_ActiveReminder="" then
							pcv_ActiveReminder="0"
						end if
						if pcv_ActiveReminder<>"1" then pcv_Active = "0"
						Set rs=Nothing
						%>
						<tr>
							<td colspan="2">
								Customer Name: <b><%=pname%>&nbsp;<%=pLastName%></b>
								<% if pcustomerCompany <> "" then %>
								 - <%=pcustomerCompany%>
								<% end if %>
								&nbsp;
								<a href="modCusta.asp?idcustomer=<%=pidcustomer%>">Edit</a>&nbsp;|&nbsp;<a href="viewCustOrders.asp?idcustomer=<%=pidcustomer%>">Order History</a><%if pcvGuest="0" then%>&nbsp;|&nbsp;<a href="adminPlaceOrder.asp?idcustomer=<%=pidcustomer%>" target="_blank">Place Order</a><%end if%><%if pcv_Active<>"0" then%>&nbsp;|&nbsp;<a href="javascript: if(confirm('Are you sure you want to send this customer an e-mail asking him or her to write a product review?')) {location='prv_AutoSendEmails.asp?idorder=<%=qry_ID%>';}">Send &quot;Write a Review&quot; Reminder</a><%end if%>
							</td>
						</tr>

						<%if pcOrderKey<>"" then%>
						<tr>
							<td colspan="2">Order Code: <strong><%=pcOrderKey%></strong><%if pcvGuest="1" then%>&nbsp;-&nbsp;Guest Checkout&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=317')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a><%end if%></td>
						</tr>
						<%end if%>

						<% os=porderStatus
						if os="2" then
							os="Pending"
						end if
						if os="3" then
							os="Processed"
						end if
						if os="4" then
							os="Shipped"
						end if
						if os="5" then
							os="Cancelled"
						end if
						if os="6" then
							os="Return"
						end if
						if os="7" then
							os="Partially Shipped"
						end if
						if os="8" then
							os="Shipping"
						end if
						if os="9" then
							os="Partially Returned"
						end if
						if os="10" then
							os="Delivered"
						end if
						if os="11" then
							os="Will Not Deliver" '// This is no longer used, but needs to remain here for backwards compatibility.
						end if
						if os="12" OR pOrdArc="1" then
							os="Archived"
						end if
						Select Case pcv_PaymentStatus
							Case "0": pcv_PayStatusName="Pending"
							Case "1": pcv_PayStatusName="Authorized"
							Case "2": pcv_PayStatusName="Paid"
							Case "3": pcv_PayStatusName="Declined"
							Case "4": pcv_PayStatusName="Cancelled"
							Case "5": pcv_PayStatusName="Cancelled By Google"
							Case "6": pcv_PayStatusName="Refunded"
							Case "7": pcv_PayStatusName="Charging"
							Case "8": pcv_PayStatusName="Voided"
							case "9": pcv_PayStatusName="Set to review by PayPal"
						End Select
						pcArrayPayment = split(ppaymentDetails,"||")
						PaymentType=pcArrayPayment(0)

						'// Google Checkout Order Status
						if pcv_strDeactivateStatus=1 then
							if os="Pending" then
								os="New (Pending)"
							end if
							if os="Processed" then
								os="Processing"
							end if
							if os="Cancelled" then
								os="Cancelled (Will Not Deliver)"
							end if
							if pcv_PaymentStatus="0" then
								pcv_PayStatusName="Reviewing"
							end if
						end if
						%>

						<tr>
							<td colspan="2">
								Current Order Status: <b><%=os%></b>
								&nbsp;&nbsp;&nbsp;
								<%
								if porderstatus="7" then %>
									<a href="#" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab2');form2.ActiveTab.value = 2">View partial shipment details</a>
									&nbsp;|&nbsp;
								<%
								end if
								if porderStatus>"1"  AND pOrdArc="0" AND pcv_strDeactivateStatus=0 then %>
									<a href="AdminEditOrder.asp?ido=<%=qry_ID%>">Edit Order</a>
								<% end if %>

								<% if pcv_strDeactivateStatus=0 then %>
									<% if porderStatus>"1" AND pOrdArc="0" then %>
										&nbsp;|&nbsp;
									<% end if %>
									<%if pOrdArc="0" then%>
									<a href="Orddetails.asp?id=<%=qry_ID%>&ActiveTab=3">
									<% if porderStatus="2" then %>Process Order<% else %>Update Order Status<% end if %> &gt;&gt;
									</a>
									<%end if%>
									<%if pOrdArc="0" then%>
										&nbsp;|&nbsp;<a href="ArcOrder.asp?id=<%=qry_ID%>&t=1">Archive this order &gt;&gt;</a>
									<%else%>
										&nbsp;|&nbsp;<a href="ArcOrder.asp?id=<%=qry_ID%>&t=2">Unarchive this order &gt;&gt;</a>
									<%end if%>

								<% end if %>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								Current Payment Status: <b><%=pcv_PayStatusName%></b>
								<% if pcv_strDeactivateStatus=0 then %>
								&nbsp;&nbsp;&nbsp;
								<a href="Orddetails.asp?id=<%=qry_ID%>&ActiveTab=5">Update Payment Status &gt;&gt;</a>
								<% end if %>
								<% if (pcv_strDeactivateStatus=1) AND (pcv_PaymentStatus="5") then %>
									<div style="padding:6px"><span class="pcCPnotes">Google Checkout has automatically cancelled this order for you. The reason might be that the order was considered fraudulent. Log into your Google Checkout account for more information.</span></div>
								<% end if %>
								<% if (pcv_strDeactivateStatus=1) AND (pcv_PaymentStatus="0") then %>
									<div style="padding:6px"><span class="pcCPnotes">You are not currently able to process this order because Google Checkout is still reviewing its validity.</span></div>
								<% end if %>
							</td>
						</tr>


						<% if pcv_haveDropPrds <> 0 then %>
						<tr>
							<td>This order contains products that will be <strong>drop-shipped</strong> by a third party. See <a href="#" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab2');form2.ActiveTab.value = 2">Product Details</a> for more information.</td>
						</tr>
						<% end if %>

						<% if pcv_haveBOPrds <> 0 then %>
						<tr>
							<td>This order contains products that might be <strong>back-ordered</strong>. See <a href="#" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab2');form2.ActiveTab.value = 2">Product Details</a> to see a list of back-ordered products.</td>
						</tr>
						<% end if %>


						<% 'Hide/show link to Help Desk if active
						If scShowHD <> 0 then %>
							<tr>
								<td colspan="2">
									<%if pcf_GetCustType(pidcustomer)=0 then%>
									Help Desk: <a href="adminviewallmsgs.asp?IDOrder=<%=qry_ID%>" target="_blank">View/Post</a>
									<%else%>
									Help Desk: <em>the Help Desk is disabled for this order: this customer is a &quot;Guest&quot; and would not be able to view/reply to the ticket. <a href="modcusta.asp?idcustomer=<%=pidcustomer%>" target="_blank">Change the customer status</a> to use the Help Desk.</em>
									<%end if%>
								</td>
							</tr>
						<% end if %>

						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">Order Name:
								<input type="text" name="ord_OrderName" size="30" maxsize="50" value="<%=pord_OrderName%>">
								<%
								'// If delivery date and/or time are shown, hide this button since another button is shown below
								if DFShow<>"1" and TFShow<>"1" then
								%>
									<input type="submit" name="Submit11" value="Update Order Name" class="submit2">
								<%
								end if
								%>
							</td>
						</tr>

						<%
						'// START - Show delivery date field
						if DFShow="1" then
						%>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2">
									Delivery Date:&nbsp;
									<% if pord_DeliveryDate<>"" and pord_DeliveryDate<>"1/1/1900" then
										Date1=month(pord_DeliveryDate) & "/" & day(pord_DeliveryDate) & "/" & year(pord_DeliveryDate)
										Date1=showDateFrmt(Date1)
									else
										Date1=""
									end if%>
									<input type="text" name="DF1" size="30" value="<%=Date1%>"> &nbsp;<a href="javascript:CalPop('document.form2.DF1');"><img SRC="../Calendar/icon_Cal.gif"></a>
								</td>
							</tr>
						<%
						'// END - Show delivery date field
						end if
						%>

						<%
						'// START - Show delivery time field
						if TFShow="1" then
						%>
							<tr>
								<td colspan="2">
									Delivery Time:&nbsp;
								<%
								if pord_DeliveryDate<>"" and pord_DeliveryDate<>"1/1/1900" then
									if cint(hour(pord_DeliveryDate))+cint(minute(pord_DeliveryDate))>0 then
										if hour(pord_DeliveryDate)<10 then
											hour1="0" & hour(pord_DeliveryDate)
										else
											hour1="" & hour(pord_DeliveryDate)
										end if
										if minute(pord_DeliveryDate)<10 then
											min1="0" & minute(pord_DeliveryDate)
										else
											min1="" & minute(pord_DeliveryDate)
										end if
										Time1=hour1 & ":" & min1
									else
										Time1=""
									end if
								else
									Time1=""
								end if%>
									<select name="TF1">
									<%if Time1="" then%>
										<option value=""><%response.write dictLanguage.Item(Session("language")&"_viewCatOrder_6")%></option>
									<%end if%>
									<% if scDateFrmt="DD/MM/YY" then %>
										<option value="7:00" <% If Time1 = "7:00" Then %>SELECTED<% End If %>>7:00</option>
										<option value="7:30" <% If Time1 = "7:30" Then %>SELECTED<% End If %>>7:30</option>
										<option value="8:00" <% If Time1 = "8:00" Then %>SELECTED<% End If %>>8:00</option>
										<option value="8:30" <% If Time1 = "8:30" Then %>SELECTED<% End If %>>8:30</option>
										<option value="9:00" <% If Time1 = "9:00" Then %>SELECTED<% End If %>>9:00</option>
										<option value="9:30" <% If Time1 = "9:30" Then %>SELECTED<% End If %>>9:30</option>
										<option value="10:00" <% If Time1 = "10:00" Then %>SELECTED<% End If %>>10:00</option>
										<option value="10:30" <% If Time1 = "10:30" Then %>SELECTED<% End If %>>10:30</option>
										<option value="11:00" <% If Time1 = "11:00" Then %>SELECTED<% End If %>>11:00</option>
										<option value="11:30" <% If Time1 = "11:30" Then %>SELECTED<% End If %>>11:30</option>
										<option value="12:00" <% If Time1 = "12:00" Then %>SELECTED<% End If %>>12:00</option>
										<option value="12:30" <% If Time1 = "12:30" Then %>SELECTED<% End If %>>12:30</option>
										<option value="13:00" <% If Time1 = "13:00" Then %>SELECTED<% End If %>>13:00</option>
										<option value="13:30" <% If Time1 = "13:30" Then %>SELECTED<% End If %>>13:30</option>
										<option value="14:00" <% If Time1 = "14:00" Then %>SELECTED<% End If %>>14:00</option>
										<option value="14:30" <% If Time1 = "14:30" Then %>SELECTED<% End If %>>14:30</option>
										<option value="15:00" <% If Time1 = "15:00" Then %>SELECTED<% End If %>>15:00</option>
										<option value="15:30" <% If Time1 = "15:30" Then %>SELECTED<% End If %>>15:30</option>
										<option value="16:00" <% If Time1 = "16:00" Then %>SELECTED<% End If %>>16:00</option>
										<option value="16:30" <% If Time1 = "16:30" Then %>SELECTED<% End If %>>16:30</option>
										<option value="17:00" <% If Time1 = "17:00" Then %>SELECTED<% End If %>>17:00</option>
										<option value="17:30" <% If Time1 = "17:30" Then %>SELECTED<% End If %>>17:30</option>
										<option value="18:00" <% If Time1 = "18:00" Then %>SELECTED<% End If %>>18:00</option>
										<option value="18:30" <% If Time1 = "18:30" Then %>SELECTED<% End If %>>18:30</option>
										<option value="19:00" <% If Time1 = "19:00" Then %>SELECTED<% End If %>>19:00</option>
										<option value="19:30" <% If Time1 = "19:30" Then %>SELECTED<% End If %>>19:30</option>
										<option value="20:00" <% If Time1 = "20:00" Then %>SELECTED<% End If %>>20:00</option>
										<option value="20:30" <% If Time1 = "20:30" Then %>SELECTED<% End If %>>20:30</option>
										<option value="21:00" <% If Time1 = "21:00" Then %>SELECTED<% End If %>>21:00</option>
									<% Else  %>
										<option value="7:00 AM" <% If Time1 = "7:00" Then %>SELECTED<% End If %>>7:00 AM</option>
										<option value="7:30 AM" <% If Time1 = "7:30" Then %>SELECTED<% End If %>>7:30 AM</option>
										<option value="8:00 AM" <% If Time1 = "8:00" Then %>SELECTED<% End If %>>8:00 AM</option>
										<option value="8:30 AM" <% If Time1 = "8:30" Then %>SELECTED<% End If %>>8:30 AM</option>
										<option value="9:00 AM" <% If Time1 = "9:00" Then %>SELECTED<% End If %>>9:00 AM</option>
										<option value="9:30 AM" <% If Time1 = "9:30" Then %>SELECTED<% End If %>>9:30 AM</option>
										<option value="10:00 AM" <% If Time1 = "10:00" Then %>SELECTED<% End If %>>10:00 AM</option>
										<option value="10:30 AM" <% If Time1 = "10:30" Then %>SELECTED<% End If %>>10:30 AM</option>
										<option value="11:00 AM" <% If Time1 = "11:00" Then %>SELECTED<% End If %>>11:00 AM</option>
										<option value="11:30 AM" <% If Time1 = "11:30" Then %>SELECTED<% End If %>>11:30 AM</option>
										<option value="12:00 PM" <% If Time1 = "12:00" Then %>SELECTED<% End If %>>12:00 PM</option>
										<option value="12:30 PM" <% If Time1 = "12:30" Then %>SELECTED<% End If %>>12:30 PM</option>
										<option value="1:00 PM" <% If Time1 = "13:00" Then %>SELECTED<% End If %>>1:00 PM</option>
										<option value="1:30 PM" <% If Time1 = "13:30" Then %>SELECTED<% End If %>>1:30 PM</option>
										<option value="2:00 PM" <% If Time1 = "14:00" Then %>SELECTED<% End If %>>2:00 PM</option>
										<option value="2:30 PM" <% If Time1 = "14:30" Then %>SELECTED<% End If %>>2:30 PM</option>
										<option value="3:00 PM" <% If Time1 = "15:00" Then %>SELECTED<% End If %>>3:00 PM</option>
										<option value="3:30 PM" <% If Time1 = "15:30" Then %>SELECTED<% End If %>>3:30 PM</option>
										<option value="4:00 PM" <% If Time1 = "16:00" Then %>SELECTED<% End If %>>4:00 PM</option>
										<option value="4:30 PM" <% If Time1 = "16:30" Then %>SELECTED<% End If %>>4:30 PM</option>
										<option value="5:00 PM" <% If Time1 = "17:00" Then %>SELECTED<% End If %>>5:00 PM</option>
										<option value="5:30 PM" <% If Time1 = "17:30" Then %>SELECTED<% End If %>>5:30 PM</option>
										<option value="6:00 PM" <% If Time1 = "18:00" Then %>SELECTED<% End If %>>6:00 PM</option>
										<option value="6:30 PM" <% If Time1 = "18:30" Then %>SELECTED<% End If %>>6:30 PM</option>
										<option value="7:00 PM" <% If Time1 = "19:00" Then %>SELECTED<% End If %>>7:00 PM</option>
										<option value="7:30 PM" <% If Time1 = "19:30" Then %>SELECTED<% End If %>>7:30 PM</option>
										<option value="8:00 PM" <% If Time1 = "20:00" Then %>SELECTED<% End If %>>8:00 PM</option>
										<option value="8:30 PM" <% If Time1 = "20:30" Then %>SELECTED<% End If %>>8:30 PM</option>
										<option value="9:00 PM" <% If Time1 = "21:00" Then %>SELECTED<% End If %>>9:00 PM</option>
									<% end if %>
									</select>
								</td>
							</tr>
							<%
							end if
							'// END - Show delivery time field

							'GGG Add-on start
							'// Show event information
							if gIDEvent<>"0" then

								query="select pcEvents.pcEv_name,pcEvents.pcEv_Date,customers.name,customers.lastname from pcEvents,Customers where Customers.idcustomer=pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
								set rs1=connTemp.execute(query)

								geName=rs1("pcEv_name")
								geDate=rs1("pcEv_Date")
								if year(geDate)="1900" then
									geDate=""
								end if
								if gedate<>"" then
									if scDateFrmt="DD/MM/YY" then
										gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
									else
										gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
									end if
								end if
								gReg=rs1("name") & " " & rs1("lastname")
								%>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td>Event Name:</td>
									<td><%=gename%></td>
								</tr>
								<tr>
									<td>Event Date:</td>
									<td><%=geDate%></td>
								</tr>
								<tr>
									<td>Registrant's Name:</td>
									<td><%=gReg%></td>
								</tr>
							<%
							end if
							'// END - Show event inforation
							'GGG Add-on end

							'// IF there is delivery date, delivery time, or event information, show Update button
							if gIDEvent<>"0" or TFShow="1" or DFShow="1" then
							%>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="submit" name="Submit11" value="Update Order Information" class="submit2">
								</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
						<%end if%>

						<%
						'SB S
						%><!--#include file="OrdDetails_SB.asp"--><%
						'SB E
						%>

						<%
						'///////////////////////////////////////////////////////
						'/// START: IF PayPal WPP
						'///////////////////////////////////////////////////////
						'SB S
						if pcv_strGUID="" then
						%><!--#include file="OrdDetails_PayPal.asp"--><%
						end if
						'SB E
						'///////////////////////////////////////////////////////
						'/// END: IF PayPal WPP
						'///////////////////////////////////////////////////////

						'///////////////////////////////////////////////////////
						'/// START: IF Google Checkout
						'///////////////////////////////////////////////////////
						if pcv_strDeactivateStatus=1 then
						%>
						<tr>
							<td colspan="2">
								Google Order ID: <strong><%=pcv_strGoogleIDOrder%></strong>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Google Checkout Order – Risk Information</th>
						</tr>
						<tr>
							<td colspan="2">
								<%
								if pcv_strCustomerIP="" then
									pcv_strCustomerIP="Not Available"
								end if
								%>
								Customer IP: <strong><%=pcv_strCustomerIP %>  </strong>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<%
								if pcv_strEligibleForProtection="" then
									pcv_strEligibleForProtection="Not Available"
								end if

								if pcv_strEligibleForProtection="true" then pcv_strEligibleForProtection="Yes"
								if pcv_strEligibleForProtection="false" then pcv_strEligibleForProtection="No"
								%>
								Eligible For Protection: <strong><%=pcv_strEligibleForProtection %></strong>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<%
								select case pcv_strAVSRespond
								case "Y": pcv_strAVSRespond="Full AVS match (address and postal code)"
								case "P": pcv_strAVSRespond="Partial AVS match (postal code only)"
								case "A": pcv_strAVSRespond="Partial AVS match (address only)"
								case "N": pcv_strAVSRespond="no AVS match"
								case "U": pcv_strAVSRespond="AVS not supported by issuer"
								case else: pcv_strAVSRespond="Not Available"
								end select
								%>
								AVS Response: <strong><%=pcv_strAVSRespond%></strong>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<%
								select case pcv_strCVNResponse
								case "M": pcv_strCVNResponse="CVN match"
								case "N": pcv_strCVNResponse="No CVN match"
								case "U": pcv_strCVNResponse="CVS not available"
								case "E": pcv_strCVNResponse="CVN error"
								case else: pcv_strCVNResponse="Not Available"
								end select
								%>
								CVN Response: <strong><%=pcv_strCVNResponse%></strong>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<%
								if pcv_strPartialCCNumber="" then
									pcv_strPartialCCNumber="Not Available"
								end if
								%>
								Partial CC Number: <strong><%=pcv_strPartialCCNumber%></strong>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<%
								if pcv_strBuyerAccountAge="" then
									pcv_strBuyerAccountAge="Not Available"
								end if
								%>
								Buyer Account Age (Days): <strong><%=pcv_strBuyerAccountAge%></strong>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<%
						'///////////////////////////////////////////////////////
						'/// END: IF Google Checkout
						'///////////////////////////////////////////////////////
						end if
						%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Administrator Comments</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">
								<% If pcomments <> "" then %><div><strong>Note</strong> that this order also contains <a href="javascript:;" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', 'current');showTab('tab7');form2.ActiveTab.value = 7">comments entered by the customer</a> while placing the order.</div><% end if %>
								<%
								Dim padminCommentsExist
								padminCommentsExist=0
								if trim(padminComments)<>"" and not IsNull(padminComments) then
									padminCommentsExist=1
									'padminComments = replace(padminComments, "vbCrLf", "<br>")
									padminComments = replace(padminComments, Chr(13), "<br>")
									%>
									<div style="width: 700px; height: auto; background-color: #f1f1f1; padding: 10px; margin-bottom: 15px; text-align: left; border: 1px solid #e1e1e1;"><%=padminComments%></div>
								<%
								end if
								%>
								<div>
									<a href="javascript: if(confirm('You are about to leave this page. Unsaved changes to this order will be lost. Would you like to continue?')) {location='OrdDetailsComments.asp?idorder=<%=qry_ID%>';}">Add<% if padminCommentsExist=1 then%>/Edit<%end if%> Administrator Comments</a>
								</div>
								<div class="pcSmallText">Any comments you add here are for administrative purposes only. They are never shown to the customer (i.e. they never appear on any order-related page in the storefront or e-mails sent to the customer).</div>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<% 'if order has any edits in the audit log, list them here
						query="SELECT pcAdminAuditLogID, idAdmin, idOrder, pcAdminAuditDate, pcAdminAuditPage FROM pcAdminAuditLog WHERE idOrder = " & qry_ID
						set rsAuditObj=connTemp.execute(query)
						if NOT rsAuditObj.eof then %>
						<tr>
							<th colspan="2">Audit Logs</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">This order has been edited on the following dates.<br><br></td>
						</tr>
						<% do until rsAuditObj.eof %>
							<tr>
								<td><strong>Date:</strong> <% = rsAuditObj("pcAdminAuditDate") %>&nbsp;&nbsp;&nbsp;&nbsp;<strong>Admin ID:</strong> <% = rsAuditObj("idAdmin") %></td>
						</tr>
						<% rsAuditObj.MoveNext
						loop
						%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<% end if %>
					</table>
					</div>

					<%
					'--------------
					' END TAB 1
					'--------------

					'--------------
					' START TAB 2
					'--------------
					%>

					<div id="tab2" class="TabbedPanes" style="<%=pcTab2Style%>">
					<table class="pcCPcontent">
						<tr>
							<th>Product Details
								<% if porderStatus>"1" AND pcv_strDeactivateStatus=0 then %>
								 - <a href="AdminEditOrder.asp?ido=<%=qry_ID%>">Edit</a>
								<% end if%>
							</th>
						</tr>
						<tr>
							<td class="pcCPspacer"></td>
						</tr>
						<tr>
							<td>
								<table width="100%" cellpadding="5" cellspacing="0" border="1" class="invoice">
									<tr style="background-color:#e1e1e1;">
										<td width="8%" class="invoice"><b>QTY</b></td>
										<td class="invoice"><b>SKU - DESCRIPTION</b></td>
										<td width="16%" class="invoice" align="right"><b>UNIT PRICE</b></td>
										<td width="12%" class="invoice" align="right"><b>TOTAL</b></td>
									</tr>
									<% query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.pcSubscription_ID, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.unitPrice, ProductsOrdered.xfdetails"
									'BTO ADDON-S
									if scBTO=1 then
										query=query&", ProductsOrdered.idconfigSession"
									end if
									'BTO ADDON-E
									query=query&", ProductsOrdered.rmaSubmitted,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts, pcDropShipper_ID, pcPrdOrd_BackOrder, pcPrdOrd_SentNotice, ProductsOrdered.pcPO_GWOpt,pcPO_GWNote,pcPO_GWPrice,pcPrdOrd_BundledDisc,pcSC_ID FROM ProductsOrdered WHERE ProductsOrdered.idOrder=" & qry_ID & ";"

									Set rs=Server.CreateObject("ADODB.Recordset")
									set rs=connTemp.execute(query)
									dim intTotalWeight
									intTotalWeight=int(0)
									Dim strCol
									strCol="#F1F1F1"
									Dim intPcCount
									intPcCount = 0
									Dim pSubscriptionID
									tmpAllPrdSubTotal=0
									Dim TotPrdSubTotal
									TotPrdSubTotal = 0
									Dim pTempSku
									pTempSku = ""
									Do until rs.eof
										pidProduct=rs("idProduct")
										pquantity=rs("quantity")
										pSubscriptionID=rs("pcSubscription_ID")
										If pSubscriptionID="" or IsNULL(pSubscriptionID) Then
											pSubscriptionID = 0
										End If

										'// Product Options Arrays
										pcv_strSelectedOptions = ""
										pcv_strOptionsPriceArray = ""
										pcv_strOptionsArray = ""
										pcv_strSelectedOptions = rs("pcPrdOrd_SelectedOptions") ' Column 11
										pcv_strOptionsPriceArray = rs("pcPrdOrd_OptionsPriceArray") ' Column 25
										pcv_strOptionsArray = rs("pcPrdOrd_OptionsArray") ' Column 4

										punitPrice=rs("unitPrice")
										pxdetails=rs("xfdetails")
										pxdetails=replace(pxdetails,"|","<br>")
										pxdetails=replace(pxdetails,"::",":")
										if scBTO=1 then
											pidConfigSession=rs("idConfigSession")
										end if
										prmaSubmitted=rs("rmaSubmitted")
										QDiscounts=rs("QDiscounts")
										ItemsDiscounts=rs("ItemsDiscounts")

										pcv_IntDropShipperId=rs("pcDropShipper_ID")
										if IsNull(pcv_IntDropShipperId) or pcv_IntDropShipperId="" then
											pcv_IntDropShipperId=0
										end if

										pcv_IntDropNotified=rs("pcPrdOrd_SentNotice")
										if IsNull(pcv_IntDropNotified) or pcv_IntDropNotified="" then
											pcv_IntDropNotified=0
										end if


										pcv_IntBackOrder=rs("pcPrdOrd_BackOrder")
										if IsNull(pcv_IntBackOrder) or pcv_IntBackOrder="" then
											pcv_IntBackOrder=0
										end if

										'GGG Add-on start
										pGWOpt=rs("pcPO_GWOpt")
										if not pGWOpt<>"" then
											pGWOpt="0"
										end if
										pGWText=rs("pcPO_GWNote")
										pGWPrice=rs("pcPO_GWPrice")
										if not pGWPrice<>"" then
											pGWPrice="0"
										end if
										'GGG Add-on end
										pcPrdOrd_BundledDisc=rs("pcPrdOrd_BundledDisc")
										pcSCID=rs("pcSC_ID")
										if pcSCID="" Or (IsNull(pcSCID)) then
											pcSCID=0
										end if

										query="SELECT description,sku,weight,pcprod_QtyToPound,stock FROM products WHERE idproduct="& pidProduct
										Set rstemp=Server.CreateObject("ADODB.Recordset")
										set rstemp=connTemp.execute(query)
										pdescription=rstemp("description")
										pSKU=rstemp("sku")
										If (pTempSku&""<>"") AND (pTempSKU<>pSKU) Then
											TotPrdSubTotal = 0
										End If
										pWeight=rstemp("weight")
										pStock=rstemp("stock")
										pcv_QtyToPound=rstemp("pcprod_QtyToPound")
										if pcv_QtyToPound>0 then
											pWeight=(16/pcv_QtyToPound)
											if scShipFromWeightUnit="KGS" then
												pWeight=(1000/pcv_QtyToPound)
											end if
										end if
										intTotalWeight=intTotalWeight+(pWeight*pquantity)
										set rstemp=nothing

										'BTO ADDON-S
										err.number=0
										TotalUnit=0
										If scBTO=1 then
											pIdConfigSession=trim(pidconfigSession)
											if pIdConfigSession<>"0" then
												query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
												set rsConfigObj=conntemp.execute(query)
												if err.number <> 0 then
													set rsConfigObj=nothing
													call closedb()
													response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails: "&err.description)
												end if
												stringProducts=rsConfigObj("stringProducts")
												stringValues=rsConfigObj("stringValues")
												stringCategories=rsConfigObj("stringCategories")
												stringQuantity=rsConfigObj("stringQuantity")
												stringPrice=rsConfigObj("stringPrice")
												ArrProduct=Split(stringProducts, ",")
												ArrValue=Split(stringValues, ",")
												ArrCategory=Split(stringCategories, ",")
												ArrQuantity=Split(stringQuantity, ",")
												ArrPrice=Split(stringPrice, ",")
												set rsConfigObj=nothing
												for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)

													query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
													set rsQ=connTemp.execute(query)
													tmpMinQty=1
													if not rsQ.eof then
														tmpMinQty=rsQ("pcprod_minimumqty")
														if IsNull(tmpMinQty) or tmpMinQty="" then
															tmpMinQty=1
														else
															if tmpMinQty="0" then
																tmpMinQty=1
															end if
														end if
													end if
													set rsQ=nothing
													tmpDefault=0
													query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
													set rsQ=connTemp.execute(query)
													if not rsQ.eof then
														tmpDefault=rsQ("cdefault")
														if IsNull(tmpDefault) or tmpDefault="" then
															tmpDefault=0
														else
															if tmpDefault<>"0" then
																tmpDefault=1
															end if
														end if
													end if
													set rsQ=nothing

													query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
													set rsConfigObj=conntemp.execute(query)
													if NOT isNumeric(ArrQuantity(i)) then
														pIntQty=1
													else
														pIntQty=ArrQuantity(i)
													end if
													if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
														if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
															if tmpDefault=1 then
																UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
															else
																UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
															end if
														else
															UPrice=0
														end if
														TotalUnit=TotalUnit+((ArrValue(i)+UPrice)*pQuantity)
													end if
													set rsConfigObj=nothing
												next
											end if
										End If
										'BTO ADDON-E

										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										' START: Get the total Price of all options
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										pOpPrices=0
										dim pcv_tmpOptionLoopCounter, pcArray_TmpCounter
										If len(pcv_strOptionsPriceArray)>0 then

											pcArray_TmpCounter = split(pcv_strOptionsPriceArray,chr(124))
											For pcv_tmpOptionLoopCounter = 0 to ubound(pcArray_TmpCounter)
												pOpPrices = pOpPrices + pcArray_TmpCounter(pcv_tmpOptionLoopCounter)
											Next

										end if

										if NOT isNumeric(pOpPrices) then
											pOpPrices=0
										end if

										'// Apply Discounts to Options Total
										'   >>> call function "pcf_DiscountedOptions(OriginalOptionsTotal, ProductID, Quantity, CustomerType)" from stringfunctions.asp
										Dim pcv_intDiscountPerUnit
										pOpPrices = pcf_DiscountedOptions(pOpPrices, pidProduct, pquantity, pcustomerType)

										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										' END: Get the total Price of all options
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


										if TotalUnit>0 then
											punitPrice1=punitPrice
											if pIdConfigSession<>"0" AND pIdConfigSession<>"" then
													pRowPrice1=Cdbl(pquantity * ( punitPrice1 )) - TotalUnit
													punitPrice1=Round(pRowPrice1/pquantity,2)
											else
													pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
											end if
										else
											punitPrice1=punitPrice
											if pIdConfigSession<>"0" AND pIdConfigSession<>"" then
													pRowPrice1=Cdbl(pquantity * ( punitPrice1 ))
											else
												pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
												punitPrice1=Round(pRowPrice1/pquantity,2)
											end if
										end if

										If strCol <> "#FFFFFF" Then
											strCol="#FFFFFF"
										Else
											strCol="#F1F1F1"
										End If

										%>

										<tr bgcolor="<%= strCol %>">
											<td class="invoice"><%=pquantity%></td>
											<td class="invoice">
												<%=psku%> - <a href="FindProductType.asp?id=<%=pidProduct%>" target="_blank"><%=pDescription%></a>
												<%
												'// Show sale icon, if applicable
												pcShowSaleIcon

												'SB S
												If pSubscriptionID>0 Then
													if pcv_strGUID<>"" then %>
														<div>
															Subscription ID: <strong><%=pcv_strGUID%></strong> -
															<a href="https://www.subscriptionbridge.com/MerchantCenter/" target="_blank">Manage</a>
														</div>
														<div>
															<%=pcv_strTerms%>
														</div>
													<% end if
												End If
												'SB E%>
											</td>
											<td class="invoice" align="right"><%=scCurSign&money((punitPrice1))%></td>
											<td class="invoice" align="right"><%=scCurSign&money(pRowPrice1)%></td>
										</tr>

										<% if pcv_IntDropShipperId <> 0 then %>
											<tr bgcolor="<%= strCol %>">
												<td class="invoice"></td>
												<td colspan="3" class="invoice"><span class="pcCPnotes">This product will be <strong>drop-shipped</strong>. <%if pcv_IntDropNotified <> 0 then%>The drop-shipping company <u>has been notified</u>.<%else%><u>Notify the drop-shipping company</u> using the &quot;Send Notification E-mail&quot; link below.<% end if %></span></td>
											</tr>
										<% end if %>

										<% if pcv_IntBackOrder <> 0 then %>
											<tr bgcolor="<%= strCol %>">
												<td class="invoice"></td>
												<td colspan="3" class="invoice"><span class="pcCPnotes">This product can be purchased when out of stock (&quot;back-ordered&quot;). <% if pStock<=0 then%>It's currently <strong>out of stock</strong><%else%>You currently have <%=pStock%> units in stock<%end if%>.</span></td>
											</tr>
										<% end if %>


										<% 'BTO ADDON-S
										if scBTO=1 then
											pIdConfigSession=trim(pidConfigSession)
											if pIdConfigSession<>"0" then
												query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
												set rsConfigObj=server.CreateObject("ADODB.RecordSet")
												set rsConfigObj=connTemp.execute(query)

												stringProducts=rsConfigObj("stringProducts")
												stringValues=rsConfigObj("stringValues")
												stringCategories=rsConfigObj("stringCategories")
												stringQuantity=rsConfigObj("stringQuantity")
												stringPrice=rsConfigObj("stringPrice")
												ArrProduct=Split(stringProducts, ",")
												ArrValue=Split(stringValues, ",")
												ArrCategory=Split(stringCategories, ",")
												ArrQuantity=Split(stringQuantity, ",")
												ArrPrice=Split(stringPrice, ",")


%>
												<tr bgcolor="<%= strCol %>">
													<td class="invoice">&nbsp;</td>
													<td class="invoice" colspan="3">
														<table width="100%" cellspacing="1" cellpadding="0" class="invoiceBto">
															<tr>
																<td colspan="3" class="invoiceNob"><u>Customizations:</u></td>
															</tr>
															<%
															for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
																query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pidProduct
																set rsQ=server.CreateObject("ADODB.RecordSet")
																set rsQ=conntemp.execute(query)

																btDisplayQF=rsQ("displayQF")
																set rsQ=nothing

																query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
																set rsQ=connTemp.execute(query)
																tmpMinQty=1
																if not rsQ.eof then
																	tmpMinQty=rsQ("pcprod_minimumqty")
																	if IsNull(tmpMinQty) or tmpMinQty="" then
																		tmpMinQty=1
																	else
																		if tmpMinQty="0" then
																			tmpMinQty=1
																		end if
																	end if
																end if
																set rsQ=nothing
																tmpDefault=0
																query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
																set rsQ=connTemp.execute(query)
																if not rsQ.eof then
																	tmpDefault=rsQ("cdefault")
																	if IsNull(tmpDefault) or tmpDefault="" then
																		tmpDefault=0
																	else
																		if tmpDefault<>"0" then
																			tmpDefault=1
																		end if
																	end if
																end if
																set rsQ=nothing

																query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
																set rsConfigObj=server.CreateObject("ADODB.RecordSet")
																set rsConfigObj=connTemp.execute(query)
																pcategoryDesc=rsConfigObj("categoryDesc")
																pdescription=rsConfigObj("description")
																psku=rsConfigObj("sku")
																pItemWeight=rsConfigObj("weight")
																if NOT isNumeric(ArrQuantity(i)) then
																	pIntQty=1
																else
																	pIntQty=ArrQuantity(i)
																end if %>
																<tr>
																	<td width="20%" valign="top" class="invoiceNob"><%=pcategoryDesc%>:</td>
																	<td width="70%" valign="top" class="invoiceNob"><%=psku%> - <%=pdescription%>
																	<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%> - QTY: <%=ArrQuantity(i)%><%end if%>
																	</td>
																	<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
																	if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
																		if tmpDefault=1 then
																			UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
																		else
																			UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
																		end if
																	else
																		UPrice=0
																	end if
																	'pfPrice=pfPrice+cdbl((ArrValue(i)+UPrice)*pQuantity) %>
																	<%end if%>
																	<td width="10%" valign="top" nowrap class="CustomizeOrdDet" align="right">
																		<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then%>
																			<%=scCurSign & money((ArrValue(i)+UPrice)*pQuantity)%>
																		<%else
																			if tmpDefault=1 then%>Included<%end if%>
																		<%end if%></td>
																</tr>
																<% intItemWeight=pItemWeight*pIntQTY*pquantity
																intTotalWeight=intTotalWeight+intItemWeight
																set rsConfigObj=nothing
															next
															set rsConfigObj=nothing %>
														</table>
													</td>
												</tr>
											<% end if %>
										<% end if
										'BTO ADDON-E %>

										<!-- start options -->
										<%
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										' START: SHOW PRODUCT OPTIONS
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
											pcv_strSelectedOptions = ""
										end if

										if len(pcv_strSelectedOptions)>0 then %>
											<tr bgcolor="<%= strCol %>" valign="top">
												<td class="invoice">&nbsp;</td>
												<td colspan="3" class="invoice">

													<table width="100%" cellspacing="0" cellpadding="0" class="invoiceBto">
														<tr>
															<td colspan="3" class="invoiceNob"><u>Options:</u></td>
														</tr>

														<%
														'#####################
														' START LOOP
														'#####################

														'// Generate Our Local Arrays from our Stored Arrays

														' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers
														pcArray_strSelectedOptions = ""
														pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))

														' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
														pcArray_strOptionsPrice = ""
														pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))

														' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
														pcArray_strOptions = ""
														pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))

														' Get Our Loop Size
														pcv_intOptionLoopSize = 0
														pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)

														' Start in Position One
														pcv_intOptionLoopCounter = 0

														' Display Our Options
														For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize %>
															<tr>
																<td class="invoiceNob"><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></td>

																<% tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
																if tempPrice="" or tempPrice=0 then %>
																<% else
																	'// Adjust for Quantity Discounts
																	tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice) %>
																	<td width="14%" class="invoiceNob" style="text-align: right;"><%=scCurSign&money(tempPrice)%></td>
																	<td width="13%" class="invoiceNob" style="text-align: right">
																	<%
																	tAprice=(tempPrice*Cdbl(pquantity))
																	response.write scCurSign&money(tAprice)
																	%>

																	</td>
																<% end if %>
															</tr>
														<% Next
														'#####################
														' END LOOP
														'#####################
														%>
													</table>
												</td>
											</tr>
										<% end if
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										' END: SHOW PRODUCT OPTIONS
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										%>
										<!-- end options -->

										<% if pxdetails<>"" then %>
											<tr bgcolor="<%= strCol %>">
												<td class="invoice">&nbsp;</td>
												<td class="invoice" align="left" valign="top" colspan="3"><%=pxdetails%></td>
											</tr>
										<% end if %>
										<!-- end of option descriptions -->

										<!-- if RMA -->
										<% if NOT isNull(prmaSubmitted) AND prmaSubmitted<>"" AND prmaSubmitted>0 then %>
											<tr bgcolor="<%= strCol %>">
												<td class="invoice"><%=prmaSubmitted%></td>
												<td class="invoice">RETURNED</td>
												<td class="invoice">&nbsp;</td>
												<td class="invoice">&nbsp;</td>
											</tr>
										<% end if	%>
										<!-- end of RMA -->
										<%'BTO ADDON-S
										pRowPrice=(punitPrice)*(pquantity)
										pExtRowPrice=pRowPrice
										Charges=0
										If scBTO=1 then
											pidConfigSession=trim(pidConfigSession)
											if pidConfigSession<>"0" then
												ItemsDiscounts=trim(ItemsDiscounts)
												if (ItemsDiscounts<>"") and (CDbl(ItemsDiscounts)<>"0") then
													%>
													<tr bgcolor="<%= strCol %>">
														<td class="invoice">&nbsp;</td>
														<td class="invoice">&nbsp;</td>
														<td class="invoice">Items Discounts:</td>
														<td class="invoice" align="right"><%=scCurSign&money(-1*ItemsDiscounts)%></td>
													</tr>
													<% pRowPrice=pRowPrice-Cdbl(ItemsDiscounts)
												end if%>

												<% 'BTO Additional Charges
												query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
												set rsConfigObj=server.CreateObject("ADODB.RecordSet")
												set rsConfigObj=connTemp.execute(query)

												stringCProducts=rsConfigObj("stringCProducts")
												stringCValues=rsConfigObj("stringCValues")
												stringCCategories=rsConfigObj("stringCCategories")
												ArrCProduct=Split(stringCProducts, ",")
												ArrCValue=Split(stringCValues, ",")
												ArrCCategory=Split(stringCCategories, ",")
												if ArrCProduct(0)<>"na" then %>
													<tr bgcolor="<%= strCol %>">
														<td class="invoice">&nbsp;</td>
														<td class="invoice" colspan="3">
															<table width="100%" cellspacing="0" cellpadding="1" class="invoiceBto">
																<tr>
																	<td colspan="3" class="invoiceNob"><u>Additional Charges</u></td>
																</tr>

																<% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
																	query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))"
																	set rsConfigObj=server.CreateObject("ADODB.RecordSet")
																	set rsConfigObj=connTemp.execute(query)
																	pcategoryDesc=rsConfigObj("categoryDesc")
																	pdescription=rsConfigObj("description")
																	psku=rsConfigObj("sku")
																	pItemWeight=rsConfigObj("weight")
																	intTotalWeight=intTotalWeight+pItemWeight
																	if (CDbl(ArrCValue(i))>0)then
																		Charges=Charges+cdbl(ArrCValue(i))
																	end if %>
																	<tr>
																		<td width="20%" class="invoiceNob" valign="top"><%=pcategoryDesc%>:</td>
																		<td width="70%" class="invoiceNob" valign="top"><%=psku%> - <%=pdescription%></td>
																		<td width="10%" class="invoiceNob" nowrap valign="top" align="right"><%if (ArrCValue(i)>0) then%><%=scCurSign & money(ArrCValue(i))%><%end if%></td>
																	</tr>
																	<% set rsConfigObj=nothing
																next
																pRowPrice=pRowPrice+Cdbl(Charges)%>
															</table>
														</td>
													</tr>
												<% end if
												'BTO Additional Charges

											end if
										end if 'BTO
										QDiscounts=trim(QDiscounts)
										if (QDiscounts<>"") and (CDbl(QDiscounts)<>"0") then %>
											<tr bgcolor="<%= strCol %>">
												<td class="invoice">&nbsp;</td>
												<td class="invoice">QUANTITY DISCOUNTS:</td>
												<td class="invoice">&nbsp;</td>
												<td class="invoice" align="right"><%=scCurSign&money(-1*QDiscounts)%></td>
											</tr>
											<% pRowPrice=pRowPrice-Cdbl(QDiscounts)
										end if %>
										<% TotPrdSubTotal = TotPrdSubTotal + pRowPrice
										if pExtRowPrice<>pRowPrice then %>
												<tr bgcolor="<%= strCol %>">
													<td class="invoice">&nbsp;</td>
													<td class="invoice">&nbsp;</td>
													<td class="invoice" align="right">Product Subtotal (<%=pSKU%>)</td>
													<td class="invoice" align="right"><%=scCurSign&money(TotPrdSubTotal)%></td>
												</tr>
											<% TotPrdSubTotal = 0
										else
											pTempSKU = pSKU
										end if
										'GGG Add-on start
										if pGWOpt<>"0" then
											query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
											set rsG=connTemp.execute(query)
											if not rsG.eof then%>
											<tr bgcolor="<%= strCol %>">
												<td class="invoice">&nbsp;</td>
												<td class="invoice" colspan="3">
													<b>Gift Wrapping:</b> <%=rsG("pcGW_OptName")%> - Price:&nbsp;<%=scCurSign & money(pGWPrice)%>
													<%if pGWText<>"" then%>
														<br>
														<b>Gift Notes:</b><br><%=pGWText%>
													<%end if%>
												</td>

											</tr>
											<%
											end if
											set rsG=nothing
										end if
										'GGG Add-on end

										tmpAllPrdSubTotal=tmpAllPrdSubTotal+CDbl(pRowPrice)

										pcPrdOrd_BundledDisc=trim(pcPrdOrd_BundledDisc)
										if (pcPrdOrd_BundledDisc<>"") and (CDbl(pcPrdOrd_BundledDisc)<>"0") then
										tmpAllPrdSubTotal=tmpAllPrdSubTotal-CDbl(pcPrdOrd_BundledDisc) %>
											<tr bgcolor="<%= strCol %>">
												<td class="invoice">&nbsp;</td>
												<td class="invoice">Bundle Discount</td>
												<td width="16%" class="invoice">&nbsp;</td>
												<td class="invoice" align="right"><%=scCurSign&money(-1*pcPrdOrd_BundledDisc)%></td>
											</tr>
										<% end if
										' MOVE TO NEXT PRODUCT
										rs.moveNext
										intPcCount = intPcCount + 1
									loop %>
									<tr>
											<td colspan="2" class="invoice">&nbsp;</td>
											<td align="right" class="invoice">
												<b>ALL PRODUCTS SUBTOTAL</b>
											</td>
											<td class="invoice" align="right"><%=scCurSign&money(tmpAllPrdSubTotal)%></td>
										</tr>


									<% 'RP ADDON-S
									If RewardsActive <> 0 Then
										if piRewardValue>0 then %>
											<tr>
												<td width="8%" class="invoice">&nbsp;</td>
												<% if RewardsLabel="" then
													RewardsLabel="Rewards Program"
												end if %>
												<td class="invoice"><%=RewardsLabel%></td>
												<td width="16%" class="invoice">&nbsp;</td>
												<td width="12%" class="invoice" align="right">-<%=scCurSign&money(piRewardValue)%></td>
											</tr>
										<% end if
									end if
									'RP ADDON-E %>

									<%'GGG Add-on start
									if pGWTotal>0 then%>
										<tr>
											<td colspan="2" class="invoice">&nbsp;</td>
											<td align="right" class="invoice">
												<b>GIFT WRAPPING</b>
											</td>
											<td class="invoice" align="right"><%=scCurSign&money(pGWTotal)%></td>
										</tr>
									<%
									end if
									'GGG Add-on end%>

									<tr>
										<%
										if (isNull(ptaxDetails) OR trim(ptaxDetails)="") OR pord_VAT>0 then
											tRowSpan=7
										else
											SpanArray=split(ptaxDetails,",")
											tRowSpan=6-(Ubound(SpanArray)-2)
										end if
										if NOT isNull(prmaCredit) AND prmaCredit<>"" AND prmaCredit>0 then
											tRowSpan = tRowSpan + 1
										end if
										PrdSales=ptotal
										%>
										<td colspan="2" rowspan="<%=tRowSpan%>" class="invoice">&nbsp;</td>
										<td align="right" class="invoice">SHIPPING</td>
										<td align="right" class="invoice"><%=scCurSign&money(postage)%></td>
									</tr>
									<%PrdSales=PrdSales-postage%>

									<% if serviceHandlingFee<>0 then %>
										<tr>
											<td align="right" class="invoice">SHIPPING &amp;<br>HANDLING FEES</td>
											<td align="right" class="invoice"><%=scCurSign&money(serviceHandlingFee)%></td>
										</tr>
									<% end if %>
									<%PrdSales=PrdSales-serviceHandlingFee%>

									<%
									payment = split(ppaymentDetails,"||")
									PaymentType=payment(0)
									on error resume next
									If trim(payment(1))="" then
										if err.number<>0 then
											PayCharge=0
										end if
										PayCharge=0
									else
										PayCharge=payment(1)
									end If

									err.clear
									%>
									<% if PayCharge>0 then %>
										<tr>
											<td class="invoice" align="right">PROCESSING FEES</td>
											<td class="invoice" align="right"><%=scCurSign&money(PayCharge)%></td>
										</tr>
									<%PrdSales=PrdSales-PayCharge%>
									<% end if %>
									<% if NOT (pord_VAT>0) then
										If ptaxVAT<>"1" Then
											if isNull(ptaxDetails) OR trim(ptaxDetails)="" then %>
												<tr>
													<td class="invoice" align="right">TAXES</td>
													<td class="invoice" align="right"><%=scCurSign&money(ptaxAmount)%></td>
												</tr>
												<%PrdSales=PrdSales-ptaxAmount%>

											<% else
												taxArray=split(ptaxDetails,",")
												for i=0 to (ubound(taxArray)-1)
													taxDesc=split(taxArray(i),"|")
													if taxDesc(0)<>"" then
													%>
													 <tr>
														<td class="invoice" align="right"><%=ucase(taxDesc(0))%></td>
														<% pDisTax=(money(taxDesc(1))) %>
														<td class="invoice" align="right"><%=scCurSign&pDisTax%></td>
													</tr>
													<%PrdSales=PrdSales-taxDesc(1)%>
													<% end if
												next
											end if
										End If
									end if
									%>
									<% if (instr(pdiscountDetails,"- ||")>0) or (pcv_CatDiscounts>"0") then
										if instr(pdiscountDetails,",") then
											DiscountDetailsArry=split(pdiscountDetails,",")
											intArryCnt=ubound(DiscountDetailsArry)
											for k=0 to intArryCnt
												if (DiscountDetailsArry(k)<>"") AND (instr(DiscountDetailsArry(k),"- ||")=0) then
													DiscountDetailsArry(k+1)=DiscountDetailsArry(k)+"," + DiscountDetailsArry(k+1)
													DiscountDetailsArry(k)=""
												end if
											next
										else
											intArryCnt=0
										end if
										dim discounts, discountType, havdis

										discount=0
										havdis=0
										for k=0 to intArryCnt
											if intArryCnt=0 then
												pTempDiscountDetails=pdiscountDetails
											else
												pTempDiscountDetails=DiscountDetailsArry(k)
											end if
											if instr(pTempDiscountDetails,"- ||") then
												havdis=havdis+1
												discounts = split(pTempDiscountDetails,"- ||")
												if havdis=1 then
													discountType = discountType&discounts(0)
												else
													discountType = discountType&"</b> and <b>"&discounts(0)
												end if
												tdiscount = discounts(1)
											else
												tdiscount=0
											end if
											if tdiscount="" OR isNULL(tdiscount)=True OR tdiscount="0" then
												tdiscount=0
											end if
											discount=discount+tdiscount
										Next
										%>
										<tr>
											<td class="invoice" align="right">DISCOUNTS</td>
											<td class="invoice" align="right">-<%=scCurSign&money(discount+pcv_CatDiscounts)%></td>
										</tr>
										<% PrdSales=PrdSales+pcv_CatDiscounts %>
									<% end if %>

									<%'GGG Add-on start
									IF GCAmount>"0" THEN%>
										<tr>
											<td align="right" class="invoice">GIFT CERTIFICATE AMOUNT</td>
											<td align="right" class="invoice"><% response.write "-"&scCurSign&money(GCAmount)%></td>
										</tr>
									<%END IF
									'GGG Add-on end%>

									<tr>
										<td class="invoice" align="right"><strong>TOTAL</strong></td>
										<td class="invoice" align="right"><strong><%=scCurSign&money(ptotal)%></strong></td>
									</tr>
									<% if pord_VAT>0 then %>
										<% if pcv_IsEUMemberState=1 then %>
										<tr>
											<td class="invoice" align="right" colspan="2"><span class="pcSmallText">Includes <%=scCurSign&money(pord_VAT)%> of VAT</span></td>
										</tr>
										<% else %>
										<tr>
											<td class="invoice" align="right" colspan="2"><span class="pcSmallText"><%=scCurSign&money(pord_VAT)%> of VAT Removed</span></td>
										</tr>
										<% end if %>
									<% end if %>
									<%
									'--------------------------------
									' START Show CREDITS applied to the order
									'--------------------------------
									if NOT isNull(prmaCredit) AND prmaCredit<>"" AND prmaCredit>0 then %>
										<tr>
											<td class="invoice" align="right"><b>CREDIT</b></td>
											<td class="invoice" align="right">-<%=money(prmaCredit)%></td>
										</tr>
									<% end if
									'--------------------------------
									' END Show CREDITS
									'--------------------------------
									%>
								</table>
							</td>
						</tr>
					</table>

					<%
					'Start SDBA
					'******************************************
					' START: Drop-Shipping & Back-ordering
					'******************************************
					IF pcv_haveBOPrds=1 OR pcv_haveDropPrds=1 THEN %>
						<a name="dropship">&nbsp;</a>
						<table class="pcCPcontent">
							<!--
								<tr>
									<th colspan="3">
										<% if pcv_haveBOPrds=1 and pcv_haveDropPrds<>1 then %>
												Back-Ordered Products
										<% elseif pcv_haveBOPrds<>1 and pcv_haveDropPrds=1 then %>
												Drop-Shipping Products
										<% else %>
												Drop-Shipping and Back-Ordered Products
										<% end if %>
									</th>
								</tr>
							-->
							<tr>
								<td colspan="3" class="pcCPspacer"></td>
							</tr>

							<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' START Have Drop-shipped products
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if pcv_haveDropPrds=1 then
								query="SELECT Products.idproduct, Products.Description, ProductsOrdered.quantity, ProductsOrdered.pcPackageInfo_ID, ProductsOrdered.pcDropShipper_ID, ProductsOrdered.pcPrdOrd_Shipped, ProductsOrdered.pcPrdOrd_BackOrder,  ProductsOrdered.pcPrdOrd_SentNotice FROM Products INNER JOIN ProductsOrdered ON (products.idproduct = ProductsOrdered.idproduct AND products.pcProd_IsDropShipped = 1) WHERE productsOrdered.idorder=" & qry_ID & " ORDER BY ProductsOrdered.pcDropShipper_ID DESC, ProductsOrdered.pcPackageInfo_ID;"
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)

								if err.number<>0 then
									call LogErrorToDatabase()
									set rs=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if

								save_dropshipperid=0
								save_isSupplier=0
								save_firsttime=1
								save_packageID=0
								count=0

								if not rs.eof then %>
									<tr>
										<th colspan="3"><strong>DROP-SHIPPED PRODUCTS</strong>
											<input type="hidden" id="pcv_DS_ID" name="pcv_DS_ID" value="">
											<input type="hidden" id="pcv_DS_IsSupplier" name="pcv_DS_IsSupplier" value="">
											<input type="hidden" id="pcv_DS_Comments" name="pcv_DS_Comments" value="">
											<input type="hidden" id="pcv_DS_IDPackage" name="pcv_DS_IDPackage" value="">
										</th>
									</tr>

									<script>
										function SetValue(DS_ID,DS_IsSupplier,DS_Comments,DS_IDPackage)
										{
											var tmp1=document.getElementById("pcv_DS_ID")
											tmp1.value=DS_ID
											var tmp1=document.getElementById("pcv_DS_IsSupplier")
											tmp1.value=DS_IsSupplier
											var tmp1=document.getElementById("pcv_DS_Comments")
											tmp1.value=document.getElementById(DS_Comments).value
											var tmp1=document.getElementById("pcv_DS_IDPackage")
											tmp1.value=DS_IDPackage
										}
									</script>

									<% do while not rs.eof
										pcv_showHeader = 0
										pcv_IDProduct=rs("idproduct")
										pcv_PrdName=rs("Description")
										pcv_PrdQty=rs("quantity")
										pcv_PackageID=rs("pcPackageInfo_ID")
										if IsNull(pcv_PackageID) or pcv_PackageID="" then
											pcv_PackageID=0
										end if
										pcv_DropShipperID=rs("pcDropShipper_ID")
										if IsNull(pcv_DropShipperID) or pcv_DropShipperID="" then
											pcv_DropShipperID=0
										end if
										pcv_Shipped=rs("pcPrdOrd_Shipped")
										if IsNull(pcv_Shipped) or pcv_Shipped="" then
											pcv_Shipped=0
										end if
										pcv_IsBackOrder=rs("pcPrdOrd_BackOrder")
										if IsNull(pcv_IsBackOrder) or pcv_IsBackOrder="" then
											pcv_IsBackOrder=0
										end if
										pcv_SentNotice=rs("pcPrdOrd_SentNotice")
										if IsNull(pcv_SentNotice) or pcv_SentNotice="" then
											pcv_SentNotice=0
										end if
										if clng(pcv_DropShipperID)>0 then
											query="SELECT pcDS_IsDropShipper FROM pcDropShippersSuppliers WHERE idproduct=" & pcv_IDProduct & ";"
											set rs1=server.CreateObject("ADODB.RecordSet")
											set rs1=conntemp.execute(query)

											pcv_IsSupplier=0
											if not rs1.eof then
												if rs1("pcDS_IsDropShipper")="1" then
													pcv_IsSupplier=1
												end if
											end if
											set rs1=nothing
										end if

										if ((clng(save_dropshipperid)<>clng(pcv_DropShipperID)) or (clng(save_isSupplier)<>clng(pcv_IsSupplier))) or (save_firsttime=1) then
											save_dropshipperid=pcv_DropShipperID
											save_isSupplier=pcv_IsSupplier
											count=count+1
											pcv_showHeader = 1
											if clng(pcv_DropShipperID)>0 then

												if pcv_IsSupplier=1 then
													query="SELECT pcSupplier_ID As DropShipperID,pcSupplier_FirstName As DropShipperFirstName,pcSupplier_LastName As DropShipperLastName,pcSupplier_Company As DropShipperCompany FROM pcSuppliers WHERE pcSupplier_ID=" & pcv_DropShipperID & ";"
												else
													query="SELECT pcDropShipper_ID As DropShipperID,pcDropShipper_FirstName  As DropShipperFirstName,pcDropShipper_LastName As DropShipperLastName,pcDropShipper_Company As DropShipperCompany FROM pcDropShippers WHERE pcDropShipper_ID=" & pcv_DropShipperID & ";"
												end if
												set rs1=server.CreateObject("ADODB.RecordSet")
												set rs1=conntemp.execute(query)

												if not rs1.eof then
													pcv_DropShipperID=rs1("DropShipperID")
													pcv_DSname=rs1("DropShipperCompany") & " (" & rs1("DropShipperFirstName") & " " & rs1("DropShipperLastName") & ")"
													query="SELECT pcPrdOrd_Shipped FROM ProductsOrdered WHERE idorder=" & qry_ID & " AND pcDropShipper_ID=" & pcv_DropShipperID & " AND pcPackageInfo_ID>0;"
													set rs1=connTemp.execute(query)
													pcv_OrderUpdated=0
													if not rs1.eof then
														pcv_OrderUpdated=1
													end if
													set rs1=nothing
													%>
													<tr>
														<td colspan="3" class="pcCPspacer"></td>
													</tr>
													<tr>
														<td colspan="3" style="background-color: #e1e1e1;">Drop-Shipper: <b><%=pcv_DSname%></b></td>
													</tr>
													<tr>
														<td>Notification Sent: <b><%if pcv_SentNotice="1" then%>Yes<%else%>No<%end if%></b></td>
														<td>Order Updated: <b><%if pcv_OrderUpdated="1" then%>Yes<%else%>No<%end if%></b></td>
														<td align="right">
														<a href="javascript:openshow<%=count%>()">Send/Resend Order Notification Email</a>
														<script>
															function openshow<%=count%>()
															{
																var tmpobj=document.getElementById("show_<%=count%>")
																tmpobj.style.display='';
															}
														</script>
														</td>
													</tr>
													<tr>
														<td colspan="3" align="right">
															<table id="show_<%=count%>" style="display:none" class="pcCPcontent">
																<tr>
																	<td align="right">&nbsp;Add any comments for the drop-shipper (optional):<br>
																		<%query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=4 AND pcDropShipper_ID=" & pcv_DropShipperID & " AND pcACom_IsSupplier=" & pcv_IsSupplier & ";"
																		set rsQ=server.CreateObject("ADODB.RecordSet")
																		set rsQ=connTemp.execute(query)

																		pcv_AdmComments=""
																		if not rsQ.eof then
																			pcv_AdmComments=rsQ("pcACom_Comments")
																		end if
																		set rsQ=nothing%>
																		<textarea id="AdmComments_<%=pcv_DropShipperID%>_<%=pcv_IsSupplier%>" name="AdmComments_<%=pcv_DropShipperID%>_<%=pcv_IsSupplier%>" cols="60" rows="5" wrap="VIRTUAL" style="background-color:#FFFFFF; margin-top:5px; margin-bottom:5px;"><%=pcv_AdmComments%></textarea><br>
																		<input type="submit" name="SubmitA3" value="Send/Resend Order Notification" onclick="javascript:SetValue(<%=pcv_DropShipperID%>,<%=pcv_IsSupplier%>,'AdmComments_<%=pcv_DropShipperID%>_<%=pcv_IsSupplier%>',0);" class="submit2">
																	</td>
																</tr>
																<tr>
																	<td class="pcCPspacer"></td>
																</tr>
															</table>
														</td>
													</tr>
												<%end if
											else 'Unknown Drop-Shipper %>
												<tr>
													<td colspan="3" class="pcCPspacer"></td>
												</tr>
												<tr>
													<td colspan="3"><b>Drop-Shipper: Not Known</b></td>
												</tr>
												<tr>
													<td colspan="3" class="pcCPspacer"></td>
												</tr>
											<% end if
											save_firsttime=0
										end if
										'of New Drop-Shipper

										'// Header Control
										if (pcv_showHeader = 1) and (pcv_PackageID="0") then %>
											<tr>
												<td style="border-bottom: 1px dashed #CCC;">Product Name</td>
												<td style="border-bottom: 1px dashed #CCC;">Qty</td>
												<td style="border-bottom: 1px dashed #CCC;">Shipment Status</td>
											</tr>
										<% end if %>
										<% if clng(pcv_PackageID)<>clng(save_packageID) then
											if pcv_PackageID<>"0" then
												query="SELECT pcPackageInfo_ShipMethod, pcPackageInfo_TrackingNumber, pcPackageInfo_ShippedDate,  pcPackageInfo_Comments, pcPackageInfo_UPSServiceCode, pcPackageInfo_UPSLabelFormat, pcPackageInfo_MethodFlag,pcPackageInfo_Endicia,pcPackageInfo_EndiciaLabelFile,pcPackageInfo_EndiciaExp FROM pcPackageInfo WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
												set rs1=server.CreateObject("ADODB.RecordSet")
												set rs1=conntemp.execute(query)

												if not rs1.eof then
													tmp_ShipMethod= rs1("pcPackageInfo_ShipMethod")
													tmp_TrackingNumber= rs1("pcPackageInfo_TrackingNumber")
													tmp_ShippedDate=rs1("pcPackageInfo_ShippedDate")
													tmp_Comments=rs1("pcPackageInfo_Comments")
													tmp_ServiceCode=rs1("pcPackageInfo_UPSServiceCode")
													tmp_LabelFormat=rs1("pcPackageInfo_UPSLabelFormat")
													tmp_MethodFlag=rs1("pcPackageInfo_MethodFlag")
													tmp_Endicia=rs1("pcPackageInfo_Endicia")
													tmpEndiciaLabel=rs1("pcPackageInfo_EndiciaLabelFile")
													tmpEndiciaExp=rs1("pcPackageInfo_EndiciaExp")
													if tmpEndiciaExp<>"" then
														tmpEndiciaExp=CDate(tmpEndiciaExp)+EDCExpDays
													end if

													pcv_HavePackage=0
													%>
													<tr>
														<td colspan="3" style="background-color: #e3e3e3">
														<%
														if tmp_ShipMethod<>"" OR tmp_TrackingNumber<>"" OR tmp_ShippedDate<>"" then
															pcv_HavePackage=1
															%>
															<strong>Shipment Information</strong>:<br>
														<%end if%>
														</td>
													</tr>
													<tr style="background-color: #ffffff">
														<td colspan="2">
														<%if tmp_ShipMethod<>"" then
														if tmp_Endicia="1" then
														Select Case tmp_ShipMethod
															Case "Express": tmp_ShipMethod="Express Mail"
															Case "First": tmp_ShipMethod="First-Class Mail"
															Case "LibraryMail": tmp_ShipMethod="Library Mail"
															Case "MediaMail": tmp_ShipMethod="Media Mail"
															Case "ParcelPost": tmp_ShipMethod="Standard Post"
															Case "ParcelSelect": tmp_ShipMethod="Parcel Select"
															Case "Priority": tmp_ShipMethod="Priority Mail"
															Case "StandardMail": tmp_ShipMethod="Standard Mail"
															Case "ExpressMailInternational": tmp_ShipMethod="Express Mail International"
															Case "FirstClassMailInternational": tmp_ShipMethod="First-Class Mail International"
															Case "PriorityMailInternational": tmp_ShipMethod="Priority Mail International"
														End Select
														end if
														%>Shipping Method: <%=tmp_ShipMethod%><br><%end if%>
														<% if tmp_MethodFlag="4" AND tmp_ServiceCode<>"" then
															select case tmp_ServiceCode
																case "E"
																	tmp_StrService="Express Mail Label"
																case "S"
																	tmp_StrService="Signature Confirmation Label"
																case "D"
																	tmp_StrService="Delivery Confirmation Label"
															end select
															%>
															<%
															set fso = server.createObject("Scripting.FileSystemObject")
															if fso.fileExists(server.mappath("USPSLabels/receipt"&tmp_TrackingNumber&"."&tmp_LabelFormat&"")) then
															strReceiptLink="&nbsp;|&nbsp;<a href='USPSLabels/receipt"&tmp_TrackingNumber&"."&tmp_LabelFormat&"' target='_blank'>"
															else
																strReceiptLink=""
															end if
															%>
															<%if tmp_StrService<>"" then%>
															Label Type: <%=tmp_StrService%><br>
															<%end if%>
															<%if tmp_TrackingNumber<>"" then%>
																Label Tracking Number: <%=tmp_TrackingNumber%>&nbsp;|&nbsp;<a href="USPSLabels/<%if tmpEndiciaLabel<>"" then%><%=tmpEndiciaLabel%><%else%>label<%=tmp_TrackingNumber%>.<%=tmp_LabelFormat%><%end if%>" target="_blank">Print Label</a><%=strReceiptLink%>&nbsp;|&nbsp;<a href="javascript:openshipwin('USPS_TrackAndConfirmRequest.asp?TN=<%=tmp_TrackingNumber%>');">Track Package</a><%if (Now()<=tmpEndiciaExp) AND (tmp_Endicia="1") then%>&nbsp;|&nbsp;<a href="javascript: if(confirm('Are you sure you want to create a refund request for this USPS label?')) {location='EDC_refund.asp?id=<%=pcv_PackageID%>';}">Refund</a><%end if%><br>
																<% tmp_TrackingNumberShown=1
															end if%>
														<% else
															tmp_TrackingNumberShown=0
														end if %>
														<%if not IsNull(tmp_ShippedDate) then%>Shipped Date: <%=ShowDateFrmt(tmp_ShippedDate)%><br><%end if%>
														<%if tmp_TrackingNumber<>"" AND tmp_TrackingNumberShown=0 then%>Tracking Number: <%=tmp_TrackingNumber%><br><%end if%>
														<%if tmp_Comments<>"" then%>Comments:<br><i><%=tmp_Comments%></i><%end if%>
														</td>
														<td align="right" valign="top">
														<script>
															function openshow<%=count%>A()
															{
																var tmpobj=document.getElementById("show_<%=count%>A")
																tmpobj.style.display='';
															}
														</script>
														<%if pcv_HavePackage=1 AND tmp_MethodFlag<>"4" then%>
															<a href="javascript:openshow<%=count%>A()">Send/Resend Partially Shipped Email</a>
														<%else%>
															&nbsp;
														<%end if%>
														</td>
													</tr>
													<%if pcv_HavePackage=1 then
														if pcv_IsSupplier="" then
															pcv_IsSupplier=0
														end if %>
														<tr valign="top">
															<td></td>
															<td colspan="2">
																<table id="show_<%=count%>A" style="display:none">
																	<tr>
																		<td>Comments:<br>
																			<%query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=2 AND pcDropShipper_ID=" & pcv_DropShipperID & " AND pcACom_IsSupplier=" & pcv_IsSupplier & ";"
																			set rsQ=server.CreateObject("ADODB.RecordSet")
																			set rsQ=connTemp.execute(query)

																			pcv_AdmComments=""
																			if not rsQ.eof then
																				pcv_AdmComments=rsQ("pcACom_Comments")
																			end if
																			set rsQ=nothing%>
																			<textarea id="AdmComments_<%=pcv_DropShipperID%>_<%=pcv_IsSupplier%>A" name="AdmComments_<%=pcv_DropShipperID%>_<%=pcv_IsSupplier%>A" cols="60" rows="5" wrap="VIRTUAL" style="font-size:11px"><%=pcv_AdmComments%></textarea><br>
																			<input type=submit name="SubmitA4" value="Continue Send/Resend Partially Shipped Email" onclick="javascript:SetValue(<%=pcv_DropShipperID%>,<%=pcv_IsSupplier%>,'AdmComments_<%=pcv_DropShipperID%>_<%=pcv_IsSupplier%>A',<%=pcv_PackageID%>);">
																		</td>
																	</tr>
																</table>
															</td>
														</tr>
													<%end if%>
													<tr style="background-color: #e5e5e5">
														<td>Product Name</td>
														<td>Qty</td>
														<td>Shipment Status</td>
													</tr>
												<%end if
												set rs1=nothing
											end if
										end if
										%>
										<tr align="top">
											<td><%=pcv_PrdName%>
												<%if pcv_IsBackOrder="1" then%><br><i>(back-ordered)</i><%end if%></td>
											<td><%=pcv_PrdQty%></td>
											<td>
												<%if pcv_Shipped="1" then%>
													<span style="color: #0000FF">Shipped</span>
												<%else%>
													<% if pcv_Shipped="2" then
														pcv_ModShipLink="USPS_ShipOrderWizard.asp?smode="&tmp_ServiceCode&"&orderID="&qry_ID&"&packID="&pcv_PackageID %>
														<span style="color: #00F;">Pending</span>&nbsp;|&nbsp;<a href="<%=pcv_ModShipLink%>">Ship &gt;&gt;</a>
													<% else %>
														<span style="color: #00F;">Pending</span>&nbsp;|&nbsp;<a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Ship &gt;&gt;</a>
													<% end if %>
												<%end if%>
											</td>
										</tr>
										<%
										if clng(pcv_PackageID)<>clng(save_packageID) then
											save_packageID=pcv_PackageID
										end if
										rs.MoveNext
									loop
									set rs=nothing
									%>
									<tr>
										<td colspan="3" class="pcCPspacer"></td>
									</tr>
								<% end if
							end if
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' END Have Drop-shipped products
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							%>

							<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' Start Have Back-Ordered products
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if pcv_haveBOPrds=1 then
								query="SELECT Products.idproduct,Products.Description,ProductsOrdered.quantity,ProductsOrdered.pcPackageInfo_ID,ProductsOrdered.pcDropShipper_ID,ProductsOrdered.pcPrdOrd_Shipped,ProductsOrdered.pcPrdOrd_BackOrder,ProductsOrdered.pcPrdOrd_SentNotice FROM Products INNER JOIN ProductsOrdered ON (products.idproduct=ProductsOrdered.idproduct AND products.pcProd_IsDropShipped<>1) WHERE productsOrdered.idorder=" & qry_ID & " AND productsOrdered.pcPrdOrd_BackOrder=1 ORDER BY ProductsOrdered.pcPackageInfo_ID;"
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								save_packageID=0
								if not rs.eof then%>
									<tr>
										<td colspan="3" class="pcCPspacer"></td>
									</tr>
									<tr>
										<th colspan="3"><strong>BACK-ORDERED PRODUCTS</strong></th>
									</tr>
									<% '// Header Control
									if NOT rs.eof then
										pcv_Shipped=rs("pcPrdOrd_Shipped")
										if IsNull(pcv_Shipped) or pcv_Shipped="" then
											pcv_Shipped=0
										end if
										pcv_showHeader = 1
										if pcv_Shipped = 1 then
											pcv_showHeader = 0
										end if
									end if
									if pcv_showHeader = 1 then
									%>
									<tr style="background-color: #e5e5e5">
										<td>Product Name</td>
										<td>Qty</td>
										<td>Shipment Status</td>
									</tr>
									<% end if %>

									<% do while not rs.eof
										pcv_IDProduct=rs("idproduct")
										pcv_PrdName=rs("Description")
										pcv_PrdQty=rs("quantity")
										pcv_PackageID=rs("pcPackageInfo_ID")
										if IsNull(pcv_PackageID) or pcv_PackageID="" then
											pcv_PackageID=0
										end if
										pcv_DropShipperID=rs("pcDropShipper_ID")
										if IsNull(pcv_DropShipperID) or pcv_DropShipperID="" then
											pcv_DropShipperID=0
										end if
										pcv_Shipped=rs("pcPrdOrd_Shipped")
										if IsNull(pcv_Shipped) or pcv_Shipped="" then
											pcv_Shipped=0
										end if
										pcv_IsBackOrder=rs("pcPrdOrd_BackOrder")
										if IsNull(pcv_IsBackOrder) or pcv_IsBackOrder="" then
											pcv_IsBackOrder=0
										end if
										pcv_SentNotice=rs("pcPrdOrd_SentNotice")
										if IsNull(pcv_SentNotice) or pcv_SentNotice="" then
											pcv_SentNotice=0
										end if

										if clng(pcv_PackageID)<>clng(save_packageID) then
											if pcv_PackageID<>"0" then
												query="SELECT pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_ShippedDate,pcPackageInfo_Comments, pcPackageInfo_UPSLabelFormat, pcPackageInfo_MethodFlag FROM pcPackageInfo WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
												set rs1=connTemp.execute(query)
												if not rs1.eof then
													tmp_ShipMethod= rs1("pcPackageInfo_ShipMethod")
													tmp_TrackingNumber= rs1("pcPackageInfo_TrackingNumber")
													tmp_ShippedDate=rs1("pcPackageInfo_ShippedDate")
													tmp_Comments=rs1("pcPackageInfo_Comments")
													tmp_LabelFormat=rs1("pcPackageInfo_UPSLabelFormat")
													tmp_MethodFlag=rs1("pcPackageInfo_MethodFlag")
													if isNULL(tmp_MethodFlag) OR tmp_MethodFlag="" then
														tmp_MethodFlag=0
													end if
													select case tmp_MethodFlag
													case 2
														pcv_MF_UPS=1
													case 3
														pcv_MF_FEDEX=1
													case 0
														pcv_MF_UPS=1
														pcv_MF_FEDEX=1
													end select
													%>
													<tr style="background-color: #cccccc">
														<td colspan="3">
														<%if tmp_ShipMethod<>"" OR tmp_TrackingNumber<>"" OR tmp_ShippedDate<>"" then%>
															<strong>Shipment Information</strong>:<br>
														<%end if%>
														</td>
													</tr>
													<tr style="background-color: #ffffff">
														<td colspan="3">
														<%if tmp_ShipMethod<>"" then%>Shipping Method: <%=tmp_ShipMethod%><br><%end if%>
														<%if tmp_TrackingNumber<>"" then%>Tracking Number: <%=tmp_TrackingNumber%><br><%end if%>
														<%if not IsNull(tmp_ShippedDate) then%>Shipped Date: <%=ShowDateFrmt(tmp_ShippedDate)%><br><%end if%>
														<%if tmp_Comments<>"" then%>Comments:<br><i><%=tmp_Comments%></i><%end if%>
														</td>
													</tr>
													<tr style="background-color: #e5e5e5">
														<td>Product Name</td>
														<td>Qty</td>
														<td>Shipment Status</td>
													</tr>
												<% end if
												set rs1=nothing
											end if
										end if
										%>
										<tr>
											<td><%=pcv_PrdName%></td>
											<td><%=pcv_PrdQty%></td>
											<td>
												<%if pcv_Shipped="1" then%>
													<span style="color: #0000FF">Shipped</span>&nbsp;(<a href="javascript:openshipwin('modifyPackInfo.asp?packID=<%=pcv_PackageID%>');">Modify shipment information</a>)
												<%else%>
													<span style="color: #00F;">Pending</span>&nbsp;|&nbsp;<a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Ship &gt;&gt;</a>
												<%end if%>
											</td>
										</tr>
										<%
										if clng(pcv_PackageID)<>clng(save_packageID) then
											save_packageID=pcv_PackageID
										end if
										%>
										<tr>
											<td colspan="3"><hr></td>
										</tr>
										<%
										rs.MoveNext
									loop
									set rs=nothing
									%>
									<tr>
										<td colspan="3" class="pcCPspacer"></td>
									</tr>
								<% end if
								'---------------------------------------------
								' - START - Single vs. Separate Shipments
								'---------------------------------------------

								if intPcCount > 1 then ' START order contains more than 1 product %>
									<tr>
										<td colspan="3"><hr></td>
									</tr>
									<tr>
										<td colspan="3"><strong>Single vs. Separate Shipments</strong></td>
									</tr>
									<tr>
										<td colspan="3">
											<% if pOrderStatus<=2 and pcv_CustRequestStr="NA" then %>
												Your store is setup to allow customers to request that back-ordered products are shipped separately. If the order contains more than one product, the customer <u>will be asked</u> to specify whether to receive one or separate shipments <u>when you process the order</u>.
											<% else
												if pcv_CustAllow="1" then %>
													The customer was notified and indicated that he/she wants <strong>separate shipments</strong>
												<% else
													if pcv_CustAllow="2" then %>
														The customer was notified and indicated that he/she wants <strong>one shipment</strong>. Therefore, wait until all products are in stock, then ship the order and update the customer.
													<% else
														if (pcv_CustAllow=0) and (pcv_CustRequestStr<>"NA") then%>
															The customer <u>has been alerted</u> to indicate whether he/she wants one or separate shipments. The system is <u>awaiting for the customer response</u>.
														<% end if
													end if
												end if
											end if
											%>
										</td>
									</tr>
								<% end if ' END order contains more than 1 product
								'---------------------------------------------
								' - START - Single vs. Separate Shipments
								'---------------------------------------------
								%>
								<tr>
									<td colspan="3"><hr></td>
								</tr>
							<% end if
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' End Back-Ordered products
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ %>
						</TABLE>
					<%	end if
					'******************************************
					' END: Drop-Shipping & Back-ordering
					'******************************************

					'******************************************
					' START: Products not drop-shipped and not back-ordered
					'******************************************

					IF varShip="0" THEN ' START - There is no shipping needed for this order

					ELSE ' This order contains product that do need to be shipped

						query="SELECT Products.idproduct, Products.Description, ProductsOrdered.quantity, ProductsOrdered.pcPackageInfo_ID, ProductsOrdered.pcDropShipper_ID, ProductsOrdered.pcPrdOrd_Shipped, ProductsOrdered.pcPrdOrd_BackOrder, ProductsOrdered.pcPrdOrd_SentNotice FROM Products INNER JOIN ProductsOrdered ON (products.idproduct=ProductsOrdered.idproduct AND products.pcProd_IsDropShipped<>1) WHERE productsOrdered.idorder=" & qry_ID & " AND productsOrdered.pcPrdOrd_BackOrder=0 AND productsOrdered.pcDropShipper_ID=0 ORDER BY ProductsOrdered.pcPackageInfo_ID;"
						set rs=connTemp.execute(query)
						save_packageID=0
						if not rs.eof then
							'..There is a product flagged as shipped already %>
							<table class="pcCPcontent">
								<tr>
									<td colspan="3" class="pcCPspacer"></td>
								</tr>
								<% if pcv_haveBOPrds<>1 and pcv_haveDropPrds<>1 then %>
									<tr>
										<th colspan="3"><strong>SHIPPING STATUS</strong></th>
									</tr>
								<% else %>
									<tr>
										<th colspan="3"><strong>PRODUCTS IMMEDIATELY AVAILABLE</strong></th>
									</tr>
								<% end if %>
								<%
								do while not rs.eof
									pcv_IDProduct=rs("idproduct")
									pcv_PrdName=rs("Description")
									pcv_PrdQty=rs("quantity")
									pcv_PackageID=rs("pcPackageInfo_ID")
									if IsNull(pcv_PackageID) or pcv_PackageID="" then
										pcv_PackageID=0
									end if
									pcv_DropShipperID=rs("pcDropShipper_ID")
									if IsNull(pcv_DropShipperID) or pcv_DropShipperID="" then
										pcv_DropShipperID=0
									end if
									pcv_Shipped=rs("pcPrdOrd_Shipped")
									if IsNull(pcv_Shipped) or pcv_Shipped="" then
										pcv_Shipped=0
									end if
									pcv_IsBackOrder=rs("pcPrdOrd_BackOrder")
									if IsNull(pcv_IsBackOrder) or pcv_IsBackOrder="" then
										pcv_IsBackOrder=0
									end if
									pcv_SentNotice=rs("pcPrdOrd_SentNotice")
									if IsNull(pcv_SentNotice) or pcv_SentNotice="" then
										pcv_SentNotice=0
									end if

									if clng(pcv_PackageID)<>clng(save_packageID) then
										'Set MethodFlags to 0
										dim pcv_MF_UPS, pcv_MF_FEDEX
										pcv_MF_UPS=0
										pcv_MF_FEDEX=0
										if pcv_PackageID<>"0" then
											query="SELECT pcPackageInfo_ShipMethod, pcPackageInfo_TrackingNumber, pcPackageInfo_ShippedDate, pcPackageInfo_Comments, pcPackageInfo_UPSServiceCode, pcPackageInfo_UPSLabelFormat, pcPackageInfo_MethodFlag,pcPackageInfo_Endicia,pcPackageInfo_EndiciaLabelFile,pcPackageInfo_EndiciaExp FROM pcPackageInfo WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
											set rs1=connTemp.execute(query)
											if not rs1.eof then
												tmp_ShipMethod= rs1("pcPackageInfo_ShipMethod")
												tmp_TrackingNumber= rs1("pcPackageInfo_TrackingNumber")
												tmp_ShippedDate=rs1("pcPackageInfo_ShippedDate")
												tmp_Comments=rs1("pcPackageInfo_Comments")
												tmp_ServiceCode=rs1("pcPackageInfo_UPSServiceCode")
												tmp_LabelFormat=rs1("pcPackageInfo_UPSLabelFormat")
												tmp_MethodFlag=rs1("pcPackageInfo_MethodFlag")
												if isNULL(tmp_MethodFlag) OR tmp_MethodFlag="" then
													tmp_MethodFlag=0
												end if
												select case tmp_MethodFlag
												case 2
													pcv_MF_UPS=1
												case 3
													pcv_MF_FEDEX=1
												case 0
													pcv_MF_UPS=1
													pcv_MF_FEDEX=1
												end select
												tmp_Endicia=rs1("pcPackageInfo_Endicia")
												tmpEndiciaLabel=rs1("pcPackageInfo_EndiciaLabelFile")
												tmpEndiciaExp=rs1("pcPackageInfo_EndiciaExp")
												if tmpEndiciaExp<>"" then
													tmpEndiciaExp=CDate(tmpEndiciaExp)+EDCExpDays
												end if
												%>
												<tr style="background-color: #cccccc">
													<td colspan="3">
													<%if tmp_ShipMethod<>"" OR tmp_TrackingNumber<>"" OR tmp_ShippedDate<>"" then%>
														<strong>Shipment Information</strong>:<br>
													<%end if%>
													</td>
												</tr>
												<tr style="background-color: #ffffff">
													<td colspan="3">
													<%if tmp_ShipMethod<>"" then
														if tmp_Endicia="1" then
														Select Case tmp_ShipMethod
															Case "Express": tmp_ShipMethod="Express Mail"
															Case "First": tmp_ShipMethod="First-Class Mail"
															Case "LibraryMail": tmp_ShipMethod="Library Mail"
															Case "MediaMail": tmp_ShipMethod="Media Mail"
															Case "ParcelPost": tmp_ShipMethod="Standard Post"
															Case "ParcelSelect": tmp_ShipMethod="Parcel Select"
															Case "Priority": tmp_ShipMethod="Priority Mail"
															Case "StandardMail": tmp_ShipMethod="Standard Mail"
															Case "ExpressMailInternational": tmp_ShipMethod="Express Mail International"
															Case "FirstClassMailInternational": tmp_ShipMethod="First-Class Mail International"
															Case "PriorityMailInternational": tmp_ShipMethod="Priority Mail International"
														End Select
														end if
														%>Shipping Method: <%=tmp_ShipMethod%><br><%end if%>
													<% if tmp_MethodFlag="4" AND tmp_ServiceCode<>"" then
														'..This is a USPS Package
														select case tmp_ServiceCode
															case "E"
																tmp_StrService="Express Mail Label"
															case "S"
																tmp_StrService="Signature Confirmation Label"
															case "D"
																tmp_StrService="Delivery Confirmation Label"
														end select
														%>
														<%	dim fso
														set fso = server.createObject("Scripting.FileSystemObject")
														if fso.FileExists(server.mappath("USPSLabels/receipt"&tmp_TrackingNumber&"."&tmp_LabelFormat&"")) then
														   strReceiptLink="&nbsp;|&nbsp;<a href='USPSLabels/receipt"&tmp_TrackingNumber&"."&tmp_LabelFormat&"' target='_blank'>Print Receipt</a>"
														else
															strReceiptLink=""
														end if
														SET fso=nothing
														%>
														<%if tmp_StrService<>"" then%>
														Label Type: <%=tmp_StrService%><br>
														<%end if%>
														<%if tmp_TrackingNumber<>"" then%>
															Label Tracking Number: <%=tmp_TrackingNumber%>&nbsp;|&nbsp;<a href="USPSLabels/<%if tmpEndiciaLabel<>"" then%><%=tmpEndiciaLabel%><%else%>label<%=tmp_TrackingNumber%>.<%=tmp_LabelFormat%><%end if%>" target="_blank">Print Label</a><%=strReceiptLink%>&nbsp;|&nbsp;<a href="javascript:openshipwin('USPS_TrackAndConfirmRequest.asp?TN=<%=tmp_TrackingNumber%>');">Track Package</a><%if (Now()<=tmpEndiciaExp) AND (tmp_Endicia="1") then%>&nbsp;|&nbsp;<a href="javascript: if(confirm('Are you sure you want to create a refund request for this USPS label?')) {location='EDC_refund.asp?id=<%=pcv_PackageID%>';}">Refund</a><%end if%><br>
															<% tmp_TrackingNumberShown=1
														end if%>
													<% else
														tmp_TrackingNumberShown=0
													end if %>
													<%if not IsNull(tmp_ShippedDate) then%>Shipped Date: <%=ShowDateFrmt(tmp_ShippedDate)%><br><%end if%>
													<%if tmp_TrackingNumber<>"" AND tmp_TrackingNumberShown=0 then%>Tracking Number: <%=tmp_TrackingNumber%><br><%end if%>
													<%if tmp_Comments<>"" then%>Comments:<br><i><%=tmp_Comments%></i><%end if%>
													</td>
												</tr>
												<tr style="background-color: #e5e5e5">
													<td>Product Name</td>
													<td>Qty</td>
													<td>Shipment Status</td>
												</tr>
											<% end if
											set rs1=nothing
										end if
									end if
									%>
									<tr>
										<td><%=pcv_PrdName%></td>
										<td><%=pcv_PrdQty%></td>
										<td>
										<% if pcv_Shipped="1" then
											if pcv_MF_UPS=1 then
												pcv_ModShipLang="View UPS shipment information"
												pcv_ModShipLink="UPS_ManageShipmentsResults.asp?id="&qry_ID
												pcv_DelShipMsg="This Shipment has a UPS label associated with it. Resetting the shipment does not void any existing UPS labels created for this order. You are about to delete all current shipping information for this package."
											else
												pcv_ModShipLang="Modify shipment information"
												pcv_ModShipLink="javascript:openshipwin('modifyPackInfo.asp?packID="&pcv_PackageID&"');"
												pcv_DelShipMsg="You are about to delete all current shipping information for this package."
											end if %>
											<span style="color: #0000FF">Shipped</span>&nbsp;(<a href="<%=pcv_ModShipLink%>"><%=pcv_ModShipLang%></a>&nbsp;|&nbsp;<a href="javascript:if (confirm('<%=pcv_DelShipMsg%> Are you sure you want to complete this action?')) location='modifyPackInfo.asp?m=del&id=<%=qry_ID%>&packID=<%=pcv_PackageID%>'">Reset shipment</a>)
												<% if pcv_MF_UPS=1 then
													pcv_ModShipLang="Modify shipment information"
													pcv_ModShipLink="javascript:openshipwin('modifyPackInfo.asp?info=UPS&packID="&pcv_PackageID&"');" %>
													 - <a href="<%=pcv_ModShipLink%>"><%=pcv_ModShipLang%></a>
												<% end if %>
										<%else
											if pcv_Shipped="2" then
												pcv_ModShipLink="USPS_ShipOrderWizard.asp?smode="&tmp_ServiceCode&"&orderID="&qry_ID&"&packID="&pcv_PackageID %>
												<span style="color: #00F;">Pending</span>&nbsp;|&nbsp;<a href="<%=pcv_ModShipLink%>">Ship &gt;&gt;</a>
											<% else %>
												<span style="color: #00F;">Pending</span>&nbsp;|&nbsp;<a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Ship &gt;&gt;</a>
											<% end if %>
										<%end if%>
										</td>
									</tr>
									<%
									if clng(pcv_PackageID)<>clng(save_packageID) then
										save_packageID=pcv_PackageID
									end if
									%>
									<tr>
										<td colspan="3"><hr></td>
									</tr>
									<%
									rs.MoveNext
								loop
								set rs=nothing
								%>
								<tr>
									<td colspan="3" class="pcCPspacer"></td>
								</tr>
							</table>
						<% end if %>

						<%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: SHIPPING DETAILS (PRE-VERSION 3)
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						if pshipvia <> "" OR ptrackingNum <> "" then
							'// if shipping info is relevant
							if (porderStatus>"3" AND porderStatus<"5") OR (porderStatus>"6" AND porderStatus<"9") then %>
								<table class="pcCPcontent">
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr style="background-color: #e5e5e5">
										<td colspan="2"><strong>Shipping Information</strong>:</td>
									</tr>
									<tr style="background-color: #FFFFFF">
										<td width="108">Shipped Via: </td>
										<td width="369">
										<input name="shipVia" type="text" size="30" value="<%=pshipvia%>"></td>
									</tr>
									<tr style="background-color: #FFFFFF">
										<td width="108">Date shipped:</td>
										<td width="369">
										<input name="shipDate" type="text" size="20" value="<%=pshipDate%>"> (Date format: <%=lcase(scDateFrmt)%>)</td>
									</tr>
									<tr style="background-color: #FFFFFF">
										<td width="108">Tracking Number:</td>
										<td width="369">
										<input name="trackingNum" type="text" size="20" value="<%=ptrackingNum%>">
										</td>
									</tr>
									<tr>
										<td colspan="2">
										<input type="checkbox" name="sendEmailShip14" value="YES">&nbsp;
										Resend <u>order shipped</u> e-mail when the shipping information is updated.
										</td>
									</tr>
									<tr>
										<td colspan="2"><input type="submit" name="Submit14" value="Update Shipping Information" class="submit14"></td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								</table>
							<% end if '// if (porderStatus>"3" AND porderStatus<"5") OR (porderStatus>"6" AND porderStatus<"9") then
						end if '// if pshipvia <> "" OR ptrackingNum <> "" then
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: SHIPPING DETAILS (PRE-VERSION 3)
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					END IF ' END - There is no shipping needed for this order

					'******************************************
					' END: Products not drop-shipped and not back-ordered
					'******************************************
					%>

				</div>

				<%
				'--------------
				' END TAB 2
				'--------------

				'--------------
				' START TAB 3
				'--------------
				%>

					<div id="tab3" class="TabbedPanes" style="<%=pcTab3Style%>">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<% if porderStatus<>"5" then %>
							<tr>
								<th colspan="2"><% if porderStatus="2" AND pcv_strDeactivateStatus=0 then %>Process or Cancel Order<% elseif pcv_strDeactivateStatus=0 then %>Update Order Status<% else %>Google Checkout Status<% end if %></th>
							</tr>
							<% end if %>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<%
							'///////////////////////////////////////////////////////
							'/// START: IF GOOGLE CHECKOUT
							'///////////////////////////////////////////////////////
							if pcv_strDeactivateStatus=1 then
								%><!--#include file="OrdDetails_GoogleCheckout.asp"--><%
							end if
							'///////////////////////////////////////////////////////
							'/// END: IF GOOGLE CHECKOUT
							'///////////////////////////////////////////////////////
							%>

							<%
							if porderStatus<>"5" AND pcv_strDeactivateStatus=0 then
							'///////////////////////////////////////////////////////
							'/// START: NORMAL ORDER - NOT CANCELLED
							'///////////////////////////////////////////////////////
								if porderStatus="2" then ' IF PENDING - START
							%>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td valign="top" align="right"><img src="images/quick.gif" width="24" height="18"></td>
										<td><strong>Process Order</strong>. &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=309')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
										<br />
										This order is currently &quot;<strong>Pending</strong>&quot;. After you have verified the accuracy and legitimacy of the order, you can update its status to &quot;processed&quot; by clicking on the button below. If you need to change the products that were added to the order, their quantity, discounts, taxes, and shipping charges, <a href="AdminEditOrder.asp?ido=<%=qry_ID%>">edit the order</a>.</td>
									</tr>
									<%
									query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=1;"
									set rsQ=connTemp.execute(query)
									pcv_AdmComments=""
									if not rsQ.eof then
										pcv_AdmComments=rsQ("pcACom_Comments")
									end if
									set rsQ=nothing%>
									<tr>
										<td></td>
										<td>Comments:<br>
										<textarea name="AdmComments1" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
										</td>
									</tr>
									<tr>
										<td align="right"><input type="checkbox" name="sendEmailConf" value="YES" checked class="clearBorder"></td>
										<td>Send <u>order confirmation</u> e-mail when the order status is updated to &quot;Processed&quot;.</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td>
											<input type="hidden" name="hidden1" id="hidden1">
											<input type="submit" onclick="this.disabled=true;hidden1.name=this.name;hidden1.value=this.value;this.form.submit()" name="Submit4" value="<% if pOrderStatus=2 then %>Process This Order<% else %>Update Order Status<% end if %>" class="submit2">
											&nbsp;
											<input type="button" onClick="location.href='batchprocessorders.asp'" value="Batch Process with Other Orders">
										</td>
									</tr>
									<tr>
										<td colspan="2"><hr></td>
									</tr>
									<tr>
										<td valign="top" align="right"><img src="images/delete2.gif" width="23" height="18"></td>
										<td valign="top">
										<strong>Cancel Order</strong>. &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=310')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
										<br />
										You may cancel this order at anytime before the order is shipped. To cancel, click on the &quot;Cancel&quot; button below.</td>
									</tr>
									<%'
									query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=5;"
									set rsQ=connTemp.execute(query)
									pcv_AdmComments=""
									if not rsQ.eof then
										pcv_AdmComments=rsQ("pcACom_Comments")
									end if
									set rsQ=nothing%>
									<tr>
										<td></td>
										<td align="top">Comments:<br>
										<textarea name="AdmComments5" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
										</td>
									</tr>
									<tr>
										<td align="right">
										<input type="checkbox" name="sendEmailCanc" value="YES" checked class="clearBorder"></td>
										<td>Send <u>order cancelled</u> e-mail when the order is cancelled.</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td><input type="submit" name="Submit6" value="Cancel Order" class="submit2"></td>
									</tr>
								<%
								end if ' IF PENDING - END

								if porderStatus>"2" then ' Order has been processed, show date
								%>
									<tr>
										<td colspan="2">This order was <b>processed</b> on: <b><%=pprocessDate%></b></td>
									</tr>
								<%
								end if

								if porderStatus ="4" then ' Order has been entirely shipped, show link
								%>
									<tr>
										<td colspan="2">This order was <b>shipped</b>: <a href="#" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab2');form2.ActiveTab.value = 2">see shipment(s) details &gt;&gt;</a></td>
									</tr>
								<%
								end if

								' IF SHIPPING - Start
								if porderstatus="7" then ' Order has been partially shipped: show links
								%>
									<tr>
										<td colspan="2"><hr></td>
									</tr>
									<tr>
										<td colspan="2">
											This order has been <strong>partially shipped</strong>.
											<ul style="list-style:circle; padding: 10px; margin-left: 15px;">
												<li><a href="#" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab2');form2.ActiveTab.value = 2">View partial shipment details</a></li>
												<li><a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Ship additional items</a></li>
											</ul>
										</td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								<%
								end if
								if ((porderStatus="3") or (porderStatus="8") or (porderStatus="13")) and (porderStatus<>"7") then
								%>
									<tr>
										<td colspan="2"><hr></td>
									</tr>
									<tr>
										<td valign="top"><img src="images/quick.gif" width="24" height="18"></td>
										<td width="95%">
											<strong>Ship Order</strong>
											<br />
											Ship this order using the <a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Shipping Wizard &gt;&gt;</a></td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td><a href="batchshiporders.asp">Batch Ship multiple orders at once</a></td>
									</tr>
									<tr>
										<td colspan="2"><hr></td>
									</tr>
									<tr>
										<td valign="top">&nbsp;</td>
										<td>If you cannot ship all products in this order, you can update the order status to &quot;<b>partially shipped</b>&quot; or &quot;<b>shipping</b>&quot;. This happens typically when the order contains back-ordered and/or drop-shipped products. You can update the order status:
										<ul style="list-style:circle; padding: 10px; margin-left: 15px;">
											<li><u>Automatically</u>, using the <a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Shipping Wizard</a> (recommended)</li>
											<li><u>Manually</u>, using the drop-down below.</li>
										</ul>
										</td>
									</tr>
									<tr>
										<td></td>
										<td>Status:
											<select name="pcv_OrderStatus">
												<option value="7" <%if porderStatus="7" then%>selected<%end if%>>Partially Shipped</option>
												<option value="8" <%if porderStatus="8" then%>selected<%end if%>>Shipping</option>
											</select>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td>
										<input type="submit" name="SubmitA2" value="Update Order Status" class="submit2">
										</td>
									</tr>
									<tr>
										<td colspan="2"><hr></td>
									</tr>
									<tr>
										<td valign="top" align="right"><img src="images/delete2.gif" width="23" height="18"></td>
										<td valign="top">
										<strong>Cancel Order</strong>. &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=310')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
										<br />
										You may cancel this order at anytime before the order is shipped. To cancel, click on the &quot;Cancel&quot; button below.</td>
									</tr>
									<%
									query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=5;"
									set rsQ=connTemp.execute(query)
									pcv_AdmComments=""
									if not rsQ.eof then
										pcv_AdmComments=rsQ("pcACom_Comments")
									end if
									set rsQ=nothing%>
									<tr>
										<td></td>
										<td align="top">Comments:<br>
										<textarea name="AdmComments5" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
										</td>
									</tr>
									<tr>
										<td align="right">
										<input type="checkbox" name="sendEmailCanc2" value="YES" checked class="clearBorder"></td>
										<td>Send <u>order cancelled</u> e-mail when the order is cancelled.</td>
									</tr>
									<tr>
										<td>&nbsp; </td>
										<td>
										<input type="submit" name="Submit10" value="Cancel Order" class="ibtnGrey"></td>
									</tr>
								<%
								end if ' IF SHIPPING - End

								if RMAVar=0 AND RMAStatus=0 AND porderStatus>"3" then
								%>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<th colspan="2">Returns</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2"><strong>Create Return Merchandise Authorization Number</strong> (&quot;RMA&quot; Number)</td>
									</tr>
									<tr>
										<td valign="top"><img src="images/quick.gif" width="24" height="18"></td>
										<td><input type="button" name="RMAInfo" value="Create RMA Number" class="submit2" onClick="location.href='genRma.asp?idOrder=<%=qry_ID%>'"></td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								<%
								end if

								if RMAVar=1 AND RMAStatus=0 then %>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2" valign="top"><strong>RMA Information</strong></td>
									</tr>
									<tr>
										<td colspan="2">
										<% if trim(prmaNumber)="" then %>
										A customer requested a RMA (Return Merchandise Authorization). Click below to generate an RMA number and send it to the customer.
										<% else %>
										A Return Merchandise Authorization number was created for this order. Here are the details:
										<% end if %>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td>
											<table class="pcCPcontent" style="width:auto;">
												<tr>
													<td nowrap="nowrap">Date Submitted: </td>
													<td nowrap="nowrap"><%=ShowDateFrmt(pRMADate)%></td>
												</tr>
												<tr>
													<td nowrap="nowrap">RMA Number:</td>
													<td nowrap="nowrap"><%=pRmaNumber%></td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td><input type="button" name="RMAInfo" value="View RMA Information" onClick="location.href='modRmaa.asp?idOrder=<%=qry_ID%>&idRMA=<%=pIdRMA%>'"></td>
									</tr>
								<% end if %>

								<% if RMAVar=1 AND RMAStatus=1 then %>
									<tr>
										<td colspan="2">&nbsp;</td>
									</tr>
									<tr>
										<td colspan="2"><b>RMA Information</b></td>
									</tr>
									<tr>
										<td valign="top">&nbsp;</td>
										<td>
										<table class="pcCPcontent" style="width:auto;">
											<tr>
												<td nowrap="nowrap">RMA Number:</td>
												<td nowrap="nowrap"><b><%=pRmaNumber%></b></td>
											</tr>
											<tr>
												<td nowrap="nowrap">Status:</td>
												<td><b><%=pRmaStatus%></b></td>
											</tr>
										</table>
										</td>
									</tr>
									<tr>
										<td valign="top">&nbsp;</td>
										<td><input type="button" name="RMAInfo" value="View RMA Information" onClick="location.href='modRmaa.asp?idOrder=<%=qry_ID%>&idRMA=<%=pIdRMA%>'"></td>
									</tr>
								<% end if

							'///////////////////////////////////////////////////////
							'/// END: NORMAL ORDER - NOT CANCELLED
							'///////////////////////////////////////////////////////
							end if
							%>

							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2">Reset Order Status</th>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td valign="top"><img src="images/move.gif" width="25" height="18"></p></td>
								<td>
								<strong>Reset Order Status</strong>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=311')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
								<br />

								<% if pcv_strDeactivateStatus=1 then %>
									<div style="padding:6px"><span class="pcCPnotes">In rare cases you may need to manually update the order status of a Google Order. <strong>Proceed with caution!</strong>  ProductCart <strong><u>can not</u> synchronize the new order status</strong> with the Google Checkout Merchant Center.</span></div>
								<% end if %>

								To reset the order status of this order, select a status from the drop-down menu below and click on &quot;Reset Order Status&quot;.

								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>
									<select name="resetstat" id="resetstat">
										<option value="2">Pending</option>
										<option value="3">Processed</option>
										<%'Start SDBA%>
										<option value="7">Partially Shipped</option>
										<option value="8">Shipping</option>
										<%'End SDBA%>
										<option value="4">Shipped</option>
										<%'Start SDBA%>
										<option value="9">Partially Return</option>
										<%'End SDBA%>
										<option value="6">Return</option>
										<option value="5">Canceled</option>
										<% if GOOGLEACTIVE=-1 then %>
										<option value="10" >Declined</option>
										<option value="12" >Archived</option>
										<% end if %>
									</select>
									<% 'hidden values to use for comparison when confirming status change, and for managing button clicks %>
									<input type="hidden" name="oldstat" value="<% =porderstatus%>">
									<input type="hidden" name="checkreset" value="">
								</td>
							</tr>

							<tr>
								<td>&nbsp;</td>
								<td>
								<!-- Display a pop-up dialog to confirm desire to reset order status; set flag to manage button clicks -->
								<input type="submit" name="Submit7" value="Reset Order Status" onclick="if(confirm('Do you really want to reset the Order Status?')) {document.form2.checkreset.value = 'true';} else {document.form2.checkreset.value = 'false';}">
								</td>
							</tr>

							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
						</table>

						<%
						if pcv_strDeactivateStatus=0 then
						'******************************************
						' START: Resend emails
						'******************************************
						%>
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2">Resend E-mail Messages</th>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2">
								<strong>Order Received E-mail</strong><br>
								Use the button below to <u>resend the Order Received e-mail</u> to this customer (e.g. the customer reports that they did not receive it).
								</td>
							</tr>
							<%
							' Resend Confirmation email
							query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=0;"
							set rsQ=connTemp.execute(query)
							pcv_AdmComments=""
							if not rsQ.eof then
								pcv_AdmComments=rsQ("pcACom_Comments")
							end if
							set rsQ=nothing%>
							<tr>
								<td colspan="2">
								<div>Comments:</div>
								<div style="padding: 5px 0 5px 0;">
									<textarea name="AdmComments0" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
								</div>
								<input type="submit" name="Submit4B" value="Resend Order Received E-mail" class="submit2">
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<%
							' Resend Confirmation email
							if ((porderStatus>"2") and (porderStatus<"5")) or ((porderStatus>"6") and (porderStatus<"9")) then%>
								<tr>
									<td colspan="2">
									<strong>Order Confirmation E-mail</strong><br>
									Use the button below to <u>resend the Order Confirmation e-mail</u> to this customer (e.g. the customer reports that they did not receive it).
									</td>
								</tr>
								<%
								query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=1;"
								set rsQ=connTemp.execute(query)
								pcv_AdmComments=""
								if not rsQ.eof then
									pcv_AdmComments=rsQ("pcACom_Comments")
								end if
								set rsQ=nothing%>
								<tr>
									<td colspan="2">
									<div>Comments:</div>
									<div style="padding: 5px 0 5px 0;">
										<textarea name="AdmComments1A" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
									</div>
									<input type="submit" name="Submit4A" value="Resend Order Confirmation e-mail" class="submit2">
									</td>
								</tr>
								<tr>
									<td colspan="2"><hr></td>
								</tr>
							<%
							end if

							'Resend order shipped e-mail
							if porderStatus="4" then

								query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=3;"
								set rsQ=connTemp.execute(query)
								pcv_AdmComments=""
								if not rsQ.eof then
									pcv_AdmComments=rsQ("pcACom_Comments")
								end if
								set rsQ=nothing
							%>
								<tr>
									<td colspan="2">
										<strong>Order Shipped E-mail</strong><br>
										Use the button below to <u>resend the Order Shipped e-mail</u> to this customer (e.g. the customer reports that they did not receive it).
									</td>
								</tr>
								<tr>
									<td colspan="2">
									<div>Comments:</div>
									<div style="padding: 5px 0 5px 0;">
										<textarea name="AdmComments3A" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
										<input name="sendEmailShip" type="hidden" value="YES">
									</div>
									<input type="submit" name="Submit5A" value="Resend Order Shipped Email" class="submit2"></td>
								</tr>
								<tr>
									<td colspan="2"><hr></td>
								</tr>
							<%
							end if

							'Resend order cancelled e-mail
							if porderStatus="5" then
								query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=5;"
								set rsQ=connTemp.execute(query)
								pcv_AdmComments=""
								if not rsQ.eof then
									pcv_AdmComments=rsQ("pcACom_Comments")
								end if
								set rsQ=nothing
							%>
								<tr>
									<td colspan="2">
										Use the button below to <u>resend the Order Cancelled email</u> to this customer (e.g. the customer reports that they did not receive it).
									</td>
								</tr>
								<tr>
									<td colspan="2">
									<div>Comments:</div>
									<div style="padding: 5px 0 5px 0;">
									<textarea name="AdmComments5A" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
									</div>
									<input type="submit" name="Submit10A" value="Resend Order Cancelled Email" class="submit2">
									</td>
								</tr>
								<tr>
									<td colspan="2"><hr></td>
								</tr>
						<%
							end if

						'******************************************
						' END: Resend emails
						'******************************************
						%>
						</table>
					<% end if %>
					</div>

				<%
				'--------------
				' END TAB 3
				'--------------

				'--------------
				' START TAB 4
				'--------------
				%>

					<div id="tab4" class="TabbedPanes" style="<%=pcTab4Style%>">
					<%

					' Was FedEx used for this order?
					Dim intShipFedExWizard
					intShipFedExWizard = 1
					intShipUPSWizard = 1
					intShipUSPSWizard = 1
					'******************************************
					' START: Show Shipping Wizard
					'******************************************
					if ( (porderStatus <= 2 AND pcv_strDeactivateStatus=0) OR (pcv_strDeactivateStatus=1 AND pcv_PaymentStatus<2) ) then ' START Order must be processed before it can be shipped
					%>
							<table class="pcCPcontent">
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td colspan="2">
									<% if pcv_strDeactivateStatus=0 then %>
									This order has not yet been processed. An order must be processed before it can be shipped. <a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', 'current');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab3');form2.ActiveTab.value = 3">Process this order &gt;&gt;</a>
									<% else %>
									This order has not yet been "Charged". An order must be "Charged" before you "Ship and confirm". <a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', 'current');change('tabs4', '');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab3');form2.ActiveTab.value = 3">Charge this order &gt;&gt;</a>
									<% end if %>
									</td>
								</tr>
							</table>
					<%
					else ' The order can be shipped. Show the shipping wizard.
						'// if porderStatus <> 10 then
					%>
								<table class="pcCPcontent">
									<tr>
										<th colspan="2">Custom Shipping Options Shipping Wizard </th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2">Use the button below to start the Shipping Wizard. This will allow you to ship your order using one of the custom shipping options that you have setup on your store. <% if pcv_strFedExEnabled or pcv_strUpsEnabled then %>If you are planning to use FedEx or UPS for your shipment use either the <strong>FedEx Shipping Wizard</strong> for FedEx shipments or the <strong>UPS Shipping Wizard</strong> for UPS shipments.<% end if %><br><br></td>
									</tr>
									<tr>
											<td colspan="2"><input type="button" name="shiporderwizard" value="Start Custom Shipping Options Shipping Wizard" onclick="location='sds_ShipOrderWizard1.asp?idorder=<%=qry_ID%>';" class="submit2"></td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"><hr></td>
									</tr>
								</table>
								<%
								dim pcv_strFedWSEnabled

								' If the FedEx Web Services integration is active change links
								query="SELECT active FROM ShipmentTypes WHERE idShipment=9"
								set rsFedEx=Server.CreateObject("ADODB.Recordset")
								set rsFedEx=connTemp.execute(query)
								if NOT rsFedEx.eof then
									pcv_strFedWSEnabled = 1
									pcv_intFedExAction = "FedExWs"
									pcv_strFedExShipURL = "FedExWS_ManageShipmentsResults.asp"
								else
									pcv_strFedWSEnabled = 0
									pcv_intFedExAction = "FedEx"
									pcv_strFedExShipURL = "FedEx_ManageShipmentsResults.asp"
								end if

								if (pcv_strFedExEnabled and intShipFedExWizard = 1) OR pcv_strFedWSEnabled = 1 then ' START show FedEx shipping wizard
								%>
								<a name="FedExshipwizard">&nbsp;</a>
								<a name="managepackages">&nbsp;</a>
								<table class="pcCPcontent">
									<tr>
										<th colspan="2">FedEx<sup>&reg;</sup> Shipping Wizard</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2"><img src="images/Clct_Prf_2c_Pos_Plt_150.png"></td>
									</tr>
									<tr>
										<td colspan="2">Use the following button(s) to access the FedEx<sup>&reg;</sup> Shipping Center</td>
									</tr>
									<tr>
										<td colspan="2">
											<input type="button" name="ShipManager" value="Start FedEx Shipping Wizard" onclick="location='sds_ShipOrderWizard1.asp?idorder=<%=qry_ID%>&PageAction=<%=pcv_intFedExAction%>';" class="submit2">
										</td>
									</tr>
									<%
									' // Show Manage Shipped Packages
									if pcv_strFedExPackagesExist AND pcv_MF_FEDEX=1 then
									%>
									<tr>
										<td colspan="2">
											<input type="button" name="managepackages" value="Manage Shipped Packages" onclick="location='<%=pcv_strFedExShipURL%>?id=<%=qry_ID%>';">
										</td>
									</tr>
									<% end if %>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2" align="center"><% = pcf_FedExWriteLegalDisclaimers %></td>
									</tr>
									<tr>
										<td colspan="2"><hr></td>
									</tr>
								</table>
							<%
							end if ' END show FedEx Shipping Wizard

							if pcv_strUPSEnabled and intShipUPSWizard = 1 then ' START show FedEx shipping wizard
							%>
								<a name="UPSshipwizard">&nbsp;</a>
								<a name="managepackages">&nbsp;</a>
								<table class="pcCPcontent">
									<tr>
										<th colspan="2">UPS OnLine&reg; Tools Shipping</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50"></td>
									</tr>
									<tr>
										<td colspan="2">Use the following button(s) to access UPS OnLine&reg; Tools Shipping</td>
									</tr>
									<tr>
										<td colspan="2">
											<input type="button" name="ShipManager" value="Start Shipping Wizard" onclick="location='sds_ShipOrderWizard1.asp?idorder=<%=qry_ID%>&PageAction=UPS&m=new';" class="submit2">
										</td>
									</tr>
									<%
									' // Show Manage Shipped Packages
									if pcv_strUPSEnabled AND pcv_strUPSPackagesExist AND pcv_MF_UPS=1 then
									%>
									<tr>
										<td colspan="2">
											<input type="button" name="managepackages" value="Manage Shipped Packages" onclick="location='UPS_ManageShipmentsResults.asp?id=<%=qry_ID%>';">
										</td>
									</tr>
									<% end if %>
									<tr>
										<td colspan="2"><hr></td>
									</tr>
									<tr>
										<td colspan="2" align="center"><% = pcf_UPSWriteLegalDisclaimersText %></td>
									</tr>
								</table>
							<%
							end if ' END show FedEx Shipping Wizard
							if pcv_strUSPSEnabled and intShipUSPSWizard = 1 then ' START show USPS shipping wizard
							%>
								<a name="USPSshipwizard">&nbsp;</a>
								<a name="managepackages">&nbsp;</a>
								<table class="pcCPcontent">
									<tr>
										<th colspan="2">U.S.P.S. Shipping Labels</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2"><img src="images/hdr_uspsLogo.jpg"></td>
									</tr>
									<% if pcv_strUSPSURLActive then %>
									<tr>
										<td colspan="2">Use the following button to access U.S.P.S. Shipping Labels Wizard. Using this wizard, you will be able to generate Delivery Confirmation, Signature Confirmation and Express Mail Labels.</td>
									</tr>
									<tr>
										<td colspan="2">
											<input type="button" name="ShipManager" value="Start Shipping Wizard" onclick="location='sds_ShipOrderWizard1.asp?idorder=<%=qry_ID%>&PageAction=USPS&m=new';" class="submit2">
										</td>
									</tr>
									<% else %>
									<tr>
										<td colspan="2">It appears that you have not set a "USPS Secured Server" URL. In order to use the USPS label wizard, you must have a value for this URL: <a href="USPS_EditLicense.asp">Set the &quot;USPS Secured Server&quot; URL now</a></td>
									</tr>
									<% end if
									' // Show Manage Shipped Packages
									if pcv_strUSPSEnabled AND pcv_strUSPSPackagesExist AND pcv_MF_USPS=1 then
									%>
									<tr>
										<td colspan="2">
											<input type="button" name="managepackages" value="Manage Shipped Packages" onclick="location='USPS_ManageShipmentsResults.asp?id=<%=qry_ID%>';">
										</td>
									</tr>
									<% end if %>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								</table>
							<%
							end if ' END show FedEx Shipping Wizard
						'// if porderStatus <> 10 then
					end if ' END Order must be processed
					%>
				</div>

				<%
				'--------------
				' END TAB 4
				'--------------

				'--------------
				' START TAB 5
				'--------------
				%>

				<div id="tab5" class="TabbedPanes" style="<%=pcTab5Style%>">
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">
							<% if pcv_strDeactivateStatus=0 then %>
							Update Payment Status
							<% else %>
							Payment Status
							<% end if %>
							</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">
							<% if pcv_strDeactivateStatus=0 then %>
							To update this order's the payment status select the new status below and click on &quot;Update Payment Status&quot;. Note that the payment gateway involved, if any, is <u>not</u> contacted when you change the payment status.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=302')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
							<% else %>
							This order was placed with Google Checkout. The payment status of this order will update as you move through the order fulfillment process. Use ProductCart for common commands, such as charging and shipping orders, and the Merchant Center for unusual events, such as initiating a refund. Google keeps order states correct no matter which method you use and then it updates ProductCart.
							<% end if %>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">
								Current Payment Status: <b><%=pcv_PayStatusName%></b>
							</td>
						</tr>

						<tr>
							<td colspan="2">
								<% if pcv_strDeactivateStatus=1 then %>
									<div style="padding:6px"><span class="pcCPnotes">In rare cases you may need to manually update the payment status of a Google Order. <strong>Proceed with caution!</strong>  ProductCart <strong><u>can not</u> synchronize the new payment status</strong> with the Google Checkout Merchant Center.</span></div>
								<% end if %>

								<select name="pcv_PaymentStatus">
									<option value="0" <%if pcv_PaymentStatus="0" then%>selected<%end if%>>Pending</option>
									<option value="1" <%if pcv_PaymentStatus="1" then%>selected<%end if%>>Authorized</option>
									<option value="2" <%if pcv_PaymentStatus="2" then%>selected<%end if%>>Paid</option>
									<option value="6" <%if pcv_PaymentStatus="6" then%>selected<%end if%>>Refunded</option>
									<option value="8" <%if pcv_PaymentStatus="8" then%>selected<%end if%>>Voided</option>
									<% if GOOGLEACTIVE=-1 then %>
									<option value="3" <%if pcv_PaymentStatus="3" then%>selected<%end if%>>Declined</option>
									<option value="4" <%if pcv_PaymentStatus="4" then%>selected<%end if%>>Cancelled</option>
									<option value="5" <%if pcv_PaymentStatus="5" then%>selected<%end if%>>Cancelled By Google</option>
									<option value="7" <%if pcv_PaymentStatus="7" then%>selected<%end if%>>Charging</option>
									<% end if %>
								</select>
							</td>
						</tr>
						<tr>
							<td colspan="2"><input type=submit name="SubmitA1" value=" Update payment status " class="submit2" onclick="if(confirm('Please note that the payment gateway involved with this order (if any), will not be contacted when you manually update the payment status. Do you really want to change the Payment Status?')) {} else {return(false);}"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<%
						'///////////////////////////////////////////////////////
						'/// START: IF EIG
						'///////////////////////////////////////////////////////
						'EIG S
						query = "SELECT idOrder FROM pcPay_EIG_Authorize WHERE idOrder="&qry_ID&";"
						set rsEIGObj=Server.CreateObject("ADODB.Recordset")
						set rsEIGObj=connTemp.execute(query)
						if NOT rsEIGObj.eof then

						%><!--#include file="OrdDetails_EIG.asp"--><%
						end if
						'EIG E
						'///////////////////////////////////////////////////////
						'/// END: IF EIG
						'///////////////////////////////////////////////////////
						%>
						<tr>
							<th colspan="2">Payment Details</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
							<tr>
								<td colspan="2">Payment Method: <%=PaymentType%></td>
							</tr>
							<%
							varTransID=""
							varTransName="Transaction ID"
							varAuthCode=""
							varAuthName="Authorization Code"
							pcv_PendingString=""
							%>
							<% if NOT isNull(pcpaymentCode) AND pcpaymentCode<>"" then
								varShowCCInfo=0
								varPayer=""
								select case pcpaymentCode
								case "LinkPoint"
									varAry=split(pcgwAuthCode,":")
									varTransName="Approval Number"
									varAuthName="Reference Number"
									varTransID=left(varAry(1),6)
									varAuthCode=right(varAry(1),10)
									varAVS=left(varAry(2),2)
									varDisplayAVS=left(varAry(2),3)

									select case varAVS
										case "YY"
											varAVSResult="Address matches, zip code matches"
										case "YN"
											varAVSResult="Address matches, zip code does not match"
										case "YX"
											varAVSResult="Address matches, zip code comparison not available"
										case "NY"
											varAVSResult="Address does not match, zip code matches"
										case "XY"
											varAVSResult="Address comparison not available, zip code matches"
										case "NN"
											varAVSResult="Address comparison does not match, zip code does 									not match"
										case "NX"
												varAVSResult="Address does not match, zip code comparison not 									available"
										case "XN"
												varAVSResult="Address comparison not available, zip code does not									match"
										case "XX"
												varAVSResult="Address comparisons not available, zip code 									comparison not available"
										case else
												varAVSResult="No AVS result found."
									end select

									varCVV=right(varAry(2),1)
									select case varCVV
										case "M"
											varCVVResult="Card Code Match"
										case "N"
											varCVVResult="Card code does not match"
										case "P"
											varCVVResult="Not processed"
										case "S"
											varCVVResult="Merchant has indicated that the card code is not present on the card"
										case "U"
											varCVVResult="Issuer is not certified and/or has not provided encryption keys"
										case "X"
											varCVVResult="No response from the credit card association was received"
										case else
											varCVVReseult="No code received"
									end select
								case "PFLink", "PFPro", "PFPRO", "PFLINK"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
									varShowCCInfo=1
									varGWInfo="P"
								case "Authorize"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
									varShowCCInfo=1
									if instr(ucase(PaymentType),"CHECK") then
										varShowCCInfo=0
									end if
									varGWInfo="A"
								case "twoCheckout"
									varTransName="2Checkout Order No"
									varTransID=pcgwTransId
								case "BOFA"
									varTransName="Order No"
									varAuthName="Authorization Code"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
								case "WorldPay"
									varTransID=""
									varAuthCode=""
								case "iTransact"
									varTransName="Transaction ID"
									varAuthName="Authorization Code"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
								case "PSI", "PSIGate"
									varTransName="Transaction ID"
									varAuthName="Authorization Code"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
								case "fasttransact", "FastTransact", "FAST","CyberSource","HSBC"
									varTransName="Transaction ID"
									varAuthName="Authorization Code"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
								case "USAePay","FastCharge"
									varTransName="Transaction reference code"
									varAuthName="Authorization code"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
								case "EIG"
									varAuthCode=pcgwAuthCode
									varTransID=pcgwTransId
									varShowCCInfo=1
									varGWInfo="EIG"
								case "PayPalExp"
									varPayer=pcOrd_Payer
								case "PxPay"
									varTransName="DPS Transaction Reference Number"
									varAuthName="Authorization code"
									varTransID=pcgwTransId
									varAuthCode=pcgwAuthCode
								case "PayPal"
									varAuthCode=pcgwAuthCode
									varTransID=pcgwTransId

									if lcase(pcgwAuthCode)="pending" then
										select case lcase(varTransID)
											case "echeck"
												pcv_PendingString="Pending Reason: eCheck - The payment is pending because it was made by an eCheck, which has not yet cleared"
											case "multi_currency"
												pcv_PendingString="Pending Reason: Multi Currency - You do not have a balance in the currency sent, and you do not have your Payment Receiving Preferences set to automatically convert and accept this payment. You must manually accept or deny this payment from your Account Overview"
											case "intl"
												pcv_PendingString="Pending Reason: International - The payment is pending because you, the merchant, hold a non-U.S. account and do not have a withdrawal mechanism. You must manually accept or deny this payment from your Account Overview"
											case "verify"
												pcv_PendingString="Pending Reason: Not Verified - The payment is pending because you, the merchant, are not yet Verified. You must Verify your account before you can accept this payment"
											case "address"
												pcv_PendingString="Pending Reason: Address Not Confirmed - The payment is pending because your customer did not include a confirmed shipping address and you, the merchant, have your Payment Receiving Preferences set such that you want to manually accept or deny each of these payments. To change your preference, go to the 'Selling Preferences' section of your Profile"
											case "upgrade"
												pcv_PendingString="Pending Reason: Account Upgrade Required - The payment is pending because it was made via credit card and you, the merchant, must upgrade your account to Premier or Business status in order to receive the funds. It may also mean that you have reached the monthly limit for transactions on your account"
											case "unilateral"
												pcv_PendingString="Pending Reason: E-mail address not registered - The payment is pending because it was made to an email address that is not yet registered or confirmed with PayPal "
											case "other"
												pcv_PendingString="Pending Reason: Unknown - The payment is pending for a unknown reason. For more information, please contact Customer Service"
											case else
												pcv_PendingString="Pending Reason: Unknown - The payment is pending for a unknown reason. For more information, please contact Customer Service"
										end select
										varAuthCode=""
										varTransID=""
									end if
								end select
								%>
								<% if varPayer<>"" then %>
									<tr>
										<td colspan="2">PayPal E-mail: <a href="mailto:<%=varPayer%>"><%=varPayer%></a></td>
									</tr>
								<% end if %>

								<% if pcpaymentCode="VM" then 
									pcvmAVSResponse = ""
									pcvmCVV2Response = ""

									'AVS Response Codes
									select case pcv_strAVSRespond
										case "A"
											pcvmAVSResponse = "Address matches - ZIP Code does not match"
										case "B"
											pcvmAVSResponse = "Street address match, Postal code in wrong format (International issuer)"
										case "C"
											pcvmAVSResponse = "Street address and postal code in wrong formats"
										case "D"
											pcvmAVSResponse = "Street address and postal code match (international issuer)"
										case "E"
											pcvmAVSResponse = "AVS Error"
										case "F"
											pcvmAVSResponse = "Address does compare and five-digit ZIP code does compare (UK only)"
										case "G"
											pcvmAVSResponse = "Service not supported by non-US issuer"
										case "I"
											pcvmAVSResponse = "Address information not verified by international issuer."
										case "M"
											pcvmAVSResponse = "Street Address and Postal code match (international issuer)"
										case "N"
											pcvmAVSResponse = "No Match on Address (Street) or ZIP"
										case "O"
											pcvmAVSResponse = "No Response sent"
										case "P"
											pcvmAVSResponse = "Postal codes match, Street address not verified due to incompatible formats"
										case "R"
											pcvmAVSResponse = "Retry, System unavailable or Timed out"
										case "S"
											pcvmAVSResponse = "Service not supported by issuer"
										case "U"
											pcvmAVSResponse = "Address information is unavailable"
										case "W"
											pcvmAVSResponse = "9-digit ZIP matches, Address (Street) does not match"
										case "X"
											pcvmAVSResponse = "Exact AVS Match"
										case "Y"
											pcvmAVSResponse = "Address (Street) and 5-digit ZIP match"
										case "Z"
											pcvmAVSResponse = "5-digit ZIP matches, Address (Street) does not match"
										case else
											pcvmAVSResponse = "No response recieved"
									end select
	
									select case pcv_strCVNResponse
										case "M"
											pcvmCVV2Response = "CVV2 Match"
										case "N"
											pcvmCVV2Response = "CVV2 No match"
										case "P"
											pcvmCVV2Response = "Not Processed"
										case "S"
											pcvmCVV2Response = "Issuer indicates that CVV2 data should be present on the card, but the merchant has indicated that the CVV2 data is not resent on the card"
										case "U"
											pcvmCVV2Response = "Issuer has not certified for CVV2 or Issuer has not provided Visa with the CVV2 encryption keys"
										case else
											pcvmCVV2Response = "No response recieved"
									end select
									%>
									<tr>
										<td colspan="2">AVS Code: <%=pcvmAVSResponse%></td>
									</tr>
									<tr>
										<td colspan="2">Card Code Information: <%=pcvmCVV2Response%></td>
									</tr>
								<% end if %>


								<% if pcpaymentCode="LinkPoint" then %>
									<tr>
										<td colspan="2">AVS Code: <%=varDisplayAVS&": "&varAVSResult%></td>
									</tr>
									<tr>
										<td colspan="2">Card Code Information: <%=varCVVResult%></td>
									</tr>
									<tr>
										<td colspan="2">LinkPoint Order ID: <%=pcgwTransId%></td>
									</tr>
								<% end if %>
								<% if varTransID<>"" then %>
									<tr>
										<td colspan="2"><%=varTransName%>: <%=varTransID%></td>
									</tr>
								<% end if %>
								<% if varAuthCode<>"" then %>
									<tr>
										<td colspan="2"><%=varAuthName%>: <%=varAuthCode%></td>
									</tr>
								<% end if %>
								<% if pcv_PendingString<>"" then %>
									<tr>
										<td colspan="2">PayPal Payment Status: Pending</td>
									</tr>
									<tr>
										<td colspan="2"><%=pcv_PendingString%></td>
									</tr>
								<% end if %>
								<% if varShowCCInfo=1 and (varGWInfo="A" OR varGWInfo="P" OR varGWInfo="EIG") then
									if varGWInfo="A" then
										query="SELECT ccnum, ccexp, pcSecurityKeyID FROM authorders WHERE idOrder=" & qry_ID & ";"
									elseif varGWInfo="EIG" then
										query="SELECT ccnum, ccexp, cctype, pcSecurityKeyID FROM pcPay_EIG_Authorize WHERE idOrder=" & qry_ID & ";"
									else
										query="SELECT acct, expdate, pcSecurityKeyID FROM pfporders WHERE idOrder=" & qry_ID & ";"
									end if
									Set rs=Server.CreateObject("ADODB.Recordset")
									set rs=connTemp.execute(query)
									if NOT rs.eof then
										if varGWInfo="A" then
											pcardNumber=rs("ccnum")
											pexpiration=rs("ccexp")
											pcv_SecurityKeyID=rs("pcSecurityKeyID")
										elseif varGWInfo="EIG" then
											pcardNumber=rs("ccnum")
											pexpiration=rs("ccexp")
											VarCCType=rs("cctype")
											pcv_SecurityKeyID=rs("pcSecurityKeyID")
										else
											pcardNumber=rs("acct")
											pexpiration=rs("expdate")
											pcv_SecurityKeyID=rs("pcSecurityKeyID")
										end if
										set rs=nothing
										if pcardNumber="*" then %>
											<tr>
												<td colspan="2">Credit Card information has been purged for this order.</td>
											</tr>
										<% else

											pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

											VarCCNum=pcardNumber
											VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)
											if varGWInfo="EIG" then
												Select Case VarCCType
													Case "V": VarCCType="Visa"
													Case "M": VarCCType="MasterCard"
													Case "A": VarCCType="American Express"
													Case "D": VarCCType="Discover"
													Case Else : VarCCType=VarCCType
												End Select
											else
											VarCCType=ShowCardType(VarCCNum2)
											end if
											VarCCNum2=ShowLastFour(VarCCNum2)
											%>
											<tr>
												<td colspan="2">Card Number: <%=VarCCNum2%></td>
											</tr>
											<tr>
												<td colspan="2">Card Type: <%=VarCCType%></td>
											</tr>
											<tr>
												<td colspan="2">Expiration Date: <%=left(pexpiration,2)&"/"&right(pexpiration,2)%></td>
											</tr>
										<% end if
									end if
								end if
							end if %>

							<%
							'=================================
							'Check for CC Order
							'====================================
							dim intShowBtn, intShowPurgeBtn
							intShowBtn = 0
							intShowPurgeBtn = 0
							query="SELECT cardType,cardNumber,expiration,comments,pcSecurityKeyID FROM creditCards WHERE idOrder=" & qry_ID & ";"
							Set rs=Server.CreateObject("ADODB.Recordset")
							set rs=connTemp.execute(query)
							if NOT rs.eof then
								intShowBtn=1
								intShowPurgeBtn = 1
								pcardType=rs("cardType")
								pcardNumber=rs("cardNumber")
								pexpiration=rs("expiration")
								pcardComments=rs("comments")
								pcv_SecurityKeyID=rs("pcSecurityKeyID")
								CCT=pcardType
								ccp="Y"
								If CCT="M" then
									CCType="MasterCard"
								end if
								If CCT="V" then
									CCType="Visa"
								end if
								If CCT="D" then
									CCType="Discover"
								end if
								If CCT="A" then
									CCType="American Express"
								end if
								If CCT="DC" then
									CCType="Diner's Club"
								end if
								%>
								<tr>
									<td>Card Type:
									<input type="hidden" name="ccp" value="<%=ccp%>">
									<input type="hidden" name="CCT" value="<%=CCT%>">
									</td>
									<td>
									<input type="text" name="CCType" value="<%=CCType%>" size="20"></td>
								</tr>

								<tr>
									<td>Card Number:</td>
									<td>
									<%
									pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

									Dim VarCCNum
									VarCCNum=pcardNumber
									VarCCNum2=enDeCrypt(VarCCNum, pcv_SecurityPass)
									%>
									<input type="text" name="CCNum" value="<%=VarCCNum2%>" size="20">
									</td>
								</tr>

								<tr>
									<td>Expiration Date:</td>
									<td>Month:
									<input type="text" name="CCexpM" value="<%=Month(pexpiration)%>" size="2">
									Year:
									<input type="text" name="CCexpY" value="<%=Year(pexpiration)%>" size="4">
									</td>
								</tr>
								<%
								' This field is no longer used in v3. Billing information is in the billing & shipping tab.
								' However, stores that upgraded from v2.76 might have older orders that have billing info saved here.
								if trim(pcardComments) <> "" then
								%>
									<tr>
										<td valign="top">Billing Information:</td>
										<td><textarea name="CCcomments" cols="30" rows="5"><%=pcardComments%></textarea></td>
									</tr>
								<%
								end if
							end if

							'=================================
							'Check for Custom Payment
							'====================================
							query="SELECT customCardOrders.idCCOrder, customCardOrders.idOrder, customCardOrders.strFormValue, customCardOrders.strRuleName, customCardOrders.idCustomCardRules FROM customCardOrders WHERE ((customCardOrders.idOrder)=" & qry_ID & ") ORDER BY customCardOrders.idCCOrder;"
							Set rs=Server.CreateObject("ADODB.Recordset")
							set rs=connTemp.execute(query)
							custcardtype=0

							if NOT rs.eof then
								intShowBtn=1
								custcardtype=1
								ccCnt=0
								do until rs.eof
									pIdCCOrder=rs("idCCOrder")
									pStrFormValue=rs("strFormValue")
									pStrRuleName=rs("strRuleName")
									pTempIdCCRules=rs("idCustomCardRules")
									'check length of field
									query="SELECT intlengthOfField, intmaxInput FROM customCardRules WHERE idcustomCardRules="&pTempIDCCRules&";"
									Set rsRulObj=Server.CreateObject("ADODB.Recordset")
									Set rsRulObj=connTemp.execute(query)
									if rsRulObj.eof then
										pLOF="20"
										pMaxInput="999"
									else
										pLOF=rsRulObj("intlengthOfField")
										pMaxInput=rsRulObj("intmaxInput")
									end if
									set rsRulObj=nothing
									'pIdCCR=rs("idcustomCardRules")
									if pMaxInput="" or pMaxInput="0" then
										pMaxInput= pLOF
									end if
									ccCnt=ccCnt+1%>
									<tr>
										<td><%=pStrRuleName&": "%></td>
										<td>
										<input type="hidden" name="CCID<%=ccCnt%>" value="<%=pIdCCOrder%>">
										<input name="CCO<%=pIdCCOrder%>" type="text" value="<%=pStrFormValue%>" size="<%=pLOF%>" maxlength="<%=pMaxInput%>">
										</td>
									</tr>
									<% rs.moveNext
								loop
								set rs=nothing %>
							<% end if %>
							<input name="ccCnt" type="hidden" value="<%=ccCnt%>">
							<input name="custcardtype" type="hidden" value="<%=custcardtype%>">

							<% if intShowBtn=1 AND pcv_strDeactivateStatus=0 then %>
								<tr>
									<td>&nbsp;</td>
									<td><input type="submit" name="Submit2" value="Update Payment Information" class="submit2"><% If intShowPurgeBtn = 1 Then %>
									<input type="button" name="PurgeCCNumber" value="Purge Credit Card Number" onclick="location='gwCCPurgeSubmit.asp?POID=<%=qry_ID%>';">
								<% End If %></td>
								</tr>

							<% end if

							'====================================
							'Check for offline payment Order
							'====================================
							query="SELECT idPayment, AccNum FROM offlinepayments WHERE idOrder=" & qry_ID & ";"
							Set rs=Server.CreateObject("ADODB.Recordset")
							Set rs=connTemp.execute(query)
							if NOT rs.eof then
								pidPayment=rs("idPayment")
								pAccNum=rs("AccNum")
								set rs=nothing
								query="SELECT CReq,Cprompt FROM Paytypes WHERE idPayment="& pidPayment
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)
								If rs.eof then
									set rs=nothing
									tempCReq="0"
								else
									tempCReq=rs("CReq")
									tempCprompt=rs("Cprompt")
									set rs=nothing
								end if %>
								<tr>
									<td colspan="2">Terms: <%=PaymentType%></td>
								</tr>
								<% if tempCReq="-1" then %>
								<tr>
									<td><%=tempCprompt%>:&nbsp;<%=pAccNum%></td>
									<td>&nbsp;</td>
								</tr>
								<% end if %>
							<% end if %>

							<% if PayCharge>0 then %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td colspan="2">Additional Fee for Payment Type: <%=money(PayCharge)%></td>
								</tr>
							<% end if %>

					</table>
				</div>

				<%
				'--------------
				' END TAB 5
				'--------------

				'--------------
				' START TAB 6
				'--------------
				%>

				<div id="tab6" class="TabbedPanes" style="<%=pcTab6Style%>">
					<table class="pcCPcontent">
						<tr>
							<td align="left" valign="top" width="50%">

								<table class="pcCPcontent">
									<tr>
										<th colspan="2">Billing Information</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<% ' Calculate customer number using sccustpre constant
									Dim pcCustomerNumber
									pcCustomerNumber = (sccustpre + int(pidcustomer))
									%>
									<tr>
										<td colspan="2"><b>Customer ID: <%=pcCustomerNumber%></b> - <a href="modCusta.asp?idcustomer=<%=pidcustomer%>" class="pcSmallText">View/Edit Customer</a></td>
									</tr>
									<tr>
										<td colspan="2">
											<%
											if pcv_strCustomerIP="" then
												pcv_strCustomerIP="Not Available"
											end if
											%>
											<strong>Customer IP: <%=pcv_strCustomerIP %></strong>
										</td>
									</tr>
									<tr>
										<td width="20%" nowrap="nowrap"><p>First Name:</p></td>
										<td width="80%"><p><input type="text" name="name" size="25" value="<%=pname%>"></p></td>
									</tr>
									<tr>
										<td nowrap="nowrap"><p>Last Name:</p></td>
										<td><p><input type="text" name="lastName" size="25" value="<%=plastName%>"></p></td>
									</tr>
									<tr>
										<td><p>Company:</p></td>
										<td><p><input type="text" name="customerCompany" size="25" value="<%=pcustomerCompany%>"></p></td>
									</tr>
									<%
									'///////////////////////////////////////////////////////////
									'// START: COUNTRY AND STATE/ PROVINCE CONFIG
									'///////////////////////////////////////////////////////////
									'
									' 1) Place this section ABOVE the Country field
									' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
									' 3) Additional Required Info

									'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
									pcv_isStateCodeRequired = pcv_isShipStateCodeRequired '// determines if validation is performed (true or false)
									pcv_isProvinceCodeRequired = pcv_isShipProvinceCodeRequired '// determines if validation is performed (true or false)
									pcv_isCountryCodeRequired = pcv_isShipCountryCodeRequired '// determines if validation is performed (true or false)

									'// #3 Additional Required Info
									pcv_strTargetForm = "form2" '// Name of Form
									pcv_strCountryBox = "CountryCode" '// Name of Country Dropdown
									pcv_strTargetBox = "stateCode" '// Name of State Dropdown
									pcv_strProvinceBox =  "state" '// Name of Province Field

									'// Set local Country to Session
									if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
										Session(pcv_strSessionPrefix&pcv_strCountryBox) = pCountryCode
									end if

									'// Set local State to Session
									if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
										Session(pcv_strSessionPrefix&pcv_strTargetBox) = pstateCode
									end if

									'// Set local Province to Session
									if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
										Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pstate
									end if
									on error resume next%>
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
										<td><p>Address:</p></td>
										<td><p><input type="text" name="address" value="<%=pAddress%>" size="25"></p></td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td><p><input type="text" name="address2" value="<%=pAddress2%>" size="25"></p></td>
									</tr>
									<tr>
										<td><p>City:</p></td>
										<td><p><input type="text" name="city" value="<%=pcity%>" size="25"></p></td>
									</tr>
									<%
									'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
									pcs_StateProvince
									%>
									<tr>
										<td><p>Postal Code:</p></td>
										<td><p><input type="text" name="zip" value="<%=pzip%>" size="5"></p></td>
									</tr>

									<tr>
										<td><p>Telephone:</p></td>
										<td><p><input type="text" name="phone" value="<%=pphone%>" size="20"></p></td>
									</tr>
									<tr>
										<td><p>E-mail:</p></td>
										<td><p><input name="email" type="text" value="<%=pemail%>" size="25"></p></td>
									</tr>
									<%
									'Start Special Customer Fields
									session("cp_nc_custfields")=""
									session("cp_nc_custfieldsExists")=""
									query="SELECT pcCField_ID,pcCField_Name,pcCField_FieldType,pcCField_Value,pcCField_Length,pcCField_Maximum,pcCField_Required,pcCField_PricingCategories,pcCField_ShowOnReg,pcCField_ShowOnCheckout,'' FROM pcCustomerFields ORDER BY pcCField_Order ASC, pcCField_Name ASC;"
									set rsQ=connTemp.execute(query)
									if not rsQ.eof then
										session("cp_nc_custfields")=rsQ.GetRows()
										session("cp_nc_custfieldsExists")="YES"
									end if
									set rsQ=nothing

									if session("cp_nc_custfieldsExists")="YES" then
										pcArr=session("cp_nc_custfields")
										For k=0 to ubound(pcArr,2)
											pcArr(10,k)=""
											query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & pidcustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
											set rsQ=connTemp.execute(query)
											if not rsQ.eof then
												pcArr(10,k)=rsQ("pcCFV_Value")
											end if
											set rsQ=nothing
										Next
										session("cp_nc_custfields")=pcArr
									end if
									'End of Special Customer Fields
									'Start Special Customer Fields
									if session("cp_nc_custfieldsExists")="YES" then
										pcArr=session("cp_nc_custfields")
										For k=0 to ubound(pcArr,2)%>
											<tr>
												<td colspan="2">
													<%=pcArr(1,k)%>:&nbsp;
													<%if pcArr(2,k)="1" then%>
														<input type="checkbox" name="custfield_<%=pcArr(0,k)%>" <%if pcArr(10,k)<>"" then%>value="<%=pcArr(10,k)%>"<%else%><%if pcArr(3,k)<>"" then%>value="<%=pcArr(3,k)%>"<%else%>value="1"<%end if%><%end if%> <%if pcArr(10,k)<>"" then%>checked<%end if%> class="clearBorder">
													<%else%>
														<input type="text" name="custfield_<%=pcArr(0,k)%>" value="<%=pcArr(10,k)%>" size="<%=pcArr(4,k)%>" <%if pcArr(5,k)>"0" then%>maxlength="<%=pcArr(5,k)%>"<%end if%>>
													<%end if%>
													<%if pcArr(6,k)="1" then%>
														<img src="images/pc_required.gif" width="9" height="9">
													<%end if%>
												</td>
											</tr>
										<% Next
									end if
									'End of Special Customer Fields
									%>
									<% IDRefer=pIDRefer
									if (IDRefer<>"") and (IDRefer<>"0") then
										query="select name from Referrer where IDRefer=" & IDRefer
										set rs=server.CreateObject("ADODB.RecordSet")
										set rs=connTemp.execute(query)
										if not rs.eof then
											pRefer=rs("name")
											set rs=nothing %>
											<tr>
												<td>Referrer:</td>
												<td><%=pRefer%></td>
											</tr>
										<% end if
									end if %>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2" align="center"><input type="submit" name="Submit1" value="Update Customer Information" class="submit2" onclick="return(Form2_Validator(document.getElementById('form2')));">
										</td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								</table>

								<script language="JavaScript">
								<!--

								function Form2_Validator(theForm)
								{
								<%'Start Special Customer Fields
									if session("cp_nc_custfieldsExists")="YES" then
										pcArr=session("cp_nc_custfields")
										For k=0 to ubound(pcArr,2)
										if pcArr(6,k)="1" then%>
											if (theForm.custfield_<%=pcArr(0,k)%>.value == "")
												{
												<%if pcArr(0,k)="1" then%>
													alert("Please select the option.");
												<%else%>
													alert("Please enter a value for this field.");
												<%end if%>
													theForm.custfield_<%=pcArr(0,k)%>.focus();
													return (false);
											}
										<%end if
										Next
									end if
								'End of Special Customer Fields%>

								return (true);
								}
								//-->
								</script>
							</td>
							<td width="50%" align="left" valign="top">

								<table class="pcCPcontent">
									<tr>
										<th colspan="2">Shipping Information</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2">
												<% if pcOrd_ShipWeight>0 then
													intTotalWeight=pcOrd_ShipWeight
												end if
												if cdbl(intTotalWeight)<1 AND cdbl(intTotalWeight)>0 then
													intTotalWeight=1
												end if
												intTotalWeight=round(intTotalWeight,0)
												if scShipFromWeightUnit="KGS" then
												pKilos=Int(intTotalWeight/1000)
												pWeight_g=intTotalWeight-(pKilos*1000) %>
												<div align="left">Total Shipping Weight:&nbsp;&nbsp;<%=pKilos&" kg "%>
												<% if pWeight_g>0 then
													response.write pWeight_g&" g"
												end if %>
												</div>
											<% else
												pPounds=Int(intTotalWeight/16)
												pWeight_oz=intTotalWeight-(pPounds*16) %>
												<div align="left">Total Shipping Weight:&nbsp;&nbsp;<%=pPounds&" lbs. "%>
												<% if pWeight_oz>0 then
													response.write pWeight_oz&" oz."
												end if %>
												</div>
											<% end if %>
										</td>
									</tr>
									<tr>
										<td colspan="2">Number of Packages:
										<input name="ordPackageNum" type="text" id="ordPackageNum" value="<%=pOrdPackageNum%>" size="4" maxlength="4"></td>
									</tr>
									<tr>
										<td colspan="2">Shipping Method:
										<% if pSRF="1" then
											response.write pshipmentDetails
										else
											if varShip<>"0" AND Service<>"" then
												response.write Service
											else
												response.write pshipmentDetails
											end if
										end if %>
										</td>
									</tr>
									<%
									if varShip<>"0" then
										if pOrdShipType=0 then
											pDisShipType="Residential"
										else
											pDisShipType="Commercial"
										end if
										%>
										<tr>
											<td colspan="2">Shipping Type:&nbsp;<%=pDisShipType%>&nbsp;[<a href="#" onclick="popwin('ordChangeShipType.asp?id=<% = qry_ID %>');return false;">Change</a>]</td>
										</tr>
										<tr>
											<td colspan="2" class="pcCPspacer"></td>
										</tr>
										<tr>
											<th colspan="2">Shipping Address</th>
										</tr>
										<tr>
											<td colspan="2" class="pcCPspacer"></td>
										</tr>
										<% if pshippingAddress<>"" then %>
											<tr>
												<td width="20%"><p>Recipient:</p></td>
												<td width="80%"><p><input type="text" name="shippingFullName" value="<%=pshippingFullName%>" size="20"></p>
												</td>
											</tr>
											<tr>
												<td><p>Company:</p></td>
												<td><p><input type="text" name="shippingCompany" value="<%=pshippingCompany%>" size="20"></p></td>
											</tr>
										<% end if %>


										<%
										'///////////////////////////////////////////////////////////
										'// START: COUNTRY AND STATE/ PROVINCE CONFIG
										'///////////////////////////////////////////////////////////
										'
										' 1) Place this section ABOVE the Country field
										' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
										' 3) Additional Required Info

										'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
										pcv_isStateCodeRequired = pcv_isShipStateCodeRequired '// determines if validation is performed (true or false)
										pcv_isProvinceCodeRequired = pcv_isShipProvinceCodeRequired '// determines if validation is performed (true or false)
										pcv_isCountryCodeRequired = pcv_isShipCountryCodeRequired '// determines if validation is performed (true or false)

										'// #3 Additional Required Info
										pcv_strTargetForm = "form2" '// Name of Form
										pcv_strCountryBox = "shippingCountryCode" '// Name of Country Dropdown
										pcv_strTargetBox = "shippingStatecode" '// Name of State Dropdown
										pcv_strProvinceBox =  "ShippingState" '// Name of Province Field

										'// Set local Country to Session
										if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
											Session(pcv_strSessionPrefix&pcv_strCountryBox) = pshippingCountryCode
										end if

										'// Set local State to Session
										if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
											Session(pcv_strSessionPrefix&pcv_strTargetBox) = pOrdshippingStateCode
										end if

										'// Set local Province to Session
										if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
											Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pshippingState
										end if

										'// Declare the instance number if greater than 1
										pcv_strFormInstance = "2"
										'///////////////////////////////////////////////////////////
										'// END: COUNTRY AND STATE/ PROVINCE CONFIG
										'///////////////////////////////////////////////////////////
										%>

										<%
										'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
										pcs_CountryDropdown
										%>

										<tr>
											<td><p>Address:</p></td>
											<td>
												<p><% if pshippingAddress="" then %>
													<input type="text" name="shippingAddress" value="<% response.write "(Same as billing address)" %>" size="25">
												<% else %>
													<input type="text" name="shippingAddress" value="<%=pshippingAddress%>" size="25">
												<% end if %></p>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td>
												<p><% if pshippingAddress="" then %>
													<input type="text" name="shippingAddress2" size="20">
												<% else %>
													<input type="text" name="shippingAddress2" value="<%=pshippingAddress2%>" size="20">
												<% end if %></p>
											</td>
										</tr>
										<tr>
											<td><p>City:</p></td>
											<td>
												<p><% if pshippingAddress="" then %>
													<input type="text" name="shippingcity" size="20">
												<% else %>
													<input type="text" name="shippingcity" value="<%=pshippingcity%>" size="20">
												<% end if %></p>
											</td>
										</tr>

										<%
										'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
										pcs_StateProvince
										%>

										<tr>
											<td><p>Postal Code:</p></td>
											<td><p>
												<% if pshippingAddress="" then %>
													<input type="text" name="shippingZip" size="5">
												<% else %>
													<input type="text" name="shippingZip" value="<%=pshippingZip%>" size="5">
												<% end if %></p>
											</td>
										</tr>

										<tr>
											<td><p>Email:</p></td>
											<td><p>
												<% if pshippingEmail="" then %> <input name="shippingEmail" type="text" size="20">
												<% else %> <input name="shippingEmail" type="text" value="<%=pshippingEmail%>" size="20">
												<% end if %></p>
											</td>
										</tr>

										<tr>
											<td><p>Telephone:</p></td>
											<td><p>
												<% if pshippingPhone="" then %> <input name="shippingPhone" type="text" size="20" maxlength="20">
												<% else %> <input name="shippingPhone" type="text" value="<%=pshippingPhone%>" size="20" maxlength="20">
												<% end if %></p>
											</td>
										</tr>

										<tr>
											<td colspan="2">&nbsp;</td>
										</tr>
										<tr>
											<td colspan="2" align="center">
												<input type="submit" name="Submit3" value="Update Shipping Information" class="submit2">
											</td>
										</tr>
									<% end if %>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
					</div>

				<%
				'--------------
				' END TAB 6
				'--------------

				'--------------
				' START TAB 7
				'--------------

				Dim intShowAddDetails
				intShowAddDetails=0
				%>

					<div id="tab7" class="TabbedPanes" style="<%=pcTab7Style%>">
						<table class="pcCPcontent">

							<%
							' START Reward Points

								If RewardsActive <> 0 And piRewardPoints > 0 Then ' The customer used Reward Points
									intShowAddDetails=1
									iDollarValue = piRewardPoints * (RewardsPercent / 100)
									%>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<th colspan="2"><%=RewardsLabel%></th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2">The customer used <b><%=piRewardPoints & " " & RewardsLabel%></b> on this purchase for a dollar value of <%=scCurSign&money(iDollarValue)%>.</td>
									</tr>
							<%
								End If

								If RewardsActive <> 0 And piRewardPointsCustAccrued > 0 Then ' The customer accrued Reward Points
									intShowAddDetails=1
							%>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<th colspan="2"><%=RewardsLabel%></th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2">The customer accrued <b><%=piRewardPointsCustAccrued & " " & RewardsLabel%></b> on this purchase.</td>
									</tr>
							<%
								End If

								' END Reward Points

							'if discount was present, show discount type here
							if discountType<>"" then
								intShowAddDetails=1
							%>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Discounts</th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td colspan="2">One or more discount codes (electronic coupons) were used for this order.</td>
								</tr>
								<tr>
									<td colspan="2">Discount Name(s): <b><%=discountType%></b></td>
								</tr>
							<% end if %>

							<% 'if category-driven quantity discounts were used, show them here
							if pcv_CatDiscounts<>"0" then
								intShowAddDetails=1
							%>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
								<tr>
									<td colspan="2">Category-based quantity discounts were applied to this order is the amount of: <%=money(pcv_CatDiscounts)%></td>
								</tr>
							<%
							end if

							'START Affiliate information, if any

							If pidaffiliate>"1" then
								intShowAddDetails=1

								Dim paffiliateName, paffiliateCommission, paffiliateCompany
								query="SELECT affiliateName, commission, affiliateCompany FROM affiliates WHERE idAffiliate =" & pidAffiliate
								Set rs=Server.CreateObject("ADODB.Recordset")
								Set rs=connTemp.execute(query)

								paffiliateName = rs("affiliateName")
								paffiliateCommission = rs("commission")
								paffiliatePayActual =  ((pTotalAdj * paffiliatePay)/pTotal)     ''''pTotalAdj * (paffiliateCommission/100)
								paffiliateCompany = rs("affiliateCompany")

								' order cancelled/returned then set paffiliatePayActual = 0
								if( porderStatus=5 OR porderStatus=6 ) then
									paffiliatePayActual = 0
								end if

								Set rs = nothing
								%>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Affiliate Information</th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td>Affiliate Name:</td>
									<td>
										<b>
										<% Response.write paffiliateName & " (ID: " & pidaffiliate & ")"
										if paffiliateCompany <> "" then
											response.Write " - " & paffiliateCompany
										end if %>
										</b>
										- <a href="modAffa.asp?idAffiliate=<%=pidaffiliate%>">Edit Affiliate</a> - <a href="srcOrdByDate.asp#aff">View Sales by Affiliate</a>
									</td>
								</tr>
								<tr>
									<td nowrap>Initial Commission on this order:</td>
									<td><%=money(paffiliatePay)%><input type="hidden" name="PrdSales" value="<%=PrdSales%>"></td>
								</tr>
								<tr>
									<td nowrap>Commission earned on this order:</td>
									<td><%=money(paffiliatePayActual)%></td>
								</tr>
								<tr>
									<td>Adjust new commission by: <a href="JavaScript:win('helpOnline.asp?ref=477')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
									<td>
										<input type="text" name="comm1" value="" size="5">
										<input type="radio" name="optByPercentAmount" checked value="1" class="clearBorder">Percent(%)
										<input type="radio" name="optByPercentAmount" value="2" class="clearBorder">Amount(<%=scCurSign%>)
								</tr>
								<tr>
									<td>
										<div>
											Admin Comments:
											<br />(Any comments you add here are for administrative purposes only. They are never shown to the customer.)
										</div>
									</td>
									<td>
										<div style="padding: 5px 0 5px 0;">
											<textarea name="adminCommentsAffliate" cols="70" rows="7" wrap="virtual" style="background-color:#FFFFFF;"><%=padminComments%></textarea>
										</div>
									</td>
								</tr>
								<tr>
									<td nowrap>&nbsp;</td>
									<td><input name="submit12" type="submit" value="Update earned commission" class="submit2"></td>
								</tr>
							<% end if %>

							<%
							If pcomments <> "" then
								intShowAddDetails=1
							%>
								<tr>
									<td colspan="2" class="pcCPspacer"><a name="ccomments"></a></td>
								</tr>
								<tr>
									<th colspan="2">Additional Comments</th>
								</tr>
								<tr>
									<td height="5" colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td><%=pcomments%></td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
							<% end if %>

							<%'GGG Add-on start
							IF (GCDetails<>"") then
								intShowAddDetails=1
							%>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Gift Certificates</th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td>The following Gift Certificate(s) were used for this order:</td>
								</tr>
								<%
								GCArry=split(GCDetails,"|g|")
								intArryCnt=ubound(GCArry)

								for k=0 to intArryCnt

								if GCArry(k)<>"" then
									GCInfo = split(GCArry(k),"|s|")
									if GCInfo(2)="" OR IsNull(GCInfo(2)) then
										GCInfo(2)=0
									end if
									pGiftCode=GCInfo(0)
									pGiftUsed=GCInfo(2)
								query="SELECT products.IDProduct,products.Description FROM pcGCOrdered,Products WHERE products.idproduct=pcGCOrdered.pcGO_idproduct AND pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
								SET rs=server.CreateObject("ADODB.RecordSet")
								SET rs=connTemp.execute(query)

								if NOT rs.eof then
									pIdproduct=rs("idproduct")
									pName=rs("Description")
									pCode=pGiftCode
									%>
									<tr>
										<td nowrap><b>Gift Certificate Product Name:</b></td>
										<td><b><%=pName%></b></td>
									</tr>
									<tr>
										<td nowrap valign="top">&nbsp;</td>
										<td valign="top">
										<% query="SELECT pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status FROM pcGCOrdered WHERE pcGO_GcCode='" & pGiftCode & "'"
										SET rsGCObj=server.CreateObject("ADODB.RecordSet")
										SET rsGCObj=connTemp.execute(query)

										if NOT rsGCObj.eof then
											pcGO_GcCode=rsGCObj("pcGO_GcCode")
											pExpDate=rsGCObj("pcGO_ExpDate")
											pGCAmount=rsGCObj("pcGO_Amount")
											pGCStatus=rsGCObj("pcGO_Status")
											%>
											Gift Certificate Code: <b><%=pcGO_GcCode%></b><br>
											Used for this order:&nbsp;<%=scCurSign & money(pGiftUsed)%><br><br>
											<% if cdbl(pGCAmount)<=0 then%>
												This Gift Certificate has been completely redeemed.
											<% else %>
												Available Amount: <b><%=scCurSign & money(pGCAmount)%></b>
												<br>
												<% if year(pExpDate)="1900" then%>
													This Gift Certificate does not expire.
												<%else
													if scDateFrmt="DD/MM/YY" then
														pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
													else
														pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
													end if %>
													Expiration Date: <font color=#ff0000><b><%=pExpDate%></b></font>
												<%end if%>
												<br>
												<% if pGCStatus="1" then%>
													Status: Active
												<%else%>
													Status: Inactive
												<%end if%>
											<%end if%>
											<tr>
												<td colspan="2" class="pcCPspacer"></td>
											</tr>
										<%end if
										set rsG=nothing
										%>
										</td>
									</tr>
								<%end if
								set rs=nothing
								end if
								Next%>
							<% END IF
							'GGG Add-on end%>

							<%
							If (pcDPs<>"") and (pcDPs="1") then
								intShowAddDetails=1
							%>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Downloadable Product(s) Information</th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<% query="select IdProduct from DPRequests WHERE IdOrder=" & pidorder &";"
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								do while not rs.eof
									pIdProduct=rs("idProduct")
									query="SELECT Description, URLExpire, ExpireDays, License, LicenseLabel1, LicenseLabel2, LicenseLabel3, LicenseLabel4, LicenseLabel5 FROM Products,DProducts WHERE products.idproduct=" & pIdProduct & " AND DProducts.idproduct=Products.idproduct AND products.downloadable=1;"
									set rstemp=server.CreateObject("ADODB.RecordSet")
									set rstemp=connTemp.execute(query)

									if not rstemp.eof then
										pName=rstemp("Description")
										pURLExpire=rstemp("URLExpire")
										pExpireDays=rstemp("ExpireDays")
										pLicense=rstemp("License")
										pLL1=rstemp("LicenseLabel1")
										pLL2=rstemp("LicenseLabel2")
										pLL3=rstemp("LicenseLabel3")
										pLL4=rstemp("LicenseLabel4")
										pLL5=rstemp("LicenseLabel5")
										set rstemp=nothing

										query="SELECT RequestSTR,StartDate FROM DPRequests WHERE idproduct=" & pIdProduct & " AND idorder=" & pidorder & " AND idcustomer=" & pidcustomer &";"
										set rstemp=server.CreateObject("ADODB.RecordSet")
										set rstemp=connTemp.execute(query)
										pdownloadStr=rstemp("RequestSTR")
										RequestSTR=pdownloadStr
										StartDate=rstemp("StartDate")
										set rstemp=nothing
										SPath1=Request.ServerVariables("PATH_INFO")
										mycount1=0
										do while mycount1<2
											if mid(SPath1,len(SPath1),1)="/" then
												mycount1=mycount1+1
											end if
											if mycount1<2 then
												SPath1=mid(SPath1,1,len(SPath1)-1)
											end if
										loop
										SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1

										if Right(SPathInfo,1)="/" then
											pdownloadStr=SPathInfo & "pc/pcdownload.asp?id=" & pdownloadStr
										else
											pdownloadStr=SPathInfo & "/pc/pcdownload.asp?id=" & pdownloadStr
										end if %>

										<tr>
											<td><b>Product Name:</b></td>
											<td><b><%=pName%></b></td>
										</tr>
										<tr>
											<td nowrap valign="top">Download URL:</td>
											<td><a href="<%=pdownloadStr%>" target="_blank"><%=pdownloadStr%></a>
												<% if (pURLExpire<>"") and (pURLExpire="1") then
													if date()-(CDate(StartDate)+pExpireDays)<0 then%>
														<br>This URL will expire in <%=(CDate(StartDate)+pExpireDays)-date()%> days<br>
													<%else
														if date()-(CDate(StartDate)+pExpireDays)=0 then%>
															<br>This URL will expire at the end of the day<br>
														<%else%>
															<br><span style="color: #FF0000"><strong>This URL expired</strong></span><br>
														<%end if
													end if%>
													<input type="button" name="resetURL" value=" Reset URL Expiration " class="ibtnGrey" onclick="location='resetURL.asp?requestSTR=<%=RequestStr%>&orderid=<%=pidorder%>';">
												<%end if%>
											</td>
										</tr>
										<% if (pLicense<>"") and (pLicense="1") then %>
											<tr>
												<td nowrap valign="top">License(s):</td>
												<td>
												<% query="SELECT Lic1, Lic2, Lic3, Lic4, Lic5 FROM DPLicenses WHERE idproduct=" & pIdProduct & " AND idorder=" & pidorder
												set rstemp=server.CreateObject("ADODB.RecordSet")
												set rstemp=connTemp.execute(query)
												do while not rstemp.eof %>
													<table width="100%" border="0" cellpadding="2" cellspacing="0">
														<%	Lic1=rstemp("Lic1")
														if Lic1<>"" then%>
															<tr><td nowarp><%=pLL1%>:</td><td><%=Lic1%></td></tr>
														<%end if
														Lic2=rstemp("Lic2")
														if Lic2<>"" then%>
															<tr><td nowarp><%=pLL2%>:</td><td><%=Lic2%></td></tr>
														<%end if
														Lic3=rstemp("Lic3")
														if Lic3<>"" then%>
															<tr><td nowarp><%=pLL3%>:</td><td><%=Lic3%></td></tr>
														<%end if
														Lic4=rstemp("Lic4")
														if Lic4<>"" then%>
															<tr><td nowarp><%=pLL4%>:</td><td><%=Lic4%></td></tr>
														<%end if
														Lic5=rstemp("Lic5")
														if Lic5<>"" then%>
															<tr><td nowarp><%=pLL5%>:</td><td><%=Lic5%></td></tr>
														<%end if%>
													</table><br>
													<%rstemp.movenext
												loop%>
												</td>
											</tr>
										<%end if
									end if
									rs.MoveNext
								loop
								set rs=nothing %>
							<% end if%>

							<%
							'***************************************
							'* START - GIFT CERTIFICATE INFO
							'***************************************
							IF (pGCs<>"") and (pGCs="1") then
								intShowAddDetails=1
							%>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Gift Certificate(s) Information</th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<%
								query="SELECT idproduct FROM ProductsOrdered WHERE idOrder="& pidorder &";"
								SET rsGCObj=server.CreateObject("ADODB.RecordSet")
								SET rsGCObj=connTemp.execute(query)
								do while not rsGCObj.eof
									pcv_tempGCProductID=rsGCObj("idProduct")
									query="SELECT products.Description,pcGCOrdered.pcGO_GcCode FROM Products,pcGCOrdered WHERE products.idproduct=" & pcv_tempGCProductID & " AND pcGCOrdered.pcGO_idproduct=Products.idproduct AND products.pcprod_GC=1 AND pcGCOrdered.pcGO_idOrder="& qry_ID
									SET rs=server.CreateObject("ADODB.RecordSet")
									set rs=connTemp.execute(query)

									if not rs.eof then ' There are "processed" Gift Certificates, show the details
										pName=rs("Description")
										pCode=rs("pcGO_GcCode")
										%>
										<tr>
											<td colspan="2" >Gift Certificate Product Name: <b><%=pName%></b></td>
										</tr>
										<tr>
											<td colspan="2">
											<%
											query="SELECT pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status FROM pcGCOrdered WHERE pcGO_idproduct=" & pcv_tempGCProductID & " and pcGO_idorder="& pidorder &";"
											SET rsGCodeObj=server.CreateObject("ADODB.RecordSet")
											set rsGCodeObj=connTemp.execute(query)
											do while not rsGCodeObj.eof
												pGCCode=rsGCodeObj("pcGO_GcCode")
												pExpDate=rsGCodeObj("pcGO_ExpDate")
												pGCAmount=rsGCodeObj("pcGO_Amount")
												pGCStatus=rsGCodeObj("pcGO_Status")
											%>
												<p>
												Gift Certificate Code: <b><%=pGCCode%></b><br>
												<% if year(pExpDate)="1900" then%>
													This Gift Certificate does not expire.
												<% else
													if scDateFrmt="DD/MM/YY" then
														pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
													else
														pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
													end if %>
													Expiration Date: <font color=#ff0000><b><%=pExpDate%></b></font>
												<% end if %>
												<br>
												<% if cdbl(pGCAmount)<=0 then %>
													This Gift Certificate has been completely redeemed.
												<% else %>
													Available Amount: <b><%=scCurSign & money(pGCAmount)%></b>
												<% end if %><br>
												<% if pGCStatus="1" then %>
													Status: Active
												<% else %>
													Status: Inactive
												<% end if %>
												</p>
									<%
											rsGCodeObj.movenext
											loop
											set rsGCodeObj=nothing
									%>
											</td>
										</tr>
										<tr>
											<td colspan="2" class="pcCPspacer"></td>
										</tr>
									<%
									end if
									set rs=nothing
									rsGCObj.MoveNext
								loop
								%>
								<%
								query="SELECT pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg FROM Orders WHERE idOrder="& pidorder &" AND pcOrd_GcReEmail<>'';"
								SET rsGCObj=server.CreateObject("ADODB.RecordSet")
								SET rsGCObj=connTemp.execute(query)
								Gc_ReName=""
								Gc_ReEmail=""
								Gc_ReMsg=""
								if not rsGCObj.eof then
								Gc_ReName=rsGCObj("pcOrd_GcReName")
								Gc_ReEmail=rsGCObj("pcOrd_GcReEmail")
								Gc_ReMsg=rsGCObj("pcOrd_GcReMsg")
								end if
								set rsGCObj=nothing
								%>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Recipient Information</th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td><b>Recipient Name</b>:</td>
									<td><input type="text" name="GC_RecName" size="30" value="<%=Gc_ReName%>"></b></td>
								</tr>
								<tr>
									<td><b>Email</b>:</td>
									<td><input type="text" name="GC_RecEmail" size="30" value="<%=Gc_ReEmail%>"></b></td>
								</tr>
								<tr>
									<td valign="top"><strong><%if (Gc_ReName<>"") AND (Gc_ReEmail<>"") then%>New<%end if%> Message</strong> (optional): <br /><br /><%if (Gc_ReName<>"") AND (Gc_ReEmail<>"") then%>The original message is not shown for privacy reasons<%end if%></td>
									<td valign="top"><textarea name="GC_RecMsg" cols="60" rows="5" wrap="VIRTUAL"></textarea></td>
								</tr>
								<tr>
									<td valign="top">&nbsp;</td>
									<td valign="top">
					<%if Gc_ReMsg<>"" then%><input type="submit" name="submitReSendGCRecA" value="Resend with original message" class="submit2">&nbsp;<%end if%>
					<input type="submit" name="submitReSendGCRec" value="<%if (Gc_ReName<>"") AND (Gc_ReEmail<>"") then%>Resend with new message<%else%>Send message<%end if%>" class="submit2">
				   </td>
								</tr>
								<%
						END IF
						'***************************************
						'* END - GIFT CERTIFICATE INFO
						'***************************************

						'***************************************
						'* START - TERMS
						'***************************************
						if scTerms=1 then
							if scTermsShown = 0 then
								query="SELECT pcCustomerTermsAgreed.idCustomer, pcCustomerTermsAgreed.idOrder, pcCustomerTermsAgreed.pcCustomerTermsAgreed_InsertDate FROM pcCustomerTermsAgreed WHERE pcCustomerTermsAgreed.idCustomer=" & pidcustomer & ";"
							else
								query="SELECT pcCustomerTermsAgreed.idCustomer, pcCustomerTermsAgreed.idOrder, pcCustomerTermsAgreed.pcCustomerTermsAgreed_InsertDate FROM pcCustomerTermsAgreed WHERE (((pcCustomerTermsAgreed.idCustomer)=" & pidcustomer & ") AND ((pcCustomerTermsAgreed.idOrder)=" & qry_ID & "));"
							end if
							set rsTerms=server.CreateObject("ADODB.RecordSet")
							set rsTerms=conntemp.execute(query)

							if NOT rsTerms.eof then
								pcv_AgreedOrder = int(rsTerms("idOrder"))
								pcv_AgreedDate = rsTerms("pcCustomerTermsAgreed_InsertDate")

									if trim(pcv_AgreedOrder) = qry_ID then
										pcv_AgreedMessage = "This customer agreed to the Terms and Conditions Agreement during the checkout process prior to placing <strong>this order</strong> on <strong>" & ShowDateFrmt(pcv_AgreedDate) & "</strong>."
										else
										pcv_AgreedMessage = "The store is currently setup so that customers are required to agree to the Terms only on their first order. This customer agreed to the Terms and Conditions Agreement during the checkout process prior to placing order #<strong><a href=OrdDetails.asp?id=" & pcv_AgreedOrder & ">" & (scpre+int(pcv_AgreedOrder)) & "</a></strong> on <strong>" & ShowDateFrmt(pcv_AgreedDate) & "</strong>."
									end if

								intShowAddDetails=1 %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Terms and Conditions Agreement</th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>

								<tr>
									<td colspan="2"><%=pcv_AgreedMessage%></td>
								</tr>
							<% end if
							set rsTerms=nothing
						end if

						'***************************************
						'* END - TERMS
						'***************************************

						' NO addition information found, show message
						if intShowAddDetails=0 then %>
							<tr>
								<td colspan="2">No additional order information is available for this order.</td>
							</tr>
						<% end if %>

					</table>
				</div>

			</td>
		</tr>
	</table>
	<%call closedb() %>
</form>

<!--#include file="AdminFooter.asp"-->