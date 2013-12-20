<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Custom Data Export" %>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->  
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/javascripts/pcDateFunctions.js"-->
<% 
on error resume next
Dim connTemp,rs,query
dim xlsTest, xlsObj
xlsObj=0
xlsTest=CreateObject("Excel.Application")
if err.number<>0 then
	xlsObj=1
	err.number=0
end if


dim strDateFormat
strDateFormat="mm/dd/yyyy"
if scDateFrmt="DD/MM/YY" then
	strDateFormat="dd/mm/yyyy"
end if
counter=0%>
<script language="JavaScript">
<!--
	function Validate_Dates(theForm)
	{
	
		if (theForm.FromDate.value == "")
		{
			alert("Please enter From Date and try again.");
			theForm.FromDate.focus();
			return (false);
		}
		
		if (theForm.ToDate.value == "")
		{
			alert("Please enter To Date and try again.");
			theForm.ToDate.focus();
			return (false);
		}
		
		if (isDate(theForm.FromDate.value,"<%=strDateFormat%>","From Date")==false)
		{
			theForm.FromDate.focus()
			return false
		}
		
		if (isDate(theForm.ToDate.value,"<%=strDateFormat%>","To Date")==false)
		{
			theForm.ToDate.focus()
			return false
		}
		
		if (CompareDates(theForm.FromDate,theForm.ToDate,"From <= To")==false)
		{
			alert("From Date should be less than To Date.")
			theForm.ToDate.focus()
			return false
		}
	return (true);
	}
//-->
</script>
<%

Dim varMonth, varDay, varYear
varMonth=Month(Date)
varDay=Day(Date)
varYear=Year(Date) 
dim dtInputStrStart, dtInputStr
dtInputStrStart=(varMonth&"/01/"&varYear)
if scDateFrmt="DD/MM/YY" then
	dtInputStrStart=("01/"&varMonth&"/"&varYear)
end if
dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
if scDateFrmt="DD/MM/YY" then
	dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
end if

%>
<TABLE class="pcCPcontent">
<tr> 
<td>
<script>
<!--
	
function FormEXP1_Validator()
{
	if (document.EXP1.ReturnAS.value=='XLSN')
		{document.EXP1.action='exportOrder.asp';}
	else {document.EXP1.action='runqueryXportOrders.asp';}
	return (true);
}

function FormEXP2_Validator()
{
	if (document.EXP2.ReturnAS.value=='XLSN')
		{document.EXP2.action='exportCustomer.asp';}
	else {document.EXP2.action='runqueryXportCustomer.asp';}
	return (true);
}

function FormEXP3_Validator()
{
	if (document.EXP3.ReturnAS.value=='XLSN')
		{document.EXP3.action='exportAffiliate.asp';}
	else {document.EXP3.action='runqueryXportAffiliate.asp';}
	return (true);
}

function FormEXP4_Validator()
{
	if (document.EXP4.ReturnAS.value=='XLSN')
		{document.EXP4.action='exportProduct.asp';}
	else {document.EXP4.action='runqueryXportProduct.asp';}
	return (true);
}

function FormEXP5_Validator()
{
	if (document.EXP5.pcv_exportType.value=='XLSN')
		{document.EXP5.action='exportOrderedProducts.asp';}
	else {document.EXP5.action='exportProductsOrdered.asp';}
	return (true);
}
//-->
</script>
<a name="top"></a>
<div style="float: left;">
        <ul class="pcListIcon">
            <li><a href="#products">Simple product export</a> (<em>small product catalogs only</em>)</li>
            <li><a href="ReverseImport_step1.asp">Advanced product export (Reverse Import Wizard)</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=438')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></li>
            <li><a href="ReverseCatImport_step1.asp">Advanced category export (Reverse Import Wizard)</a>&nbsp;<a href="http://wiki.earlyimpact.com/productcart/export-import-categories" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information about exporting categories"></a></li>
            <li><a href="exportFroogle.asp">Create data feed for Google Shopping</a>&nbsp;<a href="http://wiki.earlyimpact.com/productcart/marketing-generate_google_base_file" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" title="More information about creating a Google Shopping data feed"></a></li>
            <li><a href="pcCashback_main.asp">Create data feed for Bing cashback</a>&nbsp;<a href="http://wiki.earlyimpact.com/productcart/marketing-generate_cashback_file" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" title="More information about creating a Microsoft Bing Cashback data feed"></a></li>
            <li><a href="pcNextTag_step1.asp">Create data feed for NextTag</a>&nbsp;<a href="http://wiki.earlyimpact.com/productcart/marketing-generate_nextag_data_feed" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" title="More information about creating a NexTag product data feed"></a></li>
            <li><a href="pcYahoo_step1.asp">Create data feed for Yahoo! Search Marketing</a>&nbsp;<a href="http://wiki.earlyimpact.com/productcart/marketing-generate_yahoo_data_feed" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" title="More information about Yahoo! Search Marketing product data feeds"></a></li>
        </ul>
</div>
<div style="float: right; padding-right: 50px;">
	<ul class="pcListIcon">
        <li><a href="#orders">Export orders</a></li>
        <li><a href="#productsordered">Export ordered products information</a></li>
        <li><a href="qb_home.asp">Export orders to QuickBooks</a></li>
        <li><a href="#customers">Export customer information</a></li>
        <li><a href="#affiliates">Export affiliate information</a></li>
		<li><a href="#fedex">Export orders to FedEx</a></li>
        <li>Export addresses to <a href="#ups">UPS</a> and <a href="#usps">USPS</a></li>
      </ul>
</div>
<div style="clear: both;">&nbsp;</div>
<FORM name="EXP1" ACTION="runqueryXportOrders.asp" METHOD="POST" onSubmit="return (Validate_Dates(this) && FormEXP1_Validator());" class="pcForms">
<table class="pcCPcontent">
<tr> 
	<th colspan="2"><a name="orders"></a>Export Order Information</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_idOrder" value="1" checked>
</td>
<td width="94%">Order ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_ord_OrderName" value="1">
</td>
<td width="94%">Order Name</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_orderDate" value="1">
</td>
<td width="94%">Order Date</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_idCustomer" value="1">
</td>
<td width="94%">Customer ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_CustomerDetails" value="1">
</td>
<td width="94%">Customer Details</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_details" value="1">
</td>
<td width="94%">Order Details</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_total" value="1">
</td>
<td width="94%">Order Total</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_processDate" value="1">
</td>
<td width="94%">Processed Date</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_ShippingFullName" value="1">
</td>
<td width="94%">Shipping Name</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingCompany" value="1">
</td>
<td width="94%">Shipping Company</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingAddress" value="1">
</td>
<td width="94%">Shipping Address</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingAddress2" value="1">
</td>
<td width="94%">Shipping Address 2</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingCity" value="1">
</td>
<td width="94%">Shipping City</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingStateCode" value="1">
</td>
<td width="94%">Shipping State</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingState" value="1">
</td>
<td width="94%">Shipping Province</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingCountryCode" value="1">
</td>
<td width="94%">Shipping Country</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingZip" value="1">
</td>
<td width="94%">Shipping Zip</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shippingPhone" value="1">
</td>
<td width="94%">Shipping Phone</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shipmentDetails" value="1">
</td>
<td width="94%">Shipment Details</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_ordShiptype" value="1">
</td>
<td width="94%">Shipping Type</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_ordPackageNum" value="1">
</td>
<td width="94%">Number of packages</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shipDate" value="1">
</td>
<td width="94%">Shipping Date</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_shipVia" value="1">
</td>
<td width="94%">Shipped Via</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_trackingNum" value="1">
</td>
<td width="94%">Tracking Number</td>
</tr></font></p>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_ord_DeliveryDate" value="1">
</td>
<td width="94%">Delivery Date</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_orderStatus" value="1">
</td>
<td width="94%">Order Status</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_paymentDetails" value="1">
</td>
<td width="94%">Payment Details</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_idAffiliate" value="1">
</td>
<td width="94%">Affiliate ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_AffiliateName" value="1">
</td>
<td width="94%">Affiliate Name</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_affiliatePay" value="1">
</td>
<td width="94%">Affiliate Payment</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_iRewardPoints" value="1">
</td>
<td width="94%"><%=RewardsLabel%></td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_iRewardPointsCustAccrued" value="1">
</td>
<td width="94%">Accrued <%=RewardsLabel%></td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_IDRefer" value="1">
</td>
<td width="94%">Referrer ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_ReferName" value="1">
</td>
<td width="94%">Referrer Name</td>
</tr>

<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_rmaCredit" value="1">
</td>
<td width="94%">RMA Credit</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_gwAuthCode" value="1">
</td>
<td width="94%">Authorization Code</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_gwTransId" value="1">
</td>
<td width="94%">Transaction ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_paymentCode" value="1">
</td>
<td width="94%">Payment Gateway</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_taxAmount" value="1">
</td>
<td width="94%">Tax Amount</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_taxDetails" value="1">
</td>
<td width="94%">Tax Details</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_ord_VAT" value="1">
</td>
<td width="94%">VAT</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_pcOrd_DiscountDetails" value="1">
</td>
<td width="94%">Discount Details</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_pcOrd_CatDiscounts" value="1">
</td>
<td width="94%">Category Discounts</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_pcOrd_GiftCertificates" value="1">
</td>
<td width="94%">Redeemed Gift Certificates</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_comments" value="1">
</td>
<td width="94%">Customer Comments</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_adminComments" value="1">
</td>
<td width="94%">Admin Comments</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_returnDate" value="1">
</td>
<td width="94%">Return Date</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_returnReason" value="1">
</td>
<td width="94%">Return Reason</td>
</tr>
<tr>                          
<td width="6%">
<input type="checkbox" class="clearBorder" name="chk_DSNotify" value="1">
</td>
<td width="94%">Drop-shipper Notifications
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer">
    <hr>
    <a href="javascript: checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript: uncheckAll();">Uncheck All</a>
    
    <SCRIPT LANGUAGE="JavaScript">
		function checkAll() {
			var theForm, z = 0;
			theForm = document.EXP1;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = true;
			  }
			}
		}
		 
		function uncheckAll() {
			var theForm, z = 0;
			theForm = document.EXP1;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = false;
			  }
			}
		}
	</SCRIPT>
    
    </td>
</tr>	
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>				
<tr> 
<td colspan="2">
	From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
	&nbsp;
	To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
</td>
</tr>
<tr>
	<td class="pcSmallText" colspan="2">Export information by defining a date range and an export format.<br>
	<b>Note</b>: You must enter both dates in the format <%=strDateFormat%>.<br>
	If you do not specify a date range, the report will include all orders currently in the system.
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
<td colspan="2">
<input type="checkbox" name="includeAll" value="1" class="clearBorder"> Include <span style="font-style: italic">Incomplete, Returned, </span>and <span style="font-style: italic">Cancelled </span>orders.</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td colspan="2">Export format: 
<select name="ReturnAS" onchange="javascript:if (this.value=='XLSN') {document.EXP1.action='exportOrder.asp';} else {document.EXP1.action='runqueryXportOrders.asp';};">
<option value="HTML">HTML Table</option>
<option value="CSV">CSV</option>
<% if xlsObj=0 then %>
<option value="XLS">Excel</option>
<% end if %>
<option value="XLSN">Excel (No Driver)</option>
</select>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr> 
<td colspan="2">
	<input type="Submit" name="Submit3" value="Submit" class="submit2">
    &nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
 </td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</form>
<FORM name="EXP2" ACTION="runqueryXportCustomer.asp" METHOD="POST" onSubmit="return FormEXP2_Validator()" class="pcForms">
<table class="pcCPcontent">
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<th colspan="2"><a name="customers"></a>Export Customer Information</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="idcustomer" value="1">
</td>
<td width="94%">Customer ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="name" value="1">
</td>
<td width="94%">First Name</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="lastname" value="1">
</td>
<td width="94%">Last Name</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="customerCompany" value="1">
</td>
<td width="94%">Company</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="phone" value="1">
</td>
<td width="94%">Phone</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="email" value="1">
</td>
<td width="94%">E-mail</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="address" value="1">
</td>
<td width="94%">Address</td>
</tr>
								
<tr>
									
<td>
<input type="checkbox" class="clearBorder" name="address2" value="1">
</td>
<td>Address 2</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="city" value="1">
</td>
<td width="94%">City</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="stateCode" value="1">
</td>
<td width="94%">State</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="zip" value="1">
</td>
<td width="94%">Zip</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="CountryCode" value="1">
</td>
<td width="94%">Country</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="customerType" value="1">
</td>
<td width="94%">Customer Type</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="pcrp_accrued" value="1">
</td>
<td width="94%">
<% =RewardsLabel%>&nbsp;Accrued</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="pcrp_used" value="1">
</td>
<td width="94%"><%=RewardsLabel%>&nbsp;Used</td>
</tr>
<tr>                          
<td width="6%">
<input type="checkbox" class="clearBorder" name="pcrp_available" value="1">
</td>
<td width="94%">Available <% =RewardsLabel%>
</td>
</tr>
<%
call opendb()
fieldCount=0
query="SELECT pcCField_ID,pcCField_Name FROM pcCustomerFields;"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArr=rs.GetRows()
	intCount=ubound(pcArr,2)
	set rs=nothing
	For i=0 to intCount%>
	<tr>
		<td>
			<input type="checkbox" class="clearBorder" name="pccust_cf<%=i+1%>" value="<%=pcArr(0,i)%>">
		</td>
		<td><%=pcArr(1,i)%></td>
	</tr>
	<%Next
	fieldCount=intCount+1
end if
set rs=nothing
call closedb()
'MailUp-S
call opendb()
	pcIncMailUp=0
	query="SELECT pcMailUpSett_RegSuccess FROM pcMailUpSettings WHERE pcMailUpSett_RegSuccess=1;"
	set rs=connTemp.execute(query)
		if err.number<>0 then
			set rs = nothing
			call closedb()
			response.Redirect("upddb_MailUp.asp")
		end if
	if not rs.eof then
		pcIncMailUp=1
	end if
	set rs=nothing
	call closedb()
%>
<tr>
	<td>
		<input type="checkbox" class="clearBorder" name="pccust_recvnews" value="1">
		<input type="hidden" name="fieldCount" value="<%=fieldCount%>">
	</td>
	<td><%IF pcIncMailUp=1 THEN%>MailUp Opted-in Lists<%ELSE%>Newsletter subscriber<%END IF%></td>
</tr>
<%'MailUp-E%>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="pccust_IDRefer" value="1">
</td>
<td width="94%">Referrer ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="pccust_ReferName" value="1">
</td>
<td width="94%">Referrer Name</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer">
    <hr>
    <a href="javascript: checkAll2();">Check All</a>&nbsp;|&nbsp;<a href="javascript: uncheckAll2();">Uncheck All</a>
    
    <SCRIPT LANGUAGE="JavaScript">
		function checkAll2() {
			var theForm, z = 0;
			theForm = document.EXP2;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = true;
			  }
			}
		}
		 
		function uncheckAll2() {
			var theForm, z = 0;
			theForm = document.EXP2;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = false;
			  }
			}
		}
	</SCRIPT>
    
    </td>
</tr>
<tr> 
<td colspan="2">Export format: 
<select name="ReturnAS" onchange="javascript:if (this.value=='XLSN') {document.EXP2.action='exportCustomer.asp';} else {document.EXP2.action='runqueryXportCustomer.asp';};">
<option value="HTML">HTML Table</option>
<option value="CSV">CSV</option>
<% if xlsObj=0 then %>
<option value="XLS">Excel</option>
<% end if %>
<option value="XLSN">Excel (No Driver)</option>
</select>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr> 
<td colspan="2">
<input type="Submit" name="Submit2" value="Submit" class="submit2">
&nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</FORM>
<FORM name="EXP3" ACTION="runqueryXportAffiliate.asp" METHOD="POST" onSubmit="return FormEXP3_Validator()" class="pcForms">
<table class="pcCPcontent">
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<th colspan="2"><a name="affiliates"></a>Export Affiliates Information</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>						
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="idaffiliate" value="1">
</td>
<td width="94%">Affiliate ID</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="affiliateName" value="1">
</td>
<td width="94%">Name</td>
</tr>
<tr> 
<td width="6%">
<input type="checkbox" class="clearBorder" name="affiliateEmail" value="1">
</td>
<td width="94%">E-mail</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="commission" value="1">
</td>
<td width="94%">Commission</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td colspan="2">Export format: 
<select name="ReturnAS" onchange="javascript:if (this.value=='XLSN') {document.EXP3.action='exportAffiliate.asp';} else {document.EXP3.action='runqueryXportAffiliate.asp';};">
<option value="HTML">HTML Table</option>
<option value="CSV">CSV</option>
<% if xlsObj=0 then %>
<option value="XLS">Excel</option>
<% end if %>
<option value="XLSN">Excel (No Driver)</option>
</select>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr> 
<td colspan="2">
<input type="Submit" name="Submit5" value="Submit" class="submit2">
&nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</form>
<FORM name="EXP4" ACTION="runqueryXportProduct.asp" METHOD="POST" onSubmit="return FormEXP4_Validator()" class="pcForms">
<table class="pcCPcontent">
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<th colspan="2"><a name="products"></a>Export Product Information</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>	
<tr> 
	<td colspan="2">Only use this feature with <strong>small catalogs</strong> (1,000 products or less). With larger product catalogs (and to export more fields), use the <a href="ReverseImport_step1.asp">Advanced Product Export</a> feature (also referred to as <em>Reverse Import Wizard</em> because it prepares data for a quick re-import).</td>
</tr>	
<tr> 
	<td colspan="2"><hr></td>
</tr>							
<tr> 						
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpid" value="1"></td>
<td width="94%">Product ID</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpsku" value="1"></td>
<td width="94%">Product SKU</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpname" value="1"></td>
<td width="94%">Product Name</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpsdesc" value="1"></td>
<td width="94%">Short Description</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpdesc" value="1"></td>
<td width="94%">Description</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpprice" value="1"></td>
<td width="94%">Online Price</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fplprice" value="1"></td>
<td width="94%">List Price</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpwprice" value="1"></td>
<td width="94%">Wholesale Price</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fptype" value="1"></td>
<td width="94%">Product Type</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpweight" value="1"></td>
<td width="94%">Weight</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpstock" value="1"></td>
<td width="94%">Stock</td>
</tr>

<tr>
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpLIN" value="1"></td>
<td width="94%">Low Inventory Notification Level</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpimg" value="1"></td>
<td width="94%">General Image URL</td>
</tr>
								
<tr> 
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fptimg" value="1"></td>
<td width="94%">Thumbnail Image URL</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpdimg" value="1"></td>
<td width="94%">Detail Image URL</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpbrand" value="1"></td>
<td width="94%">Brand ID</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpactive" value="1"></td>
<td width="94%">Active</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpsavings" value="1"></td>
<td width="94%">Show savings</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpspecial" value="1"></td>
<td width="94%">Special</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpnotax" value="1"></td>
<td width="94%">No-taxable</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpnoship" value="1"></td>
<td width="94%">No shipping charge</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpnosale" value="1"></td>
<td width="94%">Not for sale</td>
</tr>
								
<tr>
									
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpnosalecopy" value="1"></td>
<td width="94%">Not for sale copy</td>
</tr>
<tr>
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpoversize" value="1"></td>
<td width="94%">Oversize</td>
</tr>
<tr>
<td width="6%">
<input type="checkbox" class="clearBorder" name="catinfor" value="1"></td>
<td width="94%">Category Assignments</td>
</tr>
<td width="6%">
<input type="checkbox" class="clearBorder" name="fpRemoved" value="1"></td>
<td width="94%">Removed/Deleted Product</td>
</tr>
<td width="6%">
<input type="checkbox" class="clearBorder" name="CSearchFields" value="1"></td>
<td width="94%">Product Search Fields</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer">
    <hr>
    <a href="javascript: checkAll3();">Check All</a>&nbsp;|&nbsp;<a href="javascript: uncheckAll3();">Uncheck All</a>
    
    <SCRIPT LANGUAGE="JavaScript">
		function checkAll3() {
			var theForm, z = 0;
			theForm = document.EXP4;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = true;
			  }
			}
		}
		 
		function uncheckAll3() {
			var theForm, z = 0;
			theForm = document.EXP4;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = false;
			  }
			}
		}
	</SCRIPT>
    
    </td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
	<tr>
		<td colspan="2">Include deleted products: 
			<input type="radio" name="includeDeleted" value="0" class="clearBorder" checked="checked"> No&nbsp;&nbsp;
			<input type="radio" name="includeDeleted" value="1" class="clearBorder"> Yes&nbsp;&nbsp;
			<input type="radio" name="includeDeleted" value="2" class="clearBorder"> Only deleted products
		</td>
	</tr>
<tr> 
<td colspan="2">Export format: 
	<select name="ReturnAS" onchange="javascript:if (this.value=='XLSN') {document.EXP4.action='exportProduct.asp';} else {document.EXP4.action='runqueryXportProduct.asp';};">
		<option value="HTML">HTML Table</option>
		<option value="CSV">CSV</option>
		<% if xlsObj=0 then %>
		<option value="XLS">Excel</option>
		<% end if %>
		<option value="XLSN">Excel (No Driver)</option>
	</select>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>				
<tr> 
<td colspan="2">
	<input type="Submit" name="Submit2" value="Submit" class="submit2">
    &nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
    </td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</FORM>

<FORM ACTION="ExportToFedEx.asp?action=add" METHOD="POST" class="pcForms" onsubmit="return Validate_Dates(this)">
<table class="pcCPcontent">
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<th colspan="2"><a name="fedex"></a>Export Orders to FedEx</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
<td colspan="2">This feature allows you to export orders in a format that can be imported into the <strong>FedEx Ship Manager</strong>. <em>Shipping Date</em> is the date the orders will be shipped.</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>							
<tr> 
<td width="6%">&nbsp;</td>
<td width="94%">
	<table>			
	<tr>
	<td>Order Date Range: From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
	&nbsp; To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
	</td>
	<tr>
		<td>Status: 
		<select name="otype">
			<option value="0" selected>All</option>
			<option value="2">Pending</option>
			<option value="3">Processed</option>
			<option value="7">Partially Shipped</option>
			<option value="8">Shipping</option>
			<option value="4">Shipped</option>
			<option value="5">Canceled</option>
			<option value="9">Partially Returned</option>
			<option value="6">Returned</option>
			<option value="1">Incomplete</option>
		</select>
		</td>
	</tr>
	<tr>
		<td>
			Shipping Date: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ShipDate">
		</td>
	</tr>		
	<tr>	
	<td><font size="1"><b>Note</b>: You must enter both dates in the format <font color="#000099"><%=strDateFormat%></font>.<br>If you do not specify a date range, the report will include all orders currently in the system.</font></td>
	</tr>
	</table>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>						
<tr> 
<td width="6%">&nbsp;</td>
<td width="94%">
<input type="Submit" name="goUPS" value="Submit" class="submit2">
&nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</form>


<FORM ACTION="ExportToUPS.asp?action=add" METHOD="POST" class="pcForms" onsubmit="return Validate_Dates(this)">
<table class="pcCPcontent">
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<th colspan="2"><a name="ups"></a>Export Addresses to UPS</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
<td colspan="2">This feature allows you to export customer addresses in a format that can be imported into the <u>Address Book</u> at UPS.com. When you import the file, select &quot;My UPS Address Book&quot; from the &quot;Original File Format&quot; drop-down menu.</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>							
<tr> 
<td width="6%" align="right"><input type="radio" class="clearBorder" name="exportType" value="0" checked></td>
<td width="94%">All customers</td>
</tr>							
<tr> 
<td align="right"><input type="radio" class="clearBorder" name="exportType" value="1"></td>
<td>Only customers that have placed an order</td>
</tr>
<tr> 
<td width="6%">&nbsp;</td>
<td width="94%">
	<table>			
	<tr>
	<td>Order Date Range: From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
	&nbsp; To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
	</td>
	</tr>		
	<tr>	
	<td><font size="1"><b>Note</b>: You must enter both dates in the format <font color="#000099"><%=strDateFormat%></font>.<br>If you do not specify a date range, the report will include all orders currently in the system.</font></td>
	</tr>
	</table>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>						
<tr> 
<td width="6%">&nbsp;</td>
<td width="94%">
<input type="Submit" name="goUPS" value="Submit" class="submit2">
&nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</form>

<FORM ACTION="ExportToUSPS.asp?action=add" METHOD="POST" class="pcForms" onsubmit="return Validate_Dates(this)">
<table class="pcCPcontent">
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<th colspan="2"><a name="usps"></a>Export Addresses to U.S.P.S.</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
<td colspan="2">This feature allows you to export customer addresses in a format that can be imported into the <a href="http://www.usps.com/shippingassistant/welcome.htm" target="_blank">USPS Shipping Assistant</a>.</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>								
<tr> 
<td width="6%" align="right"><input type="radio" class="clearBorder" name="exportType" value="0" checked></td>
<td width="94%">All customers</td>
</tr>						
<tr> 
<td align="right"><input type="radio" class="clearBorder" name="exportType" value="1"></td>
<td>Only customers that have placed an order</td></tr>
<tr> 
<td>&nbsp;</td>
<td width="94%">
	<table>			
	<tr>
	<td>Order Date Range: From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
	&nbsp; To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
	</td>
	</tr>		
	<tr>	
	<td><font size="1"><b>Note</b>: You must enter both dates in the format <font color="#000099"><%=strDateFormat%></font>.<br>If you do not specify a date range, the report will include all orders currently in the system.</font></td>
	</tr>
	</table>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>						
<tr> 
<td>&nbsp;</td>
<td><input type="Submit" name="goUPS" value="Submit" class="submit2">
    &nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</form>

<FORM name="EXP5" ACTION="exportProductsOrdered.asp" method="post" onSubmit="return (Validate_Dates(this) && FormEXP5_Validator())" class="pcForms" target="_blank">
<table class="pcCPcontent">
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<th colspan="2"><a name="productsordered"></a>Export Ordered Products Information</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_OrderID" value="1" checked>
	</td>
	<td width="94%">Order ID</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_OrderDate" value="1" checked>
	</td>
	<td width="94%">Order Date</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_PrdSKU" value="1" checked>
	</td>
	<td width="94%">Product SKU</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_PrdName" value="1" checked>
	</td>
	<td width="94%">Product Name</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_UnitPrice" value="1" checked>
	</td>
	<td width="94%">Unit Price</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_Units" value="1" checked>
	</td>
	<td width="94%">Units</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_WholesalePrice" value="1">
	</td>
	<td width="94%">WholeSale Price</td>
</tr>
<%if scBTO=1 then%>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_BTOConf" value="1">
	</td>
	<td width="94%">BTO Configuration</td>
</tr>
<%end if%>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_POptions" value="1">
	</td>
	<td width="94%">Product Options</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_CInputs" value="1">
	</td>
	<td width="94%">Custom Input Fields</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_QDiscounts" value="1">
	</td>
	<td width="94%">Quantity Discounts</td>
</tr>
<%if scBTO=1 then%>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_IDiscounts" value="1">
	</td>
	<td width="94%">Items Discounts</td>
</tr>
<%end if%>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_EventName" value="1">
	</td>
	<td width="94%">Event Name</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_GWOption" value="1">
	</td>
	<td width="94%">Gift Wrapping Option</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_GWPrice" value="1">
	</td>
	<td width="94%">Gift Wrapping Price</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_PackageID" value="1">
	</td>
	<td width="94%">Shipping Package ID</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_PCost" value="1">
	</td>
	<td width="94%">Product Cost</td>
</tr>
<tr> 
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_Margin" value="1">
	</td>
	<td width="94%">Margins</td>
</tr>
<tr>
	<td width="6%">
		<input type="checkbox" class="clearBorder" name="pcv_TotalPrice" value="1" checked>
	</td>
	<td width="94%">Total Price</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer">
    <hr>
    <a href="javascript: checkAll4();">Check All</a>&nbsp;|&nbsp;<a href="javascript: uncheckAll4();">Uncheck All</a>
    
    <SCRIPT LANGUAGE="JavaScript">
		function checkAll4() {
			var theForm, z = 0;
			theForm = document.EXP5;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = true;
			  }
			}
		}
		 
		function uncheckAll4() {
			var theForm, z = 0;
			theForm = document.EXP5;
			 for(z=0; z<theForm.length;z++){
			  if(theForm[z].type == 'checkbox'){
			  theForm[z].checked = false;
			  }
			}
		}
	</SCRIPT>
    
    </td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td colspan="2">
    <div style="float: right; margin-right: 10px;" class="pcSmallText">Export information by defining a date range and an export format.<br>
        <b>Note</b>: You must enter both dates in the format <strong><%=strDateFormat%></strong>.</div>
    From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">&nbsp;
	To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
</td>
</tr>
<tr>
	<td colspan="2">
    	Limit to products whose name include: <input type="text" value="" size="20" name="productKeywords">
    </td>
<tr> 
	<td colspan="2">Export format: 
		<select name="pcv_exportType" onchange="javascript:if (this.value=='XLSN') {document.EXP5.action='exportOrderedProducts.asp';} else {document.EXP5.action='exportProductsOrdered.asp';};">
			<option value="HTML">HTML Table</option>
			<option value="CSV">CSV</option>
			<option value="XLSN">Excel (No Driver)</option>
		</select>
	</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr> 
	<td colspan="2">
		<input type="Submit" name="Submit2" value="Submit" class="submit2">
	    &nbsp;<input type="button" value="Back to the Top" onClick="document.location.href='#top';">
	</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</FORM>

</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->