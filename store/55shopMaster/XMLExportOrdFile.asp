<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Order XML File" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,query,rs
call openDB()%>

<script language="javascript">
function CalPop(sInputName)
{
	window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
}
</script>

<form name="ajaxSearch" method="post" action="XMLExportOrdFileA.asp?action=newsrc" class="pcForms">

<table class="pcCPsearch">
	<tr>
		<td colspan="2">
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
			<h2>Find Orders</h2>
			Use the following filters to look for orders in your store to generate XML file
		</td>
	</tr>
	<tr> 
	<td valign="top" nowrap>Customer:</td>
	<td>
		<%
		query="SELECT DISTINCT Orders.idcustomer,Customers.Name,Customers.LastName FROM Customers INNER JOIN Orders ON Customers.idcustomer=Orders.idcustomer WHERE Orders.orderStatus>1 ORDER BY Customers.Name ASC,Customers.LastName ASC;"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		%>
		<select name="srcCustomerID_value">
		<option value=""></option>
		<%
		if not rstemp.eof then
			pcArr=rstemp.getRows()
			set rstemp=nothing
			intCount=ubound(pcArr,2)
			For i=0 to intCount%>
				<option value="<%=pcArr(0,i)%>"><%=pcArr(1,i) & " " & pcArr(2,i)%></option>
			<%Next
		end if
		set rstemp = nothing
		%>
		</select><br>
		<i>Note: Only customer(s) that placed order(s) are shown</i>
	</td>
</tr>

<tr> 
	<td nowrap>Customer Type:</td>
	<td>
		<select name="customerType"> 
			<option value=""></option>
			<option value='0'>Retail Customer</option>
			<option value='1'>Wholesale Customer</option>
			<% 'START CT ADD %>
			<% 'if there are PBP customer type categories - List them here
			query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
			SET rs=Server.CreateObject("ADODB.RecordSet")
			SET rs=conntemp.execute(query)
			if not rs.eof then
				pcArr=rs.getRows()
				set rs=nothing
				intCount=ubound(pcArr,2)
				For i=0 to intCount 
					intIdcustomerCategory=pcArr(0,i)
					strpcCC_Name=pcArr(1,i)
					%>
					<option value='CC_<%=intIdcustomerCategory%>'
					<%if Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory then 
						response.write "selected"
					end if%>
					><%=strpcCC_Name%></option>
				<%Next
			end if
			SET rs=nothing
			'END CT ADD %>
		</select>
	</td>
</tr>

<tr>
	<td>Order Status:</td>
	<td>
		<select name="srcOrderStatus_value">
			<option value=""></option>
			<option value="2">Pending</option>
			<option value="3">Processed</option>
			<option value="7">Partially Shipped</option>
			<option value="8">Shipping</option>
			<option value="4">Shipped</option>
			<option value="5">Canceled</option>
			<option value="9">Partially Returned</option>
			<option value="6">Returned</option>
		</select>
	</td>
</tr>

<tr>
	<td>Payment Status:</td>
	<td>
		<select name="srcPaymentStatus_value">
			<option value=""></option>
			<option value="0">Pending</option>
			<option value="1">Authorized</option>
			<option value="2">Paid</option>
		</select>
	</td>
</tr>

<tr>
	<td>Payment Type:</td>
	<td>
		<%
		query="SELECT DISTINCT idPayment,paymentDesc FROM payTypes ORDER BY paymentDesc ASC"
		Set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query) %>
		<select name="srcPaymentType_value">
		<option value=""></option>
		<%if not rs.eof then
			pcArr=rs.getRows()
			set rs=nothing
			intCount=ubound(pcArr,2)
			For i=0 to intCount
				intIdPayment=pcArr(0,i)
				strPaymentDesc=pcArr(1,i)
				 %>
				<option value="<%=strPaymentDesc%>"><%=strPaymentDesc%></option>
			<%Next
		end if
		Set rs=Nothing
		%>
		</select>
	</td>
</tr>

<tr> 
	<td nowrap>State Code:</td>
	<td> 
	<input name="srcStateCode_value" type="text" size="30" maxlength="150"></td>
</tr>

<tr> 
	<td nowrap>Country Code:</td>
	<td>
		<%
		query="SELECT CountryCode,countryName FROM countries ORDER BY countryName ASC"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		%>
		<select name="srcCountryCode_value">
		<option value=""></option>
		<%
		if not rstemp.eof then
			pcArr=rstemp.getRows()
			set rstemp=nothing
			intCount=ubound(pcArr,2)
			For i=0 to intCount%>
				<option value="<%=pcArr(0,i)%>"><%=pcArr(1,i)%></option>
			<%Next
		end if
		set rstemp = nothing
		%>
	</td>
</tr>

<tr>
	<td>Discount Code:</td>
	<td> 
		<%
		query="SELECT iddiscount,discountdesc,discountcode FROM discounts ORDER BY discountdesc asc"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		%>
		<select name="srcDiscountCode_value">
		<option value=""></option>
		<%if not rs.eof then
			pcArr=rs.getRows()
			set rs=nothing
			intCount=ubound(pcArr,2)
			For i=0 to intCount%>
				<option value="<%=pcArr(2,i)%>"><%=pcArr(1,i) & " (" & pcArr(2,i) & ")"%></option>
			<%Next
		end if
		set rs=nothing
		%>
		</select>
	</td>
</tr>

<tr>
	<td valign="top">Product Ordered:</td>
	<td>
		<%
		query="SELECT idproduct,description,sku FROM products WHERE removed=0 AND active=-1 AND sales>0 ORDER BY description ASC"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)

		intCount=CInt(-1)
		if not rstemp.eof then
			prdArray = rstemp.getRows()
			if ubound(prdArray,2) <> "" then
				intCount=ubound(prdArray,2)
			end if
		end if
		set rstemp = nothing		
		%>

		<select name="srcPrdOrderedID_value">
		<option value=""></option>
		<% for i=0 to intCount%>
			<option value="<%=prdArray(0,i)%>"><%=prdArray(1,i)%> (<%=prdArray(2,i)%>)</option>
		<% next %>
		</select><br>
		<i>Note: Only products that have been sold are shown</i>
	</td>
</tr>

<tr>
	<td colspan="2"><hr /></td>
</tr>

<tr>
	<td colspan="2"><b>Date Created</b></td>
</tr>
<tr>
	<td valign="top">From Date:</td>
	<td valign="top">
		<input type="text" name="pFromDate" value="" size="13"> <a href="javascript:CalPop('document.ajaxSearch.pFromDate');"><img SRC="../Calendar/icon_Cal.gif" border="0"></a>
	</td>
</tr>
<tr>
	<td valign="top">To Date:</td>
	<td valign="top">
		<input type="text" name="pToDate" value="" size="13"> <a href="javascript:CalPop('document.ajaxSearch.pToDate');"><img SRC="../Calendar/icon_Cal.gif" border="0"></a>
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td colspan="2">
		<b>Include/Exclude Previously Exported Orders</b>
	</td>
</tr>
<tr>
	<td colspan="2">
		Would you like to export <u><b>only</b></u> orders that <u><b>have not</b></u> previously been exported? <input type="radio" name="pHideExported" value="1" class="clearBorder" checked> Yes    <input type="radio" name="pHideExported" value="0" class="clearBorder"> No</td>
</tr>
<tr>
	<td colspan="2">
		<br>
		<u><b>Note:</b></u> Exported XML File will be saved in the folder: "<%=scPcFolder%>/xml/export" on your webserver</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<%
query="SELECT pcXP_ID,pcXP_PartnerID,pcXP_Name From pcXMLPartners WHERE pcXP_Status=1 AND pcXP_FTPHost<>'' AND pcXP_FTPDirectory<>'' AND pcXP_FTPUsername<>'' AND pcXP_FTPPassword<>'';"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
	%>
	<tr>
	<td colspan="2">
		<h3>Upload Exported XML File to FTP Server</h3>
		Please choose a XML Partner that you want to upload exported XML file to their FTP Server
	</td>
	</tr>
	<tr>
		<td>XML Partner:</td>
		<td>
			<select name="pFTPPartner">
				<option value="0"></option>
			<%For i=0 to intCount%>
				<option value="<%=pcArr(0,i)%>"><%=pcArr(1,i)%><%if trim(pcArr(2,i))<>"" then%>&nbsp;(<%=pcArr(2,i)%>)<%end if%></option>
			<%Next%>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right">
			<input type="checkbox" name="pRmvFile" value="1" class="clearBorder">
		</td>
		<td>
			Remove Exported XML File after uploading it to the Partner FTP Server
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
<%
end if
set rs=nothing
%>
<tr>
	<td colspan="2">
		<input name="runform" type="submit" value="Export Order XML File" id="searchSubmit">
	</td>
</tr>  
</table>
</form>
<!--End of Search Form-->
<%call closedb()%><!--#include file="AdminFooter.asp"-->