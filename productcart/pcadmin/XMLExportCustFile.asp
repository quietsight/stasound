<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Customer XML File" %>
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

<form name="ajaxSearch" method="post" action="XMLExportCustFileA.asp?action=newsrc" class="pcForms">
<table class="pcCPsearch">
	<tr>
		<td colspan="2">
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
			<h2>Find Customers</h2>
			Use the following filters to look for customers in your store to generate XML file
		</td>
	</tr>
<tr> 
	<td nowrap>First Name:</td>
	<td> 
	<input name="srcFirstName_value" type="text" size="30" maxlength="150"></td>
</tr>
<tr> 
	<td nowrap>Last Name:</td>
	<td> 
	<input name="srcLastName_value" type="text" size="30" maxlength="150"></td>
</tr>
<tr> 
	<td nowrap>Company:</td>
	<td> 
	<input name="srcCompany_value" type="text" size="30" maxlength="150"></td>
</tr>
<tr> 
	<td nowrap>E-mail:</td>
	<td> 
	<input name="srcEmail_value" type="text" size="30" maxlength="150"></td>
</tr>
<tr> 
	<td nowrap>City:</td>
	<td> 
	<input name="srcCity_value" type="text" size="30" maxlength="150"></td>
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
	<td nowrap>Phone:</td>
	<td><input name="srcPhone_value" type="text" size="30" maxlength="150"></td>
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
	<td nowrap>CustomerField Value:</td>
	<td> 
	<input name="srcCustomerField_value" type="text" size="30" maxlength="150"></td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td>Include locked customers:</td>
	<td><input type="checkbox" name="srcIncLocked_value" value="1" class="clearBorder"></td>
</tr>
<tr>
	<td>Include suspended customers:</td>
	<td><input type="checkbox" name="srcIncSuspended_value" value="1" class="clearBorder"></td>
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
		<b>Include/Exclude Previously Exported Customers</b>
	</td>
</tr>
<tr>
	<td colspan="2">
		Would you like to export <u><b>only</b></u> customers that <u><b>have not</b></u> previously been exported? <input type="radio" name="pHideExported" value="1" class="clearBorder" checked> Yes    <input type="radio" name="pHideExported" value="0" class="clearBorder"> No</td>
</tr>
<tr>
	<td colspan="2">
		<br>
		<u><b>Note:</b></u> Exported XML File will be saved in the folder: "<%=scPcFolder%>/xml/export" on your webserver</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<%query="SELECT pcXP_ID,pcXP_PartnerID,pcXP_Name From pcXMLPartners WHERE pcXP_Status=1 AND pcXP_FTPHost<>'' AND pcXP_FTPDirectory<>'' AND pcXP_FTPUsername<>'' AND pcXP_FTPPassword<>'';"
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
<%end if
set rs=nothing%>
<tr>
	<td colspan="2">
		<input name="runform" type="submit" value="Export Customer XML File" id="searchSubmit">
	</td>
</tr>  
</table>
</form>
<!--End of Search Form-->
<%call closedb()%><!--#include file="AdminFooter.asp"-->