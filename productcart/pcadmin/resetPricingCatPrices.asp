<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Reset Pricing Category Prices" %>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim rs, conntemp, query

if request("action")="upd" then
	pcv_idPCat=request("customerType")
	pcv_Type=request("C1")
	if pcv_idPCat="" or pcv_Type="" then
		response.redirect "resetPricingCatPrices.asp"
	end if
	
	call openDb()
	if pcv_Type="1" or pcv_Type="2" then
		query="DELETE FROM pcCC_Pricing WHERE idcustomerCategory=" & pcv_idPCat & ";"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	'BTO ADDON-S
	If scBTO=1 then
		if pcv_Type="2" then
			query="DELETE FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & pcv_idPCat & ";"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	end if
	'BTO ADDON-E
	call closeDb()
	%>
	<table class="pcCPcontent">
		<tr> 
			<td>
				<div class="pcCPmessageSuccess">Pricing Category Prices updated successfully!</div>
            </td>
		</tr>
	</table>	
<%end if%>	

<script>
function Form1_Validator(theForm)
{
	if (theForm.customerType.value == "")
	{
	    alert("Please select a pricing category");
	    theForm.customerType.focus();
	    return (false);
	}
	return (confirm('All prices that you have manually assigned to any products in this pricing category will be replaced by the default price for that pricing category. Are you sure that you want to proceed?'));
}
</script>
<form name="form1" action="resetPricingCatPrices.asp?action=upd" method="post" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Use the feature to reset the prices that you have manually overwritten for any products in your database to the default price for those products within the selected pricing category.</td>
	</tr>
	<tr>
		<td colspan="2">For example: assuming you have a pricing category that sets prices to be 10% off the Online Price and you edited N products to have the price for that pricing category be higher or lower than 10% of the Online Price, this feature will allow you to quickly remove those edited prices. All products will go back to having a price that is 10% off the Online Price for that pricing category.</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td width="15%" align="right" nowrap="nowrap"> Pricing Category:</td>
		<td>
			<select name="customerType">
				<option value=""></option>
				<% 'START CT ADD %>
				<% 'if there are PBP customer type categories - List them here 
				call openDb()
				query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
				SET rs=Server.CreateObject("ADODB.RecordSet")
				SET rs=conntemp.execute(query)
				if NOT rs.eof then 
					do until rs.eof 
						intIdcustomerCategory=rs("idcustomerCategory")
						strpcCC_Name=rs("pcCC_Name")
						%>
						<option value='<%=intIdcustomerCategory%>'
						<%if Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory then 
							response.write "selected"
						end if%>
						><%=strpcCC_Name%></option>
						<% rs.moveNext
					loop
				end if
				SET rs=nothing
				call closeDb()
				'END CT ADD %>
			</select>
		</td>
	</tr>
	<%'BTO ADDON-S
	If scBTO=1 then%>
	<tr> 
		<td align="right"><input type="radio" name="C1" value="1" checked class="clearBorder"></td>
		<td>Reset Prices</td>
	</tr>
	<tr> 
		<td align="right"><input type="radio" name="C1" value="2" class="clearBorder"></td>
		<td>Reset Prices and BTO configuration prices</td>
	</tr>
	<%
	Else
	%>
	<tr> 
		<td colspan="2"><input type="hidden" name="C1" value="1"></td>
	</tr>
	<%
	End if
	'BTO ADDON-E%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<input type="submit" name="submit1" value=" Update " class="submit2">
            &nbsp;
            <input type="button" name="back" value="Manage Pricing Categories" onClick="document.location.href='AdminCustomerCategory.asp'">
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td colspan="2">Technical Note: all records associated with the selected pricing category are removed from the <em>pcCC_Pricing</em> table (and <em>pcCC_BTO_Pricing</em> table if that option is selected on a Build To Order store).</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->