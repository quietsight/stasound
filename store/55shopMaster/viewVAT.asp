<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage VAT - View VAT Categories" %>
<% section="misc" %>
<%PmAdmin=6%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim rsTaxLoc, connTemp, strSQL, pid

pcv_intVATID=Request("VATID")

call openDb()
 
sMode=Request("action")
If sMode <> "" Then
	
	if sMode = "remove" then
		
		'// Delete the Category
		query="DELETE FROM pcVATRates WHERE pcVATRate_ID=" & pcv_intVATID & ";"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		'// Free the Products
		query="DELETE FROM pcProductsVATRates WHERE pcVATRate_ID=" & pcv_intVATID & ";"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		set rs=nothing
		
	end if
		
End If	

query="SELECT pcVATRates.pcVATRate_ID, pcVATRates.pcVATCountry_Code, pcVATRates.pcVATRate_Category, pcVATRates.pcVATRate_Rate FROM pcVATRates "
Set rsTaxPrd=Server.CreateObject("ADODB.Recordset")   
rsTaxPrd.Open query, connTemp, adOpenStatic, adLockReadOnly
%>
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="POST" action="viewVAT.asp" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="8">
				<p>You may add or remove VAT Categories, edit Vat Category Rates, and assign Products to VAT Categories.</p>
			</td>
		</tr>	
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<% If rsTaxPrd.eof Then %>
			<tr> 
				<td colspan="8">No VAT Categories have been created.</td>
			</tr>
		<% Else %>
		<tr> 
			<th align="left" nowrap>EU Member State</th>
			<th align="left" nowrap>VAT Category</th>
			<th align="left" nowrap>Rate</th>
			<th align="center" nowrap>&nbsp;</th>
		</tr>
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<% 
		Do While NOT rsTaxPrd.EOF
		
			'// Get Country
			query="SELECT pcVATCountries.pcVATCountry_State, pcVATCountries.pcVATCountry_Code "
			query=query&"FROM pcVATCountries "
			query=query&"WHERE pcVATCountries.pcVATCountry_Code='"&rsTaxPrd("pcVATCountry_Code")&"';"
			Set rs=Server.CreateObject("ADODB.Recordset")   
			rs.Open query, connTemp, adOpenStatic, adLockReadOnly
			if not rs.eof then
				pcVATCountry_State=rs("pcVATCountry_State")
			else
				pcVATCountry_State="Unassigned"
				pcHideProductOptions=1
			end if
			set rs=nothing
			%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td><%= pcVATCountry_State %></td>
				<td><%= rsTaxPrd("pcVATRate_Category") %></td>
				<td><%=rsTaxPrd("pcVATRate_Rate") %>%</td>
				<td align="right">
					<% if pcHideProductOptions<>1 then %>
					<a href="manageVATCategories.asp?VATID=<%=rsTaxPrd("pcVATRate_ID")%>" title="View products that this VAT category applies to"><img src="images/pcIconList.jpg" width="12" height="12" alt="Products"></a>
				  <% end if %>
					<a href="EditVATCategory.asp?VATID=<%=rsTaxPrd("pcVATRate_ID")%>" title="Edit this VAT category"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit"></a>
					<a href="javascript:if (confirm('You are about to remove this item from your VAT Categories. All Products will be removed from this Category as well. Are you sure you want to complete this action?')) location='viewVAT.asp?action=remove&VATID=<%=rsTaxPrd("pcVATRate_ID")%>'" title="Delete this VAT category"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Remove"></a></td>
			</tr>
	
			<%
			rsTaxPrd.MoveNext
			Loop
			set rsTaxPrd=nothing
			End If
			call closeDb()
			%>  		

		<tr>
			<td colspan="5" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="5">
				<input type="button" name="AddVATCategory" value="Add VAT Category" onclick="location='AddVATCategory.asp';" class="submit2">
                &nbsp;
				<input type="button" name="Button" value="Manage VAT Settings" onClick="location='AdminTaxSettings_VAT.asp';">
				&nbsp;
				<input type="button" name="Button" value="Back" onClick="JavaScript:history.back()" class="ibtnGrey">
			</td>
		</tr> 
	</table>
</form>
<!--#include file="AdminFooter.asp"-->