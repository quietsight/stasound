<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings - Manual Entry Method<br>Step 3: Tax by Product" %>
<% section="misc" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<%
dim query, conntemp, rstemp

call openDb()
sMode=Request.Form("Submit")

If sMode <> "" Then
	If sMode="Add" Then
		taxPerProduct=(Request.Form("taxPerProduct")/100)
		idproduct=Request.Form("idproduct")
		zipEq=Request.Form("zipEq")
		
		If zipEq="-1" Then
			zip=Request.Form("zip")
		Else
			zipEq="0"
		End If
		stateCodeEq=Request.Form("stateCodeEq")
		If stateCodeEq="-1" Then
			stateCode=Request.Form("stateCode")
			Else
			stateCodeEq="0"
		End If
		CountryCodeEq=Request.Form("CountryCodeEq")
		If CountryCodeEq="-1" Then
			CountryCode=Request.Form("CountryCode")
		Else
			CountryCodeEq="0"
		End If

		query="INSERT INTO taxPrd (idproduct, CountryCode, CountryCodeEq, stateCode, stateCodeEq, zip, zipEq, taxPerProduct) VALUES ("& Cint(idproduct) &", '"& CountryCode &"',"& Cint(CountryCodeEq) &", '"& stateCode &"',"& Cint(stateCodeEq) &",'"& zip &"',"& Cint(zipEq) &","& taxmoney(taxperproduct) &")"

		set rstemp=Server.CreateObject("ADODB.Recordset")     
		rstemp.Open query, conntemp
		
		if err.number <> 0 then
		  pcErrDescription = err.description
		  set rstemp=nothing
		  call closeDb()
		  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddTaxPerPrd.asp: "&pcErrDescription) 
		end If

	End If
	
	set rstemp=nothing
	call closeDb()	
	response.redirect "viewTax.asp"
End If

%>
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript" type="text/javascript" src="../includes/spry/xpath.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryData.js"></script>
<script type="text/javascript">
	var dscategories = new Spry.Data.XMLDataSet("pcSpryCategoriesXML.asp?CP=1&idRootCategory=<%=pcv_IdRootCategory%>", "categories/category",{sortOnLoad:"categoryDesc",sortOrderOnLoad:"ascending",useCache:true});
	dscategories.setColumnType("idCategory", "number");
	dscategories.setColumnType("idParentCategory", "number");
	var dsproducts = new Spry.Data.XMLDataSet("categories_productsxml.asp?mode=1&x_idCategory={dscategories::idCategory}", "/root/row",{useCache:true});
</script>
<form method="post" name="addtax" action="AddTaxPerPrd.asp" class="pcForms">
<table class="pcCPcontent">         
    <tr> 
        <td valign="top" nowrap>Select a category:</td>
        <td>
        <div spry:region="dscategories" id="categorySelector">
        	<div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Categories...</div>
            <div spry:state="ready">
            <select spry:repeatchildren="dscategories" name="categorySelect" onchange="document.addtax.idProduct.disabled = true; dscategories.setCurrentRowNumber(this.selectedIndex);">
				<option spry:if="{ds_RowNumber} == {ds_CurrentRowNumber}" value="{idCategory}" selected="selected">{categoryDesc} {pcCats_BreadCrumbs}</option>
				<option spry:if="{ds_RowNumber} != {ds_CurrentRowNumber}" value="{idCategory}">{categoryDesc} {pcCats_BreadCrumbs}</option>
			</select>
            </div>
		</div>
        </td>
    </tr>
    <tr>
      <td nowrap>Select a product:</td>
      <td>
        <div spry:region="dsproducts" id="productSelector">
        	<div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Products...</div>
            <div spry:state="ready">
            <select spry:repeatchildren="dsproducts" id="productSelect" name="idProduct">
                <option spry:if="{ds_RowNumber} == {ds_CurrentRowNumber}" value="{idProduct}" selected="selected">{description}</option>
                <option spry:if="{ds_RowNumber} != {ds_CurrentRowNumber}" value="{idProduct}">{description}</option>
            </select>
            </div>
        </div>        
      </td>
      <td></td>
    </tr>
    <tr> 
        <td nowrap>Enter a Tax Rate:</td>
        <td><input name="taxPerProduct" size="6" value="0.00"> % <span class="pcSmallText">(e.g. 5=5%)</span></td>
        <td></td>
    </tr>
    <tr>
    	<td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td align="right"><input type="checkbox" name="zipEq" value="-1"></td>
        <td colspan="2">Tax by Postal Code</td>
    </tr>
    <tr> 
        <td></td>
        <td colspan="2"><input name="zip" size="12" value="Postal Code"></td>
    </tr>
    <tr>
    	<td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td align="right"><input type="checkbox" name="stateCodeEq" value="-1"></td>
        <td colspan="2">Tax by State Or Province</td>
    </tr>
    <tr> 
        <td></td>
        <td colspan="2">                       
        <% 
		query="SELECT stateCode, stateName FROM states ORDER BY stateName"
		set rstemp=Server.CreateObject("ADODB.Recordset")     
        rstemp.Open query, conntemp
        if err.number <> 0 then
			pcErrDescription = err.description
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddTaxPerPrd 2: "&pcErrDescription) 
		end If
		%>
        <select name="stateCode" size="1">
            <option value="">State Code</option>
            <% do until rstemp.eof %>
            <option value="<%=rstemp("stateCode")%>"><%=rstemp("stateName")%></option>
            <% rstemp.moveNext
            loop %>
        </select>
        </td>
    </tr>
    <tr>
    	<td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
    <td align="right"><input type="checkbox" name="CountryCodeEq" value="-1"></td>
    <td colspan="2">Tax by Country</td>
    </tr>
    <tr> 
        <td width="25"></td>
        <td colspan="2"> 
        <% 
		query="SELECT CountryCode, CountryName FROM countries ORDER BY countryName"
		set rstemp=Server.CreateObject("ADODB.Recordset")     
        rstemp.Open query, conntemp
        if err.number <> 0 then
			pcErrDescription = err.description
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddTaxPerPrd 2: "&pcErrDescription) 
		end If
		%>
        <select name="CountryCode">
        <option value="">Country</option>
        <% do until rstemp.eof %>
        <option value="<%=rstemp("CountryCode")%>"><%=rstemp("countryName")%></option>
        <%
		rstemp.moveNext
		loop
		%>
        </select>
        </td>
    </tr>
    <tr>
    	<td colspan="3" class="pcCPspacer"><hr></td>
    </tr>
    <tr> 
        <td>&nbsp; </td>
        <td colspan="2"> 
        <input type="submit" name="Submit" value="Add" class="submit2">
        <input type="button" name="back" value="Back" onClick="javascript:history.back()" class="ibtnGrey">
        </td>
    </tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->