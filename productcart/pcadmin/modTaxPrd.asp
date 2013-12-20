<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings - Manual Entry Method - Edit Rate" %>
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

sMode=Request.queryString("mode")
cMode=Request.form("Submit")

If cMode="Update" Then
'update string
		idTaxPerProduct=Request.Form("idTaxPerProduct")
		idproduct=Request.Form("idproduct")
		zipEq=Request.Form("zipEq")
		taxperproduct=(Request.Form("taxperproduct")/100)
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
		query="UPDATE taxPrd SET CountryCode='"& CountryCode &"', CountryCodeEq="& Cint(CountryCodeEq) &", stateCode='"& stateCode &"', stateCodeEq="& Cint(stateCodeEq) &", zip='"& zip &"', zipEq="& Cint(zipEq) &", taxPerProduct="& taxmoney(taxperproduct) &" WHERE idTaxPerProduct="&idTaxPerProduct

		set rstemp=Server.CreateObject("ADODB.Recordset")     
		
		rstemp.Open query, conntemp
		
		if err.number <> 0 then
		  pcErrDescription = err.description
		  set rstemp=nothing
		  call closeDb()
		  response.redirect "techErr.asp?error="& Server.Urlencode("Error in modTaxPrd.asp: "&pcErrDescription) 
		end If

		set rstemp = nothing
		call closeDb()
		response.redirect "viewTax.asp"
End If

If sMode <> "" Then

	If sMode="DEL" Then
		idTaxPerProduct=Request.QueryString("idTaxPerProduct")
		query="DELETE FROM taxPrd WHERE idTaxPerProduct="&idTaxPerProduct
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		rstemp.Open query, conntemp
		
		if err.number <> 0 then
		  pcErrDescription = err.description
		  set rstemp=nothing
		  call closeDb()
		  response.redirect "techErr.asp?error="& Server.Urlencode("Error in modTaxPrd.asp: "&pcErrDescription) 
		end If
		
		set rstemp=nothing
		call closeDb()
		response.redirect "viewTax.asp"
	End If
	
	If sMode="MOD" Then
		idTaxPerProduct=Request.QueryString("idTaxPerProduct")
		query="SELECT * FROM taxPrd WHERE idTaxPerProduct="&idTaxPerProduct
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		rstemp.Open query, conntemp
		if err.number <> 0 then
		  pcErrDescription = err.description
		  set rstemp=nothing
		  call closeDb()
		  response.redirect "techErr.asp?error="& Server.Urlencode("Error in modTaxPrd.asp: "&pcErrDescription) 
		end If
		pcIntIdProduct = rstemp("idproduct")
		pcStrTaxProduct = rstemp("taxperproduct")
		pcStrStateCodeEq = rstemp("stateCodeEq")
		pcStrZipEq = rstemp("zipEq")
		pcStrCountryCodeEq = rstemp("CountryCodeEq")
		pcStrZipCode = rstemp("zip")
		pcStrStateCode = rstemp("statecode")
		pcStrCountryCode = rstemp("CountryCode")
		set rstemp = nothing
		query="SELECT * FROM products WHERE idproduct="&pcIntIdProduct
		set rsObj=Server.CreateObject("ADODB.Recordset")
		rsObj.Open query, conntemp
		pcStrPrdDescription=rsObj("description")
		set rsObj = nothing
	End If
	
End If
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" name="addtax" action="modTaxPrd.asp" class="pcForms">
<table class="pcCPcontent">
    <tr> 
        <td colspan="3"><h2>Product: <%=pcStrPrdDescription%> </h2>
        <input type="hidden" name="idTaxPerProduct" value="<%=idTaxPerProduct%>">
        </td>
    </tr>       
    <tr> 
        <td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
    <td nowrap align="right">Tax Rate:</td>
    <td colspan="2">
    <% fTaxPrd=(pcStrTaxProduct*100) %>
    <input name="taxPerProduct" size="6" value="<%=fTaxPrd%>">
    <span class="pcSmallText">(5=5%)</span></td>
    </tr>
    <tr class="normal"> 
        <td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td align="right"> 
		<% If pcStrZipEq="-1" Then %>
        <input type="checkbox" name="zipEq" value="-1" checked>
        <% Else %>
        <input type="checkbox" name="zipEq" value="-1">
        <% End If %>
        </td>
        <td>Postal Code:   
        <input name="zip" size="12" value="<%=pcStrZipCode%>">
        </td>
	</tr>
    <tr class="normal"> 
        <td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td align="right"> 
			<% If pcStrStateCodeEq="-1" Then %>
            <input type="checkbox" name="stateCodeEq" value="-1" checked>
            <% Else %>
            <input type="checkbox" name="stateCodeEq" value="-1">
            <% End If %>
        </td>
        <td>State/Province Code: 
            <% 
            query="SELECT statecode, statename FROM states ORDER BY stateName"
            set rsObj=Server.CreateObject("ADODB.Recordset")     
            rsObj.Open query, conntemp
            if err.number <> 0 then
				set rsObj = nothing
				call closeDb() 
            	response.redirect "techErr.asp?error="& Server.Urlencode("Error in modTaxPrd 4: "&Err.Description) 
            end If
            %>
            <SELECT name="stateCode" size="1">
            <OPTION value="">State Code 
            <% do until rsObj.eof %>
            <% if pcStrStateCode=rsObj("stateCode") Then %>
            <option value="<%=rsObj("stateCode")%>" selected><%=rsObj("stateName")%></OPTION>
            <% Else %>
            <option value="<%=rsObj("stateCode")%>"><%=rsObj("stateName")%></OPTION>
            <% End If %>
            <%  
            rsObj.moveNext
            loop
            %>
            </SELECT>
        </td>
    </tr>
    <tr class="normal"> 
        <td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td align="right"> 
        <% If pcStrCountryCodeEq="-1" Then %>
        <input type="checkbox" name="CountryCodeEq" value="-1" checked>
        <% Else %>
        <input type="checkbox" name="CountryCodeEq" value="-1">
        <% End If %>
        </td>
        <td>Country: 
        <% query="SELECT * FROM countries ORDER BY countryName"
        set rsObj=Server.CreateObject("ADODB.Recordset")     
        rsObj.Open query, conntemp
        if err.number <> 0 then
			set rsObj = nothing
			call closeDb() 
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in modTaxPrd 5: "&Err.Description) 
		end If
		%>
        <SELECT name="CountryCode">
        <OPTION value="">Country </option>
        <% do until rsObj.eof %>
        <% if pcStrCountryCode=rsObj("CountryCode") Then %>
        <option value="<%=rsObj("CountryCode")%>" selected><%=rsObj("countryName")%></OPTION>
        <% Else %>
        <option value="<%=rsObj("CountryCode")%>"><%=rsObj("countryName")%></OPTION>
        <% End If %>
        <% rsObj.moveNext
        loop 
		set rsObj = nothing
		call closeDb() 
		%>
        </SELECT>
        </td>
        </tr>
    <tr class="normal"> 
        <td colspan="3" class="pcCPspacer"><hr></td>
    </tr>                
    <tr> 
    <td></td>
    <td colspan="2">  
    <input type="submit" name="Submit" value="Update" class="submit2">
    &nbsp;
    <input type="button" value="Back" onClick="JavaScript:history.back()">
    </td>
    </tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->