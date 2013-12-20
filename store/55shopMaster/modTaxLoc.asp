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

sMode=Request.queryString("mode")
cMode=Request.form("Submit")
If cMode="Update" Then
	idTaxPerPlace=Request.Form("idTaxPerPlace")
	pcv_taxLoc=(Request.Form("taxLoc")/100)
	pcv_taxDesc=request.Form("taxDesc")
	pcv_taxType=request.form("taxType")
	select case pcv_taxType
		case "zip"
			pcv_zipEq="-1"
			pcv_zip=Request.Form("zip")
			pcv_CountryCode=""
			pcv_stateCode=""
			pcv_stateCodeEq="0"
			pcv_CountryCodeEq="0"
			if pcv_zip="" then
				response.redirect "AddTaxPerPlace.asp?m=z"
				response.End()
			end if
		case "state"
			pcv_zipEq="0"
			strStateCountryArray=Request.Form("stateCode")
			if instr(strStateCountryArray,"||") then
				pcv_StateCountryArray=split(strStateCountryArray,"||")
				pcv_stateCode=pcv_StateCountryArray(0)
				pcv_CountryCode=pcv_StateCountryArray(1)
				pcv_CountryCodeEq="-1"
			else
				pcv_stateCode=strStateCountryArray
				pcv_CountryCode=""
				pcv_CountryCodeEq="0"
			end if
			pcv_zip=""
			pcv_stateCodeEq="-1"
			if pcv_stateCode="" then
				response.redirect "AddTaxPerPlace.asp?m=sc"
				response.End()
			end if
		case "country"
			pcv_zipEq="0"
			pcv_CountryCode=Request.Form("CountryCode")
			pcv_stateCode=""
			pcv_zip=""
			pcv_stateCodeEq="0"
			pcv_CountryCodeEq="-1"
			if pcv_CountryCode="" then
				response.redirect "AddTaxPerPlace.asp?m=cc"
				response.End()
			end if
	end select

	call openDb()
	query="UPDATE taxLoc SET CountryCode='"& pcv_CountryCode &"', CountryCodeEq="& pcv_CountryCodeEq &", stateCode='"& pcv_stateCode &"', stateCodeEq="& pcv_stateCodeEq &", zip='"& pcv_zip &"', zipEq="& pcv_zipEq &", taxLoc="& taxmoney(pcv_taxLoc) &", taxDesc='"& pcv_taxDesc &"' WHERE idTaxPerPlace="&idTaxPerPlace

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	
	rstemp.Open query, conntemp
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

	set rstemp = nothing
	call closedb()
	response.redirect "viewTax.asp"
End If

If sMode <> "" Then
	If sMode="DEL" Then
		idTaxPerPlace=Request.QueryString("idTaxPerPlace")
		call openDb()
		query="DELETE FROM taxLoc WHERE idTaxPerPlace="&idTaxPerPlace
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		rstemp.Open query, conntemp
		
		if err.number <> 0 then
			pcvErrDescription = err.description
			set rstemp = nothing
			call closedb()
		  	response.redirect "techErr.asp?error="& Server.Urlencode("Error in modtaxLoc 1: "&pcvErrDescription) 
		end If
		
		set rstemp = nothing
		call closedb()
		response.redirect "viewTax.asp"
	End If
	
	If sMode="DELZONE" then
		idTaxPerPlace=request.QueryString("idTaxPerPlace")
		call opendb()
		query="DELETE FROM pcTaxZoneRates WHERE pcTaxZoneRate_ID="&idTaxPerPlace&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		query="DELETE FROM pcTaxZonesGroups WHERE pcTaxZoneRate_ID="&idTaxPerPlace&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		set rs=nothing
		call closedb()
		
		response.Redirect("viewTax.asp")
		
	end if
	
	If sMode="MOD" Then
		idTaxPerPlace=Request.QueryString("idTaxPerPlace")
		call openDb()
		query="SELECT CountryCode, CountryCodeEq, stateCode, stateCodeEq, zip, zipEq, taxLoc,taxDesc FROM taxLoc WHERE idTaxPerPlace="&idTaxPerPlace
		set rsTaxLoc=Server.CreateObject("ADODB.Recordset")     
		rsTaxLoc.Open query, conntemp
		
		if err.number <> 0 then
			pcvErrDescription = err.description
			set rsTaxLoc = nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in modtaxLoc 2: "&pcvErrDescription) 
		end If
		pcv_CountryCode=rsTaxLoc("CountryCode")
		pcv_CountryCodeEq=rsTaxLoc("CountryCodeEq")
		pcv_stateCode=rsTaxLoc("stateCode")
		pcv_stateCodeEq=rsTaxLoc("stateCodeEq")
		pcv_zip=rsTaxLoc("zip")
		pcv_zipEq=rsTaxLoc("zipEq")
		pcv_taxLoc=rsTaxLoc("taxLoc")
		pcv_taxLoc=pcv_taxLoc*100
		pcv_taxDesc=rsTaxLoc("taxDesc")  
	End If
End If
%>
<!--#include file="AdminHeader.asp"-->

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="post" name="addtax" action="modtaxLoc.asp" class="pcForms">
<table class="pcCPcontent">
  <tr>
    <td width="20%" align="right">Tax</td>
    <td width="80%">
        <input name="taxLoc" id="taxLoc" value="<%=pcv_taxLoc%>" size="6"> (5=5%)
        <input name="idTaxPerPlace" type="hidden" id="idTaxPerPlace" value="<%=idTaxPerPlace%>">
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>Description of tax: <input name="taxDesc" type="text" id="taxDesc" size="20" maxlength="50" value="<%=pcv_taxDesc%>"> <span class="pcSmallText">(e.g. Local Sales Tax)</span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>If you chose to show each tax separately, then a tax description is required and will be shown for each tax rule that is applied to the order.</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">
		<% if pcv_zipeq="-1" then %>
        	<input name="taxType" type="radio" id="taxType" value="zip" checked>
		<% else %>
        	<input name="taxType" type="radio" id="taxType" value="zip">
		<% end if %>
    </td>
    <td>Tax by Postal Code</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input name="zip" size="12" value="<%=pcv_zip%>"></td>
  </tr>
  <tr>
    <td align="right">
		<% if pcv_stateCodeeq="-1" then %>
        <input name="taxType" type="radio" id="taxType" value="state" checked>
		<% else %>
		<input name="taxType" type="radio" id="taxType" value="state">
		<% end if %>
    </td>
    <td>Tax by State Or Province</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
        <% 
		call openDb()
		query="SELECT states.stateCode, states.stateName, countries.countryCode, countries.countryName FROM countries INNER JOIN states ON countries.countryCode = states.pcCountryCode ORDER BY  countries.countryName Desc, states.stateName ASC;"

		set rstemp=Server.CreateObject("ADODB.Recordset")     
	
		rstemp.Open query, conntemp
	
		if err.number <> 0 then
			pcvErrDescription
			set rstemp = nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddTaxPerPlace 2: "&pcvErrDescription) 
		end If
		%>
		<SELECT name=stateCode size=1>
			<OPTION value="">State Code
			<% do until rstemp.eof 
				strTmpStateCode=rstemp("stateCode")
				strTmpStateName=rstemp("stateName")
				strTmpCountryCode=rstemp("countryCode")
				strTmpCountryName=rstemp("countryName")
				if isNULL(strTmpCountryCode) or strTmpCountryCode="" then
					strTmpStateCountryCode=strTmpStateName
				else
					strTmpStateCountryCode=strTmpStateCode&"||"&strTmpCountryCode
				end if
				if pcv_stateCode=strTmpStateCode AND pcv_CountryCode=strTmpCountryCode then %>
					<option value="<%=strTmpStateCountryCode%>" selected><%=rstemp("stateName")&" - "&strTmpCountryName%></OPTION>
				<% else %>
					<option value="<%=strTmpStateCountryCode%>"><%=strTmpStateName&" - "&strTmpCountryName%></OPTION>
				<% end if %>
				<% rstemp.moveNext
			loop %>
		</SELECT>
    </td>
  </tr>
  <tr>
    <td align="right">
		<% if pcv_CountryCodeeq="-1" AND pcv_StateCodeeq="0" then %>
        <input name="taxType" type="radio" id="taxType" value="country" checked>
		<% else %>
        <input name="taxType" type="radio" id="taxType" value="country">
		<% end if %>
    </td>
    <td><div align="left">Tax by Country</div>
    </td>
  </tr>
  <tr>
    <td width="26"></td>
    <td>
      <% query="SELECT CountryCode, countryName FROM countries ORDER BY countryName"
		set rstemp=Server.CreateObject("ADODB.Recordset")     
	
		rstemp.Open query, conntemp
	
		if err.number <> 0 then
			pcvErrDescription
			set rstemp = nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddTaxPerPrd 3: "&pcvErrDescription) 
		end If
		 %>
      <select name=CountryCode>
        <option value="">Country</option>
        <% do until rstemp.eof
				if pcv_CountryCode=rstemp("CountryCode") then %>
        <option value="<%=rstemp("CountryCode")%>" selected><%=rstemp("countryName")%></option>
				<% else %>
        <option value="<%=rstemp("CountryCode")%>"><%=rstemp("countryName")%></option>
				<% end if %>
        <% 
			rstemp.moveNext
			loop 
			set rstemp = nothing
			call closedb()
		%>
      </select>
    </td>
  </tr>
  <tr>
    <td colspan="2"><hr></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
      <input type="submit" name="Submit" value="Update" class="submit2">&nbsp;
      <input type="button" name="back" value="Back" onClick="javascript:history.back()">
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->