<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Tax Settings - Manual Entry Method - Step 3: Tax by Location" %>
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
dim query, conntemp, rs

sMode=Request.Form("Submit")

If sMode <> "" Then
	If sMode="Add" Then
		pcv_taxPerPlace=(Request.Form("taxPerPlace")/100)
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
			case "zone"
				intTaxZoneID=request.Form("TaxZoneID")
				if intTaxZoneID="" then
					response.Redirect "AddTaxPerPlace.asp?m=zone"
					response.End()
				end if
		end select
		
		call openDb()
		query="INSERT INTO taxLoc (CountryCode, CountryCodeEq, stateCode, stateCodeEq, zip, zipEq, taxLoc,taxDesc) VALUES ('"& pcv_CountryCode &"',"& Cint(pcv_CountryCodeEq) &", '"& pcv_stateCode &"',"& Cint(pcv_stateCodeEq) &",'"& pcv_zip &"',"& Cint(pcv_zipEq) &","& taxmoney(pcv_taxPerPlace) &",'"&pcv_taxDesc&"')"
		set rs=Server.CreateObject("ADODB.Recordset")     
		rs.Open query, conntemp
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		set rs=nothing
		call closedb()
	End If
	response.redirect "viewTax.asp"
End If

taxErrMsg=request.QueryString("m")
if taxErrMsg<>"" then
	select case taxErrMsg
	case "z"
		msg="Postal Code is required when you choose to set a rule by Postal Code."
	case "sc"
		msg="State/Province Code is required when you choose to set a rule by State Code."
	case "cc"
		msg="Country Code is required when you choose to set a rule by Country Code."
	case "zone"
		msg="Zone is required when you choose to set a rule by Zone."
	end select
end if
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" name="addSh" action="AddTaxPerPlace.asp" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr> 
			<td width="20%" nowrap="nowrap">Tax rate:</td>
			<td width="80%"><input name="taxPerPlace" size="6" value="0"> % <span class="pcSmallText">(e.g. 5 = 5%)</span></td>
		</tr>
		<tr>
			<td nowrap="nowrap">Description shown:</td>
			<td><input name="taxDesc" type="text" id="taxDesc" size="20" maxlength="50">&nbsp;<span class="pcSmallText">(e.g. Sale Tax)</span></td>
		</tr>
		<tr> 
			<td></td>
			<td>If you chose to show each tax separately, then a tax description is required and will be shown for each tax rule that is applied to the order.</td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr>
			<td colspan="2">You can define the location as a single location or a zone (a group of locations). For example, a zone can be a group of states or provinces that share the same tax.
            <ul class="pcListIcon">
            	<li>Define tax by <strong>zone</strong>: <a href="AddTaxPerZone.asp">Add/Edit Tax Zones</a></li>
                <li style="padding-top: 10px;">Define tax by <strong>individual location</strong>: specify the location below, then click on <em>Add</em>:</li>
            </ul>
            </td>
		</tr>
		<tr> 
			<td align="right"><input name="taxType" type="radio" id="taxType" value="zip" checked class="clearBorder"></td>
			<td>Tax by Postal Code</td>
		</tr>
                
		<tr> 
			<td></td>
			<td><input name="zip" size="12" value="Postal Code"></td>
		</tr>
		<tr> 
			<td align="right"><input name="taxType" type="radio" id="taxType" value="state" class="clearBorder"></td>
			<td>Tax by State Or Province</td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>  
				<% call openDb()
				query="SELECT states.stateCode, states.stateName, countries.countryCode, countries.countryName FROM countries INNER JOIN states ON countries.countryCode = states.pcCountryCode ORDER BY  countries.countryName Desc, states.stateName ASC;"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				%>
				<SELECT name="stateCode" size=1>
					<OPTION value="">State Code </OPTION>
						<% do until rs.eof
							strTmpStateCode=rs("stateCode")
							strTmpStateName=rs("stateName")
							strTmpCountryCode=rs("countryCode")
							strTmpCountryName=rs("countryName")
							if isNULL(strTmpCountryCode) or strTmpCountryCode="" then
								strTmpStateCountryCode=strTmpStateName
							else
								strTmpStateCountryCode=strTmpStateCode&"||"&strTmpCountryCode
							end if %>
 							<option value="<%=strTmpStateCountryCode%>"><%=strTmpStateName&" - "&strTmpCountryName%></OPTION>
							<% rs.moveNext
						loop
						set rs=nothing %>
				</SELECT>
			</td>
		</tr>
		<tr> 
			<td align="right"><input name="taxType" type="radio" id="taxType" value="country" class="clearBorder"></td>
			<td>Tax by Country</td>
		</tr>
									
		<tr> 
			<td></td>
			<td>  
				<% query="SELECT CountryCode, countryName FROM countries ORDER BY countryName"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				%>
				<select name="CountryCode">
					<option value="">Country</option>
					<% do until rs.eof %>
						<option value="<%=rs("CountryCode")%>"><%=rs("countryName")%></option>
						<% rs.moveNext
					loop %>
				</select>
			</td>
		</tr>
		<% 'If Zones exist, shown drop down 
		query="SELECT pcTaxZoneRates.pcTaxZoneRate_ID, pcTaxZoneRates.pcTaxZoneRate_Name FROM pcTaxZoneRates;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.eof then %>								
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td align="right"><input name="taxType" type="radio" id="taxType" value="zone" class="clearBorder"></td>
				<td>Tax by Zone</td>
			</tr>
			<tr>
				<td></td>
				<td>
				 <select name="TaxZoneID">
						<option value="">Select Zone</option>
						<% do until rs.eof
							intTaxZoneID=rs("pcTaxZoneRate_ID")
							strTaxZoneName=rs("pcTaxZoneRate_Name")
							 %>
							<option value="<%=intTaxZoneID%>"><%=strTaxZoneName%></option>
							<% rs.moveNext
						loop 
						%>
					</select>
				</td>
			</tr>
		<% 
		end if
		set rs=nothing
		call closedb() 
		%>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>                 
			<input type="submit" name="Submit" value="Add" class="submit2">                    
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->