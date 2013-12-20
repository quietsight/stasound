<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=6%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/languages_ship.asp" --> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->

<html>
<head>
<title><%=title%></title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:470px; background-image: none;">
	
	<% 	dim query, rs, conntemp
	
	If request("mode")<>"" then
		if request("mode")="DEL" then
			call opendb()
			intDelTaxZoneDescID=getUserInput(request("TaxZoneDescID"),0)
			if NOT validNum(intDelTaxZoneDescID) then
				response.redirect "menu.asp"
			end if
			query="DELETE FROM pcTaxGroups WHERE pcTaxZoneDesc_ID="&intDelTaxZoneDescID&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZoneDescriptions WHERE pcTaxZoneDesc_ID="&intDelTaxZoneDescID&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZonesGroups WHERE pcTaxZoneDesc_ID="&intDelTaxZoneDescID&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		
			'//Check for orphaned tax zones and delete
			query="SELECT pcTaxZone_ID FROM pcTaxZones;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			do until rs.eof
				intTempTaxZoneID=rs("pcTaxZone_ID")
				query="SELECT pcTaxGroup_ID FROM pcTaxGroups WHERE pcTaxZone_ID="&intTempTaxZoneID&";"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				if rstemp.eof then
					'//Delete orphaned tax zone
					query="DELETE FROM pcTaxZones WHERE pcTaxZone_ID="&intTempTaxZoneID&";" 
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
				end if
				set rstemp=nothing
				rs.movenext
			loop					
		end if
		call closedb()
		response.redirect "pcTaxViewZones.asp?success=1"
	else
		call openDb()
				
		query="SELECT pcTaxZoneDescriptions.pcTaxZoneDesc_ID, pcTaxZoneDescriptions.pcTaxZoneDesc FROM pcTaxZoneDescriptions;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		
		if rs.eof then
			response.redirect "AdminTaxZones.asp"		
		else %>
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="2"><p><a href="AdminTaxZones.asp">
					  <input type="button" name="New2" value="Add New Zone" class="submit2" onClick="location.href='AdminTaxZones.asp'">
					</a></p></td>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<% do until rs.eof
					intTaxZoneDescID=rs("pcTaxZoneDesc_ID")
					strTaxZoneDesc=rs("pcTaxZoneDesc")
					query="SELECT pcTaxZones.pcTaxZone_CountryCode, pcTaxZones.pcTaxZone_Province, states.stateName, countries.countryName FROM ((pcTaxGroups INNER JOIN pcTaxZones ON pcTaxGroups.pcTaxZone_ID = pcTaxZones.pcTaxZone_ID) INNER JOIN countries ON pcTaxZones.pcTaxZone_CountryCode = countries.countryCode) INNER JOIN states ON (countries.countryCode = states.pcCountryCode) AND (pcTaxZones.pcTaxZone_Province = states.stateCode) WHERE (((pcTaxGroups.pcTaxZoneDesc_ID)="&intTaxZoneDescID&"));"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query) %>
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
					<tr>
						<th width="80%"><p><%=strTaxZoneDesc%></p></th>
						<th width="20%" align="right"><span class="pcSmallText"><a href="EditTaxZones.asp?TaxZoneDescID=<%= intTaxZoneDescID %>">Edit</a> | <a href="javascript:if (confirm('You are about to permanantly delete this tax zone from the database. Are you sure you want to complete this action?')) location='pcTaxViewZones.asp?TaxZoneDescID=<%= intTaxZoneDescID %>&mode=DEL'">Delete</a></span></th>
					</tr>
					<% do until rstemp.eof
						strProvince=rstemp("stateName")
						strCountryCode=rstemp("countryName") %>
						<tr>
							<td colspan="2" class="pcCPspacer"><p><%=strCountryCode &" / "&strProvince%></p></td>
						</tr>
						
						<% rstemp.movenext
					loop
					set rstemp=nothing
					rs.movenext
				loop %>
			</table>
		<% 	end if
	end if %>
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr>
			<td colspan="2" align="center"><p><input type="button" name="New" value="Add New Zone" class="submit2" onClick="location.href='AdminTaxZones.asp'">
			&nbsp;
			<% if request("success")="1" then %>
				<input type="button" name="Complete" value="Complete Tax Zone Configuration" class="submit2" onClick="opener.location.reload(); self.close();">
			<% else %>
				<input type="button" name="Back" class="ibtnGrey" value="Close" onClick="self.close();">
			<% end if %></p>
			</td>
		</tr>
	</TABLE>
</div>
</body>
</html>
