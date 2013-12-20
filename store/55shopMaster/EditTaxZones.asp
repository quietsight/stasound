<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=6%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<% 
TaxZoneDescID=request("TaxZoneDescID")
dim conntemp, rs, query

call opendb()

if request("submit")<>"" then
	TaxZoneDesc=request("TaxZoneDesc")
	TaxZoneProvinces=request("TaxZoneProvinces")
	TaxZoneCountryCode=request("TaxZoneCountryCode")
	
	query="UPDATE pcTaxZoneDescriptions SET pcTaxZoneDesc='"&TaxZoneDesc&"' WHERE pcTaxZoneDesc_ID="&TaxZoneDescID&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	'// DELETE all from pcTaxGroups
	query="DELETE FROM pcTaxGroups WHERE pcTaxZoneDesc_ID="&TaxZoneDescID&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	arryProvinces=split(TaxZoneProvinces,", ")
	for i=0 to ubound(arryProvinces)
		'// check if it exists in pcTaxZones, if not, Add
		query1="SELECT pcTaxZone_ID FROM pcTaxZones WHERE pcTaxZone_CountryCode='"&TaxZoneCountryCode&"' AND pcTaxZone_Province='"&arryProvinces(i)&"';"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query1)
		if rstemp.eof then
			query="INSERT INTO pcTaxZones (pcTaxZone_CountryCode, pcTaxZone_Province) VALUES ('"&TaxZoneCountryCode&"', '"&arryProvinces(i)&"');"
			set rstemp2=server.CreateObject("ADODB.RecordSet")
			set rstemp2=conntemp.execute(query)
			
			set rstemp2=server.CreateObject("ADODB.RecordSet")
			set rstemp2=conntemp.execute(query1)
			intTaxZoneID=rstemp2("pcTaxZone_ID")
			set rstemp2=nothing
		else
			intTaxZoneID=rstemp("pcTaxZone_ID")
		end if
		
		query="INSERT INTO pcTaxGroups (pcTaxZoneDesc_ID, pcTaxZone_ID) VALUES ("&TaxZoneDescID&","&intTaxZoneID&");"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)
		
		set rstemp=nothing
	next
	set rs=nothing
	pcSuccess=1
end if

query="SELECT pcTaxZoneDescriptions.pcTaxZoneDesc, pcTaxZones.pcTaxZone_CountryCode, countries.countryName, pcTaxZones.pcTaxZone_Province FROM ((pcTaxZoneDescriptions INNER JOIN pcTaxGroups ON pcTaxZoneDescriptions.pcTaxZoneDesc_ID = pcTaxGroups.pcTaxZoneDesc_ID) INNER JOIN pcTaxZones ON pcTaxGroups.pcTaxZone_ID = pcTaxZones.pcTaxZone_ID) INNER JOIN countries ON pcTaxZones.pcTaxZone_CountryCode = countries.countryCode WHERE (((pcTaxZoneDescriptions.pcTaxZoneDesc_ID)="&TaxZoneDescID&"));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
pcv_TaxZoneDesc=rs("pcTaxZoneDesc")
pcv_TaxZoneCountryCode=rs("pcTaxZone_CountryCode")
pcv_TaxZoneCountryName=rs("countryName")
set rs=nothing

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

pcv_ProvinceString=""
do until rs.eof
	pcv_ProvinceString=pcv_ProvinceString&","&rs("pcTaxZone_Province")
	rs.movenext
loop
set rs=nothing
%>
<html>
<head>
<title><%=title%></title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:470px; background-image: none;">
<form action="EditTaxZones.asp" method="post" name="form1" id="form1" class="pcForms">
	<input name="TaxZoneDescID" type="hidden" value="<%=TaxZoneDescID%>">
  <table class="pcCPcontent">
    <tr>
        <td colspan="4" class="pcCPspacer"></td>
    </tr>
  <tr>
    <th colspan="2"><p>Edit Zone</p></th>
  </tr>
    <tr>
        <td colspan="4" class="pcCPspacer"></td>
    </tr>
  <tr>
    <td width="16%" valign="top"><p>Zone Name: </p></td>
    <td width="84%"><p>
      <input type="text" name="TaxZoneDesc" value="<%=pcv_TaxZoneDesc%>"/>
    </p></td>
  </tr>
  <tr>
    <td valign="top"><p>Country:</p></td>
    <td><p><%=pcv_TaxZoneCountryName%></p>
		<input type="hidden" name="TaxZoneCountryCode" value="<%=pcv_TaxZoneCountryCode%>">
		</td>
  </tr>
	<% 
	query="SELECT states.stateCode, states.stateName, states.pcCountryCode FROM states WHERE (((states.pcCountryCode)='"&pcv_TaxZoneCountryCode&"'));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	%>
  <tr>
    <td valign="top"><p>States/Provinces:</p></td>
    <td><p>
      <select name="TaxZoneProvinces" size="10" multiple="multiple">
				<% do until rs.eof 
					pcv_selected=0
					pcv_stateCode=rs("stateCode")
					pcv_stateName=rs("stateName") 
					if instr(pcv_ProvinceString, pcv_stateCode) then
						pcv_selected=1
					end if
					%>
        <option value="<%=pcv_stateCode%>" <% if pcv_selected=1 then%>selected="selected"<% end if %>><%=pcv_stateName%></option>
				<%rs.movenext
			loop
			set rs=nothing
			call closedb() 
			%>
      </select>
    </p>
    <p>&nbsp;</p></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p>
      <input type="submit" name="submit" class="submit2" value="Update">&nbsp;
			<% if pcSuccess=1 then
				strBackButton="Finished"
			else
				strBackButton="Back"
			end if %>
			<input name="button" type="button" class="ibtnGrey" onClick="location.href='pcTaxViewZones.asp?success=1'" value="<%=strBackButton%>">
      </td>
  </tr>
</table>
</form>
</div>
</body>
</html>