<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Seach Fields - Mapping Complete" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<% dim query, conntemp, rstemp

pcv_strExportType = request("export")
Select Case pcv_strExportType
	Case "f": pcv_strExportFile = "Google Shopping"
	Case "c": pcv_strExportFile = "Cashback"
End Select

call openDb()

validfields=request.form("validfields")

'/////////////////////////////////////////////////////////////////////
'// START: REMOVE MAPPINGS
'/////////////////////////////////////////////////////////////////////
query="DELETE from pcSearchFields_Mappings WHERE pcSearchFieldsFileID='"& pcv_strExportType &"' "
set rs=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
'/////////////////////////////////////////////////////////////////////
'// END: REMOVE MAPPINGS
'/////////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////////
'// START: ADD MAPPINGS
'/////////////////////////////////////////////////////////////////////
For i=1 to validfields
	if trim(ucase(request("T" & i)))<>"0" then
		pcv_intSearchField = request("T" & i)
		pcv_intSearchFieldName = request("F" & i)
		query="INSERT INTO pcSearchFields_Mappings (idSearchField, pcSearchFieldsColumn, pcSearchFieldsFileID) VALUES ("& pcv_intSearchField &",'"& pcv_intSearchFieldName &"','"& pcv_strExportType &"')"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
	end if
Next
'/////////////////////////////////////////////////////////////////////
'// END: ADD MAPPINGS
'/////////////////////////////////////////////////////////////////////
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" action="SearchFields_Export3.asp?export=<%=pcv_strExportType%>" class="pcForms">
	<table class="pcCPcontent">
    <tr>
      <td class="pcCPspacer"></td>
    </tr>	
		<tr>
    <td valign="top">
      <div class="pcCPmessageSuccess">We have successfully mapped your custom search fields to export fields.</div>        	
     </td>
    </tr> 
    <tr>
      <td class="pcCPspacer"></td>
    </tr>	                 
    <tr>
        <td>
          <% if pcv_strExportFile = "Google Shopping" then %>
            <input type=button name=backstep value="<< Return to Google Shopping data feed" onClick="location='exportFroogle.asp';" class="submit2">&nbsp; 
          <% else %>
            <input type=button name=backstep value="<< Return to Bing Shopping data feed" onClick="location='exportCashback.asp';" class="submit2">&nbsp;
          <% end if %>
            <input type=button name=backstep value="Manage Search Fields " onClick="location='ManageSearchFields.asp';">          
      </td>
    </tr>
    </table>
</form>
<% call closeDb() %>
<!--#include file="AdminFooter.asp"-->