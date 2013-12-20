<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%

dim rs, conntemp, query

call openDb()

	if request("action")="update" then
		Count=request("Count")
		For k=1 to Count
			if not request("C" & k)="1" then
				query="DELETE FROM pcVATCountries WHERE pcVATCountry_Code='" & request("pcVATCountry_Code" & k) & "';"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
		Next
	end if

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="pcVATCountry_State"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If 

query="SELECT * FROM pcVATCountries ORDER BY "& strORD &" "& strSort
set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)

%>
<% pageTitle="Manage EU Member States" %>
<% section="layout" %>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="manageEUStates.asp?action=update" class="pcForms">
<table class="pcCPcontent" style="width:auto">
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="4">To remove any records, uncheck them and click on &quot;Update&quot; &nbsp;|&nbsp;<a href="AddEUStates.asp">Add New State</a></td>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
<tr> 
	<th></th>
	<th nowrap="nowrap"><a href="manageEUStates.asp?order=pcVATCountry_State&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a><a href="manageEUStates.asp?order=pcVATCountry_State&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>&nbsp;State Name</th>
	<th nowrap="nowrap" colspan="2">
	<a href="manageEUStates.asp?order=pcVATCountry_Code&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="manageEUStates.asp?order=pcVATCountry_Code&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Country Code</th>
</tr>
                      
<%if rs.eof then%>                     
<tr> 
	<td colspan="4">No EU Member States found.</td>
</tr>                
<%end if%>
                      
<%
	Count=0
	Do While NOT rs.EOF
	Count=Count + 1
	pcVATCountry_Code=rs("pcVATCountry_Code")
	StateName=rs("pcVATCountry_State")
	CountryCode=rs("pcVATCountry_Code")
%>
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
		<td>  
		<input type="checkbox" name="C<%=Count%>" value="1" checked class="clearBorder">
		<input type="hidden" name="pcVATCountry_Code<%=Count%>" value="<%=pcVATCountry_Code%>">
		</td>
		<td><a href="EditEUStates.asp?StateCode=<%=pcVATCountry_Code%>"><%=StateName%></a></td>
		<td><%=CountryCode%></td>
		<td align="right"><a href="EditEUStates.asp?StateCode=<%=pcVATCountry_Code%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit this entry"></a></td>
	</tr>
<% 
	rs.MoveNext
	loop
	set rs=nothing
	call closeDb()
%>
<input type=hidden name=Count value="<%=Count%>">
<tr>
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="4" class="cpLinksList">
		<a href="javascript:checkAll();">Check All</a>
		&nbsp;|&nbsp;
		<a href="javascript:uncheckAll();">Uncheck All</a>
		&nbsp;|&nbsp;
		<a href="AddEUStates.asp">Add New State</a>
	</td>
</tr>
<tr>
	<td colspan="4"><hr></td>
</tr>
<tr>
	<td colspan="4">
		<input name="submit" type="submit" class="submit2" value="Remove unchecked">
		&nbsp;
		<input type="button" name="Button" value="Return to VAT Settings" onClick="javascript:location='AdminTaxSettings_VAT.asp'">
		&nbsp;
		<input type="button" name="Restore" value="Restore Default Settings" onClick="javascript:location='restoreEUStates.asp?action=update'">
	</td>
</tr>   
</table>
</form>
<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.form1.C" + j); 
if (box.checked == false) box.checked = true;
   }
}

function uncheckAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.form1.C" + j); 
if (box.checked == true) box.checked = false;
   }
}
//-->
</script>            
<!--#include file="AdminFooter.asp"-->