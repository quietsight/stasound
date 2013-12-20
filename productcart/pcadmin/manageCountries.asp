<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->    
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<%

dim rs, conntemp, query

call openDb()

if request("action")="update" then
	Count=request("Count")
	For k=1 to Count
		if not request("C" & k)="1" then
			query="delete from Countries where CountryCode='" & request("Countrycode" & k) & "'"
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
	strORD="countryname"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If

query="SELECT * FROM Countries ORDER BY "& strORD &" "& strSort
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)
%>
<% pageTitle="Manage Countries" %>
<% section="layout" %>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="manageCountries.asp?action=update" class="pcForms">
<table class="pcCPcontent">
	<tr>
        <td colspan="3" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
	</tr>
	<tr>
		<td colspan="3">Uncheck the checkboxes of the coutries that you would like to remove, then click on &quot;Update&quot; &nbsp;|&nbsp;<a href="AddCountries.asp">Add New Country</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=437')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th nowrap="nowrap" colspan="2"><a href="manageCountries.asp?order=CountryName&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a><a href="manageCountries.asp?order=countryname&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>&nbsp;Country Name</th>
		<th nowrap="nowrap"><a href="manageCountries.asp?order=countrycode&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="manageCountries.asp?order=countrycode&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Country Code</th>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
                      
<%if rs.eof then%>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<tr>                
		<td colspan="3">No Countries found.</td>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>              
<%end if%>
                      
<%
		Count=0
		Do While NOT rs.EOF
		Count=Count + 1
		CountryName=rs("CountryName")
		CountryCode=rs("CountryCode")
%>
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
		<td> 
			<input type="checkbox" name="C<%=Count%>" value="1" checked class="clearBorder">
			<input type="hidden" name="CountryCode<%=Count%>" value="<%=CountryCode%>">
		</td>
		<td><a href="EditCountries.asp?CountryCode=<%=CountryCode%>"><%=CountryName%></a></td>
		<td><a href="EditCountries.asp?CountryCode=<%=CountryCode%>"><%=CountryCode%></a></td>
	</tr>
                      
<%
	rs.MoveNext
	loop
	set rs=nothing
	call closeDb()
%>
	<tr>
		<td colspan="3"></td>
	</tr> 
	<tr>
		<td colspan="3" class="cpLinksList">
			<input type="hidden" name="Count" value="<%=Count%>">
			<a href="javascript:checkAll();">Check All</a> | <a href="javascript:uncheckAll();">Uncheck All</a> | <a href="AddCountries.asp">Add New Country</a> 
		</td>
	</tr> 
	<tr>
		<td colspan="3" class="pcCPspacer"><hr></td>
	</tr> 
	<tr>
		<td colspan="3">
			<input name="submit" type="submit" value="Update" class="submit2">&nbsp;
			<input type="button" name="Button" value="Back" onClick="javascript:history.back()">&nbsp;
			<input type="button" name="Restore" value="Restore Default Country List" onClick="javascript:location='restoreCountries.asp?action=update'">
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