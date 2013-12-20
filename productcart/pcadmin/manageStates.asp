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
				query="delete from States where StateCode='" & request("Statecode" & k) & "' AND pcCountryCode='" & request("CountryCode" & k) & "'"
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
	strORD="statename"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If 

query="SELECT * FROM States ORDER BY "& strORD &" "& strSort
set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)

%>
<% pageTitle="Manage States" %>
<% section="layout" %>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="manageStates.asp?action=update" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="4" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
		<td colspan="5">Uncheck the checkboxes of the states that you would like to remove, then click on &quot;Update&quot; &nbsp;|&nbsp;<a href="AddStates.asp">Add New State</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=437')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
<tr> 
	<th colspan="2" nowrap="nowrap"><a href="manageStates.asp?order=StateName&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a><a href="manageStates.asp?order=Statename&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>&nbsp;State Name</th>
	<th colspan="1" nowrap="nowrap">
	<a href="manageStates.asp?order=Statecode&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>
	<a href="manageStates.asp?order=Statecode&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;State Code
	</th>
	<th nowrap="nowrap">
	<a href="manageStates.asp?order=pcCountryCode&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>
	<a href="manageStates.asp?order=pcCountryCode&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Country Code
	</th>
</tr>
                      
<%if rs.eof then%>                     
<tr> 
	<td colspan="4">No States found.</td>
</tr>                
<%end if%>
                      
<%
	Count=0
	Do While NOT rs.EOF
	Count=Count + 1
	StateCode=rs("StateCode")
	StateName=rs("StateName")
	CountryCode=rs("pcCountryCode")
%>
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
		<td>  
		<input type="checkbox" name="C<%=Count%>" value="1" checked class="clearBorder">
		<input type="hidden" name="StateCode<%=Count%>" value="<%=StateCode%>">
		<input type="hidden" name="CountryCode<%=Count%>" value="<%=CountryCode%>">
		</td>
		<td><a href="EditStates.asp?StateCode=<%=StateCode%>&CountryCode=<%=CountryCode%>"><%=StateName%></a></td>
		<td><a href="EditStates.asp?StateCode=<%=StateCode%>&CountryCode=<%=CountryCode%>"><%=StateCode%></a></td>
		<td><%=CountryCode%></td>
	</tr>
<% 
	rs.MoveNext
	loop
	set rs=nothing
	call closeDb()
%>
<tr>
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="4" class="cpLinksList">
		<a href="javascript:checkAll();">Check All</a> | <a href="javascript:uncheckAll();">Uncheck All</a> | <a href="AddStates.asp">Add New State</a>
	</td>
</tr>
<tr>
	<td colspan="4"><hr></td>
</tr>
<tr>
	<td colspan="4">
    	<input type="hidden" name="Count" value="<%=Count%>">
		<input name="submit" type=submit class="submit2" value="Remove unchecked">
		&nbsp;
		<input type="button" name="Button" value="Back" onClick="javascript:history.back()">
		&nbsp;
		<input type="button" name="Restore" value="Restore Default Settings" onClick="javascript:location='restoreStates.asp?action=update'">
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