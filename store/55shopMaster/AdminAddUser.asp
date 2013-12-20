<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Users - Add New Control Panel User" %>
<% section="layout" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp"-->
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<%
Dim rs, rstemp, connTemp, query, pcvAdminName, pcvAdminEmail

' ADD NEW USER - START
if request("action")="add" then

	AdminUser=request("AdminUser")
	if not validNum(AdminUser) then
		response.redirect "AdminAddUser.asp?r=1&msg=" & Server.Urlencode("The user cannot be added to the database.")
	end if
	
	'Check if the user already exists
	call openDB()

	query="SELECT * FROM Admins WHERE IDAdmin=" & AdminUser
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	
	if not rstemp.eof then
		set rstemp=nothing
		call closeDB()
		response.redirect "AdminAddUser.asp?r=1&msg=" & Server.Urlencode("This User ID is already in use in this store")
	end if
	
	password=request("AdminPassword")
	password=enDeCrypt(password, scCrypPass)

	pcvAdminName = request("adminName")
	pcvAdminName = pcf_ReplaceCharacters(pcvAdminName)
	pcvAdminEmail = request("adminEmail")
	pcvAdminEmail = pcf_ReplaceCharacters(pcvAdminEmail)

	Count=request("Count")
	Permissions=""
	For i=1 to Count
		if request("C" & i)="1" then
			Permissions=Permissions & request("ID" & i) & "*"
		end if
	Next
	if Permissions="" then
		set rstemp=nothing
		call closeDB()
		response.redirect "AdminAddUser.asp?r=1&msg=" & Server.Urlencode("You must give access to at least one area of the Control Panel")
	end if
	
	query="INSERT INTO Admins (IDadmin,AdminName,AdminPassword,AdminLevel,adm_ContactName,adm_ContactEmail) values (" & AdminUser & ",'sub-admin','" & password & "','" & permissions & "','" & pcvAdminName & "','" & pcvAdminEmail & "')"
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
	call closeDB()
	response.redirect "AdminUserManager.asp?s=1&msg=" & Server.Urlencode("The new user was added successfully!")
end if
' ADD NEW USER - END
%>

<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--

function isDigit(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}

function testLen(s)
	{
		var test=""+s ;
		if (test.length<5)
		{
				return (false);
		}
		return (true);
	}
	
function Form1_Validator(theForm)
{

	if (theForm.AdminUser.value == "")
 	{
		    alert("Please enter a value for the User Name. It must be a number and it must contain a minimum of 5 digits.");
		    theForm.AdminUser.focus();
		    return (false);
	}
	else
	{
	if (testLen(theForm.AdminUser.value) == false)
 	{
		    alert("The User Name must contain at least 5 numbers.");
		    theForm.AdminUser.focus();
		    return (false);
	}
	}
	
	if (allDigit(theForm.AdminUser.value) == false)
			{
		    alert("The User Name must be numeric.");
		    theForm.AdminUser.focus();
		    return (false);
		    }	
	
	if (theForm.AdminPassword.value == "")
 	{
		    alert("Please enter a Password");
		    theForm.AdminPassword.focus();
		    return (false);
	}
	else
	{
	if (testLen(theForm.AdminPassword.value) == false)
 	{
		    alert("The Password must contain at least 5 characters.");
		    theForm.AdminPassword.focus();
		    return (false);
	}
	}
	
	if (theForm.C11.checked == true && theForm.C12.checked == true)
	{
		    alert("Please select only one of the two Manage Pages permissions.");
		    theForm.C11.focus();
		    return (false);
	}	
  
return (true);
}
//-->
</script>

<form name="addnew" method="post" action="AdminAddUser.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
		<td colspan="2">
		Use this feature to create new store managers that have limited access to the areas of the Control Panel that you select. For details, <a href="http://wiki.earlyimpact.com/productcart/settings-manage-users" target="_blank">see the ProductCart documentation</a>. </td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>                   
		<td width="20%" align="right" nowrap>User ID:</td>
		<td width="80%"><input type="text" name="AdminUser" size="20" maxlength="9">&nbsp;&nbsp;<i>Must be numeric, at least 5 numbers.</i></td>
	</tr>
	<tr>          
		<td align="right">Password:</td>
		<td><input type="text" name="AdminPassword" size="20" maxlength="20">&nbsp;&nbsp;<i>Must be at least 5 characters.</i></td>
	</tr>
	<tr> 
		<td width="20%" align="right" nowrap>Contact Name:</td>
		<td width="80%"><input name="AdminName" type="text" value="" size="30"> (<em>optional</em>)</td>
	</tr>
	<tr> 
		<td width="20%" align="right" nowrap>Contact Email:</td>
		<td width="80%"><input name="AdminEmail" type="text" value="" size="30"> (<em>optional</em>)</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>  
	<tr>       
		<td align="right" valign="top">Permissions:</td>
		<td valign="top">
				<%
				call openDb()
				query="SELECT * FROM Permissions ORDER BY IDPm"
				set rstemp=Server.CreateObject("ADODB.Recordset")
				set rstemp=connTemp.execute(query)
				Count=0
				do while not rstemp.eof
				Count=Count+1
				%>
				<input type="hidden" name="ID<%=Count%>" value="<%=rstemp("IDPM")%>">
				<input type="checkbox" name="C<%=Count%>" value="1" class="clearBorder">&nbsp;
				<%=rstemp("PMName")%><br>
				<%
				rstemp.MoveNext
				loop
				set rstemp = nothing
				call closeDb()
				%>
				<input type="hidden" name="Count" value="<%=Count%>">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>      
		<td>
		<input name="submit" type="submit" class="submit2" value="Add New">
		&nbsp;
		<input name="back" type="button" onClick="javascript:history.back()" value="Back"> 
		</td>
	</tr>     
</table>
</form>
<!--#include file="AdminFooter.asp"-->