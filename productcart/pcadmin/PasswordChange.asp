<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% 
pageTitle="Update Master User" 
pageIcon="pcv4_keys.png"
Section="layout" 
%>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
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

	if (theForm.CUID.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.CUID.focus();
		    return (false);
	}
	
	if (allDigit(theForm.CUID.value) == false)
	{
		    alert("Please enter an integer for this field.");
		    theForm.CUID.focus();
		    return (false);
	}	
	
	if (theForm.NUID.value != "")
	{
	if (allDigit(theForm.NUID.value) == false)
	{
		    alert("Please enter an integer for this field.");
		    theForm.NUID.focus();
		    return (false);
	}
	}
	
	if (theForm.Cpass.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.Cpass.focus();
		    return (false);
	}
	
	if (theForm.pass1.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.pass1.focus();
		    return (false);
	}
	
	if (theForm.pass2.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.pass2.focus();
		    return (false);
	}
	
		if (theForm.pass1.value != theForm.pass2.value)
 	{
		    alert("The value entered in the two New Password fields does not match.");
		    theForm.pass2.focus();
		    return (false);
	}
 
return (true);
}
//-->
</script>

<%
dim query, connTemp, rs

IF request("pass1")<>"" THEN
	
	cidAdmin=trim(request("CUID"))
	if trim(session("CUID"))<>cidAdmin then
		response.redirect "PasswordChange.asp?message="&server.URLEncode("An error occurred while updating your administrative User ID and Password. Your current User ID appeared to be incorrect.")
	end if
	nidAdmin=trim(request("NUID"))
	if nidAdmin="" then
		nidAdmin=cidAdmin
	end if
	npassword=enDeCrypt(request("pass2"), scCrypPass)
	cpassword=enDeCrypt(request("Cpass"), scCrypPass)
	
	call opendb()
	set rs=server.createobject("adodb.recordset")
	query="SELECT * FROM admins WHERE idadmin="& cidAdmin &" AND adminpassword='"&cpassword&"' and ID=" & session("IDAdmin")
	set rs=conntemp.execute(query)
	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "PasswordChange.asp?message="&server.URLEncode("An error occurred while updating your administrative User ID and Password. Your current User ID or your current Password appeared to be incorrect.")
	end if

	query="SELECT * FROM admins WHERE idadmin="& nidAdmin &" AND ID<>" & session("IDAdmin")
	set rs=conntemp.execute(query)
	if not rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "PasswordChange.asp?message="&server.URLEncode("User ID was already used in your store")
	end if
	
	query="UPDATE admins SET idadmin="&nidAdmin&", adminpassword='"& npassword &"' WHERE idadmin="& cidAdmin
	set rs=conntemp.execute(query)
	session("CUID")=nidadmin
	set rs=nothing
	call closeDb()
	response.redirect "PasswordChange.asp?s=1&msg="&Server.URLEncode("Your password has been sucessfully changed. Next time you return to your Control Panel log in page, make sure to use the new password.")

ELSE %>
		<form name="form1" action="passwordChange.asp" method="post" onSubmit="return Form1_Validator(this)" class="pcForms">
		<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer">
			<!--#include file="pcv4_showMessage.asp"-->
            </td>
		</tr>
		<tr> 
			<td colspan="2">
				<p>Use this page to update the user name and password for the master administrator. <strong>All fields are required</strong>.</p>
				<ul>
					<li>&quot;User ID&quot; must be a numeric value. </li>
					<li>&quot;Password&quot; may be alphanumeric with a maximum of 20 characters in length.</li>
				</ul>
			</td>
		</tr>
		<tr> 
			<td><div align="right">Current User ID:</div></td>
			<td><input type="text" name="CUID" size="20" maxlength="9" value=<%=session("CUID")%>></td>
		</tr>
		<tr> 
			<td width="32%"><div align="right">New User ID:</div></td>
			<td width="68%"><input type="text" name="NUID" size="20" maxlength="9"></td>
		</tr>
		<tr> 
			<td><div align="right">Current Password:</div></td>
			<td><input type="password" name="Cpass" size="20" maxlength="20"></td>
		</tr>
		<tr> 
			<td><div align="right">New Password:</div></td>
			<td><input type="password" name="pass1" size="20" maxlength="20"></td>
		</tr>
		<tr> 
			<td><div align="right">Confirm New Password:</div></td>
			<td><input type="password" name="pass2" size="20" maxlength="20"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td></td>
            <td>
            	<input name="Submit" type="submit" value="Submit" class="submit2">&nbsp;
                <input name="back" type="button" value="Start Page" onClick="document.location.href='menu.asp'">
            </td>
		</tr>
	</table>
</form>
<%
END IF
%>
<!--#include file="AdminFooter.asp"-->