<% pageTitle = "Add New Country" %>
<% section = "layout" %>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
dim rs, conntemp, query

if request("action")="add" then
	CountryName=trim(request("CountryName"))
		CountryName=replace(CountryName,"'","")
	CountryCode=trim(request("CountryCode"))
		CountryCode=replace(CountryCode,"'","")
	
	' Start - Check if the country already exists
		call openDb()
		query="select * from Countries where CountryName='" & CountryName & "'"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closeDb()
			response.redirect "AddCountries.asp?s=0&msg=This Country Name is already in use in this system"
		end if

		query="select * from Countries where CountryCode='" & CountryCode & "'"
		set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closeDb()
			response.redirect "AddCountries.asp?s=0&msg=This Country Code is already in use in this system"
		end if
	' End - Check if the country already exists


	query="insert into Countries (CountryName,CountryCode) values ('" & CountryName & "','" & CountryCode & "')"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
	response.redirect "AddCountries.asp?s=1&msg=New Country was added successfully!"
end if 

%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{

  if (theForm.CountryName.value == "")
  {
    alert("Please enter a name for this country.");
    theForm.CountryName.focus();
    return (false);
  }
if (theForm.CountryName.value.indexOf("'")>0)
	{
		alert("Please do not enter apostrophes in the country name.");
		theForm.CountryName.focus();
		return (false);
    }
if (theForm.CountryCode.value.indexOf("'")>0)
	{
		alert("Please do not enter apostrophes in the country code.");
		theForm.CountryCode.focus();
		return (false);
    }
if (theForm.CountryCode.value == "")
	{
		alert("Please enter a country code.");
		theForm.CountryCode.focus();
		return (false);
    }
return (true);
}
//-->
</script>
<form method="post" action="AddCountries.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
		<td>
		Country Name: <input type="text" name="CountryName" size="30">
		</td>
	</tr>
	<tr>
		<td>
		Country Code: <input type="text" name="CountryCode" size="30">
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		<input type="submit" name="submit" value="Add New Country" class="submit2"> 
		<input type="button" name="Button" value="Back" onClick="location='manageCountries.asp';">
		</td>
	</tr>             
</table>
</form>
<!--#include file="AdminFooter.asp"-->