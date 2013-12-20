<% pageTitle = "Edit Country" %>
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

CountryName=trim(request("CountryName"))
	CountryName=replace(CountryName,"'","")
CountryCode=trim(request("CountryCode"))
	CountryCode=replace(CountryCode,"'","")
CountryName1=trim(request("CountryName1"))
CountryCode1=trim(request("CountryCode1"))

if request("action")="update" then

		call openDb()

		if ucase(CountryName)<>ucase(CountryName1) then
			query="select * from Countries where CountryName='" & CountryName & "'"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			if not rs.eof then
				set rs=nothing
				call closeDb()
				response.redirect "EditCountries.asp?s=0&msg=This Country Name is already in use in this system&CountryCode=" & CountryCode1
				end if
			end if

		if ucase(CountryCode)<>ucase(CountryCode1) then
			query="select * from Countries where CountryCode='" & CountryCode & "'"
			set rs=connTemp.execute(query)
			if not rs.eof then
				set rs=nothing
				call closeDb()
				response.redirect "EditCountries.asp?s=0&msg=This Country Code is already in use in this system&CountryCode=" & CountryCode1
			end if
		end if

	query="update Countries set CountryName='" & CountryName & "',CountryCode='" & CountryCode & "' where CountryName='" & CountryName1 & "' and CountryCode='" & CountryCode1 & "'"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
	response.redirect "EditCountries.asp?s=1&msg=This Country was updated successfully!&CountryCode=" & CountryCode
end if

call openDb()
query="select * from Countries where CountryCode='" & request("CountryCode") & "'"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 
CountryName=rs("CountryName")
CountryCode=rs("CountryCode")
set rs=nothing
call closeDb()
%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{

  if (theForm.CountryName.value == "")
  {
    alert("Please enter a value for this field.");
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
		alert("Please enter a value for this field.");
		theForm.CountryCode.focus();
		return (false);
    }
return (true);
}
//-->
</script>
<form method="post" action="EditCountries.asp?action=update" onsubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
		<td>Country Name: <input type="text" name="CountryName" size="30" value="<%=CountryName%>"></td>
	</tr>
	<input type="hidden" name="CountryName1" value="<%=CountryName%>">
	<tr>
		<td>Country Code: <input type="text" name="CountryCode" size="30" value="<%=CountryCode%>"></td>
	</tr>
	<input type="hidden" name="CountryCode1" value="<%=CountryCode%>">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<input type="submit" name="submit" value="Update" class="submit2">		  
			<input type="button" name="Button" value="Back" onClick="location='manageCountries.asp';">
		</td>
	</tr>          
</table>  
</form>
<!--#include file="AdminFooter.asp"-->