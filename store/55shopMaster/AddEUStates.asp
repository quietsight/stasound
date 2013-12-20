<% pageTitle = "Add New EU State" %>
<% section = "layout" %>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

dim rs, conntemp, query

call openDb()

	if request("action")="add" then

		StateName=replace(request("StateName"),"'","''")		
		Country=request("Country")
		
		query="SELECT * FROM pcVATCountries WHERE pcVATCountry_State='" & StateName & "' AND pcVATCountry_Code='" & Country & "';"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closeDb()
			response.redirect "AddEUStates.asp?s=0&msg=" & Server.URLEncode("This EU Member State already exists in this system")
		end if


		query="INSERT INTO pcVATCountries (pcVATCountry_State, pcVATCountry_Code) VALUES ('" & StateName & "','" & Country & "')"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		call closeDb()		
		response.redirect "AddEUStates.asp?s=1&msg=" & Server.URLEncode("New EU Member State was added successfully!")
	
	end if 

%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{

  if (theForm.StateName.value == "")
  {
    alert("Please enter a value for the EU Member State name.");
    theForm.StateName.focus();
    return (false);
  }

if (theForm.Country.value == "")
	{
		alert("Please enter a value for the ISO Country Code.");
		theForm.Country.focus();
		return (false);
    }

return (true);
}
//-->
</script>
<form method="post" action="AddEUStates.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2">
        
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
		
		Use the form below to add a new EU Member State. <u>NOTE</u>: If you are re-entering a <strong>EU Member State</strong> that you had previously deleted, make sure to use the <a href="http://publications.europa.eu/code/pdf/370000en.htm#pays" target="_blank">official abbreviations</a>.
		</td>
	</tr>
	<tr>
		<td width="20%">State Name:</td><td width="80%"> <input type="text" name="StateName" size="30"></td>
	</tr>
	<tr>
		<td nowrap>ISO Country Code:</td><td> <input type="text" name="Country" size="30"></td>
	</tr>

	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">
			<input type="submit" name="submit" value="Add State" class="submit2">
			<input type="button" name="Button" value="Back" onClick="location='manageEUStates.asp';">
		</td>
	</tr>       
</table>
</form>
<!--#include file="AdminFooter.asp"-->