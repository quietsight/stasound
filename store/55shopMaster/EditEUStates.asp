<% pageTitle = "Edit EU Member State" %>
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

if request("action")="update" then

	StateName=replace(request("StateName"),"'","''")	
	StateName1=replace(request("StateName1"),"'","''")	
	Country=request("Country")
	Country1=request("Country1")
	pcVATCountry_Code=request("pcVATCountry_Code")
	
	'// Check For Duplicates
	query="SELECT * FROM pcVATCountries WHERE pcVATCountry_Code='" & Country & "';"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		set rs=nothing		
		if Country<>Country1 then
			set rs=nothing
			call closeDb()
			response.redirect "EditEUStates.asp?s=0&msg=This EU State Name is already in use in this system&StateCode=" & pcVATCountry_Code
		end if
	end if

	'// Update the States Table
	query="UPDATE pcVATCountries SET pcVATCountry_State='" & StateName & "',pcVATCountry_Code='" & Country & "' WHERE pcVATCountry_Code='" & pcVATCountry_Code & "';"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing		
	call closeDb()
	response.redirect "EditEUStates.asp?s=1&msg=" & server.URLEncode("This EU State was updated successfully!") & "&StateCode=" & Country
	
end if

query="SELECT * FROM pcVATCountries WHERE pcVATCountry_Code='" & request("StateCode") & "';"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 
StateName=rs("pcVATCountry_State")
pcVATCountry_Code=rs("pcVATCountry_Code")
Country=rs("pcVATCountry_Code")
set rs=nothing
call closeDb()

%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{

  if (theForm.StateName.value == "")
  {
    alert("Please enter a value for the EU State Name.");
    theForm.StateName.focus();
    return (false);
  }

if (theForm.Country.value == "")
	{
		alert("Please enter a value for the ISO Country code.");
		theForm.Country.focus();
		return (false);
    }

return (true);
}
//-->
</script>
<form method="post" action="EditEUStates.asp?action=update" onsubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2">
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
			
			Use the form below to edit this EU Member State.
		</td>
	</tr>
	<tr>
		<td>EU Member State Name:</td>
		<td> 
        <input type="text" name="StateName" size="30" value="<%=StateName%>">
        <input type="hidden" name="StateName1" value="<%=StateName%>">
        </td>
	</tr>
	<tr>
		<td>ISO Country  Code:</td>
		<td>
        <input type="text" name="Country" size="30" value="<%=Country%>">
        <input type="hidden" name="Country1" value="<%=Country%>">
		<input type="hidden" name="pcVATCountry_Code" value="<%=pcVATCountry_Code%>">
        </td>	
	</tr>                          
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">
		<input type="submit" name="submit" value="Update" class="submit2"> 
		<input type="button" name="Button" value="Back" onClick="location='manageEUStates.asp';">	
		</td>
	</tr>           
</table>
</form>
<!--#include file="AdminFooter.asp"-->