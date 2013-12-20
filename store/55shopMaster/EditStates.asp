<% pageTitle = "Edit State" %>
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

StateName=trim(request("StateName"))
	StateName=replace(StateName,"'","")
StateCode=trim(request("StateCode"))
	StateCode=replace(StateCode,"'","")
StateName1=trim(request("StateName1"))
StateCode1=trim(request("StateCode1"))
Country=trim(request("Country"))
	if Country="" then
		Country=trim(request("CountryCode"))
	end if
Country1=request("Country1")

if request("action")="update" then
	
	call openDb()
	
	'// Check For Duplicates
	query="select * from States where StateName='" & StateName & "' AND StateCode='" & StateCode & "' AND pcCountryCode='" & Country & "' "
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		set rs=nothing
		if Country<>Country1 OR StateName<>StateName1 OR StateCode<>StateCode1 then
			call closeDb()
			response.redirect "EditStates.asp?s=0&msg=This State Name is already in use in the store database&StateCode=" & StateCode1
		end if
	end if

	'// Update the States Table
	query="update States set StateName='" & StateName & "',StateCode='" & StateCode & "',pcCountryCode='" & Country & "' where StateName='" & StateName1 & "' and StateCode='" & StateCode1 & "' and pcCountryCode='" & Country1 & "' "
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing	
	
	'// Update the Country Table Relationship
	query="update countries set pcSubDivisionID=1 where countryCode='" & Country & "' "
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing

	if ucase(Country)<>ucase(Country1) then
		'// Check if this country is still a DropDown
		query="select pcCountryCode from States where pcCountryCode='" & Country1 & "' "
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if rs.eof then
		set rs=nothing
			query="update countries set pcSubDivisionID=2 where countryCode='" & Country1 & "' "
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	end if
	
	call closeDb()
	response.redirect "EditStates.asp?s=1&msg=This State was updated successfully!&StateCode=" & StateCode &"&Country=" & Country
	
end if

call openDb()
query="select * from States where StateCode='" & StateCode & "' and pcCountryCode ='" & Country &"'"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 
if rs.eof then
	set rs=nothing
	call closeDb()
	response.Redirect("manageStates.asp?s=0&message=The selected State Code was not found in the database.")
end if
StateName=rs("StateName")
StateCode=rs("StateCode")
Country=rs("pcCountryCode")
set rs=nothing
call closeDb()

%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{

  if (theForm.StateName.value == "")
  {
    alert("Please enter a value for the state name.");
    theForm.StateName.focus();
    return (false);
  }

if (theForm.StateCode.value == "")
	{
		alert("Please enter a value for the state code.");
		theForm.StateCode.focus();
		return (false);
    }
if (theForm.StateName.value.indexOf("'")>0)
	{
		alert("Please do not enter apostrophes in the state name.");
		theForm.StateName.focus();
		return (false);
    }
if (theForm.StateCode.value.indexOf("'")>0)
	{
		alert("Please do not enter apostrophes in the state code.");
		theForm.StateCode.focus();
		return (false);
    }
if (theForm.Country.value == "")
	{
		alert("Please select a Country.");
		theForm.Country.focus();
		return (false);
    }
return (true);
}
//-->
</script>
<form method="post" action="EditStates.asp?action=update" onsubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" name="StateName1" value="<%=StateName%>">
<input type="hidden" name="StateCode1" value="<%=StateCode%>">
<input type="hidden" name="Country1" value="<%=Country%>">
<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
		<td width="20%" nowrap>State Name:</td>
        <td><input type="text" name="StateName" size="30" value="<%=StateName%>"></td>
	</tr>
	<tr>
		<td>State Code:</td>
        <td><input type="text" name="StateCode" size="30" value="<%=StateCode%>"></td>	
	</tr>                          
	<tr>
		<td>Country:</td>
		<td>
		<% 
		call opendb()
		query="SELECT CountryCode,countryName FROM countries WHERE CountryCode<>'US' ORDER BY countryName ASC;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)									
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		%>
		<select name="Country">
			<option value="US" selected>UNITED STATES</option>
			<% 
			dim pcTempCountryCode
			do while not rs.eof
				pcTempCountryCode=rs("CountryCode")%>
				<option value="<%=pcTempCountryCode%>"<%
				if Country=pcTempCountryCode then
					response.write "selected"           
				end if
				%>><%=rs("countryName")%></option>
				<%
				rs.movenext
			loop
			set rs=nothing
			call closedb() 
			%>
		</select>
		</td>	
	</tr>                          
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">
		<input type="submit" name="submit" value="Update" class="submit2"> 
		<input type="button" name="Button" value="Manage States" onClick="location='manageStates.asp';">	
		</td>
	</tr>           
</table>
</form>
<!--#include file="AdminFooter.asp"-->