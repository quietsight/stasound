<% pageTitle = "Add New State" %>
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

		StateName=trim(request("StateName"))
			StateName=replace(StateName,"'","")
		StateCode=trim(request("StateCode"))
			StateCode=replace(StateCode,"'","")
		Country=trim(request("Country"))
		
		call openDb()
		query="SELECT * FROM States WHERE StateCode='" & StateCode & "' AND pcCountryCode='" & Country & "' "
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closeDb()
			response.redirect "AddStates.asp?s=0&msg=" &  Server.URLEncode("This State Code already exists in this system.")
		end if

		query="INSERT INTO States (StateName,StateCode,pcCountryCode) values ('" & StateName & "','" & StateCode & "','" & Country & "')"
		set rs=connTemp.execute(query)
		
		'// Update the Country Table Relationship
		query="UPDATE countries SET pcSubDivisionID=1 WHERE countryCode='" & Country & "' "
		set rs=connTemp.execute(query)
		set rs=nothing
		
		call closeDb()		
		response.redirect "AddStates.asp?s=1&msg=" & Server.URLEncode("New State was added successfully!")
	
	end if 

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
<form method="post" action="AddStates.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent" style="width:auto;">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
		<td colspan="2">		
		Use the form below to add a new State or Province.		
		<ul>
            <li>If you are re-entering a <strong>US state</strong> that you had previously deleted, make sure to use the <a href="http://www.usps.com/ncsc/lookups/abbreviations.html#states" target="_blank">official abbreviations</a>.</li>
            <li>If you are re-entering a <strong>Canadian province</strong> that you had previously deleted, make sure to use the <a href="http://canadaonline.about.com/library/bl/blpabb.htm" target="_blank">official abbreviations</a>.</li>
		</ul>
		</td>
	</tr>
	<tr>
		<td width="20%" nowrap>State Name:</td>
        <td width="80%"><input type="text" name="StateName" size="30"></td>
	</tr>
	<tr>
		<td nowrap>State Code:</td>
        <td><input type="text" name="StateCode" size="30"></td>
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
			<input type="submit" name="submit" value="Add State" class="submit2">&nbsp;
			<input type="button" name="Button" value="Manage States" onClick="location='manageStates.asp';">
		</td>
	</tr>  
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>     
</table>
</form>
<!--#include file="AdminFooter.asp"-->