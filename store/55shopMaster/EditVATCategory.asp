<% pageTitle = "Manage VAT - Modify Category" %>
<% section = "layout" %>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
dim rs, conntemp, query

call openDb()

pcv_intVATID=Request("VATID")

if request("action")="update" then

	CategoryName=replace(request("CategoryName"),"'","''")		
	Rate=request("Rate")
	Country=request("Country")
	CategoryName1=replace(request("CategoryName1"),"'","''")		
	Rate1=request("Rate1")
	Country1=request("Country1")
	
	'// Check For Duplicates
	query="SELECT * FROM pcVATRates WHERE pcVATRate_Category='" & CategoryName & "' AND pcVATRate_Rate=" & Rate & " AND pcVATCountry_Code='" & Country & "';"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		set rs=nothing	
		if Rate<>Rate1 OR CategoryName<>CategoryName1 OR Country<>Country1 then
			call closeDb()
			response.redirect "EditVATCategory.asp?s=0&msg=This VAT Category Name is already in use in this system&VATID=" & pcv_intVATID
		end if
	end if

	query="UPDATE pcVATRates SET pcVATRate_Category='" & CategoryName & "', pcVATRate_Rate=" & Rate & ", pcVATCountry_Code='" & Country & "' WHERE pcVATRate_ID=" & pcv_intVATID & ";"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing		
	call closeDb()
	response.redirect "EditVATCategory.asp?s=1&msg=" & server.URLEncode("This VAT Category was updated successfully!") & "&VATID=" & pcv_intVATID
	
end if

query="SELECT * FROM pcVATRates WHERE pcVATRates.pcVATRate_ID=" & pcv_intVATID & ";"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 
CategoryName=rs("pcVATRate_Category")
Rate=rs("pcVATRate_Rate")
VATCountryCode=rs("pcVATCountry_Code")
set rs=nothing
call closeDb()

%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{

  if (theForm.CategoryName.value == "")
  {
    alert("Please enter a value for the VAT Category Name.");
    theForm.CategoryName.focus();
    return (false);
  }

if (theForm.Rate.value == "")
	{
		alert("Please enter a value for the Rate.");
		theForm.Rate.focus();
		return (false);
    }

return (true);
}
//-->
</script>
<form method="post" action="EditVATCategory.asp?action=update" onsubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2">
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
			
			Use the form below to edit the VAT Category.
		</td>
	</tr>
	<tr>
		<td>EU Member State:</td>
		<td>
			<%
			call openDB()
			ttaxVATRate_State=""
			query="SELECT pcVATCountries.pcVATCountry_ID, pcVATCountries.pcVATCountry_State, pcVATCountries.pcVATCountry_Code From pcVATCountries Order By pcVATCountry_State ASC;"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			%>
			<select name="Country">
			<option value="">Select an option.</option>
			<%
			if not rs.eof then
				pcArr=rs.getRows()
				set rs=nothing
				intCount=ubound(pcArr,2)
				For i=0 to intCount
					if UCASE(VATCountryCode)=UCASE(pcArr(2,i)) then
						Country=pcArr(2,i)
					end if
					%>
					<option value="<%=pcArr(2,i)%>" <%if UCASE(VATCountryCode)=UCASE(pcArr(2,i)) then response.write "selected"%>><%=pcArr(1,i) & " (" & pcArr(2,i) & ") "%></option>
				<%Next
			end if
			set rs = nothing
			call closeDB()
			%>
			</select>&nbsp;&nbsp;&nbsp;<input type="button" name="Update" value="Manage EU States" onclick="location='ManageEUStates.asp'">
			<div><i>Note: This is the country in which the store is located.</i></div>		
		</td>
	</tr>
	<tr>
		<td>VAT Category Name:</td>
		<td><input type="text" name="CategoryName" size="30" value="<%=CategoryName%>"></td>
	</tr>
	<input type="hidden" name="CategoryName1" value="<%=CategoryName%>">
	<tr>
		<td>Rate:</td>
		<td><input type="text" name="Rate" size="30" value="<%=Rate%>">
		%</td>	
	</tr>                          
	<input type="hidden" name="Rate1" value="<%=Rate%>">
	<input type="hidden" name="Country1" value="<%=Country%>">
	<input type="hidden" name="VATID" value="<%=pcv_intVATID%>">
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">
			<input type="submit" name="submit" value="Update" class="submit2"> 
			<% if ptaxVATRate_Code<>"" then %>
			<input type="button" name="ManageVATCategories" value="Manage VAT Categories" onclick="document.location.href='viewVAT.asp';" class="ibtnGrey">&nbsp;
			<% end if %>
			<input type="button" name="Button" value="Manage VAT Settings" onClick="document.location.href='AdminTaxSettings_VAT.asp';">
			<input type="button" name="Button" value="Add/Remove Products" onClick="document.location.href='manageVATCategories.asp?VATID=<%=pcv_intVATID%>';">	
		</td>
	</tr>           
</table>
</form>
<!--#include file="AdminFooter.asp"-->