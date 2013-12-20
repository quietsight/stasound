<% pageTitle = "Manage VAT - Add VAT Category" %>
<% section = "layout" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
dim rs, conntemp, query

call openDb()

pcv_intVATID=Request("VATID")

if request("action")="add" then

	CategoryName=replace(request("CategoryName"),"'","''")		
	Rate=request("Rate")
	Country=request("Country")
	
	query="SELECT * FROM pcVATRates WHERE pcVATRate_Category='" & CategoryName & "' AND pcVATRate_Rate=" & Rate & " AND pcVATCountry_Code='" & Country & "';"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		set rs=nothing
		call closeDb()
		response.redirect "AddVATCategory.asp?s=0&msg=This VAT Category already exists in this system"
	end if


	query="INSERT INTO pcVATRates (pcVATRate_Category, pcVATRate_Rate, pcVATCountry_Code) VALUES ('" & CategoryName & "'," & Rate & ",'" & Country & "')"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	call closeDb()		
	response.redirect "AddVATCategory.asp?s=1&msg=New VAT Category was added successfully!"

end if 
%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{

  if (theForm.CategoryName.value == "")
  {
    alert("Please enter a value for the VAT Category name.");
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
<form method="post" action="AddVATCategory.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2">
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Add VAT Category</th>
	</tr>	
	<tr>
		<td colspan="2">
			Use the form below to add a new VAT Category.
			<ul>
				<li>Example 1.) United Kingdom (GB) - Exempt - 0.00%</li>
				<li>Example 2.) United Kingdom (GB) - Reduced - 8.00%</li>
			</ul>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
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
					if UCASE(ptaxVATRate_Code)=UCASE(pcArr(2,i)) then
						ttaxVATRate_State=pcArr(1,i)
					end if
					%>
					<option value="<%=pcArr(2,i)%>" <%if UCASE(ptaxVATRate_Code)=UCASE(pcArr(2,i)) then response.write "selected"%>><%=pcArr(1,i) & " (" & pcArr(2,i) & ") "%></option>
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
		<td><input type="text" name="CategoryName" size="30"></td>
	</tr>
	<tr>
		<td>Rate:</td>
		<td>
		<input type="text" name="Rate" size="30"> %
		<input type="hidden" name="VATID" value="<%=pcv_intVATID%>">
		</td>
	</tr>

	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">
			<input type="submit" name="submit" value="  Save  " class="submit2">			
			<% if ptaxVATRate_Code<>"" then %>
			<input type="button" name="ManageVATCategories" value="Manage VAT Categories" onclick="location='viewVAT.asp';" class="ibtnGrey">&nbsp;
			<% end if %>
			<input type="button" name="Button" value="Manage VAT Settings" onClick="location='AdminTaxSettings_VAT.asp';">
			<input type="button" name="Button" value="Back" onClick="JavaScript:history.back()">
		</td>
	</tr>       
</table>
</form>
<!--#include file="AdminFooter.asp"-->