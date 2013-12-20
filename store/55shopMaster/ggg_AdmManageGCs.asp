<%@ LANGUAGE="VBSCRIPT" %>
<% 'GGG Add-on ONLY FILE %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="2*3*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<% pageTitle = "Manage Generated Gift Certificate Codes" %>
<% Section = "products" %>
<%Dim connTemp,query,rstemp
call opendb()
%>
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--

function check_date(field){
var checkstr = "0123456789";
var DateField = field;
var Datevalue = "";
var DateTemp = "";
var seperator = "/";
var day;
var month;
var year;
var leap = 0;
var err = 0;
var i;
   err = 0;
   DateValue = DateField.value;
   /* Delete all chars except 0..9 */
   for (i = 0; i < DateValue.length; i++) {
	  if (checkstr.indexOf(DateValue.substr(i,1)) >= 0) {
	     DateTemp = DateTemp + DateValue.substr(i,1);
	  }
	  else
	  {
	  if (DateTemp.length == 1)
		{
    	  DateTemp = "0" + DateTemp
		}
	  else
	  {
	  	if (DateTemp.length == 3)
	  	{
	  	DateTemp = DateTemp.substr(0,2) + '0' + DateTemp.substr(2,1);
	  	}
	  }
	 }
   }
   DateValue = DateTemp;
   /* Always change date to 8 digits - string*/
   /* if year is entered as 2-digit / always assume 20xx */
   if (DateValue.length == 6) {
      DateValue = DateValue.substr(0,4) + '20' + DateValue.substr(4,2); }
   if (DateValue.length != 8) {
      return(false);}
   /* year is wrong if year = 0000 */
   year = DateValue.substr(4,4);
   if (year == 0) {
      err = 20;
   }
   /* Validation of month*/
   <%if scDateFrmt="DD/MM/YY" then%>
   month = DateValue.substr(2,2);
   <%else%>
   month = DateValue.substr(0,2);
   <%end if%>
   if ((month < 1) || (month > 12)) {
      err = 21;
   }
   /* Validation of day*/
   <%if scDateFrmt="DD/MM/YY" then%>
   day = DateValue.substr(0,2);
   <%else%>
   day = DateValue.substr(2,2);
   <%end if%>
   if (day < 1) {
     err = 22;
   }
   /* Validation leap-year / february / day */
   if ((year % 4 == 0) || (year % 100 == 0) || (year % 400 == 0)) {
      leap = 1;
   }
   if ((month == 2) && (leap == 1) && (day > 29)) {
      err = 23;
   }
   if ((month == 2) && (leap != 1) && (day > 28)) {
      err = 24;
   }
   /* Validation of other months */
   if ((day > 31) && ((month == "01") || (month == "03") || (month == "05") || (month == "07") || (month == "08") || (month == "10") || (month == "12"))) {
      err = 25;
   }
   if ((day > 30) && ((month == "04") || (month == "06") || (month == "09") || (month == "11"))) {
      err = 26;
   }
   /* if 00 ist entered, no error, deleting the entry */
   if ((day == 0) && (month == 0) && (year == 00)) {
      err = 0; day = ""; month = ""; year = ""; seperator = "";
   }
   /* if no error, write the completed date to Input-Field (e.g. 13.12.2001) */
   if (err == 0) {
	<%if scDateFrmt="DD/MM/YY" then%>
	DateField.value = day + seperator + month + seperator + year;
    <%else%>
	DateField.value = month + seperator + day + seperator + year;   
    <%end if%>
	return(true);
   }
   /* Error-message if err != 0 */
   else {
	return(false);   
   }
}

function isDigit(s)
{
var test=""+s;
if(test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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

	
function Form1_Validator(theForm)
{
	
	if (theForm.ExpDate.value != "")
  	{
		if (check_date(theForm.ExpDate) == false)
	  	{
		alert("Please enter a valid date for this field.");
	    theForm.ExpDate.focus();
	    return (false);
		}
  	}
  	
  	return (true);
}

//-->
</script>
<%
query="select IDProduct,Description from Products where pcprod_GC=1 order by Description"
set rstemp=server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if not rstemp.eof then
	iGiftCerts = 1
end if
%>
<form name="hForm" method="post" action="ggg_AdmSrcGCb.asp?action=search" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<%IF iGiftCerts <> 1 THEN%>
<tr>
	<td colspan="2">
		<p>There are currently <strong>no gift certificates</strong> in your store.</p>
		<p>To <strong>add</strong> a new Gift Certificate, <a href="addProduct.asp?prdType=std">add a new product</a> and configure it to be a Gift Certificate.</p>
	</td>
</tr>
<% ELSE %>
<tr>
	<td colspan="2">
		<ul>
		<li>To <strong>add</strong> a new Gift Certificate, <a href="addProduct.asp?prdType=std">add a new product</a> and use the Gift Certificate settings section.</li>
		<li>To <strong>edit</strong> an existing Gift Certificate, use the <a href="LocateProducts.asp?cptype=0">product search</a> feature.</li>
		<li>To <strong>generate</strong> a new Gift Certificate code, use the <a href="ggg_AdmGenGCs.asp">Gift Certificate code generator</a>.</li>
		<li>To view/search Gift Certificates <strong>puchased by customers</strong>, <a href="ggg_manageGCs.asp">click here</a>.</li>
		</ul>
	</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td colspan="2">
	<p><strong>Search previously generated Gift Certificates codes</strong>.</p>
	<p>Here you are viewing Gift Certificate codes generated by the store manager.</p>
	</td>
</tr>
<tr>
	<td nowrap="nowrap" width="20%" align="right">Gift Certificate:</td>
	<td width="80%">
		<select name="IDGC">
			<option value="0" selected>All Gift Certificate Products</option>
			<%do while not rstemp.eof%>
				<option value="<%=rstemp("IDProduct")%>"><%=rstemp("Description")%></option>
				<%rstemp.MoveNext
			loop
			set rstemp=nothing%>
		</select>
	</td>
</tr>
<tr>
	<td nowrap="nowrap" align="right">Certificate Code:</td>
	<td><input type="text" name="GCCode" size="30"></td></tr>
<tr>
	<td nowrap="nowrap" align="right">Expiration Date:</td>
	<td><input type="text" name="ExpDate" size="30">&nbsp;(<i>Format: <%if scDateFrmt="DD/MM/YY" then%>DD/MM/YY<%else%>MM/DD/YY<%end if%></i>)</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td></td>
	<td>
		<input type="hidden" name="iPageSize" value="25">
		<input name="submit1" type="submit" value="Search" class="submit2">&nbsp;
		<input name="submit2" type="button" value="View All" onclick="location='ggg_AdmSrcGCb.asp?iPageSize=99999&submit2=viewall';">
	</td>
</tr>
<%END IF 'iGiftCerts = 1%>
</table>
</form>
<%
set rstemp=nothing
call closedb()
%><!--#include file="AdminFooter.asp"-->