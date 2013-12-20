<%@ LANGUAGE="VBSCRIPT" %>
<% 'GGG Add-on ONLY FILE %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*2*3*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/GCConstants.asp"-->

<% 
pageTitle = "Manage Gift Certificates"
pageIcon = "pcv4_icon_giftcertificate.png"
Section = "layout" 

Dim connTemp,query,rstemp
call opendb()

if request("GCSettings")<>"" then
	pcGCIncludeShipping = request("GCIncludeShipping") 	
	%>
	<!--#include file="pcAdminSaveGCConstants.asp"-->
	<% response.redirect "ggg_manageGCs.asp?s=1&msg=" & server.URLEncode("Gift Certificate Settings updated successfully.")
end if
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
  	
  	if (theForm.FromDate.value != "")
  	{
		if (check_date(theForm.FromDate) == false)
	  	{
		alert("Please enter a valid date for this field.");
	    theForm.FromDate.focus();
	    return (false);
		}
  	}
  	
  	if (theForm.ToDate.value != "")
  	{
		if (check_date(theForm.ToDate) == false)
	  	{
		alert("Please enter a valid date for this field.");
	    theForm.ToDate.focus();
	    return (false);
		}
  	}  	
  	
return (true);
}

function CheckWindow() {
options = "toolbar=0,status=0,menubar=0,scrollbars=0,resizable=0,width=600,height=400";
myloc='testurl.asp?file1=' + document.hForm.producturl.value + '&file2=' + document.hForm.locallg.value + '&file3=' + document.hForm.remotelg.value;
newcheckwindow=window.open(myloc,"mywindow", options);
}

//-->
</script>
<%
pcGCIncludeShipping=GC_INCSHIPPING

query="select IDProduct,Description from Products where pcprod_GC=1 order by Description"
set rstemp=server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if not rstemp.eof then
	iGiftCerts = 1
end if
%>
<table class="pcCPcontent">
<%
	' Only show this section if the user has "Manage Products" permissions
	if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("PmAdmin")="19") then
%>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <th colspan="2">Gift Certificates Products</th>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
<%
	end if
	
	IF iGiftCerts <> 1 THEN
%>
        <tr>
            <td colspan="2">
                <p>There are currently <strong>no gift certificates</strong> in your store.</p>
                <p>To <strong>add</strong> a new Gift Certificate, <a href="addProduct.asp?prdType=std">add a new product</a> and configure it to be a Gift Certificate using the corresponding settings at the bottom of the page</p>
            </td>
        </tr>
<% ELSE

	' Only show this section if the user has "Manage Products" permissions
	if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("PmAdmin")="19") then
%>
        <tr>
            <td colspan="2">
                <ul class="pcListIcon">
                  <li><strong>Add New</strong><br>
                    To add a new Gift Certificate, <a href="addProduct.asp?prdType=std">add a new product</a> and configure it to be a Gift Certificate using the corresponding settings.</li>
                    <li><strong>Edit</strong><br>
                    To edit an existing Gift Certificate <em>product</em> (not a gift certificate that was purchased in the storefront), <a href="LocateProducts.asp?cptype=0">locate the product</a>.</li>
                </ul>
            </td>
        </tr>
<%
	end if
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Gift Certificates Settings</th>
</tr>
<tr>
	<td colspan="2">
	</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer">
    <form name="hForm" method="post" action="ggg_manageGCs.asp">
        <label><input type="checkbox" name="GCIncludeShipping" id="checkbox" value="1" <% If pcGCIncludeShipping="1" then%>checked<% End If%>>
        When redeeming Gift Certificates, allow the Gift Certificate amount to be used against any shipping charges.</label>
        <input type="hidden" name="GCSettings" value="1">
        <br><br>
        <input name="SubmitRequest" type="submit" id="SubmitRequest" value="Update Settings" class="submit2">
        <br><br>
    </form>
</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Generate &amp; Manage Certificates Codes</th>
</tr>
<tr>
	<td colspan="2">
		<ul class="pcListIcon">
			<li><a href="ggg_AdmGenGCs.asp">Generate</a> one or more Gift Certificate codes.</li>
			<li><a href="ggg_AdmManageGCs.asp">View/Edit</a> Gift Certificate codes that were  generated in the Control Panel (vs. purchased GCs).</li>
		</ul>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">View Purchased Gift Certificates</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">
        <form name="hForm" method="post" action="ggg_srcGCb.asp?action=search" onSubmit="return Form1_Validator(this)" class="pcForms">
        <table class="pcCPcontent">
            <tr>
                <td nowrap="nowrap" width="20%" align="right">Gift Certificate Product:</td>
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
                <td nowrap="nowrap" align="right">Purchase Date:</td>
                <td>From: <input type="text" name="FromDate" size="15"> To: <input type="text" name="ToDate" size="15">&nbsp;(<i>Format: <%if scDateFrmt="DD/MM/YY" then%>DD/MM/YY<%else%>MM/DD/YY<%end if%></i>)</td>
            </tr>
            <tr>
                <td colspan="2"><hr></td>
            </tr>
            <tr>
                <td></td>
                <td>
                    <input type="hidden" name="iPageSize" value="25">
                    <input name="submit1" type="submit" value="Search" class="submit2">&nbsp;
                    <input name="submit2" type="button" value="View All" onclick="location='ggg_srcGCb.asp?iPageSize=99999&submit2=viewall';" class="ibtnGrey">
                </td>
            </tr>
        </table>
        </form>
		</td>
	</tr>
<%END IF 'iGiftCerts = 1%>
</table>
<%
set rstemp=nothing
call closedb()
%><!--#include file="AdminFooter.asp"-->