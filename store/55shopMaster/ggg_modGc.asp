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
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<% pageTitle = "View/Modify Gift Certificate" %>
<% Section = "products" %>
<%Dim connTemp,query,rstemp

gcCode=request("GcCode")
IDProduct=request("IDProduct")

call opendb()

if request("submit2")<>"" then
	query="update pcGCOrdered set pcGO_Status=2,pcGO_Amount=0 where pcGO_gcCode='" & gcCode & "'"
	set rstemp=connTemp.Execute(query)
	set rstemp=nothing
	call closedb()
	response.redirect "ggg_manageGCs.asp"
else
	if request("action")="update" then
		gStatus=request("Status")
		if gStatus="" then
			gStatus="0"
		end if
		gExpDate=request("ExpDate")
		if gExpDate<>"" then
		else
			gExpDate="01/01/1900"
		end if
		gAmount=request("Amount")
		if gAmount<>"" then
		else
			gAmount="0"
		end if
		if gAmount<=0 then
			gStatus="0"
		end if
		if SQL_Format="1" then
			gExpDate=(day(gExpDate)&"/"&month(gExpDate)&"/"&year(gExpDate))
		else
			gExpDate=(month(gExpDate)&"/"&day(gExpDate)&"/"&year(gExpDate))
		end if
		if scDB="SQL" then
			query="Update pcGCOrdered set pcGO_Status=" & gStatus & ",pcGO_Amount=" & gAmount & ",pcGO_ExpDate='" & gExpDate & "' where pcGO_GcCode='" & gcCode & "'"
		else
			query="Update pcGCOrdered set pcGO_Status=" & gStatus & ",pcGO_Amount=" & gAmount & ",pcGO_ExpDate=#" & gExpDate & "# where pcGO_GcCode='" & gcCode & "'"
		end if
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		msg="This gift certificate has been updated successfully!"
		msgtype=1
	end if 'Update
end if %>

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
  	
  	if (allDigit(theForm.Amount.value) == false)
	{
    alert("Please enter a number for this field.");
    theForm.Amount.focus();
    return (false);
	}
	
	if ((theForm.Amount.value != "0") && (theForm.subdel.value == "1"))
  	{
    alert("You can delete a gift certificate only if the Available Amount is 0.");
    return (false);
  	}

	if (theForm.subdel.value == "1")
  	{
    return (confirm('You are about to remove this gift certificate from your database. Are you sure you want to complete this action?'));
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
query="select Products.Description,pcGCOrdered.pcGO_ExpDate,pcGCOrdered.pcGO_Amount,pcGCOrdered.pcGO_Status from Products,pcGCOrdered where products.pcprod_GC=1 and Products.IDProduct=" & request("IDProduct") & " and pcGCOrdered.pcGO_IDProduct=products.idproduct and pcGCOrdered.pcGO_GcCode='" & request("GcCode") & "'"
set rstemp=connTemp.execute(query)

if not rstemp.eof then
	pName=rstemp("Description")
	gExpDate=rstemp("pcGO_ExpDate")
	gExpDate=ShowDateFrmt(gExpDate)
	if year(gExpDate)="1900" then
		gExpDate=""
	end if
	gAmount=rstemp("pcGO_Amount")
	gStatus=rstemp("pcGO_Status")
	if gStatus<>"" then
	else
		gStatus="0"
	end if
end if
set rstemp=nothing%>
<form name="hForm" method="post" action="ggg_modGC.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer">
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
	</td>
</tr>

<tr>
	<td nowrap="nowrap" width="20%">Gift Certificate Name:</td>
	<td width="80%"><%=pName%></td>
</tr>
<tr>
	<td>Active:</td>
	<td><input type="checkbox" name="Status" value="1" <%if gStatus="1" then%>checked<%end if%> class="clearBorder"></td>
</tr>
<tr>
	<td nowrap="nowrap">Gift Code:</td>
	<td><b><%=request("GcCode")%></b></td>
</tr>
<tr>
	<td>Expiring on:</td>
	<td><input type="text" name="ExpDate" size="15" value="<%=gExpDate%>">&nbsp;(<i>Format: <%if scDateFrmt="DD/MM/YY" then%>DD/MM/YY<%else%>MM/DD/YY<%end if%></i>)</td>
</tr>
<tr>
	<td>Available Amount:</td>
	<td><input type="text" name="Amount" size="15" value="<%=gAmount%>"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2" align="center">
		<input type="hidden" name="subdel" value="0">
		<input type="hidden" name="GcCode" value="<%=request("GcCode")%>">
		<input type="hidden" name="IDProduct" value="<%=request("IdProduct")%>">
		<input name="submit1" type="submit" value="Update" onclick="document.hForm.subdel.value='0';" class="submit2">&nbsp;
		<input name="submit2" type="submit" value="Delete" onclick="document.hForm.subdel.value='1';" class="submit2">&nbsp;
		<input type="button" name="back" value="Locate Another" onClick="location='ggg_AdmManageGCs.asp';">
	</td>
</tr>
</table>
<br>
<%
query="SELECT idorder,ord_OrderName,pcOrd_GcUsed,pcOrd_GCDetails,pcOrd_GCAmount FROM Orders WHERE pcOrd_GCDetails LIKE '%" & request("GcCode") & "%';"
set rs=connTemp.execute(query)
IF not rs.eof then%>
<table class="pcCPcontent">
<tr>
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="4">Orders where this gift certificate was used</th>
</tr>
<tr>
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="12%" nowrap>Order ID#</td>
	<td>Order Name</td>
	<td align="right">Used</td>
	<td>&nbsp;</td>
</tr>
<%do while not rs.eof
	tmpIDOrder=rs("IDOrder")
	tmpOrderName=rs("ord_OrderName")
	if tmpOrderName<>"" then
	else
		tmpOrderName="No Name"
	end if
	'New features
	GCDetails=rs("pcOrd_GCDetails")
	if GCDetails<>"" then
		GCArry=split(GCDetails,"|g|")
		intArryCnt=ubound(GCArry)

		for k=0 to intArryCnt

		if GCArry(k)<>"" then
			GCInfo = split(GCArry(k),"|s|")
			if GCInfo(2)="" OR IsNull(GCInfo(2)) then
				GCInfo(2)=0
			end if
			pGiftCode=GCInfo(0)
			pGiftUsed=GCInfo(2)
			
			if pGiftCode=request("GcCode") then
				tmpUsed=pGiftUsed
				exit for
			end if
		end if
		
		next
	else
		tmpUsed=rs("pcOrd_GcUsed")
	end if
	if tmpUsed<>"" then
	else
		tmpUsed="0"
	end if%>
	<tr>
		<td align="center"><%=scpre+int(tmpIDOrder)%></td>
		<td><%=tmpOrderName%></td>
		<td align="Right"><%=ScCurSign & money(tmpUsed)%></td>
		<td align="center"><a href="OrdDetails.asp?id=<%=tmpIdOrder%>">Details</a></td>
	</tr>
	<%rs.MoveNext
loop
%>
</table>
<%
END IF
%>
</form>
<%
set rstemp=nothing
set rs=nothing
call closedb()
%>
<!--#include file="AdminFooter.asp"-->