<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials" 
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<% 
Dim query, conntemp, rstemp, rstemp4, rstemp5

call openDB()

CP=0
session("sm_ShowNext")=0
tmpID=session("sm_pcSaleID")
if tmpID="" then
	tmpID="0"
end if
query="SELECT TOP 1 idProduct FROM pcSales_Pending WHERE pcSales_ID=" & tmpID & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	CP=1
end if
set rs=nothing


if (CP="") OR (CP="0") then
	call closedb()
	response.redirect "sm_addedit_S1.asp"
	response.end
end if

pcSaleID=session("sm_pcSaleID")

if request("a")<>"back" then

if pcSaleID<>"" then
	query="SELECT pcSales_TargetPrice,pcSales_Type,pcSales_Relative,pcSales_Amount,pcSales_Round FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpSalesType=rs("pcSales_Type")
		if tmpSalesType="0" OR tmpSalesType="1" then
			session("sm_UP1")="1"
		else
			session("sm_UP1")="2"
		end if
		session("sm_UP2")=rs("pcSales_TargetPrice")
		session("sm_UP3")=rs("pcSales_Amount")
		if session("sm_UP1")="1" then
			session("sm_UP4")=rs("pcSales_Type")
		else
			session("sm_UP4")=rs("pcSales_Relative")
		end if
		session("sm_UP5")=rs("pcSales_Round")
	else
		pcSaleID=""
		session("sm_pcSaleID")=pcSaleID
	end if
	set rs=nothing
end if

end if

if session("sm_UP1")="" then
	session("sm_UP1")=0
end if

pageIcon="pcv4_icon_salesManager.png"
if pcSaleID="" then
	pageTitle="Sales Manager - Create New Sale - Step 2: Define Sale"
else
	pageTitle="Sales Manager - Edit Sale - Step 2: Define Sale"
end if

%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<script language="JavaScript">
<!--
	
function isDigit(s)
{
var test=""+s;
if(test=="+"||test=="-"||test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
	
function isDigit1(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit1(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit1(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}	

function Form1_Validator(theForm)
{

	if (theForm.UP1.value == "0")
 	{
		    alert("Please select one of the available price change options before proceeding.");
		    return (false);
	}
	
	
	if (theForm.UP1.value == "1")
  	{
			if (theForm.cprice.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.cprice.focus();
		    return (false);
		    }
			if (allDigit(theForm.cprice.value) == false)
			{
		    alert("Please enter a right value for this field.");
		    theForm.cprice.focus();
		    return (false);
		   }		    
	}
	if (theForm.UP1.value == "2")
  	{
			if (theForm.wprice.value == "")
			{
		    alert("Please enter a value for this field.");
		    theForm.wprice.focus();
		    return (false);
		    }
			if (allDigit(theForm.wprice.value) == false)
			{
		    alert("Please enter a right value for this field.");
		    theForm.wprice.focus();
		    return (false);
		    }		    
	}
	
return (true);
}


//-->
</script>      

<form name="UpdateForm" action="sm_addedit_S4.asp?a=post" method="post" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" name="id" value="<%=pcSaleID%>">
	<input type="hidden" name="UP1" value="<%=session("sm_UP1")%>">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Select the <strong>type of price change</strong> that you would like to apply:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="5%" align="right">
		<input type="radio" name="UP" value="1" onClick="UpdateForm.UP1.value='1';" class="clearBorder" <%if session("sm_UP1")="1" then%>checked<%end if%>>
	</td>
	<td valign="top"><strong>Direct Price Change</strong>:</td>
</tr>
	<td>&nbsp;</td>
	<td valign="top">
		Set sale price by reducing the&nbsp;            
		<select name="priceSelect" id="priceSelect" size="1">
			<option value="0" <%if session("sm_UP1")="1" then%><%if session("sm_UP2")="0" then%>selected<%end if%><%else%>selected<%end if%>>Online Price</option>
			<option value="-1" <%if session("sm_UP1")="1" then%><%if session("sm_UP2")="-1" then%>selected<%end if%><%end if%>>Wholesale Price</option>
			<%tmp_HavePricingCAT=0
			query="Select idcustomerCategory, pcCC_Name FROM pcCustomerCategories order by pcCC_Name asc"
			set rstemp4=connTemp.execute(query)
			if not rstemp4.eof then
			tmp_HavePricingCAT=1%>
			<%do while not rstemp4.eof%>
			<option value="<%=rstemp4("idcustomerCategory")%>" <%if session("sm_UP1")="1" then%><%if session("sm_UP2") & ""=rstemp4("idcustomerCategory") & "" then%>checked<%end if%><%end if%>><%=rstemp4("pcCC_Name")%></option>
			<%rstemp4.MoveNext
			loop%>
			<%end if
			set rstemp4=nothing%>
		</select>
		&nbsp;by:&nbsp;              
		<input name="cprice" type="text" id="priceChange" size="8" maxlength="150" value="<%if session("sm_UP1")="1" then%><%=session("sm_UP3")%><%end if%>">
		<select name="cpriceType" id="cpriceType" size="1">
			<option value="0" <%if session("sm_UP1")="1" then%><%if session("sm_UP4")="0" then%>selected<%end if%><%else%>selected<%end if%>>% change</option>
			<option value="1" <%if session("sm_UP1")="1" then%><%if session("sm_UP4")="1" then%>selected<%end if%><%end if%>># change</option>
		</select>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input name="cpriceRound" type="radio" id="cpriceRound" value="1" <%if session("sm_UP1")="1" then%><%if session("sm_UP5")="1" then%>checked<%end if%><%end if%> class="clearBorder">&nbsp;Round updated price to the nearest integer
		<br>
		<input name="cpriceRound" type="radio" id="cpriceRound" value="0" <%if session("sm_UP1")="1" then%><%if session("sm_UP5")="0" then%>checked<%end if%><%else%>checked<%end if%> class="clearBorder">&nbsp;Round updated price to the nearest hundredth
	</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td width="5%" align="right" valign="top">
		<input type="radio" name="UP" value="2" onClick="UpdateForm.UP1.value='2';" class="clearBorder" <%if session("sm_UP1")="2" then%>checked<%end if%>>
	</td>
	<td valign="top"><strong>Relative Price Change</strong> (relative to another price or the product's cost):</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td valign="top">
		Recalculate the&nbsp; 
		<select name="priceSelect1" size="1">
			<option value="0" <%if session("sm_UP1")="2" then%><%if session("sm_UP2")="0" then%>selected<%end if%><%else%>selected<%end if%>>Online Price</option>
			<option value="-1" <%if session("sm_UP1")="2" then%><%if session("sm_UP2")="-1" then%>selected<%end if%><%end if%>>Wholesale Price</option>
			<%tmp_HavePricingCAT=0
			query="Select idcustomerCategory, pcCC_Name FROM pcCustomerCategories ORDER BY pcCC_Name asc"
			set rstemp4=connTemp.execute(query)
			if not rstemp4.eof then
			tmp_HavePricingCAT=1%>
			<%do while not rstemp4.eof%>
			<option value="<%=rstemp4("idcustomerCategory")%>" <%if session("sm_UP1")="2" then%><%if session("sm_UP2") & ""=rstemp4("idcustomerCategory") & "" then%>selected<%end if%><%end if%>><%=rstemp4("pcCC_Name")%></option>
			<%rstemp4.MoveNext
			loop%>
			<%end if
			set rstemp4=nothing%>
		</select>
		&nbsp;as&nbsp;
		<input name="wprice" type="text" id="listPriceChange" size="8" maxlength="150" value="<%if session("sm_UP1")="2" then%><%=session("sm_UP3")%><%end if%>">
		% of the&nbsp;
		<select name="priceSelect2" size="1">
			<option value="0" <%if session("sm_UP1")="2" then%><%if session("sm_UP4")="0" then%>selected<%end if%><%else%>selected<%end if%>>Online Price</option>
			<option value="-2" <%if session("sm_UP1")="2" then%><%if session("sm_UP4")="-2" then%>selected<%end if%><%end if%>>List Price</option>
			<option value="-1" <%if session("sm_UP1")="2" then%><%if session("sm_UP4")="-1" then%>selected<%end if%><%end if%>>Wholesale Price</option>
			<%'Start SDBA%>
			<option value="-3" <%if session("sm_UP1")="2" then%><%if session("sm_UP4")="-3" then%>selected<%end if%><%end if%>>Product Cost</option>
			<%'End SDBA%>
			<%tmp_HavePricingCAT=0
			query="Select idcustomerCategory, pcCC_Name FROM pcCustomerCategories ORDER BY pcCC_Name asc"
			set rstemp4=connTemp.execute(query)
			if not rstemp4.eof then
			tmp_HavePricingCAT=1%>
			<%do while not rstemp4.eof%>
			<option value="<%=rstemp4("idcustomerCategory")%>" <%if session("sm_UP1")="2" then%><%if session("sm_UP4") & ""=rstemp4("idcustomerCategory") & "" then%>selected<%end if%><%end if%>><%=rstemp4("pcCC_Name")%></option>
			<%rstemp4.MoveNext
			loop%>
			<%end if
			set rstemp4=nothing%>
		</select>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input name="cpriceRound1" type="radio" id="cpriceRound1" value="1" class="clearBorder" <%if session("sm_UP1")="2" then%><%if session("sm_UP5")="1" then%>checked<%end if%><%end if%>>&nbsp;Round updated price to the nearest integer
		<br>
		<input name="cpriceRound1" type="radio" id="cpriceRound1" value="0" class="clearBorder" <%if session("sm_UP1")="2" then%><%if session("sm_UP5")="0" then%>checked<%end if%><%else%>checked<%end if%>>&nbsp;Round updated price to the nearest hundredth
	</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr align="center">
	<td colspan="2">
		<input type="submit" name="Preview" value="Preview Sale" class="submit2">&nbsp;
		<input type="button" name="Back" value="Review Products Selection" class="ibtnGrey" onclick="location='sm_addedit_S1.asp?a=rev';">
	</td>
</tr>
</table>
</form>
<% 
call closeDb()
set rstemp= nothing
set rstemp4= nothing
set rstemp5= nothing
%>
<!--#include file="AdminFooter.asp"-->