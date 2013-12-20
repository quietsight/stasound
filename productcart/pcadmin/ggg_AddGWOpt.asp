<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add Gift Wrapping Option" %>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim connTemp,query,rstemp
call opendb()

if request("action")="add" then

	OptName=request("OptName")

	if OptName<>"" then
		OptName=replace(OptName,"'","''")
	end if	

	OptImg=request("OptImg")
	
	OptPrice=request("OptPrice")
	if OptPrice="" then
		OptPrice="0"
	end if
	
	if scDecSign = "," then
	    OptPrice = replace(OptPrice,".","")
	    OptPrice = replace(OptPrice,",",".")
	else
	    OptPrice = replace(OptPrice,",","")
	end if
	
	OptActive=request("OptActive")
	if OptActive="" then
		OptActive="0"
	end if
	
	OptOrder=request("OptOrder")
	if OptOrder="" then
		OptOrder="0"
	end if

	query="INSERT INTO pcGWOptions (pcGW_OptName,pcGW_OptImg,pcGW_OptPrice,pcGW_OptActive,pcGW_OptOrder) values ('" & OptName & "','" & OptImg & "'," & OptPrice & "," & OptActive & "," & OptOrder & ");"
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
		
	msg="New gift wrapping option added successfully!"
	msgType=1
end if
%>		

<script language="JavaScript">
<!--
function newWindow(file,window)
{
	msgWindow=open(file,window,'resizable=no,width=400,height=500');
	if (msgWindow.opener == null) msgWindow.opener = self;
}

function chgWin(file,window)
{
	msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
	if (msgWindow.opener == null) msgWindow.opener = self;
}

function isDigit(s)
{
	var test=""+s;
	if(test==","||test=="."||test=="-"||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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

	if (theForm.OptName.value == "")
	{
		alert("Please enter a value for this field.");
		theForm.OptName.focus();
    return (false);
	}

	if (theForm.OptPrice.value != "")
  	{
		if (allDigit(theForm.OptPrice.value) == false)
		{
			alert("Please enter a valid number for this field.");
			theForm.OptPrice.focus();
	    return (false);
		}
	}

return (true);
}
//-->
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form name="hForm" method="post" action="ggg_AddGWOpt.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="23%">Option Name:</td>
	<td width="75%"><input name="OptName" type=text value=""></td>
</tr>
<tr>
	<td>Option Image <i>(optional)</i>:</td>
	<td><input name="OptImg" type=text value=""><a href="javascript:chgWin('../pc/imageDir.asp?ffid=OptImg<%=Count%>&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a href="javascript:;" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')"><img src="images/sortasc_blue.gif" alt="Upload Image"></a></td>
</tr>
<tr>
	<td>Price:</td>
	<td><input name="OptPrice" type="text" size="8" value="0"></td>
</tr>
<tr>
	<td>Active:</td>
	<td><input name="OptActive" type="checkbox" value="1" checked class="clearBorder"></td>
</tr>
<tr>
	<td>Order:</td>
	<td><input name="OptOrder" type="text" size="4" value="0"></td>
</tr>
<tr> 
	<td align="center" colspan="2">
		<br>
		<input name="add" type="submit" class="submit2" value=" Add option ">
		&nbsp;
		<input name="back" type="button" onClick="location='ggg-GiftWrapOptions.asp';" value="Back">
	</td>
</tr>
</table>
</form>
<%call closedb()%><!--#include file="AdminFooter.asp"-->