<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/languages.asp" -->
<%
pidcustomcardtype=request.QueryString("idc")
pidpayment=request.QueryString("id")
if pidcustomcardtype="" then
	pidcustomcardtype=request.form("idc")
	pidpayment=request.form("id")
end if
dim f, mySQL, conntemp, rstemp

if request.form("Add")="Add" then
	call opendb()
	ruleName=replace(request.form("ruleName"),"'","''")
	ruleRequired=request.form("ruleRequired")
	if ruleRequired="YES" then
		intruleRequired="-1"
	else
		intruleRequired="0"
	end if
	intlengthOfField=request.Form("lengthOfField")
	if intlengthOfField="" then
		intlengthOfField="20"
	end if
	intMaxinput=request.Form("maxInput")
	if intMaxinput="" then
		intMaxinput="0"
	end if
	mySQL="Insert into customCardRules (idCustomCardType,ruleName,intruleRequired, intlengthOfField,intMaxInput, intOrder) values ("&pidcustomcardtype&",'"&ruleName&"',"&intruleRequired&", "&intlengthOfField&", "&intMaxinput&", 0)"
	set rstemp=connTemp.execute(mySQL)
	set rstemp=nothing
	call closeDb()
	response.redirect "modCustomCardPaymentOpt.asp?mode=Edit&idc="&pidcustomcardtype&"&id="&pidpayment&"&gwCode=10"&pidcustomcardtype
end if
%>
<% pageTitle="Custom Payment Options - Add New Field" %>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->

<script language="JavaScript">
<!--
function isDigit(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
	if (theForm.ruleName.value == "")
  	{
			 alert("Please enter a value for the Text Label field.");
		    theForm.ruleName.focus();
		    return (false);
		}
		if (allDigit(theForm.lengthOfField.value) == false)
			{
				alert("Please enter a numeric value for this field.");
				theForm.lengthOfField.focus();
				return (false);
			}
		if (allDigit(theForm.maxInput.value) == false)
			{
				alert("Please enter a numeric value for this field.");
				theForm.maxInput.focus();
				return (false);
			}
	return (true);
}
//-->
</script>

<form method="post" name="form1" action="OptionFieldsAdd.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input name="idc" type="hidden" id="idc" value="<%=pidcustomcardtype%>">            
	<input name="id" type="hidden" id="id" value="<%=pidpayment%>">
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
		<td width="25%"><div align="right">Text Label:</div></td>
		<td width="75%"><input name="ruleName" type="text" id="ruleName" size="20" maxlength="50">
		<span class="pcSmallText">(e.g.: Name on Card) Maximum of 50 characters</span></td>
	</tr>
	<tr>
		<td align="right"><input name="ruleRequired" type="checkbox" id="ruleRequired" value="YES" class="clearBorder"></td>
		<td>This is a required field</td>
	</tr>
	<tr>
		<td align="right">Length of Field:</td>
		<td>
			<input name="lengthOfField" type="text" id="lengthOfField" value="20" size="4" maxlength="4">
      <span class="pcSmallText">Size of the input field itself</span>
		</td>
	</tr>
	<tr>
		<td align="right">Maximum Input length:</td>
		<td>
		<input name="maxInput" type="text" id="maxInput" size="4" maxlength="3">
    <span class="pcSmallText">Not to excede 250 characters</span>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 	
		<td colspan="2" align="center">
			<input name="Add" type="submit" id="Add" value="Add" class="submit2">
			&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->