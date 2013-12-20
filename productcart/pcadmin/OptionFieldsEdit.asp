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
'on error resume next
dim f, mySQL, conntemp, rstemp

call opendb()
pidcustomcardtype=request.QueryString("idc")
pidpayment=request.QueryString("id")
pidcustomcardrule=request.QueryString("idccr")
pgwCode=request.QueryString("gwCode")
if pidcustomcardtype="" then
	pidcustomcardtype=request.form("idc")
	pidpayment=request.form("id")
	pidcustomcardrule=request.form("idccr")
	pgwCode=request.form("gwCode")
end if
if isNumeric(pidcustomcardrule) AND request.QueryString("m")="del" then
	'delete from database the one option
	mySQL="DELETE FROM customCardRules WHERE idCustomCardRules="&pidcustomcardrule
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(mySQL)
	set rstemp=nothing
	call closedb()
	response.redirect "modCustomCardPaymentOpt.asp?mode=Edit&idc="&pidcustomcardtype&"&id="&pidpayment&"&idccr="&idcustomcardrule&"&gwCode="&pgwCode
end if

if isNumeric(pidcustomcardrule) AND request.form("Update")="Update" then
	pruleName=replace(request.form("ruleName"),"'","''")
	pintRuleRequired=request.form("ruleRequired")
	if pintRuleRequired="YES" then
		pintRuleRequired="-1"
	else
		pintRuleRequired="0"
	end if
	pintLengthOfField=request.Form("intLengthOfField")
	if pintlengthOfField="" then
		pintlengthOfField="20"
	end if
	pintMaxInput=request.Form("intMaxInput")
	if pintMaxInput="" then
		pintMaxInput="0"
	end if
	mySQL="UPDATE customCardRules SET ruleName='"&pruleName&"',intruleRequired="&pintRuleRequired&", intlengthOfField="&pintLengthOfField&",intMaxInput="&pintMaxInput&" WHERE idCustomCardRules="&pidcustomcardrule&";"
	set rs=connTemp.execute(mySQL)
	set rs=nothing
	call closeDb()
	response.redirect "modCustomCardPaymentOpt.asp?mode=Edit&idc="&pidcustomcardtype&"&id="&pidpayment&"&idccr="&idcustomcardrule&"&gwCode="&pgwCode
end if

if isNumeric(pidcustomcardrule) then
	mySQL="SELECT idCustomCardType, ruleName, intRuleRequired, intLengthOfField, intMaxInput FROM customCardRules WHERE idCustomCardRules="&pidcustomcardrule&";"
	set rs=connTemp.execute(mySQL)
	pruleName=rs("ruleName")
	pintRuleRequired=rs("intRuleRequired")
	pintLengthOfField=rs("intLengthOfField")
	pintMaxInput=rs("intMaxInput")
	call closedb()
	set rs=nothing
end if 

%>
	<% pageTitle="Custom Payment Options - Edit Field" %>
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
			if (allDigit(theForm.intLengthOfField.value) == false)
				{
					alert("Please enter a numeric value for this field.");
					theForm.intLengthOfField.focus();
					return (false);
				}
			if (allDigit(theForm.intMaxInput.value) == false)
				{
					alert("Please enter a numeric value for this field.");
					theForm.intMaxInput.focus();
					return (false);
				}
		return (true);
	}
	//-->
	</script>
	
	<%
	if NOT validNum(pidcustomcardrule) then 
	%>
	<div class="pcCPmessage">Invalid input, use the browser's back button and please try again.</div>
	<% else %>
    
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
	
	<form method="post" name="form1" action="OptionFieldsEdit.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
		<input name="idc" type="hidden" id="idc" value="<%=pidcustomcardtype%>">
		<input name="id" type="hidden" id="id" value="<%=pidpayment%>">
		<input name="gwCode" type="hidden" id="gwCode" value="<%=pgwCode%>">
		<input name="idccr" type="hidden" id="idccr" value="<%=pidcustomcardrule%>">
	<table class="pcCPcontent">
	<tr>           
		<td colspan="2" class="pcCPspacer"></td>
	</tr> 
	<tr>
		<td width="25%" align="right">Text Label:</td>
		<td width="75%">
			<input name="ruleName" type="text" id="ruleName" value="<%=pruleName%>" size="20" maxlength="50">
       <span class="pcSmallText">(e.g.: Name on Card) Maximum of 50 characters</span>
		</td>
	</tr>
	<tr>
		<td align="right">
			<% if pintRuleRequired="-1" then %>
				<input name="ruleRequired" type="checkbox" id="ruleRequired" value="YES" checked class="clearBorder">
			<% else %>
				<input name="ruleRequired" type="checkbox" id="ruleRequired" value="YES" class="clearBorder">
			<% end if %>
		</td>
		<td>This is a required field</td>
	</tr>
	<tr>
		<td align="right">Length of Field:</td>
		<td>
			<input name="intLengthOfField" type="text" id="intLengthOfField" value="<%=pintLengthOfField%>" size="4" maxlength="4">
    	<span class="pcSmallText">Length of the input field itself</span>
		</td>
	</tr>
	<tr>
		<td align="right">Maximum Input length:</td>
		<td>
			<input name="intMaxInput" type="text" id="intMaxInput" value="<%=pintMaxInput%>" size="4" maxlength="3">
			<span class="pcSmallText">Not to exceed 250 characters</span>
		</td>
	</tr>   
		<tr>           
			<td colspan="2" class="pcCPspacer"></td>
		</tr>    
		<tr> 		
			<td colspan="2" align="center">
			<input name="Update" type="submit" id="Update" value="Update" class="submit2">
			&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
	</form>
<% 
	end if
%>
<!--#include file="AdminFooter.asp"-->