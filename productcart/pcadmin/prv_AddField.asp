<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Add New Field for Product review pages" 
pageIcon="pcv4_icon_reviews.png"
section="reviews" 
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<% IF request("action")="add" then

	pcv_Name=request("pcv_name" & k)
	pcv_Type=request("pcv_type" & k)
	if pcv_Type="" then
		pcv_Type="0"
	end if
	pcv_Active=request("pcv_active" & k)
	if pcv_Active="" then
		pcv_Active="0"
	end if
	pcv_Required=request("pcv_required" & k)
	if pcv_Required="" then
		pcv_Required="0"
	end if
	pcv_Order="0"

	Dim rs, connTemp, query
	call openDb()

	query="INSERT INTO pcRevFields (pcRF_Name,pcRF_Type,pcRF_Active,pcRF_Required,pcRF_Order) VALUES ('" & pcv_Name & "'," & pcv_Type & "," & pcv_Active & "," & pcv_Required & "," & pcv_Order & ");"
	set rs=connTemp.execute(query)
	
	set rs=nothing
	call closedb()
	
	response.redirect "prv_FieldManager.asp?s=1&msg=" & server.URLEncode("New field added successfully.")

END IF
	
%>
<script>
function Form1_Validator(theForm)
{
  if (theForm.pcv_name.value == "")
  {
  alert("Enter a value for Field Name.");
  theForm.pcv_name.focus();
  return(false);
  }
  if (theForm.pcv_type.value == "")
  {
  alert("Enter select a value for Field Type.");
  theForm.pcv_name.focus();
  return(false);
  }
  return(true);
}
</script>

<form method="POST" action="prv_AddField.asp?action=add" name="checkboxform" onSubmit="return Form1_Validator(this)" class="pcForms">
	<table class="pcCPcontent">
    	<tr>
        	<td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td nowrap width="20%" align="right">Field Name:</td>
			<td align="left" width="80%"><input type=text size="50" name="pcv_name" value=""></td>
		</tr>
		<tr> 
			<td align="right" nowrap>Field Type:</td>
			<td align="left">
			<select name="pcv_type">
                <option value="0" selected>1-row text field</option>
                <option value="1">Text area</option>
                <option value="2">Drop-down list</option>
                <option value="3">'Feeling' Rating</option>
                <option value="4">'Mark' Rating</option>
			</select>
            </td>
		</tr>
		<tr> 
			<td align="right" nowrap>Active:</td>
			<td align="left"><input type="checkbox" name="pcv_active" value="1" checked class="clearBorder"></td>
		</tr>
		<tr> 
			<td align="right" nowrap>Required:</td>
			<td align="left"><input type="checkbox" name="pcv_required" value="1" class="clearBorder"></td>
		</tr>
    	<tr>
        	<td colspan="2" class="pcCPspacer"><hr></td>
        </tr>
		<tr>
			<td align="center" colspan="2">
			<input type="submit" value=" Add New " class="submit2">&nbsp;
			<input type="button" value=" Back " onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->