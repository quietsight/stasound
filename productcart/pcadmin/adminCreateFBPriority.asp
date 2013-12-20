<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim pageTitle, Section
pageTitle="Create New Message Priority Level"
Section="layout" %>
<!-- #Include File="Adminheader.asp" -->
<%
Dim rs, connTemp, query

if request("action")="create" then
	call openDB()
	strPName=getUserInput(request("PName"),0)
	strPImg=getUserInput(request("PImg"),0)
	query="select pcPri_ShowImg from pcPriority order by pcPri_IDPri asc"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		intPriShowImg=rs("pcPri_ShowImg")
	else
		intPriShowImg="0"
	end if
	query="Insert Into pcPriority (pcPri_Name,pcPri_Img,pcPri_ShowImg) values ('" & strPName &"','" & strPImg &"'," & intPriShowImg & ")"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closedb()
	response.redirect "adminFBPriorityManager.asp?s=1&msg=New Message Priority Level added successfully!"
end if
%>

<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{
	if (theForm.PName.value == "")
 	{
		    alert("Please enter a value for the Message Priority name.");
		    theForm.PName.focus();
		    return (false);
	}
return (true);
}
//-->
</script>
<form name="search" method="post" action="adminCreateFBPriority.asp?action=create" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
  <tr>
    <td colspan="2" class="pcCPspacer"></td>
  </tr>
  <tr>
    <td width="10%" align="right">Name:</td>
    <td width="90%"><input name="PName" type="text" value="" size="25"></td>
  </tr>
  <tr>
    <td align="right" nowrap="nowrap">Image File:</td>
    <td valign="top" nowrap="nowrap"><input name="PImg" type="text" value="" size="25"> Type in the file name. To upload an image <a href="#" onclick="window.open('adminimageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click
      here</a>
    </td>
  </tr>
  <tr>
    <td class="pcCPspacer"></td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td>
    	<input type="submit" name="Submit" value="Create" class="submit2">
		&nbsp;<input type="button" name="back" value=" Back " onClick="location='adminFBPriorityManager.asp';">
	</td>
  </tr>
</table>
</form>
<!-- #Include File="Adminfooter.asp" -->