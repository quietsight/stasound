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
<% Dim pageTitle, Section
pageTitle="Create Feedback Status"
Section="layout" %>
<!-- #Include File="Adminheader.asp" -->
<%
Dim rs, connTemp, query

if request("action")="create" then
	call openDB()
	Dim strPName,strPImg,strPBgColor,intShowImg
	strPName=getUserInput(request("PName"),0)
	strPImg=getUserInput(request("PImg"),0)
	strPBgColor=getUserInput(request("PBgColor"),0)
	query="select pcFStat_ShowImg from pcFStatus order by pcFStat_IDStatus asc"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		intShowImg=rs("pcFStat_ShowImg")
		else
		intShowImg="0"
	end if
	query="Insert Into pcFStatus (pcFStat_Name,pcFStat_Img,pcFStat_BgColor,pcFStat_ShowImg) values ('" & strPName &"','" & strPImg &"','" & strPBgColor & "'," & intShowImg & ")"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closedb()
	response.redirect "adminFBStatusManager.asp?s=1&msg=New Feedback Status added successfully!"
end if
%>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{
	if (theForm.PName.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.PName.focus();
		    return (false);
	}
return (true);
}
//-->
</script>
<form name="form1" method="post" action="adminCreateFBStatus.asp?action=create" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
  <tr>
    <td colspan="2" class="pcCPspacer"></td>
  </tr>
  <tr>
    <td width="10%" align="right">Name:</td>
    <td width="90%"><input name="PName" type="text" value="" size="25">
    </td>
  </tr>
  <tr>
    <td align="right" nowrap="nowrap">Image File:</td>
    <td valign="top" nowrap="nowrap"><input name="PImg" type="text" value="" size="25"> Type in the file name. To upload an image <a href="javascript:;" onclick="window.open('adminimageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a></td>
  </tr>
  <tr>
    <td align="right" nowrap="nowrap">Background Color:</td>
    <td valign="top" nowrap="nowrap"><input name="PBgColor" id="PBgColor" type="text" value="" size="25">
    <input type="button" value="Choose" id="Choose" onClick="PcjsColorChooser('Choose','PBgColor','value')" name="button4">  
    </td>
  </tr>
  <tr>
    <td colspan="2" class="pcCPspacer"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
			<input type="submit" name="Submit" value="Create" class="submit2">
			&nbsp;<input type="button" name="back" value="Back" onClick="location='adminFBStatusManager.asp';">
		</td>
  </tr>
</table>
</form>
<!-- #Include File="Adminfooter.asp" -->