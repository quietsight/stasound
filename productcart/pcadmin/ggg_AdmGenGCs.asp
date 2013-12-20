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
<% pageTitle = "Generate Gift Certificates codes" %>
<% Section = "products" %>
<%Dim connTemp,query,rstemp
call opendb()
%>
<!--#include file="AdminHeader.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
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
	if (theForm.IDGC.value=="")
	{
		alert("Please select a gift certificate product");
	    theForm.IDGC.focus();
	    return (false);
	}
	if (theForm.GCcount.value == "")
	{
		alert("Please enter the quantity of gift certificate codes");
	    theForm.GCcount.focus();
	    return (false);
	}
  	if (allDigit(theForm.GCcount.value) == false)
	{
		alert("Please enter the correct number for this field.");
	    theForm.GCcount.focus();
	    return (false);
	}
	if (theForm.GCcount.value < 0)
	{
		alert("Please enter the correct number greater than zero for this field.");
	    theForm.GCcount.focus();
	    return (false);
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
<form name="hForm" method="post" action="ggg_AdmGenGCsB.asp?action=gen" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">
		<p>Use this page to generate new Gift Certificate codes for one or more customers. The codes generated here work the same as the ones that are generated when customers purchase Gift Certificates in the storefront. Gift Certificates are not customer-specific. You can send a Gift Certificate code to any new or existing customer.</p>
		<% if iGiftCerts <> 1 then %>
			<p>There are currently <strong>no gift certificates</strong> in your store.</p>
		<% end if %>
	</td>
</tr>
<%IF iGiftCerts = 1 THEN%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2"><h2>Generate Gift Certificates Codes:</h2></td>
</tr>
<tr>
	<td nowrap="nowrap" width="20%" align="right">Gift Certificate Product:</td>
	<td width="80%">
		<select name="IDGC">
			<option value="" selected></option>
			<%do while not rstemp.eof%>
				<option value="<%=rstemp("IDProduct")%>"><%=rstemp("Description")%></option>
				<%rstemp.MoveNext
			loop
			set rstemp=nothing%>
		</select>
	</td>
</tr>
<tr>
	<td nowrap="nowrap" align="right">Gift Certificate Codes to be Generated:</td>
	<td><input type="text" name="GCcount" size="5" value="1"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td></td>
	<td>
		<input type="hidden" name="iPageSize" value="25">
		<input name="submit1" type="submit" value="Generate" class="submit2" onClick="pcf_Open_genGerts();">&nbsp;
		<input type="button" name="back" value=" Back " onclick="location='ggg_AdmManageGCs.asp';" class="ibtnGrey">
	</td>
</tr>
<%END IF 'iGiftCerts = 1%>
</table>
</form>
<%
set rstemp=nothing
call closedb()

'// Loading Window
'	>> Call Method with OpenHS();
	response.Write(pcf_ModalWindow("Generating Gift Certificates. This task can take some time. Please wait...", "genGerts", 300))
%>
<!--#include file="AdminFooter.asp"-->