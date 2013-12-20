<% pageTitle = "Additional Product Images Import" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--

<%if request("action")="go" then
		imType=request("imType")
		session("ii_imType")=imType
		response.redirect "iistep3.asp"
end if%>
	
function isXLS(s)
	{
		var test=""+s ;
		test1="";
		for (var k=test.length-4; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			test1 += c
		}
		if (test1==".XLS"||test1==".Xls"||test1==".xls"||test1==".XLs"||test1==".XlS"||test1==".xlS"||test1==".xLs"||test1==".xLS")
		{
			return (true);
		}
		
		return (false);
	}
	

function Form1_Validator(theForm)
{

  if (theForm.file1.value == "")
  {
    alert("You need to supply at least one file to upload.");
    theForm.file1.value == ""
    theForm.file1.focus();
    return (false);
  }
  else
  {	if (isXLS(theForm.file1.value)==false)
	{
		alert("Invalid XLS file type is not allowed to be uploaded to the server.");
		theForm.file1.value == ""
		theForm.file1.focus();
		return (false);
    }
  }
return (true);
}
//-->
</script>
   
	<br />
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
<%if request("a")="next" then%>
<form method="post" action="iistep1.asp?action=go" class="pcForms">
<%else
session("importfile")=""%>
<form method="post" enctype="multipart/form-data" action="iistep2.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
<%end if%>
<table class="pcCPcontent">
	<%if request("a")<>"next" then%>
	<tr>
    	<td colspan="2">Use the form below to upload your existing product database. <strong>Only *.xls files</strong> can be uploaded. For more information on what fields can be imported and on how to prepare your *.xls files for import, please refer to the ProductCart <a href="http://wiki.earlyimpact.com/productcart/product_images_import" target="_blank">User Guide</a>.
	</td>
	</tr>
    <tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <th colspan="2">XLS product data file</th>
    </tr>
    <tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
        <td colspan="2">Select a file to upload <input type="file" name="file1" size="30"></td>
    </tr>
	<tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
	<tr>
        <td><input type="submit" name="submit" value="Upload" class="submit2"></td>
    </tr>
	<%end if%>
	<%if request("a")="next" then%>
    <tr> 
        <th colspan="2">Are you importing new Additional Product Views or replacing existing ones?</th>
    </tr>
	<tr>
		<td width="5%" align="right"><input type="radio" name="imType" value="0" <%if session("ii_imType")<>"1" then%>checked<%end if%> class="clearBorder"></td>
		<td><strong>New</strong> - They are added to the database</td>
	</tr>
	<tr>
		<td width="5%" align="right"><input type="radio" name="imType" value="1" <%if session("ii_imType")="1" then%>checked<%end if%> class="clearBorder"></td>
		<td><strong>Replace</strong> - Existing “Additional Product Views” for those SKUs are removed, then the new ones are added</td>
	</tr>
	<tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
        <td><input type="submit" name="submit" value="Import Data" class="submit2"></td>
    </tr>
	<%end if%>
</table>
</form>
<!--#include file="AdminFooter.asp"-->