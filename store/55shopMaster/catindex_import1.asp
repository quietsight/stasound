<% pageTitle = "Category Import Wizard - Upload data file" %>
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
	
function isCSV(s)
	{
		var test=""+s ;
		test1="";
		for (var k=test.length-4; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			test1 += c
		}
		if (test1==".CSV"||test1==".Csv"||test1==".csv"||test1==".CSv"||test1==".CsV"||test1==".csV"||test1==".cSv"||test1==".cSV")
			{
				return (true);
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
  {	if (isCSV(theForm.file1.value)==false)
	{
		alert("Invalid CSV & XLS file types is not allowed to be uploaded to the server.");
		theForm.file1.value == ""
		theForm.file1.focus();
		return (false);
    }
  }
return (true);
}
//-->
</script>

    <table class="pcCPcontent">
    <tr>
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td width="5%" align="right"><img border="0" src="images/step1a.gif"></td>
        <td width="95%"><b>Select category data file</b></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step2.gif"></td>
        <td><font color="#A8A8A8">Map fields</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step3.gif"></td>
        <td><font color="#A8A8A8">Confirm mapping</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step4.gif"></td>
        <td><font color="#A8A8A8">Import results</font></td>
    </tr>
    </table>
    
	<br />
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>

<form method="post" enctype="multipart/form-data" action="catstep1a.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
    	<td colspan="2">Use the form below to upload your existing product database. <strong>Only *.csv &amp; *.xls files</strong> can be uploaded. For more information on what fields can be imported and on how to prepare your *.csv &amp; *.xls files for import, please refer to the ProductCart <a href="http://wiki.earlyimpact.com/productcart/products_import" target="_blank">User Guide</a>.
	</td>
</tr>
<tr> 
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <th>CSV/XLS category data file</th>
    </tr>
    <tr>
    	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
        <td>Select a file to upload: <input type="file" name="file1" size="30"></td>
    </tr>
    <tr>
        <td><input type="submit" name="submit" value="Upload" class="submit2"></td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->