<% pageTitle = "Import 'Order Shipped' Information - Locate data file on Web server" %>
<% section = "orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<%
on error resume next
if request("action")="check" then
FName=request("File1")
MTest=1
	if FName<>"" then
	if PPD="1" then
		FName1= "/"&scPcFolder&"/pc/catalog/" & Fname
	else
		FName1= "../pc/catalog/" & Fname
	end if
	Err.number=0
	findit = Server.MapPath(FName1)
	if Err.number>0 then
	MTest=0
	msg="File not found! Please the check file name and make sure it has been uploaded to the 'pc/catalog' directory, then try again."
	Err.number=0
	Err.Description=""
	else
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(findit, 1)
	if Err.number>0 then
	MTest=0
	msg="File not found! Please the check file name and make sure it has been uploaded to the 'pc/catalog' directory, then try again."
	Err.number=0
	Err.Description=""	
	end if
	end if
else
MTest=0
end if
if MTest=1 then
 session("importfile")=FName
 Response.redirect "ship-index_import.asp?s=1&nextstep=1&msg=ProductCart successfully located the data file " & FName & " on the Web server."
end if
end if
%> 
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
    alert("Enter the file name.");
    theForm.file1.value == ""
    theForm.file1.focus();
    return (false);
  }
  else
  {	if (isCSV(theForm.file1.value)==false)
	{
		alert("Invalid file type. Only CSV and XLS files are allowed.");
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
        <td  width="5%" align="right"><img border="0" src="images/step1a.gif"></td>
        <td width="95%"><b><font color="#000000">Upload data file</font></b></td>
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

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<div style="margin: 10px 0 10px 0;">"Enter the full name of the data file that has already been uploaded to the Web server. <strong>Only *.csv &amp; *.xls</strong> files are accepted. For more information on what fields can be imported and on how to prepare your *.csv &amp; *.xls files for import, please refer to the ProductCart <a href="http://wiki.earlyimpact.com/productcart/orders_importing_shipping_info" target="_blank">User Guide</a>.</div>

<form method="post" action="ship-index_import2.asp?action=check" onSubmit="return Form1_Validator(this)">
<div>Enter CSV/XLS file name</div>
<div style="margin: 10px 0 10px 0">File name: <input type="text" name="file1" size="30" value="<%=request("file1")%>">&nbsp;<span class="pcSmallText">(e.g. myimport.xls)</div>
	<div><strong>NOTE</strong>: enter the file name, <u>not</u> the path. The file MUST be located in the &quot;/productcart/pc/catalog&quot; folder on the Web server (e.g. myimport.xls).</div>
    <div style="margin: 10px 0 10px 0;"><input type="submit" name="submit" class="submit2"></div>
</form>
</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->