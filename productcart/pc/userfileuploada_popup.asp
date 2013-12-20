<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/rc4.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->

<% 
dim mySQL, conntemp, rstemp

on error resume next

session("UIDFeedback")=request("IDFeedback")
session("Relink")=request("ReLink")

If session("UIDFeedback")&""="" OR session("Relink")&""="" OR scShowHD = 0 Then
	response.redirect "custPref.asp"
	response.end
End If

%>
<html>
<head>
<title>Upload Data File(s)</title>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<script language="JavaScript" type="text/javascript">
<!--
	
function isCSV(s)
	{
		var test=""+s ;
		test2="";
		for (var k=test.length-4; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			test2 += c
		}
		test1=test2.toLowerCase()
		if (test1==".txt"||test1==".gif"||test1==".jpg"||test1==".htm"||test1==".zip"||test1==".pdf"||test1==".doc"||test1==".png")
			{
				return (true);
			}
		test2="";
		for (var k=test.length-5; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			test2 += c
		}
		test1=test2.toLowerCase()
		if (test1==".html"||test1==".jpeg")
			{
				return (true);
			}			
		
		return (false);
	}

function containsComma(s)
	{
	var pos=s.indexOf(",");
	if (pos>=0)
	{
		return(true);
	}
	return(false);
}
	

function Form1_Validator(theForm)
{

  if (theForm.one.value == "" && theForm.two.value == "" && theForm.three.value == "" && theForm.four.value == "" && theForm.five.value == "" && theForm.six.value == "")
  {
    alert("You need to supply at least one file to upload.");
    theForm.one.focus();
    return (false);
  }
  else
  {	

  if (theForm.one.value != "")
  {
  if (isCSV(theForm.one.value)==false)
	{
		alert("File type not allowed. The file cannot be uploaded to the server.");
		theForm.one.focus();
		return (false);
    }
  if (containsComma(theForm.one.value)==true)
	{
		alert("The file name cannot contain a comma.");
		theForm.one.focus();
		return (false);
    }
  }

  if (theForm.two.value != "")
  {
  if (isCSV(theForm.two.value)==false)
	{
		alert("File type not allowed. The file cannot be uploaded to the server.");
		theForm.two.focus();
		return (false);
    }
  if (containsComma(theForm.two.value)==true)
	{
		alert("The file name cannot contain a comma.");
		theForm.two.focus();
		return (false);
    }
  }

  if (theForm.three.value != "")
  {
  if (isCSV(theForm.three.value)==false)
	{
		alert("File type not allowed. The file cannot be uploaded to the server.");
		theForm.three.focus();
		return (false);
    }
  if (containsComma(theForm.three.value)==true)
	{
		alert("The file name cannot contain a comma.");
		theForm.three.focus();
		return (false);
    }
  }
  
  if (theForm.four.value != "")
  {
  if (isCSV(theForm.four.value)==false)
	{
		alert("File type not allowed. The file cannot be uploaded to the server.");
		theForm.four.focus();
		return (false);
    }
  if (containsComma(theForm.four.value)==true)
	{
		alert("The file name cannot contain a comma.");
		theForm.four.focus();
		return (false);
    }
  }
  
  if (theForm.five.value != "")
  {
  if (isCSV(theForm.five.value)==false)
	{
		alert("File type not allowed. The file cannot be uploaded to the server.");
		theForm.five.focus();
		return (false);
    }
  if (containsComma(theForm.five.value)==true)
	{
		alert("The file name cannot contain a comma.");
		theForm.five.focus();
		return (false);
    }
  }
  
  if (theForm.six.value != "")
  {
  if (isCSV(theForm.six.value)==false)
	{
		alert("File type not allowed. The file cannot be uploaded to the server.");
		theForm.six.focus();
		return (false);
    }
  if (containsComma(theForm.six.value)==true)
	{
		alert("The file name cannot contain a comma.");
		theForm.six.focus();
		return (false);
    }
  }
  
  }
return (true);
}
//-->
</script>
</head>
<body>
<div id="pcMain">
<form method="post" enctype="multipart/form-data" action="userfileupl_popup.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
	<table class="pcMainTable" cellpadding="4">
		<tr>
			<td colspan="2">
				<h1>Upload Data File(s)</h1>
			</td>
		</tr>
		<tr>
			<td colspan="2">Select a file using the &quot;Browse&quot; button, then click on &quot;Upload&quot;. Only *.txt, *.htm, *.html, *.gif, *.jpg, *.pdf, *.doc and *.zip file types were accepted.
			</td>
     </tr>
     <tr> 
       <td colspan="2" class="pcSpacer"></td>
     </tr>
     <tr> 
       <td width="20%" align="right">
			 	File 1:
       </td>
       <td width="80%">
			 	<input type="file" name="one" size="25">
       </td>
      </tr>
     <tr> 
       <td width="20%" align="right">
			 	File 2:
       </td>
       <td width="80%">
			 	<input type="file" name="two" size="25">
       </td>
      </tr>
     <tr> 
       <td width="20%" align="right">
			 	File 3:
       </td>
       <td width="80%">
			 	<input type="file" name="three" size="25">
       </td>
      </tr>
     <tr> 
       <td width="20%" align="right">
			 	File 4:
       </td>
       <td width="80%">
			 	<input type="file" name="four" size="25">
       </td>
      </tr>
     <tr> 
       <td width="20%" align="right">
			 	File 5:
       </td>
       <td width="80%">
			 	<input type="file" name="five" size="25">
       </td>
      </tr>
     <tr> 
       <td width="20%" align="right">
			 	File 6:
       </td>
       <td width="80%">
			 	<input type="file" name="six" size="25">
       </td>
      </tr>
     <tr> 
       <td colspan="2" class="pcSpacer"></td>
     </tr>
     <tr>
		 	<td colspan="2" align="center"> 
				<input type="submit" name="Submit" value="Upload" class="submit2">
				<input type="button" value="Close Window" onClick="javascript:window.close();">
      </td>
     </tr>
    </table>
	</form>
</div>
</body>
</html>