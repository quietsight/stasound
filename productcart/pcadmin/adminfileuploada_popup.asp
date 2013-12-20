<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->

<%
on error resume next
session("UIDFeedback")=request("IDFeedback")
session("Relink")=request("ReLink")
%>
<html>
<head>
<title>Upload Data File(s)</title>
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
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" enctype="multipart/form-data" action="adminfileupl_popup.asp" onSubmit="return Form1_Validator(this)">
        <table width="100%" border="0" cellspacing="0" cellpadding="4" align="center">
          <tr> 
            <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Upload
              Data File(s)</font></b></font></td>
          </tr>
          <tr> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="2">Only *.txt, *.htm, *.html, *.gif, *.jpg, *.pdf, *.doc and *.zip file types may be uploaded.</font></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"></td>
          </tr>
          <tr> 
            <td width="20%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">File
                1: </font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="one" size="25">
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">File
                2:</font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="two" size="25">
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">File
                3:</font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="three" size="25">
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">File
                4:</font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="four" size="25">
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">File
                5:</font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="five" size="25">
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">File
                6:</font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="six" size="25">
              </font></b></td>
          </tr>
          <tr> 
            <td colspan="2" height="15"></td>
          </tr>
          <tr> 
            <td colspan="2"> 
              <div align="left"> 
                <p><font face="Arial, Helvetica, sans-serif" size="2">
                  <input type="submit" name="Submit" value="Upload">
                  <input type="button" value="Close Window" onClick="javascript:window.close();">
                  </font></p>
              </div>
            </td>
          </tr>
        </table>
	</form>
</body>
</html>