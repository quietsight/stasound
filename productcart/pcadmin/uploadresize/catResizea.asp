<% pageTitle="Image Upload & Auto Resize" %>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../../includes/settings.asp"-->
<!--#include file="../../includes/storeconstants.asp"--> 
<!--#include file="../../includes/opendb.asp"-->
<!--#include file="../../includes/adovbs.inc"-->
<!--#include file="../../includes/productcartFolder.asp"-->

<% dim mySQL, conntemp, rstemp
on error resume next
%>

<% Dim PID, barref%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>Upload Images</title>
	<link href="../pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-image: none;">
<!--#include file="checkImgUplResizeObjs.asp"-->
<%If HaveImgUplResizeObjs=0 then%>
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessage">We are unable to find compatible Upload and/or Image Resize server components. Please consult the User Guide for detailed system requirements.</div>
	</td>
</tr>
<tr>
	<td align="center"><input type="button" name="Close" value=" Close window " onClick="javascript:window.close();" class="ibtnGrey"></td>
</tr>
</table>
<%Else%>
<script language="javascript">
 var submitted = false;
 function check_submit(theform) {
   if (submitted) return false;
   theform.Submit.disabled=true;
   theform.Submit.value="Uploading...";
   return (submitted = true);
 }
</script>
<form action="catResizeb.asp" name="MyForm" method="post" enctype="multipart/form-data" onSubmit="return check_submit(this)">
	<table width="400" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
    <tr> 
      <td> 
        <table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
          <tr> 
            <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
          </tr>
          <tr> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td height="18" colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"> 
              Select an image using the &quot;Browse&quot; button. Then click 
              on &quot;Upload&quot;. All images are automatically uploaded to 
              the &quot;<b><%= scPcFolder %>/pc/catalog</b>&quot; folder on your Web 
              server and sizes are set.</font></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"></td>
          </tr>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Normal Size: </font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="text" name="normalsize" size="3" maxlength="3" class="ibtng" value="100"> pixels
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Large Size: </font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="text" name="largesize" size="3" maxlength="3" class="ibtng" value="200"> pixels
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Image: </font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="FILE1" size="25">
              </font></b></td>
          </tr>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Resize Based On: </font></div>
            </td>
            <td width="80%"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="radio" name="resizexy" value="Width" checked> Width&nbsp;&nbsp;&nbsp;<input type="radio" name="resizexy" value="Height" > Height
              </font></td>
          </tr>			  
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Sharpen Image: </font></div>
            </td>
            <td width="80%"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="radio" name="sharpen" value="1"> Yes&nbsp;&nbsp;&nbsp;<input type="radio" name="sharpen" value="0" checked> No
              </font></td>
          </tr>		  
          <tr> 
            <td colspan="2" height="15"></td>
          </tr>

          <tr>
            <td width="20%">&nbsp;</td>		  
            <td width="80%"> 
              <div align="left"> 
               <font face="Arial, Helvetica, sans-serif" size="2"> 
                  <input type="submit" name="Submit" value="Upload" class="ibtnGrey">
                  <input type="button" value="Close Window" onClick="javascript:window.close();" class="ibtnGrey">
               </font>
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
	</table>
</form>
<%End if%>
</body>
</html>
