<% pageTitle="Image Upload & Auto Resize" %>
<%
  ' PRV41 Start
  Section="products" %>
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
	<title>Upload File (.TXT)</title>
	<link href="../pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-image: none;">
<script language="javascript">
 var submitted = false;
 function check_submit(theform) {
   if (submitted) return false;
   theform.Submit.disabled=true;
   theform.Submit.value="Uploading...";
   return (submitted = true);
 }
</script>
<form action="fileUploadc.asp" name="MyForm" method="post" enctype="multipart/form-data" onSubmit="return check_submit(this)">
	<table width="400" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
	<tr> 
      <td> 
        <table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
          <tr> 
            <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;File Upload (.TXT file)</font></b></font></td>
          </tr>
          <%
          If InStr(request("msg"),"success")>0 Then %>
            <script>
            function fillparentform() {
            parent.opener.document.hForm.sendreviewremindertemplate.value = '<%= replace(request("f"),"'","\'") %>'
            }

            fillparentform();

            </script>
          <%
             response.write "<tr><td colspan=""3""><br /><strong>" & request("msg") & "</strong><br /><br /></td></tr>"
          End if
          %>
          <tr> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td height="18" colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"> 
              Select a TEXT file (with a .TXT extension) using the &quot;Browse&quot; button. Then click 
              on &quot;Upload&quot;. All files are automatically uploaded to 
              the &quot;<b><%= scPcFolder %>/pc/library</b>&quot; folder on your Web 
              server.</font></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"></td>
          </tr>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Text File: </font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="one" size="25">
              </font></b></td>
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
</body>
</html>
<% 'PRV41 end %>