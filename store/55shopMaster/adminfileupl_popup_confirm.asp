<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%session("uploaded")="1"%>
<html>
<title>Upload Files to Your Store</title>
<script language="Javascript">
function winclose()
{
opener.document.hForm.uploaded.value="1";
opener.document.hForm.submit();
self.close();
}
</script>
<body onUnload="javascript:winclose();" style="margin: 0; padding: 0; background-color:#FFF;">
	<table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><strong>Upload
            Data File(s) to Your Store</strong></font></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"></td>
        </tr>
        <tr> 
          <td colspan="2">
            <p>&nbsp;</p>
            <p align="center"><b><font face="Arial, Helvetica, sans-serif" size="3">File(s) 
              uploaded<br>
              successfully!</font></b></p>
            <p>&nbsp;</p>
          </td>
        </tr>
        <tr> 
          <td colspan="2"> 
            <div align="left"> 
              <p align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                <input type="button" value="Close Window" onClick="javascript:winclose();">
                </font></p>
              <p align="center">&nbsp;</p>
              <p align="center">&nbsp;</p>
              <p align="center">&nbsp;</p>
              <p>&nbsp;</p>
            </div>
          </td>
        </tr>
      </table>

</body>
</html>