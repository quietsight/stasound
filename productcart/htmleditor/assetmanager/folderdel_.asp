<%
If Not request("inpCurrFolder")="" Then
  sFolder = request("inpCurrFolder")

  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  If fso.FolderExists(sFolder) Then
    sDestination = fso.GetParentFolderName(sFolder)

    fso.DeleteFolder(sFolder)
    sMsq = "<script>document.write(getTxt('Folder deleted.'))</script>"
  Else
    sMsq = "<script>document.write(getTxt('Folder does not exist.'))</script>"
  End If
  Set fso = Nothing
End If
%>
<base target="_self">

<!--#Include File="../../includes/storeconstants.asp"-->
<!--#Include File="../../includes/productcartFolder.asp"-->
<%
response.Expires=0
if session("admin")<>"-1" then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
	response.redirect(tempURL)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
<script>
  var sLang=parent.oUtil.langDir;
  document.write("<scr"+"ipt src='language/"+sLang+"/folderdel_.js'></scr"+"ipt>");
</script>
<script>writeTitle()</script>
<script>
function refresh()
  {
    (opener?opener:openerWin).refreshAfterDelete(document.getElementById("inpDest").value);
  }
</script>
</head>
<body onload="loadTxt()" style="overflow:hidden;margin:0px;">

<table width=100% height=100% align=center style="" cellpadding=0 cellspacing=0 ID="Table1">
<tr>
<td valign=top style="padding-top:5px;padding-left:15px;padding-right:15px;padding-bottom:12px;height:100%">

  <br>
  <input type="hidden" ID="inpDest" NAME="inpDest" value="<%=sDestination%>">
  <div><b><%=sMsq%>&nbsp;</b></div>

</td>
</tr>
<tr>
<td class="dialogFooter" align="right">
  <table cellpadding="1" cellspacing="0">
    <tr>
    <td>
      <input type="button" name="btnCloseAndRefresh" id="btnCloseAndRefresh" value="close & refresh" onclick="refresh();if(self.closeWin)self.closeWin();else self.close();" class="inpBtn" onmouseover="this.className='inpBtnOver';" onmouseout="this.className='inpBtnOut'">
    </td>
    </tr>
  </table>
</td>
</tr>
</table>


</body>
</html>