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
<base target="_self">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
<script>
  var sLang=parent.oUtil.langDir;
  document.write("<scr"+"ipt src='language/"+sLang+"/folderdel.js'></scr"+"ipt>");
</script>
<script>writeTitle()</script>
<script>
function del()
  {
  var Form1 = document.forms.Form1;
  Form1.elements.inpCurrFolder.value=(opener?opener:openerWin).document.getElementById("selCurrFolder").value;
  Form1.submit();
  }
</script>
</head>
<body onload="loadTxt()" style="overflow:hidden;margin:0px;">

<table width=100% height=100% align=center style="" cellpadding=0 cellspacing=0>
<tr>
<td valign=top style="padding-top:5px;padding-left:15px;padding-right:15px;padding-bottom:12px;height:100%">

<form method=post action="folderdel_.asp"  name="Form1" id="Form1">
  <br>
  <input type="hidden" ID="inpCurrFolder" NAME="inpCurrFolder">
  <div><b><span id=txtLang>Are you sure you want to delete this folder?</span></b></div>
</form>

</td>
</tr>
<tr>
<td class="dialogFooter" align=right>
  <table cellpadding="1" cellspacing="0">
    <tr>
    <td>
      <input type="button" name="btnClose" id="btnClose" value="close" onclick="if(self.closeWin)self.closeWin();else self.close();" class="inpBtn" onmouseover="this.className='inpBtnOver';" onmouseout="this.className='inpBtnOut'">&nbsp;<input type="button" name="btnDelete" id="btnDelete" value="delete" onclick="del()" class="inpBtn" onmouseover="this.className='inpBtnOver';" onmouseout="this.className='inpBtnOut'">
    </td>
    </tr>
  </table>
</td>
</tr>
</table>

</body>
</html>