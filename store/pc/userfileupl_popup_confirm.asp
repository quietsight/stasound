<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="CustLIv.asp"-->
<%session("uploaded")="1"%>
<html>

<head>
<title>Upload Data File(s)</title>
<script language="Javascript">
function winclose()
{
opener.document.hForm.uploaded.value="1";
opener.document.hForm.submit();
self.close();
}
</script>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>

<body onUnload="javascript:winclose();">
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<h1>Upload Data File(s)</h1>
			</td>
		</tr>
		<tr>
			<td>
				<div class="pcInfoMessage">
				File(s) uploaded successfully!
				</div>
       </td>
			</tr>
      <tr> 
        <td>
					<input type="button" value="Close Window" onClick="javascript:winclose();">
         </td>
      </tr>
  </table>
</div>
</body>
</html>