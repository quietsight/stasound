<html>
<head>
<link rel="stylesheet" type="text/css" href="style.css">
<title>Show Picture</title>

<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
var isNav4, isIE4;
if (parseInt(navigator.appVersion.charAt(0)) >= 4) {
isNav4 = (navigator.appName == "Netscape") ? 1 : 0;
isIE4 = (navigator.appName.indexOf("Microsoft") != -1) ? 1 : 0;
}
function fitWindowSize() {
if (isNav4) {
window.innerWidth = document.layers[0].document.images[0].width;
window.innerHeight = document.layers[0].document.images[0].height;
}
if (isIE4) {
window.resizeTo(800, 600);
width = (800 - (document.body.clientWidth -  document.images[0].width)) + 25;
height = (600 - (document.body.clientHeight -  document.images[0].height)) + 55;
window.resizeTo(width, height);
   }
}
//  End -->
</script>

</head>

<BODY onLoad="fitWindowSize()">

<center>
<%
ShowPic = request("ShowPic")
ShowPic = Replace(ShowPic, " ", "%20")
Response.Write "<img src='" & ShowPic & "' border=0><BR>"
response.write "<a href='JavaScript:onClick= window.close()'><font face='verdana' size=1 color='#336699'>Close Window</font></A>"
%>
</center>
</body>
</html>

