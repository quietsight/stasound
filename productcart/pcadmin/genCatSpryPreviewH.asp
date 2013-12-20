<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<!--Link the Spry Manu Bar JavaScript library-->
<script src="../includes/spry/SpryMenuBar.js" type="text/javascript"></script>
<!--Link the CSS style sheet that styles the menu bar. You can select between horizontal and vertical-->
<link href="../includes/spry/SpryMenuBarHorizontal.css" rel="stylesheet" type="text/css" />
<style>
* {
	font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 12px;
}
</style>
</head>
<body>
<div style="margin: 20px 0 20px 0;">
For more information on how to use the Spry Menu Widget, please see <a href="http://labs.adobe.com/technologies/spry/articles/menu_bar/index.html" target="_blank">Working with the Menu Bar Widget</a>.
</div>
<!--Create a Menu bar widget and assign classes to each element-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../pc/inc_catsmenu.asp"-->
<!--Initialize the Menu Bar widget object-->
<script type="text/javascript">
	var menubar99 = new Spry.Widget.MenuBar("menubar99", {imgDown:"images/SpryMenuBarDownHover.gif", imgRight:"images/SpryMenuBarRightHover.gif"});
</script>
</body>
</html>