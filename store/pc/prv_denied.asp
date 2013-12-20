<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languages.asp"-->
<html>
	<head>
		<title><%=dictLanguage.Item(Session("language")&"_prv_9")%></title>
		<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
	</head>
	<body>
		<div id="pcMain">
			<table class="pcMainTable">
				<tr>
					<td>
						<h1><%=dictLanguage.Item(Session("language")&"_prv_9")%></h1>
						<p>&nbsp;</p>
						<p><%=dictLanguage.Item(Session("language")&"_prv_8")%></p>
						<p>&nbsp;</p>
						<form class="pcForms">
							<input type="button" value="Close window" onclick="window.close();" class="submit2">
						</form>
					</td>
				</tr>
			</table>
		</div>
	</body>
</html>