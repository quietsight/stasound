<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/MailUpFunctions.asp"-->
<%Dim connTemp,rs,query

call opendb()

ListID=getUserInput(request("ListID"),0)
if not IsNumeric(ListID) then
	response.redirect "CustPref.asp"
end if

tmpRes=RegisterSingleUser(Session("idCustomer"),"",ListID,session("SF_MU_URL"),session("SF_MU_Auto"))
%>
<html>
<head>
<title>Re-send subscription confirmation message</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 0;">
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td align="center">
			<div class="pcErrorMessage">
			<%if tmpRes="1" then%>
				<%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel3")%>
			<%else%>
				<%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel4")%>
				<%if MU_ErrMsg<>"" then%>
					<br>
					<%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel4a")%><%=MU_ErrMsg%>
				<%end if%>
			<%end if%>
			</div>
		</td>
	</tr>
	<tr> 
		<td colspan="2" align="right" style="padding-top: 20px;">
		<input type="image" value="Close Window" src="images/close.gif" width="32" height="25" onClick="self.close()">
    </td>
</table>
</div>
</body>
</html>
