<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->

<%
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwProtx_3DCallback.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwProtx_3DCallback.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If

'**************************************************************************************************
' VSP Direct Kit 3D Redirection inline frame
'**************************************************************************************************

'**************************************************************************************************
' Change history
' ==============
'
' 13/09/2007 - Mat Peck - Original Version
'**************************************************************************************************
response.clear
strACSURL=Session("ACSURL")
strPAReq=Session("PAReq")
strMD=Session("MD")
strVendorTxCode=Session("VendorTxCode")
Session("PAReq")=""
%>

<SCRIPT LANGUAGE="Javascript"> function OnLoadEvent() { document.form.submit(); } </SCRIPT>
<HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>3D-Secure Redirect</title>
</head>

<body OnLoad="OnLoadEvent();">
<%  
	response.write "<FORM name=""form"" action=""" & strACSURL &""" method=""POST"" target=""3DIFrame""/>"
	response.write "<input type=""hidden"" name=""PaReq"" value=""" & strPAReq &"""/>"
	response.write "<input type=""hidden"" name=""TermUrl"" value=""" & tempURL & "?VendorTxCode=" & strVendorTxCode & """/>"
	response.write "<input type=""hidden"" name=""MD"" value=""" & strMD &"""/>"
	response.write "<NOSCRIPT>" 
	response.write "<center><p>Please click button below to Authenticate your card</p><input type=""submit"" value=""Go""/></p></center>"
	response.write "</NOSCRIPT>" 
	response.write "</form>"
%>
</body>
</html>
