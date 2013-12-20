<%
'**************************************************************************************************
' VSP Direct Kit 3D Callback Redirection page
'**************************************************************************************************

'**************************************************************************************************
' Change history
' ==============
'
' 13/09/2007 - Mat Peck - Original Version
'**************************************************************************************************
response.clear
strPaRes=request.form("PaRes")
strMD=request.form("MD")
strVendorTxCode=cleaninput(request.querystring("VendorTxCode"),"VendorTxCode")
Session("VendorTxCode")=strVendorTxCode
%>

<SCRIPT LANGUAGE="Javascript"> function OnLoadEvent() { document.form.submit(); } </SCRIPT> 
<HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>3D-Secure Redirect</title>
</head>

<body OnLoad="OnLoadEvent();">
<%
	response.write "<FORM name=""form"" action=""gwProtx_3DComplete.asp"" method=""POST"" target=""_top""/>"
	response.write "<input type=""hidden"" name=""PARes"" value=""" & strPaRes &"""/>"
	response.write "<input type=""hidden"" name=""MD"" value=""" & strMD &"""/>"
	response.write "<NOSCRIPT>" 
	response.write "<center><p>Please click button below to Authorise your card</p><input type=""submit"" value=""Go""/></p></center>"
	response.write "</NOSCRIPT>" 
	response.write "</form>"


'** Filters unwanted characters out of an input string.  Useful for tidying up FORM field inputs
public function cleanInput(strRawText,strType)

	if strType="Number" then
		strClean="0123456789."
		bolHighOrder=false
	elseif strType="VendorTxCode" then
		strClean="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_."
		bolHighOrder=false
	else
  		strClean=" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,'/{}@():?-_&£$=%~<>*+""" & vbCRLF
		bolHighOrder=true
	end if

	strCleanedText=""
	iCharPos = 1

	do while iCharPos<=len(strRawText)
    	'** Only include valid characters **
		chrThisChar=mid(strRawText,iCharPos,1)

		if instr(StrClean,chrThisChar)<>0 then 
			strCleanedText=strCleanedText & chrThisChar
		elseif bolHighOrder then
			'** Fix to allow accented characters and most high order bit chars which are harmless **
			if asc(chrThisChar)>=191 then strCleanedText=strCleanedText & chrThisChar
		end if

      	iCharPos=iCharPos+1
	loop       
  
	cleanInput = trim(strCleanedText)

end function
%>
</body>
</html>
