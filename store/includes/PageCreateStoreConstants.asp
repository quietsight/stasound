<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<!--#include file="productcartFolder.asp"-->
<!--#include file="ppdstatus.inc"-->
<!--#include file="validation.asp"--> 
<!--#include file="rc4.asp" -->

<% Dim pCrypPass, pDSN, pDB, pStoreURL, puid, ppass,pmode

pmode=request("mode")
pCrypPass=Session("pcv_KeyID")
pDSN=Session("pcv_DSN")
pDB=Session("pcv_DBType")
pStoreURL=Session("StoreURL")
puid=Session("pcv_StoreUID")
ppass=Session("pcv_StorePWD")
pIntRes=request("intRes")

Session("pcv_KeyID")=""
Session("pcv_DSN")=""
Session("pcv_DBType")=""
Session("StoreURL")=""
Session("pcv_StoreUID")=""
Session("pcv_StorePWD")=""

'// Trap Page Errors
'on error resume next

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso, Body2
' request values
q=Chr(34)
PageName="storeconstants.asp"
findit=Server.MapPath(PageName)

Body=CHR(60)&CHR(37)&"private const scCrypPass="&q&scCrypPass&q&CHR(10)
Body=Body & "private const scID="&q&scCrypPass&q&CHR(10)
Body=Body & "private const scDSN="&q&scDSN&q&CHR(10)
Body=Body & "private const scDB="&q&scDB&q&CHR(10)
Body=Body & "private const scStoreURL="&q&scStoreURL&q&CHR(37)&CHR(62) 

Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&pAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Constants. Please refer to Chapter 2 of the User Guide for more information about ""Getting Started"".")
end if
Set f=nothing

Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body

Function encodeString(input)
	Dim newStr : newStr = ""
	for i = 1 to len(input)
		newStr = newStr & chr((asc(mid(input,i,1))+8)) 
	next
	encodeString = newStr
End Function

'// Create new file with encrypted storeURL
PageName2="pcSurlLvs.asp"
findit2=Server.MapPath(PageName2)

pSURLRes=replace(pStoreURL,".","")
pSURLRes=replace(pSURLRes,"http://","")
pSURLRes=replace(pSURLRes,"https://","")
pSURLRes=replace(pSURLRes,":","")
pSURLRes=replace(pSURLRes,"\","")
pSURLRes=replace(pSURLRes,"/","")
pSURLRes=encodeString(pSURLRes)

Body2=CHR(60)&CHR(37)& vbCrLf
Body2=Body2&"private const pcv_SURLResponse = """&pSURLRes&"""" & vbCrLf
Body2=Body2&"private const pcv_ITCResponse = """&pIntRes&"""" & vbCrLf
Body2=Body2&CHR(37)&CHR(62)& vbCrLf


Set f=fso.CreateTextFile(findit2, True, false)
f.Write Body2

Set fso=nothing
Set f=nothing

'encrypt password
tpass= enDeCrypt(ppass, pCrypPass)
Dim connTemp, rs, mySQL
set connTemp=server.createobject("adodb.connection")

Err.number=0
'Open your connection
connTemp.Open pDSN
if Err.number <> 0 then
	response.redirect "../setup/techErr.asp?error="&Server.Urlencode("Error while opening database. Please refer to Chapter 2 of the User Guide for more information about ""Getting Started"".")
end if

Err.number=0
'check if id in admins already exists, update if it is else insert.
set rs=server.createobject("adodb.recordset")
mySQL="SELECT * FROM admins WHERE idadmin="&puid
set rs=conntemp.execute(mySQL)

if err.number <> 0 then
	if err.description="Invalid object name 'admins'." then
		response.redirect "../setup/techErr.asp?error="&Server.Urlencode("Before you run the setup wizard you must run the default SQL database script. Please refer to Chapter 2 of the User Guide for more information about ""Getting Started"".")
		response.End()
	else
		response.redirect "../setup/techErr.asp?error="&Server.Urlencode("An error occurred while trying to insert your registration information into your database. This is usually caused by the connection string not being correct or because you have not yet set up your DSN for your database. Please refer to Chapter 2 of the User Guide for more information about ""Getting Started"".")
		response.End()
	end if
end if

if rs.eof then
	mySQL="INSERT INTO admins (idadmin, adminname, adminlevel, adminpassword) VALUES ("& puid &", 'storeowner', '19', '"&tpass&"')"
	set rs=conntemp.execute(mySQL)
Else
	mySQL="UPDATE admins SET adminpassword='"& tpass &"',adminlevel='19' WHERE idadmin="& puid
	set rs=conntemp.execute(mySQL)
End If

set rs=nothing
connTemp.Close
set connTemp=nothing

Body=CHR(60)&CHR(37)&"private const scCrypPass="&q&pCrypPass&q&CHR(10)
Body=Body & "private const scDSN="&q&pDSN&q&CHR(10)
Body=Body & "private const scDB="&q&pDB&q&CHR(10)
Body=Body & "private const scStoreURL="&q&pStoreURL&q&CHR(37)&CHR(62) 

' create the file using the FileSystemObject
on error resume next
Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&pAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Constants. Please refer to Chapter 2 of the User Guide for more information about ""Getting Started"".")
end if

Set f=nothing

Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body
f.Close

'rename setup folder
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function
pSetupName=randomNumber(99999999)

if PPD="1" then
	findit=Server.MapPath("/"&scPcFolder&"/setup/")
else
	findit=Server.MapPath("../setup/")
end if

Set folderObject = fso.GetFolder(findit)
folderObject.Name = pSetupName

Set folderObject = Nothing
Set fso = Nothing
Set f=nothing

response.redirect "FirstPageCreateSettings.asp"
%>
</body>
</html>
