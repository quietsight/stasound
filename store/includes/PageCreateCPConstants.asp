<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<% 

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="CPconstants.asp"
findit=Server.MapPath(PageName)

CP_Height=Session("ship_CP_Height")
CP_Width=Session("ship_CP_Width")
CP_Length=Session("ship_CP_Length")

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const CP_Height="&q&CP_Height&q&CHR(10)
Body=Body & "private const CP_Width="&q&CP_Width&q&CHR(10)
Body=Body & "private const CP_Length="&q&CP_Length&q&CHR(37)&CHR(62)

' create the file using the FileSystemObject
on error resume next
Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&scAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Constants")
end if
Set f=nothing

Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body
f.Close
Set fso=nothing
Set f=nothing

if request.QueryString("refer")<>"" then
	response.redirect "../"&scAdminFolderName&"/"&request.QueryString("refer")
else
	response.redirect "../"&scAdminFolderName&"/4_Step6.asp"
end if
%>