<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<% 
'// Check permissions on include folder
Dim q, PageName, findit, Body, f, fso

'// Request values
q=Chr(34)
PageName="CashbackConstants.asp"
findit=Server.MapPath(PageName)

tLSCB_KEY=Session("LSCB_KEY")
tLSCB_STATUS=Session("LSCB_STATUS")

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const LSCB_KEY="&q&tLSCB_KEY&q&CHR(10)
Body=Body & "private const LSCB_STATUS="&q&tLSCB_STATUS&q&CHR(37)&CHR(62) 

'// Create the file using the FileSystemObject
'// On Error Resume Next
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
	response.redirect "../"&scAdminFolderName&"/pcCashback_main.asp?msg=success"
end if
%>