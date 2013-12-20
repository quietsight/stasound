<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<% 
'// Check permissions on include folder
Dim q, PageName, findit, Body, f, fso

'// Request values
q=Chr(34)
PageName="SearchConstants.asp"
findit=Server.MapPath(PageName)

tSRCH_MAX=Session("SRCH_MAX")
tSRCH_CSFON=Session("SRCH_CSFON")
tSRCH_CSFRON=Session("SRCH_CSFRON")
tSRCH_WAITBOX=Session("SRCH_WAITBOX")
tSRCH_SUBS=Session("SRCH_SUBS")

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const SRCH_MAX="&q&tSRCH_MAX&q&CHR(10)
Body=Body & "private const SRCH_CSFON="&q&tSRCH_CSFON&q&CHR(10)
Body=Body & "private const SRCH_CSFRON="&q&tSRCH_CSFRON&q&CHR(10)
Body=Body & "private const SRCH_SUBS="&q&tSRCH_SUBS&q&CHR(10)
Body=Body & "private const SRCH_WAITBOX="&q&tSRCH_WAITBOX&q&CHR(37)&CHR(62) 

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
	response.redirect "../"&scAdminFolderName&"/SearchOptions.asp?msg=success"
end if
%>