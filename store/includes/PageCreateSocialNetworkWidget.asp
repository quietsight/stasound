<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<% 
'// Check permissions on include folder
Dim q, PageName, findit, Body, f, fso

'// Request values
q=Chr(34)
PageName="SocialNetworkWidgetConstants.asp"
findit=Server.MapPath(PageName)

tSNW_TYPE=Session("SNW_TYPE")
tSNW_CATEGORY=Session("SNW_CATEGORY")
tSNW_MAX=Session("SNW_MAX")
tSNW_AFFILIATE=Session("SNW_AFFILIATE")

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const SNW_TYPE="&q&tSNW_TYPE&q&CHR(10)
Body=Body & "private const SNW_MAX="&q&tSNW_MAX&q&CHR(10)
Body=Body & "private const SNW_AFFILIATE="&q&tSNW_AFFILIATE&q&CHR(10)
Body=Body & "private const SNW_CATEGORY="&q&tSNW_CATEGORY&q&CHR(37)&CHR(62) 

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
	response.redirect "../"&scAdminFolderName&"/genSocialNetworkWidget.asp?msg=success"
end if
%>