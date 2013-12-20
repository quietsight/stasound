<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<% 

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="USPSconstants.asp"
findit=Server.MapPath(PageName)

tUSPS_EM_PACKAGE=Session("ship_USPS_EM_PACKAGE")
tUSPS_PM_PACKAGE=Session("ship_USPS_PM_PACKAGE")
tUSPS_HEIGHT=Session("ship_USPS_HEIGHT")
tUSPS_WIDTH=Session("ship_USPS_WIDTH")
tUSPS_LENGTH=Session("ship_USPS_LENGTH")
tUSPS_EM_FREWeightLimit=Session("ship_USPS_EM_FREWeightLimit")
tUSPS_EM_FREOption=Session("ship_USPS_EM_FREOption")
tUSPS_PM_FREWeightLimit=Session("ship_USPS_PM_FREWeightLimit")
tUSPS_PM_FREOption=Session("ship_USPS_PM_FREOption")


Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const USPS_EM_PACKAGE="&q&tUSPS_EM_PACKAGE&q&CHR(10)
Body=Body & "private const USPS_PM_PACKAGE="&q&tUSPS_PM_PACKAGE&q&CHR(10)
Body=Body & "private const USPS_HEIGHT="&q&tUSPS_HEIGHT&q&CHR(10)
Body=Body & "private const USPS_WIDTH="&q&tUSPS_WIDTH&q&CHR(10)
Body=Body & "private const USPS_LENGTH="&q&tUSPS_LENGTH&q&CHR(10)
Body=Body & "private const USPS_EM_FREWeightLimit="&q&tUSPS_EM_FREWeightLimit&q&CHR(10)
Body=Body & "private const USPS_EM_FREOption="&q&tUSPS_EM_FREOption&q&CHR(10)
Body=Body & "private const USPS_PM_FREWeightLimit="&q&tUSPS_PM_FREWeightLimit&q&CHR(10)
Body=Body & "private const USPS_PM_FREOption="&q&tUSPS_PM_FREOption&q&CHR(37)&CHR(62) 

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
	response.redirect "../"&scAdminFolderName&"/2_Step6.asp"
end if
%>