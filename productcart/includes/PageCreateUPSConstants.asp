<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<% 

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="UPSconstants.asp"
findit=Server.MapPath(PageName)

UPS_PICKUP_TYPE=Session("ship_UPS_PICKUP_TYPE")
UPS_PACKAGE_TYPE=Session("ship_UPS_PACKAGE_TYPE")
UPS_CLASSIFICATION_TYPE=Session("ship_UPS_CLASSIFICATION_TYPE")
UPS_HEIGHT=Session("ship_UPS_HEIGHT")
UPS_WIDTH=Session("ship_UPS_WIDTH")
UPS_LENGTH=Session("ship_UPS_LENGTH")
UPS_DIM_UNIT=Session("ship_UPS_DIM_UNIT")

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const UPS_PICKUP_TYPE="&q&UPS_PICKUP_TYPE&q&CHR(10)
Body=Body & "private const UPS_PACKAGE_TYPE="&q&UPS_PACKAGE_TYPE&q&CHR(10)
Body=Body & "private const UPS_CLASSIFICATION_TYPE="&q&UPS_CLASSIFICATION_TYPE&q&CHR(10)
Body=Body & "private const UPS_HEIGHT="&q&UPS_HEIGHT&q&CHR(10)
Body=Body & "private const UPS_WIDTH="&q&UPS_WIDTH&q&CHR(10)
Body=Body & "private const UPS_LENGTH="&q&UPS_LENGTH&q&CHR(10)
Body=Body & "private const UPS_DIM_UNIT="&q&UPS_DIM_UNIT&q&CHR(37)&CHR(62) 

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
	response.redirect "../"&scAdminFolderName&"/1_Step6.asp"
end if
%>