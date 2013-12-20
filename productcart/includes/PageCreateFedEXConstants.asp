<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<% 

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="FedEXconstants.asp"
findit=Server.MapPath(PageName)
FEDEX_FEDEX_PACKAGE=Session("ship_FEDEX_FEDEX_PACKAGE")
FEDEX_HEIGHT=Session("ship_FEDEX_HEIGHT")
FEDEX_WIDTH=Session("ship_FEDEX_WIDTH")
FEDEX_LENGTH=Session("ship_FEDEX_LENGTH")
FEDEX_DROPOFF_TYPE=Session("ship_FEDEX_DROPOFF_TYPE")
FEDEX_DIM_UNIT=Session("ship_FEDEX_DIM_UNIT")
FEDEX_LISTRATE=Session("ship_FEDEX_LISTRATE")
FDX_DYNAMICINSUREDVALUE=Session("ship_FEDEX_DYNAMICINSUREDVALUE")
FDX_INSUREDVALUE=Session("ship_FEDEX_INSUREDVALUE")

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const FEDEX_FEDEX_PACKAGE="&q&FEDEX_FEDEX_PACKAGE&q&CHR(10)
Body=Body & "private const FEDEX_HEIGHT="&q&FEDEX_HEIGHT&q&CHR(10)
Body=Body & "private const FEDEX_WIDTH="&q&FEDEX_WIDTH&q&CHR(10)
Body=Body & "private const FEDEX_LENGTH="&q&FEDEX_LENGTH&q&CHR(10)
Body=Body & "private const FEDEX_DROPOFF_TYPE="&q&FEDEX_DROPOFF_TYPE&q&CHR(10)
Body=Body & "private const FEDEX_LISTRATE="&q&FEDEX_LISTRATE&q&CHR(10)
Body=Body & "private const FDX_INSUREDVALUE="&q&FDX_INSUREDVALUE&q&CHR(10)
Body=Body & "private const FDX_DYNAMICINSUREDVALUE="&q&FDX_DYNAMICINSUREDVALUE&q&CHR(10)
Body=Body & "private const FEDEX_DIM_UNIT="&q&FEDEX_DIM_UNIT&q&CHR(37)&CHR(62)

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

Session("ship_FEDEX_FEDEX_PACKAGE")=""
Session("ship_FEDEX_HEIGHT")=""
Session("ship_FEDEX_WIDTH")=""
Session("ship_FEDEX_LENGTH")=""
Session("ship_FEDEX_DROPOFF_TYPE")=""
Session("ship_FEDEX_DIM_UNIT")=""

if request.QueryString("refer")<>"" then
	response.redirect "../"&scAdminFolderName&"/"&request.QueryString("refer")
else
	response.redirect "../"&scAdminFolderName&"/3_Step6.asp"
end if
%>