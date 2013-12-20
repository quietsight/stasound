<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<%

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="FedEXWSconstants.asp"
findit=Server.MapPath(PageName)
FEDEXWS_FEDEX_PACKAGE=Session("ship_FEDEXWS_FEDEX_PACKAGE")
FEDEXWS_HEIGHT=Session("ship_FEDEXWS_HEIGHT")
FEDEXWS_WIDTH=Session("ship_FEDEXWS_WIDTH")
FEDEXWS_LENGTH=Session("ship_FEDEXWS_LENGTH")
FEDEXWS_DROPOFF_TYPE=Session("ship_FEDEXWS_DROPOFF_TYPE")
FEDEXWS_DIM_UNIT=Session("ship_FEDEXWS_DIM_UNIT")
FEDEXWS_LISTRATE=Session("ship_FEDEXWS_LISTRATE")
FEDEXWS_SATURDAYDELIVERY=Session("ship_FEDEXWS_SATURDAYDELIVERY")
FEDEXWS_SATURDAYPICKUP=Session("ship_FEDEXWS_SATURDAYPICKUP")
FDXWS_DYNAMICINSUREDVALUE=Session("ship_FEDEXWS_DYNAMICINSUREDVALUE")
FDXWS_INSUREDVALUE=Session("ship_FEDEXWS_INSUREDVALUE")
FDXWS_SMHUBID=Session("ship_FEDEXWS_SMHUBID")
FEDEXWS_ADDDAY=Session("ship_FEDEXWS_ADDDAY")

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const FEDEXWS_FEDEX_PACKAGE="&q&FEDEXWS_FEDEX_PACKAGE&q&CHR(10)
Body=Body & "private const FEDEXWS_HEIGHT="&q&FEDEXWS_HEIGHT&q&CHR(10)
Body=Body & "private const FEDEXWS_WIDTH="&q&FEDEXWS_WIDTH&q&CHR(10)
Body=Body & "private const FEDEXWS_LENGTH="&q&FEDEXWS_LENGTH&q&CHR(10)
Body=Body & "private const FEDEXWS_DROPOFF_TYPE="&q&FEDEXWS_DROPOFF_TYPE&q&CHR(10)
Body=Body & "private const FEDEXWS_LISTRATE="&q&FEDEXWS_LISTRATE&q&CHR(10)
Body=Body & "private const FEDEXWS_SATURDAYDELIVERY="&q&FEDEXWS_SATURDAYDELIVERY&q&CHR(10)
Body=Body & "private const FEDEXWS_SATURDAYPICKUP="&q&FEDEXWS_SATURDAYPICKUP&q&CHR(10)
Body=Body & "private const FDXWS_INSUREDVALUE="&q&FDXWS_INSUREDVALUE&q&CHR(10)
Body=Body & "private const FDXWS_SMHUBID="&q&FDXWS_SMHUBID&q&CHR(10)
Body=Body & "private const FDXWS_DYNAMICINSUREDVALUE="&q&FDXWS_DYNAMICINSUREDVALUE&q&CHR(10)
Body=Body & "private const FEDEXWS_DIM_UNIT="&q&FEDEXWS_DIM_UNIT&q&CHR(10)
Body=Body & "private const FEDEXWS_ADDDAY="&q&FEDEXWS_ADDDAY&q&CHR(37)&CHR(62)

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

Session("ship_FEDEXWS_FEDEX_PACKAGE")=""
Session("ship_FEDEXWS_HEIGHT")=""
Session("ship_FEDEXWS_WIDTH")=""
Session("ship_FEDEXWS_LENGTH")=""
Session("ship_FEDEXWS_DROPOFF_TYPE")=""
Session("ship_FEDEXWS_DIM_UNIT")=""
Session("ship_FEDEXWS_ADDDAY")=""

if request.QueryString("refer")<>"" then
	response.redirect "../"&scAdminFolderName&"/"&request.QueryString("refer")
else
	response.redirect "../"&scAdminFolderName&"/3_Step6.asp"
end if
%>