<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="shipFromSettings.asp"-->
<!--#include file="secureadminfolder.asp"-->
<!--#include file="rc4.asp"--> 

<% Dim pmode

pmode=request("mode")
pShipFromWeightUnit=Session("pcv_WeightUnit")
if pShipFromWeightUnit="" then
	pShipFromWeightUnit="LBS"
end if

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="shipFromSettings.asp"
findit=Server.MapPath(PageName)

Body=CHR(60)&CHR(37)&"private const scShipFromName="&q&q&CHR(10)
Body=Body & "private const scShipFromAddress1="&q&q&CHR(10)
Body=Body & "private const scShipFromAddress2="&q&q&CHR(10)
Body=Body & "private const scShipFromAddress3="&q&q&CHR(10)
Body=Body & "private const scShipFromCity="&q&q&CHR(10)
Body=Body & "private const scShipFromState="&q&q&CHR(10)
Body=Body & "private const scShipFromPostalCode="&q&q&CHR(10)
Body=Body & "private const scShipFromZip4="&q&q&CHR(10)
Body=Body & "private const scAlwAltShipAddress="&q&"0"&q&CHR(10)
Body=Body & "private const scComResShipAddress="&q&"0"&q&CHR(10)
Body=Body & "private const scAlwNoShipRates="&q&"0"&q&CHR(10)
Body=Body & "private const scShipFromPostalCountry="&q&q&CHR(10)
Body=Body & "private const scShowProductWeight="&q&"0"&q&CHR(10)
Body=Body & "private const scShowCartWeight="&q&"0"&q&CHR(10)

'Start SDBA
Body=Body & "private const scShipNotifySeparate=" &q& "0" &q& CHR(10)
'End SDBA

Body=Body & "private const scShipFromWeightUnit="&q&pShipFromWeightUnit&q&CHR(10)&CHR(37)&CHR(62)

'on error resume next
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
Session("admin")=0
Response.Cookies("AgreeLicense")="Setup"
Response.Cookies("AgreeLicense").Expires=Date()+1
MyCookiePath=Request.ServerVariables("PATH_INFO")
do while not (right(MyCookiePath,1)="/")
MyCookiePath=mid(MyCookiePath,1,len(MyCookiePath)-1)
loop
MyCookiePath=replace(MyCookiePath,"includes/","")
Response.Cookies("AgreeLicense").Path=MyCookiePath&scAdminFolderName&"/"
response.redirect "../"&scAdminFolderName&"/login_1.asp"
%>