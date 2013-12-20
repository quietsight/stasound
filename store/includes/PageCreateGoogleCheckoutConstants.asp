<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="shipFromSettings.asp"-->
<!--#include file="secureadminfolder.asp"-->
<!--#include file="rc4.asp"--> 

<% 
'// Check permissions on include folder
Dim q, PageName, findit, Body, f, fso

'// Request values
q=Chr(34)
PageName="GoogleCheckoutConstants.asp"
findit=Server.MapPath(PageName)

'// Form the Body
Body=CHR(60)&CHR(37)&"private const GOOGLEBTNSIZE="&q& "small" &q&CHR(10)
Body=Body & "private const GOOGLECURRENCY="&q& Session("pcAdminGoogleCurrency") &q&CHR(10)
Body=Body & "private const GOOGLEEXPIREDAYS="& 60 &CHR(10)
Body=Body & "private const GOOGLETAXSHIPPING="&q& Session("pcAdminGoogleTaxShipping") &q&CHR(10)
Body=Body & "private const GOOGLELOGGING="&q& "3" &q&CHR(10)
if Session("pcAdminDelete")="YES" then
	pcv_strGOOGLEACTIVE = 0
else
	pcv_strGOOGLEACTIVE = -1
end if
Session("pcAdminDelete")=""
Body=Body & "private const GOOGLEACTIVE="& pcv_strGOOGLEACTIVE &CHR(10)
Body=Body & "private const GOOGLEMERCHANTID="&q& Session("pcAdminmerchantID") &q&CHR(10)
Body=Body & "private const GOOGLEMERCHANTKEY="&q& Session("pcAdminmerchantKey") &q&CHR(10)
Body=Body & "private const GOOGLESANDBOXID="&q& Session("pcAdminSandboxMerchantID") &q&CHR(10)
Body=Body & "private const GOOGLESANDBOXKEY="&q& Session("pcAdminSandboxMerchantKey") &q&CHR(10)
Body=Body & "private const GOOGLETESTMODE="&q& Session("pcAdminGoogleTestMode") &q&CHR(10)
Body=Body & "private const GOOGLEPROCESS="& Session("pcAdminpcv_processOrder") &CHR(10)
Body=Body & "private const GOOGLEPAYSTATUS="& Session("pcAdminpcv_setPayStatus") &CHR(10)&CHR(37)&CHR(62)	

'// Write the Body
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

response.redirect "../"&scAdminFolderName&"/ConfigureGoogleCheckout2.asp?msg="&strErr
%>