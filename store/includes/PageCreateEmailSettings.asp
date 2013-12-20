<!--#include file="adminv.asp"-->
<!--#include file="settings.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<!--#include file="openDb.asp"--> 
<%
dim query, conntemp, rs
'on error resume next
call openDb()

Dim PageName, Body
Dim FS, f, findit
dim tLocalOrRemote, tPort, tCustServEmail

' request values
tLocalOrRemote = Request.form("optLocalRemote")
tPort = Request.form("optPort")
q=Chr(34)
townerEmail=request.form("ownerEmail")
tfrmEmail=request.form("frmEmail")
tCustServEmail=request.form("CustServEmail")
if trim(tCustServEmail)="" then tCustServEmail=tfrmEmail

tNoticeNewCust=request.form("NoticeNewCust")
if tNoticeNewCust="" then
	tNoticeNewCust="0"
end if


tConfirmEmail=replace(request.form("ConfirmEmail"),"""","&quot;")
tConfirmEmail=replace(tConfirmEmail,"'","''")
tConfirmEmail=replace(tConfirmEmail, vbCrLf, "<br>")

tReceivedEmail=replace(request.form("ReceivedEmail"),"""","&quot;")
tReceivedEmail=replace(tReceivedEmail,"'","''")
tReceivedEmail=replace(tReceivedEmail, vbCrLf, "<br>")
If tReceivedEmail="" then
	tReceivedEmail=""
End if
tShippedEmail=replace(request.form("ShippedEmail"),"""","&quot;")
tShippedEmail=replace(tShippedEmail,"'","''")
tShippedEmail=replace(tShippedEmail, vbCrLf, "<br>")

tCancelledEmail=request.form("CancelledEmail")
If tCancelledEmail="" then
	tCancelledEmail="This message is to inform you that order number <ORDER_ID> that you submitted in this store on <ORDER_DATE> has been cancelled."
End if
tCancelledEmail=replace(tCancelledEmail,"""","&quot;")
tCancelledEmail=replace(tCancelledEmail,"'","''")
tCancelledEmail=replace(tCancelledEmail, vbCrLf, "<br>")
tCancelledEmail=replace(tCancelledEmail, vbCrLf, "<br>")
tEmailComObj=request.form("EmailComObj")

tSMTPAuthenticationTemp=request.form("SmtpAuth")
	if tSMTPAuthenticationTemp=1 then
		tSMTPAuthentication="Y"
	else
		tSMTPAuthentication="N"
	end if
	
tSMTPUID=request.Form("SmtpAuthUID")
tSMTPPWD=request.Form("SmtpAuthPWD")	

tSMTP=request.form("SMTP")

tPayPalEmail=replace(request.form("PayPalEmail"),"'","''")
If tPayPalEmail="" then
	tPayPalEmail="We have received your order and we are awaiting payment confirmation from PayPal, the payment option that you selected. As soon as payment confirmation is received, your order will be processed and you will receive an order receipt at this email address."
end if
tPayPalEmail=replace(tPayPalEmail, vbCrLf, "<br>")
query="UPDATE emailsettings SET ownerEmail='"&townerEmail&"',frmEmail='"&tfrmEmail&"',ConfirmEmail='"&tConfirmEmail&"',ReceivedEmail='"&tReceivedEmail&"',ShippedEmail='"&tShippedEmail&"',CancelledEmail='"&tCancelledEmail&"',PayPalEmail='"&tPayPalEmail&"' WHERE id=1"

set rs=Server.CreateObject("ADODB.Recordset")     
set rs=conntemp.execute(query)

if err.number <> 0 then
    response.write "Error in PageCreateEmailSettings.asp: "&Err.Description
end if

PageName=request.form("page_name")
findit=Server.MapPath(PageName)
Body=CHR(60)&CHR(37)&"private const scEmail="&q&townerEmail&q&CHR(10)
Body=Body & "private const scFrmEmail="&q&tfrmEmail&q&CHR(10)
Body=Body & "private const scCustServEmail="&q&tCustServEmail&q&CHR(10)
Body=Body & "private const scEmailComObj="&q&tEmailComObj&q&CHR(10)
Body=Body & "private const scSMTP="&q&tSMTP&q&CHR(10)
Body=Body & "private const scLocalOrRemote="&q&tLocalOrRemote&q&CHR(10)
Body=Body & "private const scPort="&q&tPort&q&CHR(10)
Body=Body & "private const scSMTPAuthentication="&q&tSMTPAuthentication&q&CHR(10)
Body=Body & "private const scSMTPUID="&q&tSMTPUID&q&CHR(10)
Body=Body & "private const scSMTPPWD="&q&tSMTPPWD&q&CHR(10)
Body=Body & "private const scConfirmEmail="&q&tConfirmEmail&q&CHR(10)
Body=Body & "private const scReceivedEmail="&q&tReceivedEmail&q&CHR(10)
Body=Body & "private const scShippedEmail="&q&tShippedEmail&q&CHR(10)
Body=Body & "private const scNoticeNewCust="&q&tNoticeNewCust&q&CHR(10)
Body=Body & "private const scCancelledEmail="&q&tCancelledEmail&q&CHR(37)&CHR(62)

' create the file using the FileSystemObject
on error resume next
Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&scAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Email")
end if
Set f=nothing

Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body
f.Close
Set fso=nothing
Set f=nothing

response.redirect "../"&scAdminFolderName&"/emailsettings.asp?s=1&message="&Server.URLEncode("The e-mail settings were updated successfully.")
%>
