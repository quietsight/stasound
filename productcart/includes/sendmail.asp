<%
Function sendMail(fromName, from, rcpt, subject, body)

body = pcf_RestoreCharacters(body)

'// E-mails are always sent from company address
'// ReplyTo uses the information submitted via the form
replyToAddress = from
replyToName = fromName
if replyToName<>"" then
	ReplyTo ="""" & replyToName & """ <" & replyToAddress & ">"
else
	ReplyTo = replyToAddress
end if
fromAddress = """" & scCompanyName & """ <" & scEmail & ">"


if scEmailComObj="CDOSYS" then
	Dim mail 
	Dim iConf 
	Dim Flds
	Dim localOrRemote 

	on error resume next 
	
	localOrRemote = scLocalOrRemote
	if(localOrRemote = "") then
	    localOrRemote = "1"
	end if
	
	Set mail = CreateObject("CDO.Message") 'calls CDO message COM object
	Set iConf = CreateObject("CDO.Configuration") 'calls CDO configuration COM object
	Set Flds = iConf.Fields
	Flds( "http://schemas.microsoft.com/cdo/configuration/sendusing") = localOrRemote   ' "1" tells cdo we're using the local smtp service, use "2" if not local
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = scSMTP
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = scPort
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 20
	Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup" 'verify that this path is correct
	Flds.Update 'updates CDO's configuration database
	'if smtp authentication is required
	'==================================
	if scSMTPAuthentication="Y" then
		Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 ' cdoBasic 
		Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = scSMTPUID
		Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = scSMTPPWD
		Flds.Update 'updates CDO's configuration database
	end if
'==================================
	Set mail.Configuration = iConf 'sets the configuration for the message
	mail.To = rcpt
	mail.From = fromAddress
	mail.ReplyTo = ReplyTo
	mail.Subject = subject 
	If session("News_MsgType")="1" Then
			HTMLBody=HTML
			mail.HTMLBody = body
	Else
			TextBody=Plain
			mail.TextBody = body
	End If
	mail.Send 'commands CDO to send the message
	if err then
		pcv_errMsg = err.Description
	end if
	set mail=nothing 
End If

if scEmailComObj="ABMailer" then
	on error resume next 
	objMail.clear
	set objMail = Server.CreateObject("ABMailer.Mailman")
	objMail.ServerAddr = scSMTP
	'if authentication is used
	if scSMTPAuthentication="Y" then
		objMail.ServerPort = scPort
		objMail.ServerLoginUserID = scSMTPUID
	end if
	objMail.SendTo = rcpt 
	objMail.FromAddress = fromAddress
	objMail.ReplyTo = ReplyTo
	objMail.MailSubject = subject
	objMail.MailMessage = body
	objMail.SendMail 
	if err then
		pcv_errMsg = err.Description
	end if
	set objMail=nothing 
End If

if scEmailComObj="Bamboo" then
	on error resume next 
	set objMail = Server.CreateObject("Bamboo.SMTP") 
	objMail.Server = scSMTP
	objMail.Rcpt = rcpt 
	objMail.From = scEmail	
	objMail.FromName = scCompanyName
	objMail.Subject = subject
	If session("News_MsgType")="1" Then
	objMail.ContentType = "text/html"
	else
	objMail.ContentType = "text/plain"
	End if
	objMail.Message = body
	objMail.Send
	if err then
		pcv_errMsg = err.Description
	end if
	set objMail=nothing 
End If


if scEmailComObj="PersitsASPMail" then
	on error resume next 
	
	'session.codepage = 65001 'UTF-8 code  'uncomment for UTF-8 code
	set objMail = server.CreateObject("Persits.MailSender")
	
	objMail.Host = scSMTP
	objMail.Port = scPort
	'if authentication is used
	if scSMTPAuthentication="Y" then
		objMail.Username = scSMTPUID
		objMail.Password = scSMTPPWD
	end if
	objMail.From = scEmail
	objMail.FromName = scCompanyName 'comment out for UTF-8 code
	'objMail.FromName 	= objMail.EncodeHeader(fromName,"utf-8")  'uncomment for UTF-8 code
	objMail.AddAddress rcpt
	objMail.AddReplyTo ReplyTo
	objMail.Subject = subject  'comment out for UTF-8 code
	'objMail.Subject 	= objMail.EncodeHeader(subject, "utf-8")  'uncomment for UTF-8 code
	objMail.Body 	= body
	
	'UTF-8 parameters
	'objMail.CharSet = "UTF-8" 'uncomment for UTF-8 code
	'objMail.ContentTransferEncoding = "Quoted-Printable" 'uncomment for UTF-8 code
	
	If session("News_MsgType")="1" Then
		objMail.IsHTML = True
	End If
	objMail.Send
	if err then
		pcv_errMsg = err.Description
	end if
	set objMail=nothing 
End If

if scEmailComObj="JMail3" then
	on error resume next 
 	Set objMail = Server.CreateObject("JMail.SMTPMail")
 	'objMail.ServerAddress=scSMTP
	objMail.Sender=fromAddress
	'objMail.SenderName=fromName
	objMail.Subject= subject
	If session("News_MsgType")="1" Then
		objMail.ContentType = "text/html"
	else
		objMail.ContentType = "text/plain"
	End if
	objMail.AddRecipient rcpt
	objMail.ReplyTo = ReplyTo
	objMail.Body	= body
	objMail.Priority = 3
	objMail.Execute
	if err then
		pcv_errMsg = err.Description
	end if
	set objMail=nothing 
End If

if scEmailComObj="JMail4" then
	on error resume next 
 	Set objMail=Server.CreateOBject( "JMail.Message" )
	objMail.Logging = true
	objMail.silent = true
	'if authentication is used
	if scSMTPAuthentication="Y" then
		objMail.MailServerPassword=scSMTPPWD
		objMail.MailServerUserName=scSMTPUID
	end if
	objMail.From = fromAddress
	'objMail.FromName = fromName
	objMail.AddRecipient rcpt, rcpt
	objMail.ReplyTo = ReplyTo
	objMail.Subject = subject
	If session("News_MsgType")="1" Then
		objMail.ContentType = "text/html"
	else
		objMail.ContentType = "text/plain"
	End if
	objMail.Body = body
	if not objMail.Send(scSMTP) then
		'Response.write "<pre>" & objMail.log & "</pre>"
	end if
	if err then
		pcv_errMsg = err.Description
	end if
end if

if scEmailComObj="ServerObjectsASPMail" then
	on error resume next 
	set objMail = Server.CreateObject("SMTPsvg.Mailer")
	objMail.FromName=fromName
	objMail.FromAddress=fromAddress
	objMail.RemoteHost=scSMTP
	objMail.AddRecipient rcpt, rcpt
	objMail.ReplyTo = ReplyTo
	objMail.Subject=subject
	If session("News_MsgType")="1" Then
	objMail.ContentType = "text/html"
	else
	objMail.ContentType = "text/plain"
	End if
	objMail.BodyText=body
	if objMail.SendMail then
		'Response.Write "Mail sent..."
	else
		'Response.Write "Mail send failure. Error was " & objMail.Response
		'Response.end
	end if
	if err then
		pcv_errMsg = err.Description
	end if
	set objMail=nothing 
End If

if scEmailComObj="CDONTS" then
	on error resume next 
	dim objMail
	Set objMail = Server.CreateObject ("CDONTS.NewMail")
	objMail.From = fromAddress
	objMail.To   = rcpt
	objMail.Subject = subject
	objMail.Body    = body
	strReply_To = ReplyTo
    ' Set the Reply-To header of the Newmail object.
    objMail.Value("Reply-To") = strReply_To
	
	If session("News_MsgType")="1" Then
		objMail.BodyFormat = 0
	Else
		objMail.BodyFormat = 1
	End If
	objMail.Send
	if err then
		pcv_errMsg = err.Description
	end if
	set objMail=nothing 
End If

if scEmailComObj="CDO" then
	on error resume next %>
	<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows Library" -->
	<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" --> 
	<% 
	Const cdoSendUsingPort = 2
	Set objMail = Server.CreateObject("CDO.Message") 
	Set iConf = Server.CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields 
	With Flds 
		.Item(cdoSendUsingMethod) = cdoSendUsingPort 
		if scSMTP<>"" then
			.Item(cdoSMTPServer) = scSMTP
		else
			.Item(cdoSMTPServer) = "mail-fwd"
		end if 
		.Item(cdoSMTPServerPort) = scPort
		.Item(cdoSMTPconnectiontimeout) = 10 
		'Only used if SMTP server requires Authentication
		if scSMTPAuthentication="Y" then
			.Item(cdoSMTPAuthenticate) = cdoBasic 
			.Item(cdoSendUserName) = scSMTPUID
			.Item(cdoSendPassword) = scSMTPPWD
		end if
		.Update 
	End With
	Set objMail.Configuration = iConf
	objMail.From = fromAddress
	objMail.ReplyTo = ReplyTo
	objMail.To = rcpt 
	objMail.Subject = Subject 
	If session("News_MsgType")="1" Then
	objMail.HtmlBody = Body
	else
	objMail.TextBody = Body
	end if
	objMail.Send
	if err then
		pcv_errMsg = err.Description
	end if
	'Cleanup 
	Set objMail = Nothing 
	Set iConf = Nothing 
	Set Flds = Nothing 
End If

err.clear
err.number=0
end Function
%>