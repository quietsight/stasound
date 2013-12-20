<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/ppdstatus.inc" -->
<!--#include file="../includes/productcartFolder.asp" -->
<!--#include file="pcCPLog.asp" -->
<% 
Function makePassword(byVal maxLen)
			
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
			
	For intCounter = 1 To maxLen
		whatsNext = Int((1 - 0 + 1) * Rnd + 0)
		If whatsNext = 0 Then
			'character
			upper = 90
			lower = 65
		Else
			upper = 57
			lower = 48
		End If
		strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	makePassword = strNewPass
	
End function
	
Dim SPath
SPath=Request.ServerVariables("PATH_INFO")
SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
	strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
Else
	strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
End if
           
' attack was not submitted from the forgot_password page   close them out  
if Session("cp_Forgotpassword")<>"1" then		  
	Session("cp_Forgotpassword")=""
	if session("ForgotAttackCount")="" then
	session("ForgotAttackCount")=0
	end if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1	
	
	response.redirect "forgot_password.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end
end if
		
' attack was not submitted from this site  close them out 
if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "forgot_password.asp")) <>1 then
	Session("cp_Forgotpassword")=""
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end  if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1		
	response.redirect "forgot_password.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end			
end if
	
IF session("ForgotAttackCount") => 5 THEN 
	response.redirect "forgot_password.asp"
	response.end()
END IF    

dim pemail, ppassword, pAdminPassword

pAdminUser=getUserInput(request.querystring("user"),150)

dim query, conntemp, rs 
call openDb()
err.clear
' authenticated and charge session
query="SELECT IDAdmin FROM admins WHERE IDAdmin=" & pAdminUser  &" and AdminLevel='19';" 
set rs=server.CreateObject("ADODB.RecordSet")		
set rs=conntemp.execute(query)
		
if err.number>0 then
	set rs=nothing
	call closeDb()			
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1								
	response.redirect "forgot_password.asp?msg=" & dictLanguage.Item(Session("language")&"_security_2") 
	response.end()
end if

if rs.eof then
	set rs=nothing
	call closeDb()			
	if session("ForgotAttackCount")="" then
		session("ForgotAttackCount")=0
	end if
	session("ForgotAttackCount")=session("ForgotAttackCount")+1		
	response.redirect "forgot_password.asp?msg=" & dictLanguageCP.Item(Session("language")&"_forgotpasswordadminerror") 
	response.end()
else		
	Dim IDAdmin,fromName,from,rcpt,subject,body
	session("ForgotAttackCount")=0
	pAdminPassword  = makePassword(8)
	pAdminPassForDB= enDeCrypt(pAdminPassword, scCrypPass)				
	err.clear
	' authenticated and charge session
	query="update admins SET adminpassword ='"& pAdminPassForDB  &"' WHERE IDAdmin=" &pAdminUser &" And AdminLevel='19'"                
	set rs=server.CreateObject("ADODB.RecordSet")		
	set rs=conntemp.execute(query)
				
	if err.number>0 then
		set rs=nothing
		call closeDb()			
		if session("ForgotAttackCount")="" then
			session("ForgotAttackCount")=0
		end if
		session("ForgotAttackCount")=session("ForgotAttackCount")+1
							
		response.redirect "forgot_password.asp?msg=" & dictLanguageCP.Item(Session("language")&"_forgotpasswordadminDBerror")& err.number
		response.end()
	end if
	call closedb()
	
	fromName = dictLanguageCP.Item(Session("language")&"_forgotpasswordadminmailfrom")  		
	from = scFrmEmail
	rcpt = scEmail
	subject = dictLanguageCP.Item(Session("language")&"_forgotpasswordadminmailsubject") 
	body = Replace(dictLanguageCP.Item(Session("language")&"_forgotpasswordadminmailbody1"),"#password",pAdminPassword ) 	
	
	call sendMail (fromName, from, rcpt, subject, body)
	
	' SEnd an email to the store Admin 
	if Session("RedirectURL")<>"" then
		RedirectURL=Session("RedirectURL")
		Session("RedirectURL")=""
		response.redirect RedirectURL
	else
		response.redirect "login_1.asp?s=1&msg=" & server.URLEncode(dictLanguageCP.Item(Session("language")&"_forgotpasswordadminsuccess"))
	end if
end if 
%>