<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/ppdstatus.inc" -->
<!--#include file="../includes/productcartFolder.asp" -->
<!--#include file="../includes/pcSurlLvs.asp" -->
<!--#include file="../includes/contactEmail.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="pcCPLog.asp" -->
<% 'on error resume next
Dim SPath
SPath=Request.ServerVariables("PATH_INFO")
SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
	strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
Else
	strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
End if

IF scSecurity=1 THEN
	if scAdminLogin=1 then
		pcv_Test=0
		if pcv_Test=0 then
			if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL))<>1 then
				Session("cp_Adminlogin")=""
				pcv_Test=1
			end if
		end if
		
		if pcv_Test=1 then
			if session("AttackCount")="" then
				session("AttackCount")=0
			end if
			session("AttackCount")=session("AttackCount")+1
			if session("AttackCount")>=scAttackCount then
				session("AttackCount")=0
				If scAlarmMsg=1 then%>
					<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if
				response.write dictLanguage.Item(Session("language")&"_security_2")
				response.end()	
			End if
		end if
	end if
END IF

Session("cp_Adminlogin")=""
Session("cp_postnum")=""
Session("cp_num")=""

dim mySQL, conntemp, rstemp, pemail, ppassword, pAdminPassword

' form parameters
pIdAdmin=replace(request.querystring("IdAdmin"),"'","''")
pIdAdmin=replace(pIdAdmin,"--","")
If NOT isNumeric(pIdAdmin) then
	response.redirect "msg.asp?message=1"
	response.end()
end if
pAdminPassword=Decrypt(request.querystring("password"), 9286803311968)
pAdminPassword=enDeCrypt(pAdminPassword, scCrypPass)
Session("pcAuditAdmin") = pIdAdmin

call openDb()
err.clear

Dim strAltURL
strAltURL = "menu.asp"
%>
<!--#include file="AdminLoginInclude.asp"-->
<!--#include file="AdminFooter.asp"-->