<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<% On Error Resume Next
dim query, conntemp, rs

call openDb()

Dim pcEmail, pcPassword, securityCode, CAPTCHA_Postback

pcEmail=getUserInput(request("email"),0)
pcPassword=getUserInput(request("password"),0)
securityCode=getUserInput(request("securityCode"),0)
CAPTCHA_Postback=getUserInput(request("CAPTCHA_Postback"),0)

pcErrMsg=""

if scSecurity=1 AND (scUserLogin=1 OR scUserReg=1) then
	pcv_Test=0
	'// Remote access attempt
	if (session("store_userlogin")<>"1") AND (session("store_adminre")<>"1") then
		session("store_userlogin")=""
		session("store_adminre")=""
		pcv_Test=1
	end if
	if pcv_Test=0 AND scUseImgs=1 then %>
		<!-- Include file for CAPTCHA configuration -->
		<!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
		 
		<!-- Include file for CAPTCHA form processing -->
		<!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
	<%	
		If not blnCAPTCHAcodeCorrect then
			pcv_Test=2
			pcErrMsg=dictLanguage.Item(Session("language")&"_security_3")
		end if
	end if

	if pcv_Test=1 then
		If scAlarmMsg=1 then
			if session("AttackCount")="" then
				session("AttackCount")=0
			end if
			session("AttackCount")=session("AttackCount")+1
			if session("AttackCount")>=scAttackCount then
				session("AttackCount")=0%>
				<!--#include file="../includes/sendAlarmEmail.asp" -->
			<%end if	
		End if
		pcErrMsg=dictLanguage.Item(Session("language")&"_security_2")
	end if					
end if

if pcEmail="" OR pcPassword="" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checklogin_1")
else
	pcStrLoginPassword=enDeCrypt(pcPassword, scCrypPass)
	query="SELECT idcustomer,suspend,pcCust_Locked,pcCust_Guest FROM customers WHERE email like '" & pcEmail & "' AND password='" & pcStrLoginPassword & "' AND pcCust_Guest<>1;"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checklogin_2")
	else
		if rs("suspend")="1" then
			pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_3")
		end if
		if rs("pcCust_Locked")="1" then
			pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_4")
		end if
	end if
end if
set rs=nothing
if pcErrMsg="" then
	session("pcSFLoginEmail")=pcEmail
	session("pcSFLoginPassword")=pcPassword
	erypassword=encrypt(pcPassword, 9286803311968)
	session("pcSFPassWordExists")="YES"
	session("pcSFEryPassword")=erypassword
	response.redirect "login.asp?lmode=0&opc=1"
end if
%>
<%response.write pcErrMsg%>
<%
call closeDb()
%>

