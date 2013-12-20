<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/validation.asp" --> 
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<%
on error resume next

Dim SPath
SPath=Request.ServerVariables("PATH_INFO")
SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
	strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
Else
	strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
End if

IF scSecurity=1 THEN
	if scAffLogin=1 then
		pcv_Test=0
		if Session("store_afflogin")<>"1" then
			Session("store_afflogin")=""
			pcv_Test=1
		end if
		if pcv_Test=0 then
			if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "AffiliateLogin.asp"))<>1 then
				Session("store_afflogin")=""
				pcv_Test=1
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
			response.write dictLanguage.Item(Session("language")&"_security_2")
			response.end
		end if
	end if
END IF

dim mySQL, conntemp, rstemp, pEmail, pPassword

pEmail=session("email")
pPassword=session("erypassword")
pPassword=Decrypt(pPassword, 9286803311968)
pPassword=enDeCrypt(pPassword, scCrypPass)
pRedirectUrl=session("redirectUrlLI")


'open database
call openDB()

' verify password for that email
mySQL="SELECT idAffiliate, pcAff_Active, affiliateName FROM affiliates WHERE affiliateEmail='" &pEmail& "' AND [pcAff_password]='" &pPassword& "'"
set rstemp=conntemp.execute(mySQL)

if err.number <> 0 then
	call closeDb()
  response.redirect "techErr.asp?error="&Server.Urlencode(err.description)
end If

if rstemp.eof then
	call closeDb()
	If (scSecurity=1) and (scUserLogin=1) and (scAlarmMsg=1) then
		if session("AttackCount")="" then
			session("AttackCount")=0
		end if
		session("AttackCount")=session("AttackCount")+1
		if session("AttackCount")>=scAttackCount then
		session("AttackCount")=0%>
		<!--#include file="../includes/sendAlarmEmail.asp" -->
		<%end if	
	End if
 	response.redirect "msg.asp?message=91"        
end if

pcv_idAffiliate=rstemp("idAffiliate")
pcv_pcAff_Active=rstemp("pcAff_Active")
pcv_affiliateName=rstemp("affiliateName")

if (pcv_pcAff_Active="0") or (pcv_pcAff_Active="") then
	call closeDb()
 	response.redirect "msg.asp?message=92"           
end if

Session("store_afflogin")=""
Session("store_affpostnum")=""
Session("store_affnum")=""
session("redirectUrlLI")=""
session("pc_idAffiliate")=pcv_idAffiliate
session("pc_AffiliateName")=pcv_affiliateName

call closeDB()
set rstemp=nothing

if pRedirectUrl="" then
  response.redirect "AffiliateMain.asp"
else
  response.redirect pRedirectUrl
end if
%>