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
<!--#include file="../includes/ErrorHandler.asp"-->
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
	pcv_Test=0
	if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "sds_Login.asp"))<>1 then
		pcv_Test=1
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
END IF

dim query, conntemp, rstemp

sds_username=session("sds_username")
sds_Password=session("sds_erypassword")
pPassword=Decrypt(sds_Password, 9286803311968)
pPassword=enDeCrypt(pPassword, scCrypPass)
pRedirectUrl=session("redirectUrlLI")


'open database
call openDB()

' verify password for that username
query="SELECT pcDropShipper_ID As idsds, pcDropShipper_FirstName As FirstName, pcDropShipper_LastName As LastName, pcDropShipper_Company As Company,0 As IsDropShipper FROM pcDropShippers WHERE pcDropShipper_Username='" & sds_username & "' AND pcDropShipper_Password='" &pPassword& "' UNION SELECT pcSupplier_ID,pcSupplier_FirstName,pcSupplier_LastName, pcSupplier_Company,pcSupplier_IsDropShipper FROM pcSuppliers WHERE pcSupplier_Username='" & sds_username & "' AND pcSupplier_Password='" &pPassword& "' AND pcSupplier_IsDropShipper=1"
set rstemp=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	call closeDb()
	If (scSecurity=1) and (scAlarmMsg=1) then
		if session("AttackCount")="" then
			session("AttackCount")=0
		end if
		session("AttackCount")=session("AttackCount")+1
		if session("AttackCount")>=scAttackCount then
		session("AttackCount")=0%>
		<!--#include file="../includes/sendAlarmEmail.asp" -->
		<%end if	
	End if
 	response.redirect "msg.asp?message=202"        
end if

session("redirectUrlLI")=""
session("pc_idsds")=rstemp("idsds")
session("pc_sdsName")=rstemp("FirstName") & " " & rstemp("LastName")
session("pc_sdsCompany")=rstemp("Company")
session("pc_sdsIsDropShipper")=rstemp("IsDropShipper")

call closeDB()
set rstemp=nothing

if pRedirectUrl="" then
  response.redirect "sds_MainMenu.asp"
else
  response.redirect pRedirectUrl
end if
%>