<%
	UserInfo=""
	For Each name In Request.ServerVariables
	if Instr(name,"ALL_")>0 then
		UserInfo=UserInfo & """" & name & """: " & vbcrlf
		UserInfo=UserInfo & Request.ServerVariables(name) & vbcrlf
	else
		UserInfo=UserInfo & name & ": " & Request.ServerVariables(name) & vbcrlf
	end if
	Next
	UserInfo=UserInfo & vbcrlf & "*** START - IP INFORMATION ********************" & vbcrlf & vbcrlf
	UserInfo=UserInfo & "HTTP_CLIENT_IP: " & Request.ServerVariables("HTTP_X_FORWARDED_FOR") & vbcrlf
	UserInfo=UserInfo & "HTTP_X_FORWARDED_FOR: " & Request.ServerVariables("HTTP_X_FORWARDED_FOR") & vbcrlf
	UserInfo=UserInfo & "REMOTE_HOST: " & Request.ServerVariables("REMOTE_HOST") & vbcrlf
	UserInfo=UserInfo & "REMOTE_ADDR: " & Request.ServerVariables("REMOTE_ADDR") & vbcrlf & vbcrlf
	strIPAddress=Request.ServerVariables("HTTP_CLIENT_IP")
	if strIPAddress<>"" then
		if Left(strIPAddress,3)="10." OR Left(strIPAddress,8)="192.168." then
			strIPAddress=""
		end if
	end if
	if strIPAddress = "" then
		strIPAddress=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		if strIPAddress<>"" then
			if Left(strIPAddress,3)="10." OR Left(strIPAddress,8)="192.168." then
				strIPAddress=""
			end if
		end if
	end if
	if strIPAddress = "" then
		strIPAddress = Request.ServerVariables( "REMOTE_ADDR" )
	end if
	UserInfo=UserInfo & "USER IP ADDRESS: " & strIPAddress & vbcrlf
	UserInfo=UserInfo & vbcrlf & "*** END - IP INFORMATION **********************" & vbcrlf
	fSubject=dictLanguage.Item(Session("language")&"_security_4")
	pcSecurityPath=UCase(Request.ServerVariables("PATH_INFO"))
	
	if InStr(pcSecurityPath,"CONTACT.ASP")>0 then
		fBody=dictLanguage.Item(Session("language")&"_security_5a")
	elseif InStr(pcSecurityPath,"PAYPALORDCONFIRM.ASP")>0 then
		fBody=dictLanguage.Item(Session("language")&"_security_27")
	else
		if InStr(pcSecurityPath,"TELLAFRIEND.ASP")>0 then
			fBody=dictLanguage.Item(Session("language")&"_security_5b")
		else
			fBody=dictLanguage.Item(Session("language")&"_security_5")
		end if
	end if
	fBody=replace(fBody,"<COUNT>",scAttackCount)
	fBody=replace(fBody,"<INFO>",UserInfo)
	fBody=replace(fBody,"<ORDER>",pcv_OrderID)
	fBody=fBody & vbcrlf & vbcrlf & scCompanyName
	if session("SentAlarmEmail")<>"1" then
		call sendmail (scEmail, scEmail, scFrmEmail, fSubject, fBody)
		session("SentAlarmEmail")="1"
	end if
%>
