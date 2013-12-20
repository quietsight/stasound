<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"--> 
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<%
pcStrPageName = "tellafriend.asp"
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<%

dim query, conntemp, pc_fromname, pc_fromemail, pc_toname, pc_toemail, pc_pname, pc_subject, pc_message, pc_pid

Dim TurnOnSecurity

'1 - Turn On (Default)
'0 - Turn Off

TurnOnSecurity=1

Dim pcSecurityPath, strSiteSecurityURL

IF TurnOnSecurity=1 THEN
	pcSecurityPath=Request.ServerVariables("PATH_INFO")
	pcSecurityPath=mid(pcSecurityPath,1,InStrRev(pcSecurityPath,"/")-1)
	If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
		strSiteSecurityURL="http://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
	Else
		strSiteSecurityURL="https://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
	End if
END IF

pcv_isNameRequired=True
pcv_isEmailRequired=True
pcv_isFriendRequired=True
pcv_isFEmailRequired=True
pcv_isMsgRequired=False

' Send the email
IF request.Form("sendmessage")="yes" THEN

	'//set error to zero
	pcv_intErr=0
	
	pcs_ValidateTextField	"yourname", pcv_isNameRequired, 0
	pcs_ValidateEmailField	"youremail", pcv_isEmailRequired, 0
	pcs_ValidateTextField	"friendsname", pcv_isFriendRequired, 0
	pcs_ValidateEmailField	"friendsemail", pcv_isFEmailRequired, 0
	pcs_ValidateTextField	"message", pcv_isMsgRequired, 0
	
	pc_pid=getUserInput(request.form("idproduct"),10)
	if not validNum(pc_pid) then
		response.redirect "msg.asp?message=207"
	end if

	IF TurnOnSecurity=1 THEN
		pcv_Test=0
		if InStr(ucase(Request.ServerVariables("HTTP_REFERER")),ucase(strSiteSecurityURL & "tellafriend.asp"))<>1 then
			pcv_test=1
		end if
		
		if pcv_Test=1 then
			if session("AttackCount")="" then
				session("AttackCount")=0
			end if
			session("AttackCount")=session("AttackCount")+1
			if session("AttackCount")>=scAttackCount then
					session("AttackCount")=0%>
					<!--#include file="../includes/sendAlarmEmail.asp" -->
			<%end if	
			response.redirect pcStrPageName & "?emailSent=no&idproduct="&pc_pid & "&msg=1"
			response.end
		end if
		
		if pcv_Test=0 then %>
            <!-- Include file for CAPTCHA configuration -->
            <!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
             
            <!-- Include file for CAPTCHA form processing -->
            <!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
			<%	
            If not blnCAPTCHAcodeCorrect then
				pcv_test=1
			end if
		end if

		if pcv_Test=1 then
			if session("AttackCount")="" then
				session("AttackCount")=0
			end if
			session("AttackCount")=session("AttackCount")+1
			if session("AttackCount")>=scAttackCount then
					session("AttackCount")=0%>
					<!--#include file="../includes/sendAlarmEmail.asp" -->
			<%end if	
			response.redirect pcStrPageName & "?emailSent=no&idproduct="&pc_pid & "&msg=2"
			response.end
		end if
	END IF
	
	'//Email error for page
	If Session("Erryouremail")="" OR isNULL(Session("Erryouremail")) Then Session("Erryouremail")=0
	if Session("Erryouremail")=1 then
		pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_sendpassword_1"))
	else	
		'//generic error for page
		pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	end if
	
	If Session("Errfriendsemail")="" OR isNULL(Session("Errfriendsemail")) Then Session("Errfriendsemail")=0
	if Session("Errfriendsemail")=1 then
		pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_sendpassword_1"))
	else	
		'//generic error for page
		pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	end if
	
	IF pcv_intErr>0 THEN
		response.redirect pcStrPageName & "?msgerr="&pcv_strGenericPageError&"&emailSent=no&idproduct="&pc_pid
	ELSE
	pc_fromname=request.form("yourname")
	pc_fromemail=request.form("youremail")
	pc_toname=request.form("friendsname")
	pc_toemail=request.form("friendsemail")
	pc_pname=request.form("pname")
	pc_pname=ClearHTMLTags2(pc_pname,0)
	pc_subject = pc_fromname & dictLanguage.Item(Session("language")&"_tellafriend_20") & " - " & scCompanyName
	pc_subject = replace(pc_subject,"&quot;","'")
	pc_message=request.form("message")
		
	dim tempURL

	pcStrPrdLinkCan=pc_pname & "-p" & pc_pid & ".htm"
    pcStrPrdLinkCan=removeChars(pcStrPrdLinkCan)
    if scSeoURLs<>1 then
        pcStrPrdLinkCan="viewPrd.asp?idproduct="&pc_pid
    end if
	
	
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
	
	
	emailbody=dictLanguage.Item(Session("language")&"_referral_1")&pc_toname&","&vbcrlf&vbcrlf
	emailbody=emailbody&dictLanguage.Item(Session("language")&"_referral_2")&scCompanyName&vbcrlf&vbcrlf
	emailbody=emailbody&pc_message&vbcrlf&vbcrlf
	
	cid=session("idCustomer")
	
	if RewardsActive <> 0 AND (cid<>"" AND cid<>"0") then
		emailbody=emailbody&tempURL&"viewPrd.asp?idproduct="&pc_pid&"&refby="&cid&vbcrlf&vbcrlf
	else
		emailbody=emailbody&tempURL&pcStrPrdLinkCan&vbcrlf&vbcrlf
	end if
	
	emailbody=emailbody&pc_fromname
	emailbody=replace(emailbody,"'","")
	emailbody=replace(emailbody,"&quot;","'")
	
	call sendmail (pc_fromname, pc_fromemail, pc_toemail, pc_subject, emailbody)

	if err then
		response.Write err.Description 
	end if
	
	'// Send Thank You to the customer
	pc_subject = dictLanguage.Item(Session("language")&"_tellafriend_24") & " - " & scCompanyName
	pc_subject = replace(pc_subject,"&quot;","'")
	MsgBody=dictLanguage.Item(Session("language")&"_tellafriend_21") & vbcrlf & vbcrlf
	MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_tellafriend_22") & pc_pname & vbcrlf
	MsgBody=MsgBody & tempURL&pcStrPrdLinkCan & vbcrlf & vbcrlf
	MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_tellafriend_23") & pc_toname & " - " & pc_toemail  & vbcrlf & vbcrlf
	MsgBody=MsgBody & scCompanyName & vbcrlf & scStoreURL
	MsgBody=replace(MsgBody,"&quot;","'")
	call sendmail (scCompanyName, scEmail, pc_Fromemail, pc_subject, MsgBody)

	if err then
		response.Write err.Description 
	end if

	'// Send notification to the store administrator
	pc_subject = dictLanguage.Item(Session("language")&"_tellafriend_25")
	MsgBody=dictLanguage.Item(Session("language")&"_tellafriend_26") & vbcrlf & vbcrlf
	MsgBody=MsgBody & pc_pname & vbcrlf
	MsgBody=MsgBody & tempURL&pcStrPrdLinkCan & vbcrlf & vbcrlf
	MsgBody=replace(MsgBody,"&quot;","'")
	call sendmail (scCompanyName, scEmail, scFrmEmail, pc_subject, MsgBody)
	
	if err then
		response.Write err.Description 
	else
		response.redirect pcStrPageName & "?emailSent=yes&idproduct="&pc_pid
	end if
	END IF

ELSE
	
	pc_pid=getUserInput(request("idproduct"),10)
	pIdCustomer=session("idCustomer")

	if not validNum(pc_pid) then
		response.redirect "msg.asp?message=207"
	end If
	
	if not validNum(pIdCustomer) then
		pIdCustomer=0
	end If
	
	'open database
	call openDB()
	
		' get product values
		query="SELECT description, active FROM products WHERE idproduct="&pc_pid
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)	
		if err.number <> 0 then
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="&Server.Urlencode("Error: "&err.description)
		end If	
		If NOT rs.EOF Then
		
			productname=rs("description")
			pcIntProductStatus = rs("active")
			
			pcStrPrdLinkCan=productname & "-p" & pc_pid & ".htm"
			pcStrPrdLinkCan=removeChars(pcStrPrdLinkCan)
			if scSeoURLs<>1 then
				pcStrPrdLinkCan="viewPrd.asp?idproduct="&pc_pid
			end if
			
			if pcIntProductStatus=0 or isNull(pcIntProductStatus) or pcIntProductStatus="" then
				set rs = nothing
				call closeDb()
				response.redirect "msg.asp?message=95"
			end if
	
			if pIdCustomer>0 then
				
				query="SELECT name, lastName, email FROM customers WHERE idCustomer=" &pIdCustomer
				set rs2=Server.CreateObject("ADODB.Recordset")
				set rs2=conntemp.execute(query)				
				if err.number <> 0 then
					call closeDb()
					set rs2 = nothing
					response.redirect "techErr.asp?error="&Server.Urlencode("Error: "&err.description)
				end If				
				If NOT rs2.eof Then
					CustName=rs2("name") & " " & rs2("lastname")
					Session("pcSFyourname")=CustName
					CustEmail=rs2("email")
					Session("pcSFyouremail")=CustEmail
				End If
				set rs2 = nothing 
				
			end if
		
		Else
				set rs = nothing
				call closeDb()
				response.redirect "msg.asp?message=95"
		End If
		set rs = nothing
		call closeDb()
	%>
  
	<script language="JavaScript">
  <!--
    
  function Form1_Validator(theForm)
  {
    if (theForm.yourname.value == "")
      {
        alert("<%response.write dictLanguage.Item(Session("language")&"_security_20")%>");
          theForm.yourname.focus();
          return (false);
    }
  
    if (theForm.youremail.value == "")
      {
        alert("<%response.write dictLanguage.Item(Session("language")&"_security_21")%>");
          theForm.youremail.focus();
          return (false);
    }
    
    if (theForm.friendsname.value == "")
      {
        alert("<%response.write dictLanguage.Item(Session("language")&"_security_24")%>");
          theForm.friendsname.focus();
          return (false);
    }
  
    if (theForm.friendsemail.value == "")
      {
        alert("<%response.write dictLanguage.Item(Session("language")&"_security_25")%>");
          theForm.friendsemail.focus();
          return (false);
    }
    
    if (theForm.message.value == "")
      {
        alert("<%response.write dictLanguage.Item(Session("language")&"_security_23")%>");
          theForm.message.focus();
          return (false);
    }
		
    if (theForm.postnum.value == "")
      {
        alert("<%response.write dictLanguage.Item(Session("language")&"_security_6")%>");
          theForm.postnum.focus();
          return (false);
    }

	
  return (true);
  }
  //-->
  </script>

	<div id="pcMain">
		<table class="pcMainTable">
			<tr>
				<td>
					<h1><%response.write dictLanguage.Item(Session("language")&"_tellafriend_1")%></h1>
				</td>
			</tr>
		
		<% if request.QueryString("emailSent")="yes" then %>
				<tr>
					<td>
					<div class="pcSuccessMessage"><%response.write dictLanguage.Item(Session("language")&"_tellafriendthanks_2")%></div>
					<p>&nbsp;</p>
					<p><a href="viewprd.asp?idproduct=<%=pc_pid%>"><img src="<%=rslayout("back")%>" alt="Back to Product Page"></a></p>
					<p>&nbsp;</p>
				</td>
			</tr>
			<% 
			' End success message
			' Show tell-a-friend form
			else%>
			<%if request("msg")<>"" then%>
				<tr>
					<td>
						<div class="pcErrorMessage"><%if request("msg")="1" then%><%response.write dictLanguage.Item(Session("language")&"_security_2")%><%else%><%response.write dictLanguage.Item(Session("language")&"_security_6")%><%end if%></div>
					</td>
				</tr>
			<%end if%>
			<%if request("msgerr")<>"" then%>
				<tr>
					<td>
						<div class="pcErrorMessage"><%=getUserInput(request("msgerr"),0)%></div>
					</td>
				</tr>
			<%end if%>
			<tr>
				<td> 
				<div class="pcPageDesc"><%response.write dictLanguage.Item(Session("language")&"_tellafriend_2")%></div>
				<h3><%response.write dictLanguage.Item(Session("language")&"_tellafriend_9")%><%=productname%></h3>
				<form name="request" action="tellafriend.asp" method="POST" class="pcForms" onSubmit="return Form1_Validator(this)">
				<input type="hidden" name="idproduct" value="<%=pc_pid%>">
				<input type="hidden" name="pname" value="<%=productname%>">
				<input type="hidden" name="sendmessage" value="yes">
					<table class="pcShowContent">
						<tr>
							<td width="20%">
								<p><%response.write dictLanguage.Item(Session("language")&"_tellafriend_3")%></p>
							</td>
							<td width="80%">
								<input type="text" size="18" name="yourname" value="<%=pcf_FillFormField ("yourname", pcv_isNameRequired) %>">
								<% pcs_RequiredImageTag "yourname", pcv_isNameRequired %>
							</td>
						</tr>
						<tr>
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_tellafriend_4")%></p>
							</td>
							<td>
								<input type="text" size="18" name="youremail" value="<%=pcf_FillFormField ("youremail", pcv_isEmailRequired) %>">
								<% pcs_RequiredImageTag "FromEmail", pcv_isEmailRequired %>
							</td>
						</tr>
						<tr> 
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_tellafriend_5")%></p>
							</td>
							<td>
								<input type="text" size="18" name="friendsname" value="<%=pcf_FillFormField ("friendsname", pcv_isFriendRequired) %>">
								<% pcs_RequiredImageTag "friendsname", pcv_isFriendRequired %>
							</td>
						</tr>
						<tr> 
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_tellafriend_6")%></p>
							</td>
								<td>
									<input type="text" size="18" name="friendsemail" value="<%=pcf_FillFormField ("friendsemail", pcv_isFEmailRequired) %>">
									<% pcs_RequiredImageTag "friendsemail", pcv_isFEmailRequired %>
								</td>
							</tr>
							<tr> 
								<td>
									<p><%response.write dictLanguage.Item(Session("language")&"_tellafriend_7")%></p>
								</td>
								<td>
									<textarea rows="5" cols="30" name="message"><%=pcf_FillFormField ("message", pcv_isMsgRequired) %></textarea>
									<% pcs_RequiredImageTag "message", pcv_isMsgRequired %>
								</td>
							</tr>
							<%IF TurnOnSecurity=1 THEN
								Session("store_postnum")=""
								session("store_num")="      " 
							%>
							<tr>
								<td colspan="2">
									<table width="100%" id="show_security" border="0" cellpadding="0" cellspacing="0" style="display:<%=pcShowLoginStyle%>">
									<tr>
                  						<td width="20%"></td>
										<td width="80%"><%response.write dictLanguage.Item(Session("language")&"_security_26")%></td>
									</tr>
									<tr>
                  						<td></td>
										<td style="padding: 6px 0 6px 3px;"><!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" --></td>
									</tr>
									</table>
								</td>
							</tr>
							<%END IF%>
              <tr>
              	<td colspan="2">&nbsp;</td>
              </tr>
							<tr> 
              	<td></td>
								<td>
                	<input type="image" src="<%=rslayout("submit")%>" border="0" name="submit" value="<%response.write dictLanguage.Item(Session("language")&"_tellafriend_8")%>" id="submit">
                  &nbsp;
                  <a href="<%=pcStrPrdLinkCan%>"><img src="<%=rslayout("back")%>" border="0" alt="<%response.write dictLanguage.Item(Session("language")&"_altTag_15")%>"></a>&nbsp;
								</td>
							</tr>
						</table>
					</form>
				</td>
			</tr>
			<%
			end if
			' End show tell-a-friend form
			%>
		</table>
	</div>
<% END IF %><!--#include file="footer.asp"-->