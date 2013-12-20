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
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%
'// START - Check for SSL and redirect to SSL login if not already on HTTPS
	If scSSL="1" And scIntSSLPage="1" Then
		If (Request.ServerVariables("HTTPS") = "off") Then
		Dim xredir__, xqstr__
		xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
		Request.ServerVariables("SCRIPT_NAME")
		xqstr__ = Request.ServerVariables("QUERY_STRING")
		if xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
		Response.redirect xredir__
		End if
	End If
'// END - check for SSL

'Capture any redirects
dim pcRequestRedirect
pcRequestRedirect=getUserInput(request("redirectUrl"),250)

if Session("SFStrRedirectUrl")<>"" AND pcRequestRedirect="" then
else
	Session("SFStrRedirectUrl")=pcRequestRedirect
end if

session("REGidCustomer")=""

dim pcPageMode
dim query, conntemp, rs

pcPageMode=request("cmode")
if pcPageMode="" then
	pcPageMode=0
else
	if NOT validNum(pcPageMode) then
		pcPageMode=0
	end if
end if

'// Check if only PayPal Express is enabled - begin                                                         
if session("customerType")=1 then
    query="SELECT gwcode FROM paytypes WHERE active=-1;"
else
    query="SELECT gwcode FROM paytypes WHERE active=-1 and Cbtob=0;"
end if
call opendb()

set rsPPObj=server.CreateObject("ADODB.RecordSet")
set rsPPObj=conntemp.execute(query)
dim intPPECheck, intPPECnt, intPPEOnly
intPPECheck=0
intPPECnt=0
intPPEOnly=0
do until rsPPObj.eof
    PPE_GwCode=rsPPObj("gwcode")
    if PPE_GwCode="999999" then
        intPPECheck=1
    end if
    intPPECnt=intPPECnt+1
    rsPPObj.movenext
loop
set rsPPObj=nothing
if intPPECnt=1 AND intPPECheck=1 then
    intPPEOnly=1
end if
call closedb()

IF pcPageMode=0 AND request("EmailNotFound")="" THEN
    if intPPEOnly = 1 then
        response.redirect "viewcart.asp"
    end if
	response.redirect "onepagecheckout.asp"
END IF
'// Check if only PayPal Express is enabled - end

'pcPageMode
'0=checkout
'1=login
'2=retreive password
'3=autologin
'4=retreive order code(s)

	call opendb()
	Dim strCCSLCheck
	strCCSLcheck = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)
	If Len(Trim(strCCSLCheck))>0 Then
		response.redirect "viewcart.asp"
	End If
	call closedb()

if pcPageMode=2 then
	pcFromPageMode=getUserInput(request("fmode"),1)
	if pcFromPageMode="" then
		pcFromPageMode=2
	else
		if not validNum(pcFromPageMode) then
			pcFromPageMode=0
		end if
	end if
end if

if pcPageMode=4 then
	pcFromPageMode=getUserInput(request("fmode"),1)
	if pcFromPageMode="" then
		pcFromPageMode=4
	else
		if not validNum(pcFromPageMode) then
			pcFromPageMode=0
		end if
	end if
end if

If (Session("SFStrRedirectUrl")<>"" AND pcPageMode<>0) AND (session("idCustomer")<>0 and session("idCustomer")<>"") then
	response.redirect "Login.asp?lmode=2"
end if

session("pcSFCMode")=pcPageMode

'Get path for Advanced Security
if scSecurity=1 then
	Dim pcSecurityPath, strSiteSecurityURL

	pcSecurityPath=Request.ServerVariables("PATH_INFO")
	pcSecurityPath=mid(pcSecurityPath,1,InStrRev(pcSecurityPath,"/")-1)
	If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
		strSiteSecurityURL="http://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
	Else
		strSiteSecurityURL="https://" & Request.ServerVariables("HTTP_HOST") & pcSecurityPath & "/"
	End if
end if

'check if email is passed to retrieve password
if pcPageMode=2 AND request("SubmitPM.y")<>"" then
	pcv_intErr=0 'set to zero
	pcs_ValidateEmailField	"LoginEmail", true, 250
	pcStrEmail = Session("pcSFLoginEmail")

	query="SELECT idcustomer, name, lastname, email, [password] from customers WHERE email='" &pcStrEmail& "' AND (pcCust_Guest=0 OR pcCust_Guest=2)"
	
	call opendb()
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)	
	if not rs.eof then
		pcIntCustomerID=rs("idcustomer")
		pcStrName=rs("name")
		pcStrLastName=rs("lastname")
		pcStrEmail=rs("email")
		pcStrPassword=enDeCrypt(rs("Password"),scCrypPass)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// START No password, add now
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		if trim(pcStrPassword)="" or IsNull(pcStrPassword) then
			' Generate random passwords:
			function randomNumber(limit)
				randomize
				randomNumber=int(rnd*limit)+2
			end function
			pcStrCustomerPassword=randomNumber(99999999)
			pcStrCustomerPassword=enDeCrypt(pcStrCustomerPassword, scCrypPass)
			query="UPDATE customers SET [password]='"&pcStrCustomerPassword&"' WHERE idCustomer="& pcIntCustomerID
			set rsTemp=server.CreateObject("ADODB.RecordSet")
			set rsTemp=conntemp.execute(query)
			set rsTemp=nothing	
			pcStrPassword=enDeCrypt(pcStrCustomerPassword,scCrypPass)					
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// START No password, add now
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
		pcStrSubject=dictLanguage.Item(Session("language")&"_forgotpasswordmailsubject")
		pcStrBody=dictLanguage.Item(Session("language")&"_forgotpasswordmailbody1")
		pcStrBody=replace(pcStrBody,"#password",pcStrPassword)	
		pcStrBody=replace(pcStrBody,"#firstname",pcStrName)      
		pcStrBody=replace(pcStrBody,"#lastname",pcStrLastName)
		call sendmail (scEmail, scEmail, pcStrEmail, pcStrSubject, pcStrBody) 
		set rs=nothing
		call closedb()
		if pcFromPageMode=2 then
			response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=&EmailNotFound=0"
		else
			response.redirect "checkout.asp?cmode="&pcFromPageMode&"&EmailNotFound=0"
		end if
			
	else
		'password not found..
		set rs=nothing
		query="SELECT idcustomer from customers WHERE email='" &pcStrEmail& "' AND pcCust_Guest=1;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closedb()
			if pcFromPageMode=2 then
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=&msgmode=7"
			else
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&msgmode=7"
			end if
		else
			set rs=nothing
			call closedb()
			if pcFromPageMode=2 then
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=&EmailNotFound=1"
			else
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&EmailNotFound=1"
			end if
		end if
	end if
end if

if pcPageMode=4 AND request("SubmitPM.y")<>"" then
	pcv_intErr=0 'set to zero
	pcs_ValidateEmailField	"LoginEmail", true, 250
	pcStrEmail = Session("pcSFLoginEmail")

	query="SELECT idcustomer, name, lastname, email, [password], pcCust_Guest from customers WHERE email='" &pcStrEmail& "';"
	
	call opendb()
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)	
	if not rs.eof then
		pcIntCustomerID=rs("idcustomer")
		pcStrName=rs("name")
		pcStrLastName=rs("lastname")
		pcStrEmail=rs("email")
		pcStrPassword=enDeCrypt(rs("Password"),scCrypPass)
		pcv_Guest=rs("pcCust_Guest")
		
		if pcv_Guest<>"1" then
			set rs=nothing
			call closedb()
			if pcFromPageMode=4 then
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=4&ENotFound=3"
			else
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&ENotFound=3"
			end if
		end if

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// START No password, add now
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		if trim(pcStrPassword)="" or IsNull(pcStrPassword) then
			' Generate random passwords:
			function randomNumber(limit)
				randomize
				randomNumber=int(rnd*limit)+2
			end function
			pcStrCustomerPassword=randomNumber(99999999)
			pcStrCustomerPassword=enDeCrypt(pcStrCustomerPassword, scCrypPass)
			query="UPDATE customers SET [password]='"&pcStrCustomerPassword&"' WHERE idCustomer="& pcIntCustomerID
			set rsTemp=server.CreateObject("ADODB.RecordSet")
			set rsTemp=conntemp.execute(query)
			set rsTemp=nothing	
			pcStrPassword=enDeCrypt(pcStrCustomerPassword,scCrypPass)					
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// START No password, add now
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		query="SELECT pcOrd_OrderKey FROM Orders WHERE idCustomer=" & pcIntCustomerID & " AND OrderStatus>1;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pcArr=rs.getRows()
			intCount=ubound(pcArr,2)
			tmpOrderCodes=""
			For i=0 to intCount
				tmpOrderCodes=tmpOrderCodes & pcArr(0,i) & vbcrlf
			Next
			pcStrSubject=dictLanguage.Item(Session("language")&"_forgotordercodesmailsubject")
			pcStrBody=dictLanguage.Item(Session("language")&"_forgotordercodesmailbody")
			pcStrBody=replace(pcStrBody,"#ordercodes",tmpOrderCodes)	
			pcStrBody=replace(pcStrBody,"#firstname",pcStrName)      
			pcStrBody=replace(pcStrBody,"#lastname",pcStrLastName)
			call sendmail (scEmail, scEmail, pcStrEmail, pcStrSubject, pcStrBody) 
			set rs=nothing
			call closedb()
			if pcFromPageMode=4 then
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=4&ENotFound=0"
			else
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&ENotFound=0"
			end if
		else
			set rs=nothing
			call closedb()
			if pcFromPageMode=4 then
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=4&ENotFound=2"
			else
				response.redirect "checkout.asp?cmode="&pcFromPageMode&"&ENotFound=2"
			end if
		end if
			
	else
		'customer not found..
		set rs=nothing
		call closedb()
		if pcFromPageMode=4 then
			response.redirect "checkout.asp?cmode="&pcFromPageMode&"&fmode=4&ENotFound=1"
		else
			response.redirect "checkout.asp?cmode="&pcFromPageMode&"&ENotFound=1"
		end if
	end if
end if

session("availableShipStr")=""
session("provider")=""
pcAutoLoginAllowed=0

if (request.form("SubmitCO.y")<>"") or (pcPageMode=3) then
	pcv_intErr=0 'set to zero
	
	'Autologin
	if pcPageMode=3 then
		'check if admin is logged in
		if session("admin")=-1 then
			pcAutoLoginAllowed=1
		end if
		
		if pcAutoLoginAllowed=1 then
			'// Request "LoginPassword", trim, and set to Session
			pcStrLoginEmail=getUserInput(request("LoginEmail"),250)
			session("pcSFLoginEmail")=pcStrLoginEmail

			'// Request "LoginPassword", trim, and set to Session
			pcStrLoginPassword = session("ppassword")
			if len(pcStrLoginPassword)>0 then
				session("pcSFPassWordExists")="YES"
				session("pcSFLoginPassword") = pcStrLoginPassword
				session("pcSFLoginPassword")=Decrypt(session("pcSFLoginPassword"),9286803311968)
			end if
			session("ppassword") = ""
			if len(session("pcSFLoginEmail"))<1 AND session("idCustomer")=0 then
				response.redirect("checkout.asp?cmode=1&msgcode=1")
			end if
		else
			response.redirect("checkout.asp?cmode=1")
		end if
		'end Autologin
	else
		
		pcs_ValidateEmailField	"LoginEmail", true, 0
		'pcs_ValidateEmailField	"LoginEmail", true
		'pcStrLoginEmail=replace(request.form("LoginEmail"),"'","''")
		'session("pcSFLoginEmail")=pcStrLoginEmail
		'if pcStrLoginEmail="" then
		'	pcv_intErr=pcv_intErr+1
		'End if
		
		'// Request "LoginPassword", trim, and set to Session
		pcs_ValidateTextField "LoginPassword", false, 0
		'pcStrLoginPassword=request.form("LoginPassword")		
		'session("pcSFLoginPassword")=pcStrLoginPassword

		pcs_ValidateTextField "PassWordExists", false, 0
		'session("pcSFPassWordExists")=request.Form("PassWordExists")
		
		'if pcStrLoginPassword="" AND session("pcSFPassWordExists")="YES" then
		if session("pcSFLoginPassword")="" AND session("pcSFPassWordExists")="YES" then		
			pcv_intErr=pcv_intErr+1
		End if

		'if len(pcStrLoginEmail)<1 AND session("idCustomer")=0 then
		if len(session("pcSFLoginEmail"))<1 AND session("idCustomer")=0 then
			response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=1")
		end if		
	end if
	
	if session("ErrLoginEmail")="" AND pcAutoLoginAllowed=0 then
		if scSecurity=1 AND ((scUserLogin=1 AND session("pcSFPassWordExists")="YES") OR (scUserReg=1 AND session("pcSFPassWordExists")<>"YES")) then
			pcv_Test=0
			if (session("store_userlogin")<>"1") AND (session("store_adminre")<>"1") then
				session("store_userlogin")=""
				session("store_adminre")=""
				pcv_test=1
			end if
			if pcv_Test=0 AND session("store_adminre")<>"1" then
				if InStr(ucase(Request.ServerVariables("HTTP_REFERER")),ucase(strSiteSecurityURL & "checkout.asp"))<>1 then
					session("store_userlogin")=""
					session("store_adminre")=""
					pcv_test=1
				end if
				session("store_adminre")=""
			end if
			if pcv_Test=0 AND scUseImgs=1 then %>
                <!-- Include file for CAPTCHA configuration -->
                <!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
                 
                <!-- Include file for CAPTCHA form processing -->
                <!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
            <%	
                If not blnCAPTCHAcodeCorrect then
					session("store_userlogin")=""
					session("store_adminre")=""
					pcv_test=1
					response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=6")
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

				response.redirect("checkout.asp?cmode="&pcPageMode&"&msgmode=4")
			end if					
		end if

	end if
	
	if pcv_intErr=0 then
		erypassword=encrypt((session("pcSFLoginPassword")), 9286803311968)
		session("pcSFEryPassword")=erypassword
		if pcPageMode=0 then
			response.redirect "onepagecheckout.asp"
		else
			'just logging in
			response.redirect "login.asp?lmode=2"
		end if
	else
		'// handle error
	end if
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Section C - Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

pcs_JavaTextField	"LoginEmail", pcv_isLoginEmailRequired, dictLanguage.Item(Session("language")&"_validate_1")
	
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: FORM VALIDATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<div id="pcMain">
	<div id="GlobalAjaxErrorDialog" title="Communication Error" style="display:none">
		<div class="pcErrorMessage">
			Can not connect to server to exchange information. Please contact store owner or try again later
		</div>
	</div>
    
	<% 
	msgMode=getUserInput(request.querystring("msgmode"),1)
    select case msgMode
        case "2"
            msg=dictLanguage.Item(Session("language")&"_validate_2")
			msgClass="pcErrorMessage"
        case "3"
            msg=dictLanguage.Item(Session("language")&"_validate_3")
			msgClass="pcErrorMessage"
        case "4"
            msg=dictLanguage.Item(Session("language")&"_validate_2")
			msgClass="pcErrorMessage"
        case "5"
            msg=dictLanguage.Item(Session("language")&"_validate_4")
			msgClass="pcInfoMessage"
        case "6"
            msg=dictLanguage.Item(Session("language")&"_security_3")
			msgClass="pcErrorMessage"
        case "7"
            msg=dictLanguage.Item(Session("language")&"_validate_5")
			msgClass="pcInfoMessage"
            tmpemail=pcf_FillFormField("LoginEmail", true)
		case "8"
			msg=dictLanguage.Item(Session("language")&"_validate_6")
			msgClass="pcErrorMessage"
            tmpemail=pcf_FillFormField("LoginEmail", true)
        case else
            msg=""
    end select
    
    if msg="" then
        msg=getUserInput(request.querystring("msg"),0)
    end if
    If msg<>"" then	%>
        <div class="<%=msgClass%>">
        <%=msg%>
        </div>
    <% 
	end if 
	%>
    
    <table class="pcMainTable">
    	<tr>
        	<td width="50%" valign="top">
                <form name="LoginForm" method="post" action="checkout.asp" onSubmit="return Form1_Validator(this)" class="pcForms">                
                <input type="hidden" name="cmode" value="<%=pcPageMode%>">
                <table class="pcShowContent">
                    <%
                    if pcPageMode=2 then
                        pcPageTitle=dictLanguage.Item(Session("language")&"_checkout_22")
                    else
                        if pcPageMode=4 then
                            pcPageTitle=dictLanguage.Item(Session("language")&"_checkout_29")
                        else
                            pcPageTitle=dictLanguage.Item(Session("language")&"_checkout_23")
                        end if
                    end if
                    %>				
                    <tr>
                        <td colspan="2"><h1><%=pcPageTitle%></h1></td>
                    </tr>
                    <%if pcPageMode=2 then
                    pcIntEmailNotFound=getUserInput(request("EmailNotFound"),1)
                    if Not ValidNum(pcIntEmailNotFound) then
                        pcIntEmailNotFound=""
                    end if
                    if pcIntEmailNotFound<>"" then %>
                    <tr>
                        <td colspan="2">
                            <% if pcIntEmailNotFound=1 then %>
                                <div class="pcErrorMessage">
                                    <% response.write dictLanguage.Item(Session("language")&"_forgotpasswordexec_2") %>
                                </div>
                          <% else %>
                                <div class="pcSuccessMessage">
                                    <%response.write dictLanguage.Item(Session("language")&"_checkout_11")%>
                                </div>
                            <% end if %>
                        </td>
                    </tr>
                    <%end if
                    end if%>
                    <tr>
                        <td colspan="2">
                            <table class="pcShowContent">
                                <%                                
                                
                                '// Reset to zero to show "Account Login" details even "Only PayPal Express is Enabled"
                                intPPEOnly = 0  
                                
                                if intPPEOnly=0 then %>
                                    <%pcIntOPCEmailNotFound=getUserInput(request("ENotFound"),1)
                                    if session("ErrLoginEmail")<>"" then
                                    session("PCErrLoginEmail")="1" %>
                                        <tr> 
                                            <td colspan="2">
                                                <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_Custmoda_16")%></div>
                                            </td>
                                        </tr>
                                    <%else
                                    if pcPageMode=4 then
                                    if Not ValidNum(pcIntOPCEmailNotFound) then
                                        pcIntOPCEmailNotFound=""
                                    end if
                                    if pcIntOPCEmailNotFound<>"" then%>
                                    <tr>
                                        <td colspan="2">
                                        <% Select Case pcIntOPCEmailNotFound
                                        Case 1:%>
                                            <div class="pcErrorMessage">
                                            <% response.write dictLanguage.Item(Session("language")&"_checkout_30") %>
                                            </div>
                                        <%Case 2:%>
                                                <div class="pcErrorMessage">
                                                <% response.write dictLanguage.Item(Session("language")&"_checkout_31") %>
                                                </div>
                                        <%Case 3:%>
                                                <div class="pcErrorMessage">
                                                <% response.write dictLanguage.Item(Session("language")&"_checkout_35") %>
                                                </div>
                                        <%Case Else:%>
                                                <div class="pcSuccessMessage">
                                                <%response.write dictLanguage.Item(Session("language")&"_checkout_32")%>
                                                </div>
                                        <%End Select%>
                                        </td>
                                    </tr>
                                    <% end if
                                    end if
                                    end if %>
                                    
                                    <%if pcIntEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"3" then%>
                                    <tr>
                                        <td colspan="2">
                                        <p><%=dictLanguage.Item(Session("language")&"_Custmoda_4")%> <input type="text" name="LoginEmail" value="<%=pcf_FillFormField("LoginEmail", true)%>" size="30"><%pcs_RequiredImageTag "LoginEmail", true%></p>
                                        </td>
                                    </tr>
                                    <%end if%>
                                <%
                                end if
                                if pcPageMode=2 OR pcPageMode=4 then %>
                                	<tr>
                                    	<td colspan="2"><hr /></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2">
                                            <input type="hidden" name="fmode" value="<%=pcFromPageMode%>">
                                            <%if pcIntEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"0" AND pcIntOPCEmailNotFound<>"3" then%>
                                            <input type="image" src="<%=rslayout("submit")%>" name="SubmitPM" id="SubmitPM" class="submit">
                                            <%end if%>
                                            &nbsp;<a href="javascript:<%if pcFromPageMode=2 then%>location='onepagecheckout.asp';<%else%><%if pcFromPageMode=4 then%>location='checkout.asp?cmode=1';<%else%>history.go(-1);<%end if%><%end if%>"><img src="<%=rslayout("back")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_15")%>"></a>
                                        </td>
                                    </tr>
                                <%
                                else
                                    if intPPEOnly=0 then %>
                                        <tr>
                                            <td colspan="2" class="pcSpacer"></td>
                                        </tr>
                                        <tr>
                                            <td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_checkout_25")%></p></td>
                                        </tr>
                                        <% if scSecurity=1 AND scUserReg=1 then
                                            pcShowStyle=""
                                        else
                                            pcShowStyle="none"
                                        end if 
                                        if scSecurity=1 AND scUserLogin=1 then
                                            pcShowLoginStyle="" 
                                        else
                                            pcShowLoginStyle="none"
                                        end if %>
                                        <tr>
                                            <td align="right">
                                                <input name="PassWordExists" type="radio" value="YES" checked="checked" <% if scUseImgs=1 then%>onClick="document.getElementById('show_security').style.display='<%=pcShowLoginStyle%>'"<% end if%> class="clearBorder">
                                            </td>
                                            <td width="90%"><p><%=dictLanguage.Item(Session("language")&"_checkout_26")%><input type="password" name="LoginPassword" size="20" onFocus="document.LoginForm.PassWordExists[0].checked=true"></p>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="right">
                                                <input name="PassWordExists" type="radio" value="NO" onFocus="document.LoginForm.LoginPassword.value=''" <% if scUseImgs=1 then%>onClick="document.getElementById('show_security').style.display='<%=pcShowStyle%>'"<% end if %> class="clearBorder">
                                            </td>
                                            <td width="90%">
                                                <p><%=dictLanguage.Item(Session("language")&"_checkout_27")%></p>
                                            </td>
                                        </tr>
                            
                                        <% 'If Advanced Security is turned on
                                        if scSecurity=1 then
                                            Session("store_userlogin")="1"
                                            session("store_adminre")="1"	
                                            if (scUserLogin=1 OR scUserReg=1) and (scUseImgs=1) then %>
                                                <tr>
                                                    <td colspan="2" class="pcSpacer"></td>
                                                </tr>
                                                <tr>
                                                    <td colspan="2">
                                                    <table width="100%" id="show_security" border="0" cellpadding="4" cellspacing="0" style="display:<%=pcShowLoginStyle%>">
                                                    <tr>
                                                        <td><!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" --></td>
                                                    </tr>
                                                </table></td></tr>
                                            <% else 
                                                response.write "<div id=""show_security""></div>"
                                            end if %>
                                        <% else
                                            response.write "<div id=""show_security""></div>"
                                        end if %>
                                        <tr>
                                            <td colspan="2"><hr /></td>
                                        </tr>
                                        <tr>
                                            <td colspan="2">
                                            <% if pcPageMode=1 OR pcPageMode=3 then %>
                                                <input type="image" src="<%=rslayout("submit")%>" name="SubmitCO" id="submit">
                                            <% else %>
                                                <input type="image" src="<%=rslayout("login_checkout")%>" name="SubmitCO" id="submit">
                                            <% end if %>
                                            </td>
                                        </tr>
            
                                    <% 
                                    end if
                                    
                                    
                                    call opendb()
                                    if session("customerType")=1 then
                                        query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 AND (gwCode=999999 OR gwCode=46 OR gwCode=53) ORDER by paymentPriority;"
                                    else
                                        query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 and Cbtob=0 AND (gwCode=999999 OR gwCode=46 OR gwCode=53) ORDER by paymentPriority;"
                                    end if	
                                    set rs=server.CreateObject("ADODB.RecordSet")
                                    set rs=conntemp.execute(query)							
                                    If NOT rs.eof Then
                                        intPayPalExp=1
                                        '// Determine which API to use (US or UK)
                                        query="SELECT pcPay_PayPal.pcPay_PayPal_Partner, pcPay_PayPal.pcPay_PayPal_Vendor FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
                                        set rsPayPalType=Server.CreateObject("ADODB.Recordset")
                                        set rsPayPalType=conntemp.execute(query)
                                        pcPay_PayPal_Partner=rsPayPalType("pcPay_PayPal_Partner")
                                        pcPay_PayPal_Vendor=rsPayPalType("pcPay_PayPal_Vendor")
                                        if pcPay_PayPal_Partner<>"" AND pcPay_PayPal_Vendor<>"" then  
                                            pcPay_PayPal_Version = "UK"			
                                        else
                                            pcPay_PayPal_Version = "US"						
                                        end if
                                        set rsPayPalType=nothing					
                                    Else
                                        intPayPalExp=0
                                    End If
                                    set rs=nothing
                                    
                                    '====================================
                                    ' START: PayPal Express
                                    '====================================						
                                    if intPayPalExp=1 AND pcPageMode<>1 then %>
                                        <tr class="pcSectionTitle"> 
                                            <td colspan="2"><b>Fast, Secure Checkout with PayPal</b></td>
                                        </tr>
                                        <tr valign="top"> 
                                            <td colspan="2" class="pcSpacer"></td>
                                        </tr>
                                        <tr valign="top"> 
                                            <td>
                                                <% '// Display the API Button Code
                                                if pcPay_PayPal_Version = "US" then %>
                                                    <div style="padding-top: 12px;">
                                                        <a href="pcPay_ExpressPay_Start.asp"><img src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" border="0" alt="Acceptance Mark"></a>
                                                    </div>
                                                <% else %>
                                                    <div style="padding-top: 12px;">
                                                        <a href="pcPay_ExpressPayUK_Start.asp"><img src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" border="0" alt="Acceptance Mark"></a>
                                                    </div>
                                                <% end if %>
                                            </td>
                                            <td><p>Save time, Checkout securely. Pay without sharing your financial information.</p></td>
                                        </tr>
                                        <tr valign="top"> 
                                            <td colspan="2" class="pcSpacer"></td>
                                        </tr>
                                    <% end if
                                    '====================================
                                    ' END: PayPal Express
                                    '====================================
                                    %>
            
                                <% end if %>
                            </table>
                        </td>
                    </tr>
                </table>
                </form>
                
		<%
        if pcPageMode<>2 AND pcPageMode<>4 then
            pcIntEmailNotFound=getUserInput(request("EmailNotFound"),1)
            if Not ValidNum(pcIntEmailNotFound) then
                pcIntEmailNotFound=""
            end if
            
            '------------------------------
            ' START: password reminder
            '------------------------------
            
			%>
			<% if pcIntEmailNotFound<>"" AND session("PCErrLoginEmail")<>"1" then %>
					<% if pcIntEmailNotFound=1 then %>
						<div class="pcErrorMessage">
							<% response.write dictLanguage.Item(Session("language")&"_forgotpassworderror") %>
						</div>
				  <% else %>
						<div class="pcSuccessMessage">
							<%response.write dictLanguage.Item(Session("language")&"_checkout_11")%>
						</div>
					<% end if %>
			<% else %>
				<p style="margin-top: 20px;"><img src="images/pcv4_st_icon_info.png" alt="<%response.write dictLanguage.Item(Session("language")&"_Custva_3")%>" title="<%response.write dictLanguage.Item(Session("language")&"_Custva_3")%>" style="margin-right: 5px;"><%response.write dictLanguage.Item(Session("language")&"_Custva_3")%><br />
				<a href="checkout.asp?cmode=2&fmode=<%=pcPageMode%>"><%response.write dictLanguage.Item(Session("language")&"_Custva_8")%></a></p>
			<% end if
			session("PCErrLoginEmail")="" %>
			<% 
            end if
            '------------------------------
            ' END: password reminder
            '------------------------------	
        %>
        </td>
        <%
        if (not (pcPageMode=4 and pcFromPageMode=1)) AND (not (pcPageMode=4 and pcFromPageMode=4)) AND (not (pcPageMode=2 and pcFromPageMode=1)) AND (not (pcPageMode=2 and pcFromPageMode=2)) AND (request("orderReview")<>"no") then
        %>
    	<td width="50%" valign="top">
        <form id="ORVForm" name="ORVForm" class="pcForms">
			<table class="pcShowContent">
                <tr>
                    <td colspan="2"><h1><%=dictLanguage.Item(Session("language")&"_opc_checkout_1")%></h1></td>
                </tr>
                <tr>
                    <td colspan="2" style="padding-bottom: 10px;"><%=dictLanguage.Item(Session("language")&"_opc_checkout_2")%></td>
                </tr>
                <tr>
                    <td width="15%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_5")%></td>
                    <td><input type="text" name="custemail" id="custemail" value="<%=tmpemail%>" size="30" /></td>
                </tr>
                <tr>
                    <td width="15%" valign="middle" nowrap><%=dictLanguage.Item(Session("language")&"_opc_checkout_3")%></td>
                    <td><input type="text" name="ordercode" id="ordercode" value="" size="30" /></td>
                </tr>
                <tr>
                    <td colspan="2">
                        <div name="ORVLoader" id="ORVLoader" style="display:none">
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="2"><hr /></td>
                </tr>
                <tr>
                    <td colspan="2">
                        <input type="image" src="<%=rslayout("submit")%>" name="ORVSubmit" id="ORVSubmit" class="submit">
                    </td>
                </tr>
             </table>
        </form>
        
			<%
            if pcPageMode<>2 AND pcPageMode<>4 then
                pcIntEmailNotFound=getUserInput(request("ENotFound"),1)
                if Not ValidNum(pcIntEmailNotFound) then
                    pcIntEmailNotFound=""
                end if
                
                '------------------------------
                ' START: order code(s) reminder
                '------------------------------
                
            %>
                <table class="pcMainTable">			
                    <tr> 
                        <td>
                        <p style="margin-top: 20px;"><img src="images/pcv4_st_icon_info.png" alt="<%response.write dictLanguage.Item(Session("language")&"_checkout_33")%>" title="<%response.write dictLanguage.Item(Session("language")&"_checkout_33")%>" style="margin-right: 5px;"><%response.write dictLanguage.Item(Session("language")&"_checkout_33")%><br />
                        <a href="checkout.asp?cmode=4&fmode=4"><%response.write dictLanguage.Item(Session("language")&"_checkout_34")%></a></p>
                        </td>
                    </tr>
                </table>
            <% 
                '------------------------------
                ' END: order code(s) reminder
                '------------------------------
            end if
            %>
        </td>
<%
	end if
%>
	</tr>
</table>

<script>
$(document).ready(function()
{
	jQuery.validator.setDefaults({
		success: function(element) {
			$(element).parent("td").children("input, textarea").addClass("success")
		}
	});
	
	//*Ajax Global Settings
	$("#GlobalAjaxErrorDialog").ajaxError(function(event, request, settings){
		$(this).dialog('open');
		$("#ORVLoader").hide();
	});

	
	//*Dialogs
	$("#GlobalAjaxErrorDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 450,
			height: 230,
			modal: true,
			buttons: {
				' OK ': function() {
						$(this).dialog('close');
					}
			}
	});
	
	//*Validate Order Review Form
	$("#ORVForm").validate({
		rules: {
			custemail: 
			{
				required: true,
				email: true
			},
			ordercode: "required"
		},
		messages: {
			custemail: {
				required: "<%=dictLanguage.Item(Session("language")&"_opc_js_2")%>",
				email: "<%=dictLanguage.Item(Session("language")&"_opc_js_3")%>"
			},
			ordercode: {
				required: "<%=dictLanguage.Item(Session("language")&"_opc_checkout_4")%>"
			}
		}
	});
	$('#ORVSubmit').click(function(){
		if ($('#ORVForm').validate().form())
		{
			$("#ORVLoader").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_checkout_5")%>');
			$("#ORVLoader").show();	
			$.ajax({
				type: "POST",
				url: "opc_checkORV.asp",
				data: $('#ORVForm').formSerialize(),
				timeout: 5000,
				success: function(data, textStatus){
					if (data.indexOf("OK")>=0)
					{
						var tmpArr=data.split("|*|")
						$("#ORVLoader").html('<div class=pcSuccessMessage><%=dictLanguage.Item(Session("language")&"_opc_checkout_6")%></div>');
						var callbackBill=function (){setTimeout(function(){$("#ORVLoader").hide();},1000);}
						$("#ORVLoader").effect('',{},500,callbackBill);
						location=tmpArr[1];
					}
					else
					{
						$("#ORVLoader").html('<div class=pcErrorMessage> '+data+' </div>');
						var callbackBill=function (){setTimeout(function(){$("#ORVLoader").hide();},1000);}
						$("#ORVLoader").effect('',{},500,callbackBill);
					}
				}
	 		});
			return(false);
		}
		return(false);
	});
});
</script>
	
</div>

<% 
'// Managed Form Sessions Auto-Cleared
'session("ErrLoginEmail")=""
'session("pcSFLoginEmail")=""

'// Clear Un-Managed Sessions
session("pcSFLoginPassword")=""
session("pcSFPassWordExists")=""
session("pcSFEryPassword")=""
call closedb()
%><!--#include file="footer.asp"-->