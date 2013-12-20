<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/validation.asp" --> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="pcStartSession.asp"-->
<!--#include file="DBsv.asp"--> 
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="header.asp"-->
<%
dim rs, fName, fLastname, fEmail, fPassword, fFrom, fFromName, fSubject, fBody, ftmp,ftmp1,ftmp2

fEmail=replace(trim(request.querystring("email")),"'","''")
redirectUrl= server.HTMLEncode(Session("pcSF_redirectUrl"))
Session("pcSF_redirectUrl")=""
frURL=server.HTMLEncode(Session("pcSF_pcfrUrl"))
Session("pcSF_pcfrUrl")=""
mySQL="SELECT Affiliatename, AffiliateEmail, [pcAff_Password] from Affiliates WHERE AffiliateEmail='" &fEmail& "'"

call opendb()
set rs=conntemp.execute(mySQL)	
if not rs.eof then
	fName=rs("Affiliatename")
	fEmail=rs("Affiliateemail")
	fPassword=enDeCrypt(rs("pcAff_Password"),scCrypPass)		
	fSubject=dictLanguage.Item(Session("language")&"_forgotpasswordmailsubject")
	fBody=dictLanguage.Item(Session("language")&"_forgotpasswordmailbody2")

	fBody=replace(fBody,"#password",fPassword)	
	fBody=replace(fBody,"#name",fName)      
	
	call sendmail (scEmail, scEmail, fEmail, fSubject, fBody) 
%>
	<div id="pcMain">
		<div class="pcErrorMessage">
			<%response.write dictLanguage.Item(Session("language")&"_checkout_11")%>
			<p>
			<%
			if frURL<>"" then
				call closeDb()%>
				<a href="<%=frURL&"?redirectUrl="&Server.Urlencode(redirectUrl)&"&s=1"%>"><img src="<%=rslayout("submit")%>"></a> 
			<%else
				call closeDb()%>
				<a href="AffiliateLogin.asp?s=1"><img src="<%=rslayout("submit")%>"></a>
			<%end if
			call clearLanguage()%>
			</p>
		</div>
	</div>
<% else %>
	<%
	call closeDb()
	response.redirect "msg.asp?message=2"
	%>
<% end if %>
<!--#include file="footer.asp"-->