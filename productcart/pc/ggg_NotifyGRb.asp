<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% 
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

dim pc_fromname, pc_fromemail, pc_toname, pc_toemail, pc_pname, pc_subject, pc_message, pc_pid
if request("action")="send" then
	pc_fromname=request.form("yourname")
	pc_fromemail=request.form("youremail")
	maillist=request.form("friendsemail")
	pc_subject=request.form("title") 
	pc_message=request.form("message")

	pc_toemail=split(maillist,vbcrlf)

	For k=lbound(pc_toemail) to ubound(pc_toemail)
		if pc_toemail(k)<>"" then
			call sendmail (pc_fromname, pc_fromemail, pc_toemail(k), pc_subject, pc_message)
		end if
	Next
	response.redirect "msg.asp?message=96"
end if
%>