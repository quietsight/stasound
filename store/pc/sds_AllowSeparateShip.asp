<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languages_ship.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

Dim connTemp,query,rs

pcv_request=getUserInput(request("req"),0)
if pcv_request="" then
	response.redirect "default.asp"
end if

call opendb()%>
<!--#include file="header.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td colspan="2"> 
			<h1><%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_1")%></h1>
		</td>
	</tr>
	<%IF request("action")="upd" THEN
		pcv_CustAllow=getUserInput(request("R1"),0)
		if (pcv_CustAllow="") OR (not IsNumeric(pcv_CustAllow)) then
			pcv_CustAllow="1"
		end if
		query="SELECT idorder FROM Orders WHERE pcOrd_CustRequestStr='" & pcv_request &"';"
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if rs.eof then%>
		<tr>
			<td colspan="2">
				<div class="pcErrorMessage">
					<%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_5")%>
				</div>
			</td>
		</tr>
		<%response.end
		else
			pcv_IDOrder=rs("idorder")
		end if
		set rs=nothing
		
		query="UPDATE Orders SET pcOrd_CustAllowSeparate=" & pcv_CustAllow & " WHERE pcOrd_CustRequestStr='" & pcv_request &"';"
		set rs=connTemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		set rs=nothing		
		
		'Send Notification E-mail to Store Owner
		pcv_AdmSbj=replace(ship_dictLanguage.Item(Session("language")&"_notifyseparate_sbj_1"),"<ORDER_ID>",(scpre + int(pcv_IDOrder)))
		pcv_AdmMail=""
		if pcv_CustAllow="1" then
			pcv_AdmMail=replace(ship_dictLanguage.Item(Session("language")&"_notifyseparate_msg_1"),"<ORDER_ID>",(scpre + int(pcv_IDOrder))) & vbcrlf
		else
			pcv_AdmMail=replace(ship_dictLanguage.Item(Session("language")&"_notifyseparate_msg_2"),"<ORDER_ID>",(scpre + int(pcv_IDOrder))) & vbcrlf
		end if
				
		strPath=Request.ServerVariables("PATH_INFO")
		iCnt=0
		do while iCnt<2
			if mid(strPath,len(strPath),1)="/" then
				iCnt=iCnt+1
			end if
			if iCnt<2 then
				strPath=mid(strPath,1,len(strPath)-1)
			end if
		loop
	
		strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
				
		if Right(strPathInfo,1)="/" then
		else
			strPathInfo=strPathInfo & "/"
		end if
		
		strPathInfo=strPathInfo & scAdminFolderName & "/OrdDetails.asp?id=" & pcv_IDOrder
		pcv_AdmMail=pcv_AdmMail & strPathInfo
		call sendmail (scCompanyName, scEmail, scFrmEmail, pcv_AdmSbj, pcv_AdmMail)
		'End of Send Notification E-mail to Store Owner
		%>
		<tr>
			<td colspan="2">
				<div class="pcErrorMessage">
					<%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_7")%>
					<br /><br />
					<a href="default.asp"><%=dictLanguage.Item(Session("language")&"_titles_5")%></a>
				</div>
			</td>
		</tr>
	<%ELSE
		query="SELECT idorder,pcOrd_CustAllowSeparate FROM Orders WHERE pcOrd_CustRequestStr='" & pcv_request &"';"
		set rs=connTemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		if rs.eof then%>
		<tr>
			<td colspan="2">
				<div class="pcErrorMessage">
					<%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_5")%>
				</div>
			</td>
		</tr>
		<%else
			pcv_IDOrder=rs("idorder")
			pcv_CustAllow=rs("pcOrd_CustAllowSeparate")
			if IsNull(pcv_CustAllow) or pcv_CustAllow="" then
				pcv_CustAllow=0
			end if
			if pcv_CustAllow>0 then%>
				<tr>
					<td colspan="2">
						<div class="pcErrorMessage">
							<%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_6")%>
						</div>
					</td>
				</tr>
			<%else%>
			<form method="post" action="sds_AllowSeparateShip.asp?action=upd" name="form1">
			<tr>
				<td colspan="2"><%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_2")%>&nbsp;<b><%=(scpre + int(pcv_IDOrder))%></b></td>
			</tr>
			<tr>
				<td colspan="2"><%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_2a")%></td>
			</tr>
			<tr>
				<td colspan="2" class="pcSpacer"></td>
			</tr>
			<tr>
				<td><input type="radio" name="R1" value="1" checked class="clearBorder"></td>
				<td><%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_3")%></td>
			</tr>
			<tr>
				<td><input type="radio" name="R1" value="2" class="clearBorder"></td>
				<td><%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_4")%>
				<input type="hidden" name="req" value="<%=pcv_request%>"></td>
			</tr>
			<tr>
				<td colspan="2" class="pcSpacer"></td>
			</tr>
			<tr>
				<td colspan="2"><input src="<%=rslayout("submit")%>" type="image" name="Confirm" id="submit"></td>
			</tr>
			</form>
			<%end if
		end if
		set rs=nothing
	END IF%>
</table>
</div>
<%call closedb()%><!--#include file="footer.asp"-->