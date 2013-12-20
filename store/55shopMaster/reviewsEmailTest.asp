<%@Language="VBScript"%>
<%
' PRV41 Start
On Error goto 0
 Dim rs, connTemp, query
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
<TITLE>Product Reviews E-mail Test</TITLE>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
<style>
td { font-size: 12px}
</style>
</HEAD>
<body style="background-image:none;">
<form action="ReviewsEmailTest.asp" method="post" class="pcForms">
<input type="hidden" name="pcFormAction" value="send">
<table class="pcCPcontent" style="width: 100%;">
    <tr> 
        <th colspan="2">Test the Product Review Reminder E-mail</th>
    </tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>

<% dim pcv_fromname, pcv_fromemail, pcv_toname, pcv_toemail, pcv_subject, pcv_message, pcv_success

pcv_FormAction=request.Form("pcFormAction")

if pcv_FormAction = "send" then

	pcv_fromname=scCompanyName
	pcv_fromemail=getUserInput(request.form("pcFromEmail"),0)
	pcv_toname="Store Administrator"
	pcv_toemail=getUserInput(request.form("pcAdminEmail"),0)
	pcv_subject="ProductCart Reviews Email Test Message" & " - " & replace(scCompanyName,"'","") 
	pcv_errMsg=""


	'''''''''''''''''''''''''''''''''''
	Call opendb()

    If request.Form("format")="0" Or request.Form("format")="1" Then
       connTemp.execute "UPDATE pcRevSettings SET pcRS_sendReviewReminderFormat=" & CLng(request.Form("format"))
       response.write "<script>window.opener.document.hForm.sendreviewreminderformat[0].checked="
       If request.Form("format")="0" Then response.write "true" Else response.write "false"
       response.write ";window.opener.document.hForm.sendreviewreminderformat[1].checked="
       If request.Form("format")="1" Then response.write "true" Else response.write "false"
       response.write ";</script>"
    End if

	set rs=server.CreateObject("ADODB.RecordSet")
	query = "SELECT TOP 1 pcRS_RewardForReview, pcRS_sendReviewReminderTemplate, pcRS_sendReviewReminderType, pcRS_sendReviewReminderFormat, pcRS_RewardForReviewURL, pcRS_RewardForReviewFirstPts FROM pcRevSettings"
	Set rs = connTemp.execute(query)
	If rs.eof = False Then
	
	   pcV_sendReviewReminderFormat = RS("pcRS_sendReviewReminderFormat")
	   pcV_RewardForReviewFirstPts = RS("pcRS_RewardForReviewFirstPts")
	   	if not validNum(pcV_RewardForReviewFirstPts) or pcV_RewardForReviewFirstPts=0 then
			pcV_RewardForReviewFirstPts=5
		end if

       ' If admin has not defined a custom template, use what's in languages.asp
       strMessage = dictLanguage.Item(Session("language")&"_prv_28")
       If pcV_sendReviewReminderFormat = 0 Then
		 if RewardsActive <> 0 then	   	
          strMessage = strMessage & dictLanguage.Item(Session("language")&"_prv_42") & vbCRLF & vbCRLF
		 end if
	      strMessage = strMessage & dictLanguage.Item(Session("language")&"_prv_40")          
       else
		 if RewardsActive <> 0 then
          strMessage = strMessage & dictLanguage.Item(Session("language")&"_prv_29")
		 end if
	      strMessage = strMessage & dictLanguage.Item(Session("language")&"_prv_38")
       End if

	   If Len(Trim(rs("pcRS_SendReviewReminderTemplate")&""))>0 Then
	      Dim strTempRead, pcIntUseTemplate
		  strTempRead = strReadAll(server.mappath("../pc/Library/" & RS("pcRS_SendReviewReminderTemplate")))
	      If Len(strTempRead) > 0 Then 
		  	strMessage = strTempRead
			pcIntUseTemplate = 1
		  else
		    pcIntUseTemplate = 0
		  end if		  
	   End If

	   If pcIntUseTemplate = 0 then
		   If pcV_sendReviewReminderFormat = 0 Then
				strMessage = strmessage & dictLanguage.Item(Session("language")&"_prv_41")
		   else
				strMessage = strMessage & dictLanguage.Item(Session("language")&"_prv_32")
		   End if
	   End if
	   pcv_RewardForReviewURL = rs("pcRS_RewardForReviewURL")
	Else
	   pcV_sendReviewReminderFormat = 1
	   pcv_RewardForReviewURL = ""
    End If
	rs.close
	Set rs = nothing
	Call closeDB()


    dim strPath, iCnt, strPathInfo

    strPathInfo=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
    strPathInfo=replace(strPathInfo,"https:/","https://")
    strPathInfo=replace(strPathInfo,"http:/","http://")

    
	pcv_message = strMessage
	pcv_message = Replace(pcv_message, "<CUSTOMER_NAME>","Jane Customer",1,-1,vbTextCompare)
	pcv_message = Replace(pcv_message, "<NUMBER_OF_POINTS>",pcV_RewardForReviewFirstPts,1,-1,vbTextCompare)
	pcv_message = Replace(pcv_message, "<REWARD_POINTS_LABEL>",RewardsLabel,1,-1,vbTextCompare)
	pcv_message = Replace(pcv_message,"<RFR_PAGE>", "<a href=""" & pcV_RewardForReviewURL & """>",1,-1,vbTextCompare)
	pcv_message = Replace(pcv_message,"</RFR_PAGE>", "</a>",1,-1,vbTextCompare)

	pcv_message = Replace(pcv_message,"<RFR_PAGE_TEXT>", pcv_RewardForReviewURL,1,-1,vbTextCompare)

    pcv_message = Replace(pcv_message,"<POST_REVIEW_LINK>", "<a href=""" & strPathInfo & "prv_Vieworder.asp?uid=" & strNewGUID & """>",1,-1,vbTextCompare)
    pcv_message = Replace(pcv_message,"</POST_REVIEW_LINK>", "</a>",1,-1,vbTextCompare)
    pcv_message = Replace(pcv_message,"<POST_REVIEW_LINK_TEXT>", strPathInfo & "prv_Vieworder.asp?uid=" & strNewGUID,1,-1,vbTextCompare)

    pcv_message = Replace(pcv_message,"<STORE_NAME>", scCompanyName,1,-1,vbTextCompare)
	pcv_message = Replace(pcv_message,"<PRV_UNSUBSCRIBE>", strPathInfo & "prv_unsubscribe.asp?uid=" & strNewGUID,1,-1,vbTextCompare)

    session("News_MsgType")=pcV_sendReviewReminderFormat

    If pcV_sendReviewReminderFormat=0 Then 
		pcv_message=Replace(pcv_message, "<br />", vbCRLF)
		pcv_message=Replace(pcv_message, "<br>", vbCRLF)
    End if

	'''''''''''''''''''''''''''''''''''
	
	call sendmail (pcv_fromname, pcv_fromemail, pcv_toemail, pcv_subject, pcv_message)
	session("News_MsgType")=""
	if pcv_errMsg<>"" then
		pcv_err = InStr(1,pcv_errMsg,"Object required",1) %>
			 <tr> 
				<td><img src="images/pcv4_icon_alert.gif"></td>
                <td>
				<% if pcv_err > 0 then %>
					You have selected an email component that is not supported on this server
				<% else
					response.write pcv_errMsg
				end if %>					
				</td>
			</tr>
	<% else %>
             <tr> 
                <td colspan="2" align="center"><div class="pcCPmessageSuccess">Message sent successfully!<br><br>Check your e-mail to see how it looks. Please note that the links included in the message <strong>will not work</strong> as the product and customer ID will not be included. This is just a sample message.</td>
            </tr>
	<% end if	
else %>

		<tr> 
			<td align="right" nowrap>Email Component:</td>
			<td><%=scEmailComObj%></td>
		</tr>
		<tr> 
            <td align="right" nowrap>SMTP Server:</td>
			<td><%=scSMTP%></td>
		</tr>
		<tr> 
            <td align="right">From Email:</td>
			<td><input type="text" value="<%=scEmail%>" name="pcFromEmail"></td>
		</tr>
		<tr> 
			<td align="right">Admin Email:</td>
			<td><input type="text" value="<%=scFrmEmail%>" name="pcAdminEmail"></td>
		</tr>
        <tr valign="top">
           <td align="right" nowrap>Current Format:</td>
           <td><%
	          Call opendb()
              set rs=server.CreateObject("ADODB.RecordSet")
	          query = "SELECT TOP 1 pcRS_RewardForReview, pcRS_sendReviewReminderTemplate, pcRS_sendReviewReminderType, pcRS_sendReviewReminderFormat, pcRS_RewardForReviewURL, pcRS_sendReviewReminderFormat FROM pcRevSettings"
	          Set rs = connTemp.execute(query)
 	          If rs.eof = False Then
	             pcV_sendReviewReminderFormat = RS("pcRS_sendReviewReminderFormat")
              else
	             pcV_sendReviewReminderFormat = 1
              End If
              rs.close
              Set rs = Nothing
              Call closedb()

              response.write "<input type=""radio"" name=""format"" value=""0"""
              If pcV_sendReviewReminderFormat<>1 Then
                 response.write " CHECKED"
              End If
              response.write "> Text Only&nbsp;&nbsp;"
              response.write "<input type=""radio"" name=""format"" value=""1"""
              If pcV_sendReviewReminderFormat=1 Then
                 response.write " CHECKED"
              End If
              response.write "> HTML"
           %>
           </td>
        </tr>
		<tr valign="top"> 
       	  <td colspan="2" align="center"><div class='pcCPmessage'>You can change <strong>format</strong> on the <strong><em> Settings</em></strong> page or using the selection above. The message sent will either be the default product review message, or a custom message (if you have uploaded your own .TXT file)</div></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"><input type="submit" value="Send Test Message" class="submit2"></td>
		</tr>
<% end if %>
</table>
</form>
<div align="center" style="margin-top: 10px;"><a href="#" onClick="self.close();">Close Window</a> : <a href="http://wiki.earlyimpact.com/productcart/products_reviews#write_a_review_reminder" target="_blank">Documentation</a></div>
</BODY>
</HTML>
<% 'PRV41 end %>