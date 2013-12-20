<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle = "" 
pageIcon = ""
%>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/UpdateVersionCheck.asp"-->
<!--#include file="../includes/PPDStatus.inc"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<%
   Dim pSendReviewReminderDays, pSendReviewReminderType, pSendReviewReminderFormat, pSendReviewReminderTemplate
   Dim pRewardForReview, pRewardForReviewFirstPts, pRewardForReviewAdditionalPts, pRewardForReviewURL
   Dim compareDate, flgDataFound
   Dim connTemp,query,rs,flgEmailRunNeeded
   flgEmailRunNeeded = False
   
   Dim AutoSendActionURL
   AutoSendActionURL=""
   
   '// Manual, order specific request
   Dim pcIntManualRequest, pcIntOrderID
   pcIntManualRequest = 0
   pcIntOrderID = getUserInput(request("idOrder"),10)
   if not validNum(pcIntOrderID) then pcIntOrderID = 0
   if pcIntOrderID > 0 then
   	pcIntManualRequest = 1
   end if
   

   Call openDB()
	
    query = "SELECT TOP 1 pcRS_Active, pcRS_LastRunDate FROM pcRevSettings"
	Set rs = connTemp.execute(query)
	If NOT rs.eof Then
		pcv_Active=rs("pcRS_Active")
		if isNull(pcv_Active) or pcv_Active="" then
			pcv_Active="0"
		end if

	   	If pcv_Active="1" and IsDate(rs("pcRS_LastRunDate")) Then
	      	If DateValue(rs("pcRS_LastRunDate")) < DateValue(now) Then
		     	flgEmailRunNeeded = True
		  	End if
	   	Else
			If pcv_Active="1" AND (IsNull(rs("pcRS_LastRunDate")) OR (rs("pcRS_LastRunDate")="")) Then
				flgEmailRunNeeded = True
			Else
	      		flgEmailRunNeeded = False
			End If
	   End If
	End If
	Set rs = Nothing
	
	'// Override when there is manual, order-specific request
	if pcIntManualRequest = 1 then
		flgEmailRunNeeded = True
	end if

If flgEmailRunNeeded = True then
	
   query = "UPDATE pcRevSettings SET pcRS_LastRunDate=" & formatDateForDB(now)
   connTemp.execute query

   query = "SELECT TOP 1 pcRS_SendReviewReminderDays, pcRS_SendReviewReminderType, pcRS_SendReviewReminderFormat, pcRS_SendReviewReminderTemplate, pcRS_rewardForReview, pcRS_RewardForReviewFirstPts, pcRS_RewardForReviewAdditionalPts, pcRS_RewardForReviewURL FROM pcRevSettings WHERE pcRS_SendReviewReminder=1"
   Set rs = connTemp.execute(query)
   If rs.eof Then
   	  set rs=nothing
	  call closedb()
		if pcIntManualRequest = 1 then
			response.Redirect "ordDetails.asp?id=" & pcIntOrderID & "&msg=" & Server.URLEncode("Message not sent. The 'Write a Review' reminder feature is not active in this store. You can activate it using the Products Reviews Settings.")
			response.End()
		else
	  		response.Write("0")
			response.End   ' This is an automated process with no UI, so we're just going to end
		end if
   End If
   
   pSendReviewReminderDays = rs("pcRS_SendReviewReminderDays")
   pSendReviewReminderType = rs("pcRS_SendReviewReminderType")
   pSendReviewReminderFormat = rs("pcRS_SendReviewReminderFormat")
   pSendReviewReminderTemplate = rs("pcRS_SendReviewReminderTemplate")
   pRewardForReview = rs("pcRS_RewardForReview")
   pRewardForReviewFirstPts = rs("pcRS_RewardForReviewFirstPts")
   pRewardForReviewAdditionalPts = rs("pcRS_RewardForReviewAdditionalPts")
   pRewardForReviewURL = rs("pcRS_RewardForReviewURL")
   compareDate = DateAdd("d", pSendReviewReminderDays*-1, now)
   rs.close
   Set rs = nothing

   flgDataFound = False
   query="SELECT orders.idOrder, customers.Name, customers.lastName, customers.email, orders.idCustomer FROM orders, customers WHERE orders.orderStatus in (3,4,7,8,10,12) "

   If pSendReviewReminderType=0 Then
      query = query & " AND (processDate<=" & formatDateForDB(DateValue(comparedate) & " 23:59:59 PM") & " AND processDate>=" & formatDateForDB(DateValue(DateAdd("d",-45,comparedate)) & " 00:00:01 AM") & ") "
   Else
      query = query & " AND ( "
      query = query & "(shipDate<=" & formatDateForDB(DateValue(comparedate) & " 23:59:59 PM") & " AND shipDate>=" & formatDateForDB(DateValue(DateAdd("d",-45,comparedate)) & " 00:00:01 AM") & ") "
      query = query & "OR orders.idOrder IN (SELECT idOrder FROM pcPackageInfo WHERE pcPackageInfo_ShippedDate<=" & formatDateForDB(DateValue(comparedate) & " 23:59:59 PM") & " AND pcPackageInfo_ShippedDate>=" & formatDateForDB(DateValue(DateAdd("d",-45,comparedate)) & " 00:00:01 AM") & ") "
      query = query & ") "
   End If

   query = query & " AND orders.idOrder NOT IN (SELECT DISTINCT pcRN_idOrder FROM pcReviewNotifications) AND customers.idCustomer=orders.idCustomer and customers.pcCust_AllowReviewEmails<>0"
   
	'// Override when there is manual, order-specific request
	if pcIntManualRequest = 1 then
	   query="SELECT orders.idOrder, customers.Name, customers.lastName, customers.email, orders.idCustomer FROM orders, customers WHERE customers.idCustomer=orders.idCustomer AND orders.idOrder = " & pcIntOrderID
	end if

	set rs=conntemp.execute(query)

	If rs.eof = False then	' Do a getRows and close the recordset, since sending emails can take so long sometimes, and we don't want to hold it open...
	   aryRows = rs.getRows()
	   flgDataFound = True
	End if
	
	rs.close
	Set rs = Nothing

	If flgDataFound Then

	   Dim strMessage

       ' If admin has not defined a custom template, use what's in languages.asp
       strMessage = dictLanguage.Item(Session("language")&"_prv_28")
       If pSendReviewReminderFormat = 0 Then
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

	   If Len(Trim(pSendReviewReminderTemplate&""))>0 Then
	      Dim strTempRead
		  strTempRead = strReadAll(server.mappath("../pc/Library/" & pSendReviewReminderTemplate))
	      If Len(strTempRead) > 0 Then 
		  	strMessage = strTempRead
			pcIntUseTemplate = 1
		  else
		    pcIntUseTemplate = 0
		  end if	
	   End If

	   If pcIntUseTemplate = 0 then
			If pSendReviewReminderFormat = 0 Then
				strMessage = strmessage & dictLanguage.Item(Session("language")&"_prv_41")
			Else
				strMessage = strMessage & dictLanguage.Item(Session("language")&"_prv_32")
			End if	
	   End if
	   
	   session("News_MsgType") = pSendReviewReminderFormat

	   Dim intX, strNewMessage
	   Dim strPath

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
	
		if Right(strPathInfo,1)<>"/" then
			strPathInfo=strPathInfo & "/"
		end if
		msgErr=""
	   For intX = 0 To UBound(aryRows, 2)
	   
		'Check availabe products
		queryQ="SELECT products.idProduct FROM Products INNER JOIN ProductsOrdered ON Products.idProduct=ProductsOrdered.idProduct WHERE ProductsOrdered.idOrder=" & aryRows(0, intX) & " AND Products.active<>0 AND Products.removed=0 AND (Products.idProduct NOT IN (SELECT pcRE_IDProduct FROM pcRevExc));"
		set rsQ=connTemp.execute(queryQ)
		if not rsQ.eof then
			set rsQ=nothing
		  Dim strNewGUID
		  strNewGuid = genGUID()
		  query = "INSERT INTO pcReviewNotifications (pcRN_idCustomer, pcRN_idOrder, pcRN_UniqueID, pcRN_DateSent) values (" & aryRows(4, intX) & "," & aryRows(0, intX) & ",'" & strNewGUID & "','" & Now & "')"
		  connTemp.execute query


		  strNewMessage = Replace(strMessage,"<CUSTOMER_NAME>", properCase(aryRows(1, intX)) & " " & properCase(aryRows(2, intX)),1,-1,vbTextCompare)

          Set rs = connTemp.execute("SELECT count(*) FROM pcReviewPoints WHERE pcRP_IDCustomer=" & aryRows(4, intX))
		  If CLng(rs(0))=0 then
	         strNewMessage = Replace(strNewMessage,"<NUMBER_OF_POINTS>", pRewardForReviewFirstPts,1,-1,vbTextCompare)
		  Else
	         strNewMessage = Replace(strNewMessage,"<NUMBER_OF_POINTS>", pRewardForReviewAdditionalPts,1,-1,vbTextCompare)
		  End If
		  
	      strNewMessage = Replace(strNewMessage,"<REWARD_POINTS_LABEL>", RewardsLabel,1,-1,vbTextCompare)

		  If InStr(1,strMessage,"<RFR_PAGE>",vbTextCompare) Then
		     If Len(Trim(pRewardForReviewURL&""))>0 Then
			 	strNewMessage = Replace(strNewMessage,"<RFR_PAGE>", "<a href=""" & pRewardForReviewURL & """>",1,-1,vbTextCompare)
	            strNewMessage = Replace(strNewMessage,"</RFR_PAGE>", "</a>",1,-1,vbTextCompare)
		     End if
		  End If

		  If InStr(1,strMessage,"<RFR_PAGE_TEXT>",vbTextCompare) Then
		  	strNewMessage = Replace(strNewMessage,"<RFR_PAGE_TEXT>", pRewardForReviewURL,1,-1,vbTextCompare)
		  End If
		  
		  If InStr(1,strMessage,"<POST_REVIEW_LINK>",vbTextCompare) Then

	         	strNewMessage = Replace(strNewMessage,"<POST_REVIEW_LINK>", "<a href=""" & strPathInfo & "pc/prv_Vieworder.asp?uid=" & strNewGUID & """>",1,-1,vbTextCompare)
	         	strNewMessage = Replace(strNewMessage,"</POST_REVIEW_LINK>", "</a>",1,-1,vbTextCompare)
		  End if
		  
		  If InStr(1,strMessage,"<POST_REVIEW_LINK_TEXT>",vbTextCompare) Then
		  		strNewMessage = Replace(strNewMessage,"<POST_REVIEW_LINK_TEXT>", strPathInfo & "pc/prv_Vieworder.asp?uid=" & strNewGUID,1,-1,vbTextCompare)
		  End if
				
	      strNewMessage = Replace(strNewMessage,"<STORE_NAME>", scCompanyName,1,-1,vbTextCompare)

		  If InStr(1,strMessage,"<PRV_UNSUBSCRIBE>",vbTextCompare) Then
	      	strNewMessage = Replace(strNewMessage,"<PRV_UNSUBSCRIBE>", strPathInfo & "pc/prv_unsubscribe.asp?uid=" & strNewGUID,1,-1,vbTextCompare)		 
		  End If
          
          If pSendReviewReminderFormat=0 Then 
			strNewMessage=Replace(strNewMessage, "<br />", vbCRLF)
			strNewMessage=Replace(strNewMessage, "<br>", vbCRLF)
          End if

		  Dim pcvEmailSubject
		  pcvEmailSubject = scCompanyName & " - " & dictLanguage.Item(Session("language")&"_prv_44")
		  pcvEmailSubject = replace(pcvEmailSubject,",","")

          call sendmail(scCompanyName, scEmail, aryRows(3, intX), pcvEmailSubject, strNewMessage)
		else
			msgErr="Cannot send 'Write a Review' reminder because all products of this order cannot be reviewed."  
		end if
		set rsQ=nothing
	   next
	End if 'flgDataFound

End if 'flgEmailRunNeeded
session("News_MsgType")=""
Call closedb()  

'// Redirect to order details page when there is manual, order-specific request
if pcIntManualRequest=1 then
	if msgErr<>"" then
		response.Redirect "ordDetails.asp?id=" & pcIntOrderID & "&s=0&msg=" & Server.URLEncode(msgErr)
	else
		response.Redirect "ordDetails.asp?id=" & pcIntOrderID & "&s=1&msg=" & Server.URLEncode("'Write a Review' reminder successfully sent.")
	end if
	response.End()
end if

If len(AutoSendActionURL)=0 Then
	response.Write("0")
Else
	response.Write(AutoSendActionURL)
End If
response.End()
%>