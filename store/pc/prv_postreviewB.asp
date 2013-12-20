<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<% 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/securitysettings.asp" -->

<% 
' Check if the store is on. If store is turned off display store message
If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If

Function TestRegCust(tmpValue)
Dim tmp1
	tmp1=0
	if tmpValue="0" then
		if session("REGidCustomer")>"0" then
			tmp1=int(session("REGidCustomer"))
		end if
	else
		tmp1=tmpValue
	end if
	TestRegCust=int(tmp1)
End Function

Dim connTemp, rs, query
Dim BadList,intBadCount
call opendb()
%>
<!--#include file="prv_getsettings.asp"-->
<!--#include file="prv_recalc.asp"-->
<%
if pcv_Active<>"1" then
	call closedb()
	response.redirect "default.asp"
end if

pcv_IDProduct=GetUserInput(request("IDProduct"),0)
if not validNum(pcv_IDProduct) then 
	call closedb()
	response.redirect "msg.asp?message=207"
end if

pIdCustomer=GetUserInput(request("idcustomer"),0)
	if not validNum(pIdCustomer) then
		call closedb()
		response.redirect "msg.asp?message=210"
	end if
query="SELECT pcRE_IDProduct FROM pcRevExc WHERE pcRE_IDProduct=" & pcv_IDProduct
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if not rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "default.asp"
end if

pcv_IPAddress=Request.ServerVariables("REMOTE_ADDR")

query="SELECT pcRev_IDReview FROM pcReviews where pcRev_IP='" & pcv_IPAddress & "' and pcRev_IDProduct=" & pcv_IDProduct
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

Count=0

do while not rs.eof
	Count=Count+1
	rs.MoveNext
loop

Count1=getUserInput(Request.Cookies("Prd" & pcv_IDProduct),0)
if Count1="" then
	Count1=0
end if

IF (clng(Count)>=clng(pcv_PostCount)) and (pcv_LockPost="0") THEN
	set rs=nothing
	call closedb()
	response.redirect "prv_denied.asp"
END IF

IF (clng(Count1)>=clng(pcv_PostCount)) and (pcv_LockPost="1") THEN
	set rs=nothing
	call closedb()
	response.redirect "prv_denied.asp"
END IF

IF ((clng(Count)>=clng(pcv_PostCount)) or (clng(Count1)>=clng(pcv_PostCount))) and (pcv_LockPost="2") THEN
	set rs=nothing
	call closedb()
	response.redirect "prv_denied.asp"
END IF

pcv_Feel=GetUserInput(request("feel"),0)
if pcv_Feel="" then
	pcv_Feel="0"
end if
pcv_Rate=GetUserInput(request("rate"),0)
if pcv_Rate="" then
	pcv_Rate="0"
end if

query="SELECT pcRBW_word FROM pcRevBadWords"
set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

intBadCount=-1

if not rs.eof then
	BadList=rs.getRows()
	intBadCount=ubound(BadList,2)
end if

Function CheckBadW(strTest)
	Dim k,tmpstr
	tmpstr=strTest
	For k=0 to intBadCount
		if instr(1,tmpstr,BadList(0,k),1)>0 then
			tmpstr=replace(tmpstr,BadList(0,k),"****",1,-1,1)
		end if
	Next
	CheckBadW=tmpstr
End Function
	
query="SELECT pcRS_FieldList FROM pcReviewSpecials WHERE pcRS_IDProduct=" & pcv_IDProduct
set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

Dim Fi(100)
Dim FType(100)

if not rs.eof then
	pcv_FieldList=split(rs("pcRS_FieldList"),",")

	FCount=0
	For i=0 to ubound(pcv_FieldList)
		if pcv_FieldList(i)<>"" then
			Fi(FCount)=pcv_FieldList(i)
		
			query="SELECT pcRF_Type FROM pcRevFields WHERE pcRF_IDField=" & Fi(FCount)
			set rs=connTemp.execute(query)
			
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
	
			FType(FCount)=rs("pcRF_Type")
			FCount=FCount+1
		end if
	Next
else

	query="SELECT pcRF_IDField,pcRF_Type FROM pcRevFields WHERE pcRF_Active=1 order by pcRF_Order asc"
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if not rs.eof then
		pcArray=rs.getRows()
		intCount=ubound(pcArray,2)
	
		FCount=0
	
		For i=0 to intCount
			Fi(FCount)=pcArray(0,i)
			FType(FCount)=pcArray(1,i)
			FCount=FCount+1
		Next
		
	end if

end if

IF (FCount>0) and (request("action")="add") THEN
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Run Code
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Dim SPath
	SPath=Request.ServerVariables("PATH_INFO")
	SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
	If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
		strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
	Else
		strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
	End if
	
	IF scSecurity=1 THEN
		if scReview=1 then
			pcv_Test=0
			if Session("store_ReviewReg")<>"1" then
				Session("store_ReviewReg")=""
				Session("store_ReviewRegpostnum")=""
				Session("store_ReviewRegnum")=""
				pcv_Test=1
			end if
			if pcv_Test=0 then
				if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "prv_postreview.asp"))<>1 then
					Session("store_ReviewReg")=""
					Session("store_ReviewRegpostnum")=""
					Session("store_ReviewRegnum")=""
					pcv_Test=1
				end if
			end if
			if pcv_Test=0 then %>
                <!-- Include file for CAPTCHA configuration -->
                <!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
                 
                <!-- Include file for CAPTCHA form processing -->
                <!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
				<%	
                If not blnCAPTCHAcodeCorrect then				
					Session("store_ReviewReg")=""
					pcv_Test=1
					response.redirect "prv_postreview.asp?IDPRoduct="&pcv_IDProduct&"&msg="& Server.Urlencode(dictLanguage.Item(Session("language")&"_security_3"))
				end if
			end if
			
			if pcv_Test=1 then
				If scAlarmMsg=1 then
					if session("AttackCount")="" then
						session("AttackCount")=0
					end if
					session("AttackCount")=session("AttackCount")+1
					if session("AttackCount")>=scAttackCount then%>
					<!--#include file="../includes/sendAlarmEmail.asp" -->
					<%end if	
				End if
				response.write dictLanguage.Item(Session("language")&"_security_2")
				response.end
			end if
		end if
	END IF
	
	Session("store_ReviewReg")=""
	Session("store_ReviewRegpostnum")=""
	Session("store_ReviewRegnum")=""
	
	query="SELECT description FROM products WHERE idproduct=" & pcv_IDProduct
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	pcv_PrdName=rs("description")

	if pcv_NeedCheck="0" then
		RevActive="1"
	else
		RevActive="0"
	end if

	dim dtTodaysDate
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=Day(dtTodaysDate)&"/"&Month(dtTodaysDate)&"/"&Year(dtTodaysDate)
	else
		dtTodaysDate=Month(dtTodaysDate)&"/"&Day(dtTodaysDate)&"/"&Year(dtTodaysDate)
	end if
	
	dtTime=Time()
	
	dtTodaysDate=dtTodaysDate & " " & dtTime

	if scDB="Access" then
		dtTodaysDate="#" & dtTodaysDate & "#"
	else
		dtTodaysDate="'" & dtTodaysDate & "'"
	end if
	
	tmpIdOrder=0
	if request.form("xrv")<>"" then
		tmpxrv=getUserInput(request.form("xrv"),0)
		if IsNumeric(tmpxrv) then
			query="SELECT IdOrder FROM ProductsOrdered WHERE idProductOrdered=" & tmpxrv & ";"
			set rs=connTemp.execute(query)
			if not rs.eof then
				tmpIdOrder=rs("IdOrder")
			end if
			set rs=nothing
		end if
	end if


    ' PRV41 begin
	query="INSERT INTO pcReviews (pcRev_IDProduct,pcRev_Active,pcRev_IP,pcRev_Date,pcRev_MainRate,pcRev_MainDRate, pcRev_IDCustomer,pcRev_IdOrder) VALUES (" & pcv_IDProduct & "," & RevActive & ",'" & pcv_IPAddress & "'," & dtTodaysDate & "," & pcv_Feel & "," & pcv_Rate & "," & TestRegCust(fnZeroIfNull(pIdCustomer)) & "," & tmpIdOrder & ")"
	' PRV41 end
	set rs=connTemp.execute(query)
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

	query="SELECT pcRev_IDReview FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " order by pcRev_IDReview desc"
	set rs=connTemp.execute(query)
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

	pcv_IDReview=rs("pcRev_IDReview")

	For m=0 to FCount-1
		Rev_IDField=Fi(m)
		if FType(m)="3" then
			Rev_Feel=GetUserInput(request("Field" & Fi(m)),0)
			if Rev_Feel="" then
				Rev_Feel="0"
			end if
			Rev_Rate="0"
			Rev_Com=""
		end if
		if FType(m)="4" then
			Rev_Feel="0"
			Rev_Rate=GetUserInput(request("Field" & Fi(m)),0)
			if Rev_Rate="" then
				Rev_Rate="0"
			end if
			Rev_Com=""
		end if
		if FType(m)<"3" then
			Rev_Feel="0"
			Rev_Rate="0"
			Rev_Com=CheckBadW(GetUserInput(ClearHTMLTags2(request("Field" & Fi(m)),0),0))
		end if
	
		query="INSERT INTO pcReviewsData (pcRD_IDReview,pcRD_IDField,pcRD_Feel,pcRD_Rate,pcRD_Comment) VALUES (" & pcv_IDReview & "," & Rev_IDField & "," & Rev_Feel & "," & Rev_Rate & ",'" & Rev_Com & "')"
		set rs=connTemp.execute(query)

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
	Next
	'PRV41 begin
	' If review auto-approval is on, then we should update the product's average rating now
	query = "SELECT TOP 1 pcRS_NeedCheck FROM pcRevSettings"
	Set rs = connTemp.execute(query)
	If rs.eof = False Then
	   If rs("pcRS_NeedCheck")=0 then
	      query = "UPDATE Products SET pcProd_AvgRating = " & GetOverallProductRating(pcv_IDProduct) & " WHERE idProduct=" & pcv_IDProduct
	      connTemp.execute query
	   End If
	End if

	If TestRegCust(fnZeroIfNull(pIdCustomer))<>0 then
		' If Rewards For Reviews is 'on', and reviews are set to auto-approve, then 
		' we need to add reward points here (both to Customers.iRewardPointsAccrued 
		' and also to pcReviewPoints)
		query = "SELECT TOP 1 pcRS_NeedCheck, pcRS_RewardForReview, pcRS_RewardForReviewURL, pcRS_RewardForReviewFirstPts, pcRS_RewardForReviewAdditionalPts,pcRS_RewardForReviewMaxPts FROM pcRevSettings WHERE pcRS_RewardForReview=1 AND pcRS_NeedCheck<>1 AND (pcRS_RewardForReviewFirstPts>0 OR pcRS_RewardForReviewAdditionalPts>0)"

		Set rs = connTemp.execute(query)
		pcv_MaxPoints=0
		if not rs.EOF then
		
			pcv_MaxPoints=rs("pcRS_RewardForReviewMaxPts")
			if IsNull(pcv_MaxPoints) OR pcv_MaxPoints="" then
				pcv_MaxPoints=0
			end if

		   ' OK, so we know that we're supposed to award points automatically, let's find out if 
		   ' this is the user's first review or not, and award appropriate points
		   Dim rs2, ptsToAward, pIntExecuteRPtasks, prv_TotalPoints
		   ptsToAward=0
		   pIntExecuteRPtasks=0
		   
		   query = "SELECT count(*) as ct FROM pcReviewPoints WHERE pcRP_IDCustomer=" & TestRegCust(fnZeroIfNull(pIdCustomer))
		   Set rs2 = connTemp.execute(query)
		   If CLng(rs2("ct"))=0 Then
			  ptsToAward = fnZeroIfNull(rs("pcRS_RewardForReviewFirstPts"))
		   Else
			  ptsToAward = fnZeroIfNull(rs("pcRS_RewardForReviewAdditionalPts"))
		   End If
		   Set rs2 = nothing
		   
		   If ptsToAward>0 Then
		   
		   		pIntExecuteRPtasks=1 ' There are points to award. Set flag to execute tasks.
				
		   		'// Maximum RP - START
				'// Check against maximum Reward Points that can be awarded
				prv_TotalPoints=0
				if pIdCustomer <>"" AND pIdCustomer <>"0" then
					query="SELECT Sum(pcRP_PointsAwarded) AS TotalPoints FROM pcReviewPoints WHERE pcRP_IDCustomer=" & pIdCustomer & ";"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						prv_TotalPoints=rsQ("TotalPoints")
						if IsNull(prv_TotalPoints) OR prv_TotalPoints="" then
							prv_TotalPoints=0
						end if
					end if
					set rsQ=nothing
				
					if CLng(pcv_MaxPoints)>0 then
						if CLng(prv_TotalPoints)+CLng(ptsToAward)>CLng(pcv_MaxPoints) then
							ptsToAward=CLng(pcv_MaxPoints)-CLng(prv_TotalPoints)
							if Clng(ptsToAward)<=0 then
								pIntExecuteRPtasks=0 ' Customer has already been awarded the max. Remove flag.
							end if
						end if
					end if
				end if
		   		'// Maximum RP - END
			
			end if
			
			if pIntExecuteRPtasks<>0 then ' Execute tasks only if there are points to award and the max has not been reached yet
			
			  query = "INSERT INTO pcReviewPoints (pcRP_IDReview, pcRP_IDCustomer, pcRP_PointsAwarded, pcRP_DateAwarded) VALUES (" &  pcv_IDReview & "," & TestRegCust(fnZeroIfNull(pIdCustomer)) & "," & ptsToAward & "," & formatDateForDB(now) & ")"
			  connTemp.execute query

			  query = "UPDATE customers SET iRewardPointsAccrued=iRewardPointsAccrued+" & ptsToAward & " WHERE idCustomer=" & TestRegCust(fnZeroIfNull(pIdCustomer))
			  connTemp.execute query
			  
			  '// Thank You message to customer - START
			  
			  	'// Load customer information
				Dim pcStrCustName, pcStrCustEmail, pcIntSendMessage, pcStrRewardForReviewURL
				pcStrRewardForReviewURL=rs("pcRS_RewardForReviewURL")
				pcIntSendMessage=0
				query = "SELECT name, lastName, email FROM customers WHERE idCustomer = " & TestRegCust(fnZeroIfNull(pIdCustomer))
				set rs2 = Server.CreateObject("ADODB.Recordset")
				set rs2 = conntemp.execute(query)
				if rs2.eof then
					pcIntSendMessage=0
				else
					pcIntSendMessage=1
					pcStrCustName = rs2("name") & " " & rs2("lastName")
					pcStrCustEmail = rs2("email")
				end if
					
				if pcIntSendMessage=1 then
				
					'// Load product information
					Dim pcStrProductName
					query = "SELECT description FROM products WHERE idproduct = " & pcv_IDProduct
					set rs2 = conntemp.execute(query)
					pcStrProductName = rs2("description")
					
					'// Build message
					Dim strNewMessage
					strNewMessage = dictLanguage.Item(Session("language")&"_prv_37")
					strNewMessage = Replace(strNewMessage,"<CUSTOMER_NAME>", pcStrCustName,1,-1,vbTextCompare)
					strNewMessage = Replace(strNewMessage,"<PRODUCT_NAME>", pcStrProductName,1,-1,vbTextCompare)
					strNewMessage = Replace(strNewMessage,"<REWARD_POINTS_LABEL>", RewardsLabel,1,-1,vbTextCompare)
					strNewMessage = Replace(strNewMessage,"<NUM_POINTS>",ptsToAward,1,-1,vbTextCompare)
					strNewMessage = Replace(strNewMessage,"<REWARD_REVIEWS_URL>", pcStrRewardForReviewURL,1,-1,vbTextCompare)
					strNewMessage = Replace(strNewMessage,"<STORE_NAME>",scCompanyName,1,-1,vbTextCompare)
					
					'// Build subject
					Dim strNewSubject 
					strNewSubject = dictLanguage.Item(Session("language")&"_prv_45")
					strNewSubject = Replace(strNewSubject,"<REWARD_POINTS_LABEL>", RewardsLabel,1,-1,vbTextCompare)
					strNewSubject = Replace(strNewSubject,"<STORE_NAME>",scCompanyName,1,-1,vbTextCompare)

					'// Send message
					 call sendmail(scCompanyName, scEmail, pcStrCustEmail, strNewSubject, strNewMessage)
					 
				end if
				 
			  '// Thank You message to customer - END
		   
		   End if

		End If
		Set rs = Nothing
	End if

	' PRV41 end



	'[ClearHTMLTags2]
	
	'Coded by Jóhann Haukur Gunnarsson
	'joi@innn.is
	
	'	Purpose: This function clears all HTML tags from a string using Regular Expressions.
	'	 Inputs: strHTML2;	A string to be cleared of HTML TAGS
	'		 intWorkFlow2;	An integer that if equals to 0 runs only the regEx2p filter
	'							  .. 1 runs only the HTML source render filter
	'							  .. 2 runs both the regEx2p and the HTML source render
	'							  .. >2 defaults to 0
	'	Returns: A string that has been filtered by the function
	
	
	function ClearHTMLTags2(strHTML2, intWorkFlow2)
		
		'Variables used in the function
		
		dim regEx2, strTagLess2
		
		'---------------------------------------
		strTagLess2 = strHTML2
		'Move the string into a private variable
		'within the function
		'---------------------------------------
		
		'---------------------------------------
		'NetSource Commerce codes
		IF strTagLess2<>"" THEN
			strTagLess2=replace(strTagLess2,"<br>"," ")
			strTagLess2=replace(strTagLess2,"<BR>"," ")
			strTagLess2=replace(strTagLess2,"<p>"," ")
			strTagLess2=replace(strTagLess2,"<P>"," ")
			strTagLess2=replace(strTagLess2,"</p>"," ")
			strTagLess2=replace(strTagLess2,"</P>"," ")
			strTagLess2=replace(strTagLess2,vbcrlf," ")
			strTagLess2=trim(strTagLess2)
			do while instr(strTagLess2,"  ")>0
				strTagLess2=replace(strTagLess2,"  "," ")
			loop
		END IF
		'Modify the string to a friendly ONLY 1 LINE string
		'---------------------------------------
		
		IF strTagLess2<>"" THEN
			'regEx2 initialization
			'---------------------------------------
			set regEx2 = New regExp 
			'Creates a regEx2p object		
			regEx2.IgnoreCase = True
			'Don't give frat about case sensitivity
			regEx2.Global = True
			'Global applicability
			'---------------------------------------
			'Phase I
			'	"bye bye html tags"
	
			if intWorkFlow2 <> 1 then
			
				'---------------------------------------
				regEx2.Pattern = "<[^>]*>"
				'this pattern mathces any html tag
				strTagLess2 = regEx2.Replace(strTagLess2, "")
				'all html tags are stripped
				'---------------------------------------
							
			end if
			'Phase II
			'	"bye bye rouge leftovers"
			'	"or, I want to render the source"
			'	"as html."
	
			'---------------------------------------
			'We *might* still have rouge < and > 
			'let's be positive that those that remain
			'are changed into html characters
			'---------------------------------------	
			if intWorkFlow2 > 0 and intWorkFlow2 < 3 then
				regEx2.Pattern = "[<]"
				'matches a single <
				strTagLess2 = regEx2.Replace(strTagLess2, "&lt;")
	
				regEx2.Pattern = "[>]"
				'matches a single >
				strTagLess2 = regEx2.Replace(strTagLess2, "&gt;")
				'---------------------------------------
			end if
			
			'Clean up
			'---------------------------------------
			set regEx2 = nothing
			'Destroys the regEx2p object
			'---------------------------------------	
		END IF 'vefiry strTagLess2 (null strings)
			
		'---------------------------------------
		ClearHTMLTags2 = strTagLess2
		'The results are passed back
		'---------------------------------------
			
	end function
	%>
	<%IF pcv_NeedCheck="0" THEN
	ELSE
	'Send a notification e-mail to administrator
	SPathInfo=scStoreURL
	if Right(SPathInfo,1)="/" then
	else
	SPathInfo=SPathInfo&"/"
	end if
	SPathInfo=SPathInfo&scPcFolder&"/"&scAdminFolderName&"/prv_ManageReviews.asp?IDProduct=" & pcv_IDProduct & "&nav=1"
	
	strNewOrderSubject=dictLanguage.Item(Session("language")&"_prv_17") & pcv_IDProduct
	storeAdminEmail="" & dictLanguage.Item(Session("language")&"_prv_18") & vbcrlf & vbcrlf
	storeAdminEmail=storeAdminEmail&dictLanguage.Item(Session("language")&"_prv_19") & pcv_IDProduct & vbcrlf
	storeAdminEmail=storeAdminEmail&dictLanguage.Item(Session("language")&"_prv_20") & vbcrlf & vbcrlf
	storeAdminEmail=storeAdminEmail&SPathInfo & vbcrlf & vbcrlf
	storeAdminEmail=storeAdminEmail&scCompanyName
	
	call sendmail (scCompanyName, scEmail, scFrmEmail, strNewOrderSubject, replace(storeAdminEmail,"&quot;", chr(34)))

	END IF%>
	<html>
		<head>
			<title><%=dictLanguage.Item(Session("language")&"_prv_9")%></title>
			<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
			<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
		</head>
	<body>
	<div id="pcMain">
		<table class="pcMainTable">
			<tr>
				<td>
					<h1><%=dictLanguage.Item(Session("language")&"_prv_9")%></h1>
					<h2><%=dictLanguage.Item(Session("language")&"_prv_10")%> <%=pcv_PrdName%></h2>
					<p>
						<b><%if pcv_NeedCheck="0" then%>
							<%=dictLanguage.Item(Session("language")&"_prv_14")%>
							<%else%>
							<%=dictLanguage.Item(Session("language")&"_prv_16")%>
							<%end if%>
						</b>
					</p>
					<p>&nbsp;</p>
					<%
					' prv41 begin
					If IsNumeric(request.form("xrv")) Then
						%>
						<script language="JavaScript">
						<!--
						if (window.opener.document.getElementById('xrv<% = clng(request.form("xrv")) %>')){
						   window.opener.document.getElementById('xrv<% = clng(request.form("xrv")) %>').innerHTML='<% = replace(dictLanguage.Item(Session("language")&"_prv_26"),"'","\'") %>';
						}
						//-->
						</script>
						<%
					End if
					' prv41 end
					%>
					<form class="pcForms">
						<p align="center"><input type="button" value="Close window" onclick="window.close();" class="submit2"></p>
					</form>
				</td>
			</tr>
		</table>
	</div>
	</body>
	</html>

	<% PostC=getUserInput(Request.Cookies("Prd" & pcv_IDProduct),0)
	if PostC<>"" then
	else
		PostC=0
	end if
	Response.Cookies("Prd" & pcv_IDProduct)=PostC+1
	Response.Cookies("Prd" & pcv_IDProduct).Expires=Date() + 365
	MyCookiePath=Request.ServerVariables("PATH_INFO")
	do while not (right(MyCookiePath,1)="/")
		MyCookiePath=mid(MyCookiePath,1,len(MyCookiePath)-1)
	loop
	Response.Cookies("Prd" & pcv_IDProduct).Path=MyCookiePath
END IF

set rs=nothing
call closedb()
%>