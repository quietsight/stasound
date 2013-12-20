<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pcv_RevName=request("nav")
if pcv_RevName="" then
pcv_RevName="2"
end if
%>
<% section="reviews" %>
<%PmAdmin=2%>
<% response.Buffer=true %><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<% 'PRV41 start %>
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<% 'PRV41 end %>

<%
Dim connTemp,rs, query
Dim BadList,intBadCount
call opendb()
%>
<!--#include file="../pc/prv_getsettings.asp"-->
<!--#include file="../pc/prv_recalc.asp"-->
<%
pcv_IDProduct=request("IDProduct")
pcv_IDReview=request("IDReview")

if request("action")="delete" then
	query="DELETE FROM pcReviews WHERE pcRev_IDReview=" & pcv_IDReview
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	query="DELETE FROM pcReviewsData WHERE pcRD_IDReview=" & pcv_IDReview
	set rs=connTemp.execute(query)
	set rs=Nothing
	
	connTemp.execute "UPDATE Products SET pcProd_AvgRating = " & GetOverallProductRating(pcv_IDProduct) & " WHERE idProduct=" & pcv_IDProduct

	call closedb()
	response.redirect "prv_ManageReviews.asp?IDProduct=" & pcv_IDProduct & "&nav=" & pcv_RevName & "s=1&msg=" & Server.URLEncode("Product review deleted successfully!")
end if

pcv_Feel=request("feel")
if pcv_Feel="" then
	pcv_Feel="0"
end if
pcv_Rate=request("rate")
if pcv_Rate="" then
	pcv_Rate="0"
end if

query="SELECT pcRBW_word FROM pcRevBadWords"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

intBadCount=-1

if not rs.eof then
	BadList=rs.getRows()
	intBadCount=ubound(BadList,2)
end if
set rs=nothing

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
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

Dim Fi(100)
Dim FType(100)

if not rs.eof then
	pcv_FieldList=split(rs("pcRS_FieldList"),",")
	set rs=nothing
	FCount=0
	For i=0 to ubound(pcv_FieldList)
		if pcv_FieldList(i)<>"" then
			Fi(FCount)=pcv_FieldList(i)
		
			query="SELECT pcRF_Type FROM pcRevFields WHERE pcRF_IDField=" & Fi(FCount)
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
	
			FType(FCount)=rs("pcRF_Type")
			FCount=FCount+1
			set rs=nothing
		end if
	Next
else
	query="SELECT pcRF_IDField,pcRF_Type FROM pcRevFields WHERE pcRF_Active=1 order by pcRF_Order asc"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

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
	set rs=nothing
end if
call closedb()

IF (FCount>0) and (request("action")="add") THEN

	RevActive=request("active")
	
	if RevActive="" then
		RevActive="0"
	end if
	call opendb()
	query="UPDATE pcReviews SET pcRev_Active=" & RevActive &",pcRev_MainRate=" & pcv_Feel & ",pcRev_MainDRate=" & pcv_Rate & " WHERE pcRev_IDReview=" & 	pcv_IDReview
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing
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
			tmp_info=request("Field" & Fi(m))
			if tmp_info<>"" then
				tmp_info=replace(tmp_info,"'","''")
			end if
			Rev_Com=CheckBadW(tmp_info)
		end if
		
		query="SELECT pcRD_IDField FROM pcReviewsData WHERE pcRD_IDReview=" & pcv_IDReview & " and pcRD_IDField=" & Rev_IDField
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		IF not rs.eof THEN
		
		query="UPDATE pcReviewsData SET pcRD_Feel=" & Rev_Feel & ",pcRD_Rate=" & Rev_Rate & ",pcRD_Comment='" & Rev_Com & "' WHERE pcRD_IDReview=" & pcv_IDReview & " and pcRD_IDField=" & Rev_IDField
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=Nothing

		
		ELSE
		
		query="INSERT INTO pcReviewsData (pcRD_IDReview,pcRD_IDField,pcRD_Feel,pcRD_Rate,pcRD_Comment) VALUES (" & pcv_IDReview & "," & Rev_IDField & "," & Rev_Feel & "," & Rev_Rate & ",'" & Rev_Com & "')"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=Nothing
		
		END IF
	Next
	'PRV41 begin
	If RevActive<>"0" then
		connTemp.execute "UPDATE Products SET pcProd_AvgRating = " & GetOverallProductRating(pcv_IDProduct) & " WHERE idProduct=" & pcv_IDProduct
		query = "SELECT pcRev_IDCustomer FROM pcReviews where pcRev_IDReview=" & pcv_IDReview & " and pcRev_IDCustomer<>0 and pcRev_IDCustomer is not null"
		Set rs0 = connTemp.execute(query)
		If rs0.eof = False Then
			pIDCustomer = rs0("pcRev_IDCustomer")

			' If Rewards For Reviews is 'on', and reviews are NOT set to auto-approve, then 
			' we need to add reward points here (both to Customers.iRewardPointsAccrued 
			' and also to pcReviewPoints)
			query = "select top 1 pcRS_NeedCheck, pcRS_RewardForReview, pcRS_RewardForReviewURL, pcRS_RewardForReviewFirstPts, pcRS_RewardForReviewAdditionalPts,pcRS_RewardForReviewMaxPts from pcRevSettings where pcRS_RewardForReview=1 and pcRS_NeedCheck=1 and (pcRS_RewardForReviewFirstPts>0 or pcRS_RewardForReviewAdditionalPts>0)"
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

			   query = "SELECT count(*) as ct FROM pcReviewPoints WHERE pcRP_IDCustomer=" & pIDCustomer
			   Set rs2 = connTemp.execute(query)
			   If CLng(rs2("ct"))=0 Then
				  ptsToAward = fnZeroIfNull(rs("pcRS_RewardForReviewFirstPts"))
			   Else
				  ptsToAward = fnZeroIfNull(rs("pcRS_RewardForReviewAdditionalPts"))
			   End If
			   Set rs2 = nothing
			   
			   If ptsToAward>0 Then
			   
		   		pIntExecuteRPtasks=1 ' There are points to award. Set flag to execute tasks.
				
				'Has Reward points for this Review ID# or NOT?
				query = "SELECT pcRP_IDReview, pcRP_IDCustomer, pcRP_PointsAwarded FROM pcReviewPoints WHERE pcRP_IDReview=" &  pcv_IDReview & " AND pcRP_IDCustomer=" & pIDCustomer & ";"
				set rs2=connTemp.execute(query)
				if not rs2.eof then
					pIntExecuteRPtasks=0
				end if
				set rs2=nothing				
				
		   		'// Maximum RP - START
				'// Check against maximum Reward Points that can be awarded
			   	prv_TotalPoints=0
				if pIDCustomer <>"" AND pIDCustomer <>"0" then
				   	query="SELECT Sum(pcRP_PointsAwarded) AS TotalPoints FROM pcReviewPoints WHERE pcRP_IDCustomer=" & pIDCustomer & ";"
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

				  query = "INSERT INTO pcReviewPoints (pcRP_IDReview, pcRP_IDCustomer, pcRP_PointsAwarded, pcRP_DateAwarded) VALUES (" &  pcv_IDReview & "," & pIDCustomer & "," & ptsToAward & "," & formatDateForDB(now) & ")"
				  connTemp.execute query

				  query = "UPDATE customers SET iRewardPointsAccrued=iRewardPointsAccrued+" & ptsToAward & " WHERE idCustomer=" & pIDCustomer
				  connTemp.execute query

				  '// Thank You message to customer - START
				  
					'// Load customer information
					Dim pcStrCustName, pcStrCustEmail, pcIntSendMessage, pcStrRewardForReviewURL
					pcStrRewardForReviewURL=rs("pcRS_RewardForReviewURL")
					pcIntSendMessage=0
					query = "SELECT name, lastName, email FROM customers WHERE idCustomer = " & pIDCustomer
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
		End If
		Set rs0 = nothing
    End if
	' PRV41 end
	
	call closedb()
	response.redirect "prv_EditReview.asp?IDProduct=" & pcv_IDProduct & "&IDReview=" & pcv_IDReview & "&nav=" & pcv_RevName & "&s=1&msg=" & Server.URLEncode("Review updated successfully!")
END IF
%>