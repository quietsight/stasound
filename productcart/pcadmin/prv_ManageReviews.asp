<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pcv_IDCust = fnzeroifnull(request("IDCustomer"))
pcv_RevName=request("nav")
if pcv_RevName="" then
	pcv_RevName="2"
end If
If pcv_IDCust>0 Then
   pageTitle="Manage All Reviews by a Customer"
else
   if pcv_RevName="1" then
	 pageTitle="Manage Pending Reviews"
   else
	 pageTitle="Manage Product Reviews"
   end If
End if
%>
<% section="reviews" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
Dim rs, connTemp, strSQL, pid
pcv_IDProduct=getUserInput(request("IDProduct"),10)
if not validNum(pcv_IDProduct) then pcv_IDProduct=0
call opendb()
%>
<!--#include file="../pc/prv_getsettings.asp"-->
<!--#include file="prv_incfunctions.asp"-->
<!--#include file="../pc/prv_recalc.asp"-->
<%
Call CreateList()

IF request("action")="resetrating" and pcv_IDProduct>0 then
	query="SELECT pcRev_IDReview FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcArray=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArray,2)
		For k=0 to intCount	
			query="DELETE FROM pcReviewsData WHERE pcRD_IDReview=" & pcArray(0,k)
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			set rs=nothing
		Next
	end if
	
	query="DELETE FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=Nothing
	
	connTemp.execute "UPDATE Products SET pcProd_AvgRating = 0 WHERE idProduct=" & pcv_IDProduct
	
	call closedb()
	response.redirect "prv_ManageRevPrds.asp?nav=" & pcv_RevName
END IF

query = "SELECT top 1 pcRS_RewardForReview FROM pcRevSettings"
Set rs = connTemp.execute(query)
If rs.eof = False Then
   If IsNull(rs("pcRS_RewardForReview")) Then
      pcv_RewardForReview = 0
   Else
      pcv_RewardForReview = CLng(rs("pcRS_RewardForReview"))
   End if
End If
Set rs = nothing

IF request("action")="update" then

	Count=request("count")
	if (Count>"0") and (IsNumeric(Count)) then
		For k=1 to Count
			if request("C" & k)="1" then
				pcv_ID=request("ID" & k)
				if request("submit1")<>"" then
					if pcv_RevName="1" then
						query1=" pcRev_Active=1 "
					else
						query1=" pcRev_Active=0 "
					end if
					query="UPDATE pcReviews SET " & query1 & " WHERE pcRev_IDReview=" & pcv_ID
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
                    ' PRV41 begin
                    Dim pcv_IDCustomer
					pcv_IDCustomer = 0
					query = "SELECT pcRev_IDCustomer FROM pcReviews WHERE pcRev_IDReview=" & pcv_ID
					Set rs = connTemp.execute(query)
					If rs.eof = False Then
					   pcv_IDCustomer = fnZeroIfNull(rs("pcRev_IDCustomer"))
					End if
					set rs=Nothing
					' PRV41 end
					
					if pcv_RevName="1" then
						msg="Selected reviews were approved!"
						msgType=1
						' PRV41 begin
						' If the review was just approved, and Rewards For Reviews is 'on', and
						' reviews are NOT set to auto-approve, then we need to add reward points
						' here (both to Customers.iRewardPointsAccrued and also to pcReviewPoints)
						query = "select top 1 pcRS_NeedCheck, pcRS_RewardForReview, pcRS_RewardForReviewURL, pcRS_RewardForReviewFirstPts, pcRS_RewardForReviewAdditionalPts,pcRS_RewardForReviewMaxPts from pcRevSettings where pcRS_RewardForReview=1 and pcRS_NeedCheck=1 and (pcRS_RewardForReviewFirstPts>0 or pcRS_RewardForReviewAdditionalPts>0)"						
						Set rs = connTemp.execute(query)
						pcv_MaxPoints=0
						if not rs.EOF then
						
							pcv_MaxPoints=rs("pcRS_RewardForReviewMaxPts")
							if IsNull(pcv_MaxPoints) OR pcv_MaxPoints="" then
								pcv_MaxPoints=0
							end if

						   ' OK, so we know that we're supposed to award points on approval, let's find out if 
						   ' this is the user's first review or not, and award appropriate points
						   Dim rs2, ptsToAward, pIntExecuteRPtasks, prv_TotalPoints
						   ptsToAward=0
						   pIntExecuteRPtasks=0
			
						   query = "SELECT count(*) as ct FROM pcReviewPoints WHERE pcRP_IDCustomer=" & pcv_IDCustomer
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
									if pcv_IDCustomer <>"" AND pcv_IDCustomer <>"0" then
									   	query="SELECT Sum(pcRP_PointsAwarded) AS TotalPoints FROM pcReviewPoints WHERE pcRP_IDCustomer=" & pcv_IDCustomer & ";"
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

						      query = "INSERT INTO pcReviewPoints (pcRP_IDReview, pcRP_IDCustomer, pcRP_PointsAwarded, pcRP_DateAwarded) VALUES (" & pcv_ID & "," & pcv_IDCustomer & "," & ptsToAward & "," & formatDateForDB(now) & ")"
							  connTemp.execute query

							  query = "UPDATE customers SET iRewardPointsAccrued=iRewardPointsAccrued+" & ptsToAward & " WHERE idCustomer=" & pcv_IDCustomer
							  connTemp.execute query
							  
							  '// Thank You message to customer - START
							  
								'// Load customer information
								Dim pcStrCustName, pcStrCustEmail, pcIntSendMessage, pcStrRewardForReviewURL
								pcStrRewardForReviewURL=rs("pcRS_RewardForReviewURL")
								pcIntSendMessage=0
								query = "SELECT name, lastName, email FROM customers WHERE idCustomer = " & pcv_IDCustomer
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

						' PRV41 end
					else
						msg="Selected reviews were hidden!"
						msgType=1
					end if
				end if
				if request("submit2")<>"" then
					query="DELETE FROM pcReviews WHERE pcRev_IDReview=" & pcv_ID
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					query="DELETE FROM pcReviewsData WHERE pcRD_IDReview=" & pcv_ID
					set rs=connTemp.execute(query)
					msg="Selected reviews were deleted successfully!"
					msgType=1
					set rs=nothing
				end if
			end if
		Next
	end If
	
	'PRV41 begin
	connTemp.execute "UPDATE Products SET pcProd_AvgRating = " & GetOverallProductRating(pcv_IDProduct) & " WHERE idProduct=" & pcv_IDProduct
	' PRV41 end
	
END IF

if pcv_IDProduct>0 then
	query="SELECT description FROM Products WHERE idproduct=" & pcv_IDProduct
	set rs=connTemp.execute(query)
	pcv_PrdName=rs("description")
	set rs=nothing
end if

If pcv_IDCust>0 Then
  query1 = "pcRev_IDCustomer=" & pcv_IDCust
Else
  if pcv_RevName="1" then
	 query1=" pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=0 "
  else
	 query1=" pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 "
  end If   
End if
	
query="SELECT pcRev_IDReview,pcRev_Date,pcRev_MainRate,pcRev_MainDRate, pcRev_Active, pcRev_IDProduct FROM pcReviews WHERE " & query1 & " order by pcRev_Date desc"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
	
if rs.eof then
	DataEmpty=1
else
	DataEmpty=0
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
end if
set rs=nothing
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="POST" action="prv_ManageReviews.asp?action=update<%
' PRV41 begin
If pcv_IDCust>0 Then
   response.write "&idcustomer=" & pcv_IDCust
End if
' PRV41 end
%>" name="checkboxform" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="6">
			<%
			' PRV41 begin
			If pcv_IDCust>0 Then
			   query = "SELECT Name, LastName FROM customers WHERE idCustomer=" & pcv_IDCust
			   Set rs = connTemp.execute(query)
			   If rs.eof = False Then
			      response.write "<h2>All reviews by: <a href='modCusta.asp?idcustomer=" & pcv_IDCust & "' target='_blank'>" & rs("Name") & " " & rs("LastName") & "</a></h2>"
			   End If
			   Set rs = nothing
			else
			%>
            	<h2><%if pcv_RevName="1" then%>Pending reviews<%else%>Products reviews<%end if%> for: <b><%=pcv_PrdName%></b> - <a href="prv_ManageRevPrds.asp?nav=<%=pcv_RevName%>">View all <% if pcv_RevName="1" then%>pending reviews<%else%>live reviews<%end if%></a></h2>
			<%
			End If
			' PRV41 end
			%>
            </a>
		</tr>
		<tr>
			<th>ID</th>
			<%
			' PRV41 begin
			If pcv_IDCust>0 Then
			   %><th>Approved?</th><%
			End if
			' PRV41 end
			%>
			<th width="50%" nowrap>Review Title</th>
			<th nowrap>Customer Name</th>
			<th nowrap>Posted Date</th>
			<th nowrap>Rating</th>
			<%
			'PRV41 begin
			   If pcv_IDCust=0 Then
			 ' PRV41 end %>
			<th nowrap>Select</th>
			<%
			End If
			' PRV41 end
			%>
		</tr>
		<tr>
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<% If DataEmpty=1 Then %>
                      
			<tr> 
				<td colspan="6">No <%
				' PRV41 begin
				If pcv_IDCust>0 Then
				   response.write "reviews found by this customer."
				else
				   if pcv_RevName="1" then%>pending<%else%>live<%end if%> product reviews found for this product.<%
                   if pcv_RevName="1" and validNum(pcv_IDProduct) and pcv_IDProduct>0 then%> <a href="prv_ManageReviews.asp?IDProduct=<%=pcv_IDProduct%>&nav=2">View live reviews</a>.<%end if%>
				 <%
				end If
				' PRV41 end
				%></td>
			</tr>
                      
		<% Else 
			Count=0

			For k=0 to intCount
				Count=Count+1

				pcv_ID=pcArray(0,k)
				Rev_Date=pcArray(1,k)
				Rev_MainRate=pcArray(2,k)
				Rev_MainDRate=pcArray(3,k)
				' PRV41 begin
				Rev_Approved=pcArray(4,k)
				Rev_IDProduct= pcArray(5,k)
				' PRV41 end
				
				
				query="SELECT pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & pcv_ID & " and pcRD_IDField=1"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
	
	            ' PRV41 begin
				If rs.eof = False then
				   Rev_CustName=rs("pcRD_Comment")
				Else
				   Rev_CustName=""
				End If
				' PRV41 end
	
				query="SELECT pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & pcv_ID & " and pcRD_IDField=2"
				set rs=connTemp.execute(query)
	
	            ' PRV41 begin
				If rs.eof = False then
				   Rev_Name=rs("pcRD_Comment")
				Else
				   Rev_CustName=""
				End If
				' PRV41 end
				set rs=nothing
				
				%>
                      
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td><%=pcv_ID%></td>
					<%
					' PRV41 begin
					If pcv_IDCust>0 Then
					   %>
					   <td align="center"><%
					   If Rev_Approved<>0 Then
					      response.write "Yes"
					   Else
					      response.write "No"
					   End if
					   %>
					   </td>
					   <%
					End if
					' PRV41 end
					%>
					<td><a href="prv_EditReview.asp?IDReview=<%=pcv_ID%>&IDProduct=<%=Rev_IDProduct%>&nav=<%=pcv_RevName%><%
					' PRV41 begin
					If pcv_idcust>0 Then
					   response.write "&idcustomer=" & pcv_idcust
					End if
					' PRV41 end
					%>"><%=Rev_Name%></a></td>
					<td><%=Rev_CustName%></td>
					<td>
					<% If scDateFrmt="DD/MM/YY" then 
						Rev_Date = day(Rev_Date) & "/" & month(Rev_Date) & "/" & year(Rev_Date)
					Else
						Rev_Date = month(Rev_Date) & "/" & day(Rev_Date) & "/" & year(Rev_Date)
					End If %>
					<%=Rev_Date%>
                    </td>
					<td nowrap>
					<%if pcv_RatingType="0" then%>
						<%if Rev_MainRate>"0" then%><img src="../pc/catalog/<%if Rev_MainRate="2" then%><%=pcv_Img1%><%else%><%=pcv_Img2%><%end if%>" align="absbottom" alt="<%if Rev_MainRate="2" then%><%=pcv_MainRateTxt2%><%else%><%=pcv_MainRateTxt3%><%end if%>">
						<%else%>No rating<%end if%>
					<%end if%>
					<%if pcv_RatingType="1" then
						if pcv_CalMain="1" then
							Rev_Rating=Rev_MainDRate
						else
							For m=0 to FCount-1
								if (Fi(m)<>"1") and (Fi(m)<>"2") then
									query="SELECT pcRD_Feel,pcRD_Rate,pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & pcv_ID & " and pcRD_IDField=" & Fi(m)
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=connTemp.execute(query)
	
									IF not rs.eof then
	
									if FType(m)<"3" then
										FValue(m)=rs("pcRD_Comment")
									end if
									if FType(m)="3" then
										FValue(m)=rs("pcRD_Feel")
									end if
									if FType(m)="4" then
										FValue(m)=rs("pcRD_Rate")
									end if
					
									ELSE
					
									if FType(m)<"3" then
										FValue(m)=""
									end if
									if FType(m)="3" then
										FValue(m)=0
									end if
									if FType(m)="4" then
										FValue(m)=0
									end if
					
									END IF
								end if
								set rs=nothing
							Next
							tmp1=0
							tmp2=0
							For m=0 to FCount-1
								if FType(m)="4" then
									if FValue(m)>"0" then
										tmp1=tmp1+1
										tmp2=tmp2+clng(FValue(m))
									end if
								end if
							Next
							if tmp2>"0" then
								Rev_Rating=tmp2/tmp1
							else
								Rev_Rating=0
							end if
						end if
	
						Call WriteStar(Rev_Rating,0)
	
					end if%>
					</td>
					<%
					'PRV41 begin
			           If pcv_IDCust=0 Then
					 ' PRV41 end %>
					<td align="center">
					<input type="checkbox" size="3" name="C<%=count%>" value="1" class="clearBorder">
					<input type="hidden" name="ID<%=count%>" value="<%=pcv_ID%>">
					</td>
					<%
					'PRV41 begin
			         End if
					 ' PRV41 end %>
				</tr>
			<% Next
		End If %>

		<%
		' PRV41 begin
		if DataEmpty<>1 And pcv_IDCust=0 Then
		' PRV41 end %>
			<tr>
				<td colspan="6" class="cpLinksList" align="right">
					<script language="JavaScript">
                    <!--
                    function checkAll() {
                    for (var j = 1; j <= <%=count%>; j++) {
                    box = eval("document.checkboxform.C" + j); 
                    if (box.checked == false) box.checked = true;
                         }
                    }
                    
                    function uncheckAll() {
                    for (var j = 1; j <= <%=count%>; j++) {
                    box = eval("document.checkboxform.C" + j); 
                    if (box.checked == true) box.checked = false;
                         }
                    }
                    
                    //-->
                    </script>
					<input type=hidden name=count value=<%=count%>>
					<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
				</td>
			</tr>
		<%end if%>
		<tr>
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="6">
			<%
			'PRV41 begin
			If pcv_IDCust=0 Then
			' PRV41 end %>
				<%if DataEmpty<>1 then%>
				<input type="submit" value=" <%if pcv_RevName="1" then%>Approve<%else%>Hide<%end if%> selected " name="submit1"<%
				' PRV41 begin
				If pcv_RevName<>"1" And pcv_RewardForReview Then
				strPrvPrompt = Replace(dictLanguage.Item(Session("language")&"_prv_36"),"'","\'")
				strPrvPrompt = Replace(strPrvPrompt, "<REWARD_POINTS_LABEL>", RewardsLabel, 1, -1, vbTextCompare)			
				%>
				onclick="return(confirm('You are about to hide selected product reviews from your database. <% = strPrvPrompt %>\nAre you sure you want to complete this action?'));"
				<%
				End if
				' PRV41 end
				%>>&nbsp;
				<input type="button" value="Reset Rating" onclick="javascript:if (confirm('You are about to reset the rating for this product. All customer reviews for this product will be deleted from the database. Are you sure you want to complete this action?')) location='prv_ManageReviews.asp?action=resetrating&IDProduct=<%=pcv_IDProduct%>&nav=<%=pcv_RevName%>'">&nbsp;
				<input type="submit" value="Delete selected" name="submit2" class="ibtnGrey" <%
				' PRV41 begin
				If pcv_RewardForReview Then 
				Dim strPrvPrompt
				strPrvPrompt = Replace(dictLanguage.Item(Session("language")&"_prv_36"),"'","\'")
				strPrvPrompt = Replace(strPrvPrompt, "<REWARD_POINTS_LABEL>", RewardsLabel, 1, -1, vbTextCompare)
				%>
				onclick="return(confirm('You are about to remove selected product reviews from your database. <% = strPrvPrompt %>\nAre you sure you want to complete this action?'));"
				<% Else %>			
				onclick="return(confirm('You are about to remove selected product reviews from your database. Are you sure you want to complete this action?'));"
				<% end if
				' PRV41 end %>>
				<%end if%>
				</td>
			</tr>
			<tr>
				<td colspan="6">
                <hr>
				<%if DataEmpty=1 then%>
				<input type="button" value="Other Pending Reviews" onClick="location='prv_ManageRevPrds.asp?nav=1';">&nbsp;
				<%else%>
				<input type="button" value="Bad Words List" onClick="location='prv_ManageBadWords.asp';">&nbsp;
				<%end if%>
			<%
			' PRV41 begin
			End If
			' PRV41 end
			%>
			<input type="button" value="Back" onClick="location='prv_ManageRevPrds.asp?nav=<%=pcv_RevName%>';">
			</td>
		</tr>
	</table>
	<input type="hidden" name="nav" value="<%=request("nav")%>">
	<input type="hidden" name="IDProduct" value="<%=request("IDProduct")%>">
</form>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->