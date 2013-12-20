<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pcv_RevName=request("nav")
if pcv_RevName="" then
	pcv_RevName="2"
end if
pageTitle="View/Edit Product Reviews"
%>
<% section="reviews" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
Dim rs, connTemp, query
call opendb()
pcv_IDProduct=request("IDProduct")
pcv_IDReview=request("IDReview")
%>
<!--#include file="../pc/prv_getsettings.asp"-->
<!--#include file="prv_incfunctions.asp"-->
<%

Call CreateList()

query="SELECT description FROM products WHERE idproduct=" & pcv_IDProduct
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

pcv_PrdName=rs("description")
set rs=nothing

'PRV41 begin
Dim strPrvPrompt
strPrvPrompt = Replace(dictLanguage.Item(Session("language")&"_prv_36"),"'","\'")
strPrvPrompt = Replace(strPrvPrompt, "<REWARD_POINTS_LABEL>", RewardsLabel, 1, -1, vbTextCompare)			

' PRV41 end

IF FCount>0 THEN
	pcv_showNote=0
	' PRV41 begin
	query="SELECT pcRev_IDReview,pcRev_Date,pcRev_MainRate,pcRev_MainDRate,pcRev_Active, pcRev_IDCustomer FROM pcReviews where pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_IDReview=" & pcv_IDReview
	' PRV41 end
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	pcArrayR=rs.getRows(rs.PageSize)
	set rs=nothing
	
	intReCount=Ubound(pcArrayR,2)
		
	Rev_ID=pcArrayR(0,0)
	Rev_Date=pcArrayR(1,0)
	pcv_Feel=pcArrayR(2,0)
	pcv_Rate=pcArrayR(3,0)
	Rev_Active=pcArrayR(4,0)
	' PRV41 begin
	Rev_IDCustomer = fnzeroifnull(pcArrayR(5,0))
	' PRV41 end

	query="SELECT pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & Rev_ID & " and pcRD_IDField=1"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	Rev_CustName=rs("pcRD_Comment")
	set rs=nothing
	
	query="SELECT pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & Rev_ID & " and pcRD_IDField=2"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	Rev_Title=rs("pcRD_Comment")
	set rs=nothing
	
	For m=0 to FCount-1
		if (Fi(m)<>"1") and (Fi(m)<>"2") then
			query="SELECT pcRD_Feel,pcRD_Rate,pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & Rev_ID & " and pcRD_IDField=" & Fi(m)
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
			set rs=nothing
		end if
	Next
	
	%>
	<% ' START show message, if any %>
		<!--#include file="pcv4_showMessage.asp"-->
	<% 	' END show message %>
	<h2>Product Name: <strong><%=pcv_PrdName%></strong>&nbsp;|&nbsp;Review ID#: <%=pcv_IDReview%></h2>
    
	<script>

	function newWindow(file,window) {
			msgWindow=open(file,window,'resizable=no,width=400,height=500');
			if (msgWindow.opener == null) msgWindow.opener = self;
	}
	
	function Form1_Validator(theForm)
	{
		<%For i=0 to FCount-1
		if FRe(i)="1" then%>
		if <%if FType(i)<"3" then%>(theForm.Field<%=Fi(i)%>.value == "")<%else%>((theForm.Field<%=Fi(i)%>.value == "") || (theForm.Field<%=Fi(i)%>.value == "0"))<%end if%>
		{
		alert("<%if FType(i)<"3" then%>Enter<%else%>Select<%end if%> a value for '<%=FName(i)%>'");
		<%if FType(i)<"3" then%>
		theForm.Field<%=Fi(i)%>.focus();
		<%end if%>
		return(false);
		}
		<%end if
		Next%>
		<%if pcv_RatingType="0" then%>
		if ((theForm.feel.value == "") || (theForm.feel.value == "0"))
		{
		alert("Select a value for <%=dictLanguage.Item(Session("language")&"_prv_5")%>");
		return(false);
		}
		<%else
		if pcv_CalMain="1" then%>
		if ((theForm.rate.value == "") || (theForm.rate.value == "0"))
		{
		alert("Select a value for <%=dictLanguage.Item(Session("language")&"_prv_5")%>");
		return(false);
		}
		<%end if
		end if%>
	
		return(true);
	}
	
	function TestURL(tmpField)
	{
		var tmp1=tmpField.value;
		tmp1=tmp1.toUpperCase();
		
		if ((tmp1.indexOf('HTTP://')==0) || (tmp1.indexOf('HTTPS://')==0))
		{
			document.getElementById("URL_" + tmpField.name).style.display='';
		}
		else
		{
			document.getElementById("URL_" + tmpField.name).style.display='none';
		}
	}
	
	function genURL(tmpField)
	{
		tmpField.value='<a href="' + tmpField.value + '" target="_blank">' + tmpField.value + '</a>';
	}
	
	</script>
	
	<form name="rating" method="post" action="prv_EditReviewB.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
		<input type="hidden" name="IDProduct" value="<%=pcv_IDProduct%>">
		<input type="hidden" name="IDReview" value="<%=pcv_IDReview%>">
		<input type="hidden" name="nav" value="<%=pcv_RevName%>">
		<table class="pcCPcontent">
		    <%
		    ' PRV41 begin
			If Rev_IDCustomer>1 Then  ' If there's no customer ID, then no reason to show info about rewards
			
				'// Retrieve customer name
				query = "SELECT name, lastName, email FROM customers WHERE idCustomer = " & Rev_IDCustomer
				set rs2 = Server.CreateObject("ADODB.Recordset")
				set rs2 = conntemp.execute(query)
				if not rs2.eof then
					pcStrCustName = rs2("name") & " " & rs2("lastName")
				end if
				set rs2 = nothing
				
			%>
            	<tr>
                	<td align="right" valign="top">Customer:</td>
                    <td><a href="modCusta.asp?idcustomer=<%=Rev_IDCustomer%>" target="_blank"><%=pcStrCustName%></a></td>
                </tr>
            <%

				Dim strRewards, pcv_MaxPoints, pcv_FirstPoints, pcv_AddPoints, pcv_PointsToAdd
				query = "SELECT top 1 pcRS_RewardForReview, pcRS_RewardForReviewFirstPts, pcRS_RewardForReviewAdditionalPts, pcRS_RewardForReviewMaxPts FROM pcRevsettings"
				Set rs = connTemp.execute(query)
				If rs.eof = False Then
				   If fnzeroifnull(rs("pcRS_RewardForReview"))=1 Then
				      pcv_FirstPoints=rs("pcRS_RewardForReviewFirstPts")
					  pcv_AddPoints=rs("pcRS_RewardForReviewAdditionalPts")
				   	  pcv_MaxPoints=rs("pcRS_RewardForReviewMaxPts")
					  pcv_PointsToAdd=0
					  
						query="SELECT Sum(pcRP_PointsAwarded) AS TotalPoints FROM pcReviewPoints WHERE pcRP_IDCustomer=" & Rev_IDCustomer & ";"
						set rsQ=Server.CreateObject("ADODB.Recordset")
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							prv_TotalPoints=rsQ("TotalPoints")
							if IsNull(prv_TotalPoints) OR prv_TotalPoints="" then
								prv_TotalPoints=0
							end if
						end if
						set rsQ=nothing
						
						'// First review for which points are awarded or not?
						if prv_TotalPoints=0 then
							pcv_PointsToAdd=pcv_FirstPoints
							else
							pcv_PointsToAdd=pcv_AddPoints
						end if
						
						query = "SELECT pcRP_PointsAwarded FROM pcReviewPoints WHERE pcRP_IDReview=" & Rev_ID
						Set rs2 = connTemp.execute(query)
						If rs2.eof = False Then
							 strRewards = "<strong>" & fnzeroifnull(rs2("pcRP_PointsAwarded")) & "</strong> " & RewardsLabel & " were awarded for this review."
						else
							if Rev_Active="1" then
								strRewards = RewardsLabel & " were <strong>not</strong> awarded for this review."
							else
								if CLng(pcv_MaxPoints)>0 then
									if CLng(prv_TotalPoints)>=CLng(pcv_MaxPoints) then
										strRewards = RewardsLabel & " for Review (points) will not be awarded because the customer has already reached the maximum amount allowed."
									else
										strRewards = pcv_PointsToAdd & " " & RewardsLabel & " for Review (points) will be awarded if you approve this review."
									end if
								else
									strRewards = pcv_PointsToAdd & " " & RewardsLabel & " for Review (points) will be awarded if you approve this review."
								end if
							end if
						end if
						
						strRewards=strRewards & "<br><br>So far <strong>" & prv_TotalPoints & "</strong> " & RewardsLabel & " have been awarded to this customer for writing reviews."
	
						if CLng(pcv_MaxPoints)>0 then
							if CLng(prv_TotalPoints)>=CLng(pcv_MaxPoints) then
								strRewards=strRewards & "<br>The customer has reached the <strong>maximum</strong> amount of " & RewardsLabel & " that can be accrued. Therefore, future reviews will <u>not</u> grant additional " & RewardsLabel & "."
							else
								strRewards=strRewards & "<br>The customer will accrue " & RewardsLabel & " until the <strong>maximum</strong> amount of '" & RewardsLabel & " for Reviews' is reached. The maximum is <strong>" & pcv_MaxPoints & " " & RewardsLabel & "</strong>."
							end if
						end if					  
					  
				   End if
				End If
				Set rs = Nothing

				If Len(Trim(strRewards))>0 Then 
				   %>
				   <tr>
                   	  <td align="right" valign="top"><%=RewardsLabel%> for Reviews Program</td>
					  <td><%=strRewards%></td>
				   </tr>
				   <%
				End If
				
				%>
				   <tr>
				      <td>&nbsp;</td>
					  <td><a href="prv_ManageReviews.asp?IDProduct=<% = pcv_IDProduct %>&nav=<% = pcv_RevName %>&idcustomer=<% = Rev_IDCustomer %>">See all reviews by this customer</a></td>
				   </tr>
                   <tr>
                   		<td colspan="2"><hr></td>
                   </tr>
				<%
			End If
		    ' PRV41 end
			%>
			<tr>
			<td width="30%" nowrap align="right">Is the review active?</td>
			<% 'PRV41 start %>
			<td width="70%"><input type="checkbox" name="active" value="1" <%if Rev_Active="1" then%>checked<%end if%> <% If Rev_Active="1" Then 
			%> onclick="if(!this.checked){alert('<% = Replace(strPrvPrompt,"'","\'") %>');}"<% End If %> class="clearBorder"></td><% 'PRV41 end %>
			</tr>
			<tr>
				<td nowrap align="right">
					Customer Name:
				</td>
				<td><input type="text" size="45" name="Field1" value="<%=Rev_CustName%>">
					<%For i=0 to FCount-1
						if Fi(i)="1" then
							if FRe(i)="1" then%>
							<img src="images/pc_required.gif" width="9" height="9" alt="Required">
							<%end if
							exit for
						end if
					Next%>
				</td>
			</tr>
			<tr>
				<td nowrap align="right">
					Review Title:
				</td>
				<td><input type="text" size="45" name="Field2" value="<%=Rev_Title%>">
					<%For i=0 to FCount-1
						if Fi(i)="2" then
							if FRe(i)="1" then%>
								<img src="images/pc_required.gif" width="9" height="9" alt="Required">
							<%end if
							exit for
						end if
					Next%></td>
			</tr>
			<%if pcv_RatingType="0" then%>
				<tr>
					<td nowrap align="right">
					<%=dictLanguage.Item(Session("language")&"_prv_5")%>
					</td>
					<td><input name="feel" type="hidden" value="<%=pcv_feel%>"><img src="../pc/catalog/<%=pcv_Img1%>" border="0" align="absbottom"><input name="feel1" value="2" type="radio" onclick="document.rating.feel.value='2';" <%if pcv_feel="2" then%>checked<%end if%> class="clearBorder"><%=pcv_MainRateTxt2%> <img src="../pc/catalog/<%=pcv_Img2%>" border="0" align="absbottom"><input name="feel1" value="1" type="radio" onclick="document.rating.feel.value='1';" <%if pcv_feel="1" then%>checked<%end if%> class="clearBorder"><%=pcv_MainRateTxt3%>
					<img src="images/pc_required.gif" width="9" height="9" alt="Required">
					</td>
				</tr>
			<%else
				if pcv_CalMain="1" then%>
					<tr>
						<td nowrap align="right">
						<%=dictLanguage.Item(Session("language")&"_prv_5")%>
						</td>
						<td><input name="rate" type="hidden" value="<%=pcv_rate%>"><%if pcv_CalMain="1" then%><%pcv_showNote=1%><%for k=1 to pcv_MaxRating%><input name="rate1" value="<%=k%>" type="radio" onclick="document.rating.rate.value='<%=k%>';" <%if pcv_rate<>"" then%><%if clng(k)=clng(pcv_rate) then%>checked<%end if%><%end if%> class="clearBorder">&nbsp;<%next%> <%=dictLanguage.Item(Session("language")&"_prv_13")%><%=pcv_MaxRating%><%=dictLanguage.Item(Session("language")&"_prv_13a")%><%end if%>
						<img src="images/pc_required.gif" width="9" height="9" alt="Required">
						</td>
					</tr>
				<%end if
			end if%>
			<%For i=0 to FCount-1
				if (Fi(i)<>"1") and (Fi(i)<>"2") then%>
					<tr>
						<td align="right" valign="top"><%=FName(i)%>:</td>
						<td>
						<%IF FType(i)="0" THEN%>
							<input type="text" size="45" name="Field<%=Fi(i)%>" value="<%=Server.HTMLEncode(FValue(i))%>" onchange="javascript:TestURL(this);" onfocus="javascript:TestURL(this);" onblur="javascript:TestURL(this);" onclick="javascript:TestURL(this);" onmouseout="javascript:TestURL(this);" onmouseover="javascript:TestURL(this);" onkeypress="javascript:TestURL(this);"> <span id="URL_Field<%=Fi(i)%>" <%if instr(ucase(FValue(i)),"HTTP://")<>1 and instr(ucase(FValue(i)),"HTTPS://")<>1 then%>style="display:none"<%end if%>><input type="button" name="genField<%=Fi(i)%>" value="Generate URL" onclick="javascript:genURL(document.rating.Field<%=Fi(i)%>);"></span>
						<%END IF%>
						<%IF FType(i)="1" THEN%>
							<textarea cols="50" rows="5" name="Field<%=Fi(i)%>"><%=Server.HTMLEncode(FValue(i))%></textarea>
							<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?iform=rating&fi=Field<%=Fi(i)%>','window2')" class="ibtnGrey">
						<%END IF%>
						<%IF FType(i)="2" THEN%>
							<select name="Field<%=Fi(i)%>">
							<%if FRe(i)<>"1" then%>
								<option value=""></option>
							<%end if%>
							<% 
							query="SELECT pcRL_Name,pcRL_Value FROM pcRevLists WHERE pcRL_IDField=" & Fi(i)
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)
							
							if not rs.eof then

							pcArray=rs.getRows()
							set rs=nothing
							
							intCount=ubound(pcArray,2)
							For j=0 to intCount %>
								<option value="<%=pcArray(1,j)%>" <%if FValue(i)&""=pcArray(1,j)&"" then%>selected<%end if%>><%=pcArray(0,j)%></option>
							<%Next
							end if%>
							</select>
							<% 
						END IF%>
						
						<%IF FType(i)="3" THEN%>
							<input name="Field<%=Fi(i)%>" type="hidden" value="<%=FValue(i)%>"><img src="../pc/catalog/<%=pcv_Img1%>" border="0" align="absbottom"><input name="Field<%=Fi(i)%>a" value="2" type="radio" onclick="document.rating.Field<%=Fi(i)%>.value='2';" <%if FValue(i)="2" then%>checked<%end if%> class="clearBorder"><%=pcv_SubRateTxt1%> <img src="../pc/catalog/<%=pcv_Img2%>" border="0" align="absbottom"><input name="Field<%=Fi(i)%>a" value="1" type="radio" onclick="document.rating.Field<%=Fi(i)%>.value='1';" <%if FValue(i)="1" then%>checked<%end if%> class="clearBorder"><%=pcv_SubRateTxt2%>
						<%END IF%>
						<%IF FType(i)="4" THEN%>
							<input name="Field<%=Fi(i)%>" type="hidden" value="<%=FValue(i)%>"><%for k=1 to pcv_MaxRating%><input name="Field<%=Fi(i)%>a" value="<%=k%>" type="radio" onclick="document.rating.Field<%=Fi(i)%>.value='<%=k%>';" <%if FValue(i)&""=k&"" then%>checked<%end if%> class="clearBorder">&nbsp;<%next%>
							<%if pcv_showNote=0 then
								pcv_showNote=1%>
								<%=dictLanguage.Item(Session("language")&"_prv_13")%><%=pcv_MaxRating%><%=dictLanguage.Item(Session("language")&"_prv_13a")%>
							<%end if%>
						<%END IF%>
						<%if FRe(i)="1" then%>
							<img src="images/pc_required.gif" width="9" height="9" alt="Required">
						<%end if%>
						</td>
					</tr>
				<%end if
			Next%>
            <tr>
            	<td colspan="2"><hr></td>
            </tr>
			<tr>
				<td align="right" valign="top">&nbsp;</td>
				<td>
				<input type="submit" name="submit" value="Update Review" class="submit2">&nbsp;<input type="button" name="Del" value="Delete Review" onclick="javascript:if (confirm('You are about to remove this product review from your database. <% = Replace(strPrvPrompt,"'","\'") %>\nAre you sure you want to complete this action?')) location='prv_EditReviewB.asp?action=delete&IDReview=<%=pcv_IDReview%>&IDProduct=<%=pcv_IDProduct%>&nav=<%=pcv_RevName%>'">&nbsp;<input type="button" name="Back" value="Back" onclick="location='prv_ManageReviews.asp?IDProduct=<%=pcv_IDProduct%>&nav=<%=pcv_RevName%><%
				' PRV41 begin
				If fnzeroifnull(request("idcustomer"))>0 Then
				   response.write "&idcustomer=" & fnzeroifnull(request("idcustomer"))
				End if
				' PRV41 end
				%>'"></td>
			</tr>
		</table>
	</form>
<%END IF%>
<!--#include file="AdminFooter.asp"-->