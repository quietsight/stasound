<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<% 
' Check if the store is on. If store is turned off display store message
If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If

Dim connTemp,rs
call opendb()
%>
<!--#include file="prv_getsettings.asp"-->
<%
if pcv_Active<>"1" then
	call closedb()
	response.redirect "default.asp"
end if

pcv_IDProduct=GetUserInput(request("IDProduct"),0)
	if not validNum(pcv_IDProduct) then
		call closedb()
		response.redirect "msg.asp?message=210"
	end if
	
pIdCustomer=GetUserInput(request("idcustomer"),0)
If NOT len(pIdCustomer)>0 Then
	pIdCustomer=Session("idCustomer")
End If
if not validNum(pIdCustomer) then
	call closedb()
	response.redirect "msg.asp?message=210"
end if

'// Check product exclusion
	query="SELECT pcRE_IDProduct FROM pcRevExc WHERE pcRE_IDProduct=" & pcv_IDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)	
	if not rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "default.asp"
	end if

'// Check customer eligibility to write a review
	pcv_IPAddress=Request.ServerVariables("REMOTE_ADDR")
	
	query="SELECT pcRev_IDReview FROM pcReviews where pcRev_IP='" & pcv_IPAddress & "' and pcRev_IDProduct=" & pcv_IDProduct
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
	
'// Get customer name, if any
if validNum(pIdCustomer) and pIdCustomer>0 then
	query = "SELECT name, lastName, email FROM customers WHERE idCustomer = " & pIdCustomer
	set rs = conntemp.execute(query)
	pcStrCustName = rs("name") & " " & rs("lastName")
	session("pcStrCustName") = pcStrCustName
end if

pcv_Feel=GetUserInput(request("feel"),0)
pcv_Rate=GetUserInput(request("rate"),0)

query="SELECT description FROM products WHERE idproduct=" & pcv_IDProduct
set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

pcv_PrdName=rs("description")

query="SELECT pcRS_FieldList,pcRS_FieldOrder,pcRS_Required FROM pcReviewSpecials WHERE pcRS_IDProduct=" & pcv_IDProduct
set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

Dim Fi(100)
Dim FS(100)
Dim FRe(100)
Dim FName(100)
Dim FType(100)

if not rs.eof then

	pcv_FieldList=split(rs("pcRS_FieldList"),",")
	pcv_FieldOrder=split(rs("pcRS_FieldOrder"),",")
	pcv_Required=split(rs("pcRS_Required"),",")
	
	FCount=0
	For i=0 to ubound(pcv_FieldList)
		if pcv_FieldList(i)<>"" then
			Fi(FCount)=pcv_FieldList(i)
			FS(FCount)=pcv_FieldOrder(i)
			FRe(FCount)=pcv_Required(i)
				
			query="SELECT pcRF_Type,pcRF_Name FROM pcRevFields WHERE pcRF_IDField=" & Fi(FCount)
			set rs=connTemp.execute(query)
			
			FType(FCount)=rs("pcRF_Type")
			FName(FCount)=rs("pcRF_Name")
			FCount=FCount+1
		end if
	Next

	For i=0 to FCount-1
		For j=i+1 to FCount-1
			if FS(i)>FS(j) then
				tmpC=FS(j)
				FS(j)=FS(i)
				FS(i)=tmpC
				
				tmpC=Fi(j)
				Fi(j)=Fi(i)
				Fi(i)=tmpC
				
				tmpC=FRe(j)
				FRe(j)=FRe(i)
				FRe(i)=tmpC
				
				tmpC=FType(j)
				FType(j)=FType(i)
				FType(i)=tmpC
				
				tmpC=FName(j)
				FName(j)=FName(i)
				FName(i)=tmpC
			end if
		Next
	Next

else

	query="SELECT pcRF_IDField,pcRF_Name,pcRF_Type,pcRF_Required,pcRF_Order FROM pcRevFields WHERE pcRF_Active=1 order by pcRF_Order asc"
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
			FName(FCount)=pcArray(1,i)
			FType(FCount)=pcArray(2,i)
			FRe(FCount)=pcArray(3,i)
			FS(FCount)=pcArray(4,i)
			FCount=FCount+1
		Next
		
	end if

end if

IF FCount>0 THEN
	pcv_showNote=0
	
	Set conlayout=Server.CreateObject("ADODB.Connection")
	conlayout.Open scDSN
	Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
	Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
	%>
	<html>
		<head>
			<title><%=dictLanguage.Item(Session("language")&"_prv_9")%></title>
			<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
			<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
			<script>
			function Form1_Validator(theForm)
			{
				
				<%For i=0 to FCount-1
				if FRe(i)="1" then%>
					if <%if FType(i)<"3" then%>(theForm.Field<%=Fi(i)%>.value == "")<%else%>((theForm.Field<%=Fi(i)%>.value == "") || (theForm.Field<%=Fi(i)%>.value == "0"))<%end if%>
					{
					alert("Please <%if FType(i)<"3" then%>enter<%else%>select<%end if%> a value for '<%=FName(i)%>'");
					<%if FType(i)<"3" then%>
					theForm.Field<%=Fi(i)%>.focus();
					<%end if%>
					return(false);
					}
					<%if FType(i)="0" AND UCase(FName(i))="CUSTOMER NAME" then%>
						if (theForm.Field<%=Fi(i)%>.value.length < 3)
						{
							alert("The field '<%=FName(i)%>' does not contain enough characters");
							theForm.Field<%=Fi(i)%>.focus();
							return(false);
						}
					<%end if%>
					<%if FType(i)="0" AND UCase(FName(i))="TITLE" then%>
						if (theForm.Field<%=Fi(i)%>.value.length < 3)
						{
							alert("The field '<%=FName(i)%>' does not contain enough characters");
							theForm.Field<%=Fi(i)%>.focus();
							return(false);
						}
					<%end if%>
					<%if FType(i)="1" then%>
					<%if pcv_RewardForReviewMinLength>"0" then%>
						if (theForm.Field<%=Fi(i)%>.value.length < <% = pcv_RewardForReviewMinLength %>)
						{
							alert("The field '<%=FName(i)%>' does not contain enough characters. Must be at least <% = pcv_RewardForReviewMinLength %> characters in length.");
							theForm.Field<%=Fi(i)%>.focus();
							return(false);
						}
					<%end if%>
					<%end if%>
				<%end if
				Next%>
				<%if pcv_RatingType="0" then%>
				if (theForm.feel.value == "")
				{
				alert("Select a value for <%=dictLanguage.Item(Session("language")&"_prv_5")%>");
				return(false);
				}
				<%else
				if pcv_CalMain="1" then%>
				if (theForm.rate.value == "")
				{
				alert("Select a value for <%=dictLanguage.Item(Session("language")&"_prv_5")%>");
				return(false);
				}
				<%end if
				end if%>
			
				return(true);
			}
			</script>
		</head>
	<body>
	<div id="pcMain">
	<form name="rating" method="post" action="prv_postreviewB.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
	<%
	' PRV41 begin
	If Len(Trim(request("xrv")))>0 Then
	   response.write "<input type=""hidden"" name=""xrv"" value=""" & CLng(request("xrv")) & """>"
	End if
	' PRV41 end
	%>
	<input type="hidden" name="IDProduct" value="<%=pcv_IDProduct%>">
    <input type="hidden" name="IDCustomer" value="<%=pIdCustomer%>">
		<table class="pcMainTable">
			<tr>
				<td>
					<h1><%=dictLanguage.Item(Session("language")&"_prv_9")%></h1>
					<h2><%=dictLanguage.Item(Session("language")&"_prv_10")%> <%=pcv_PrdName%></h2>
				</td>
			</tr>
			<tr>
				<td>
					<table class="pcShowContent">
                    	<% if msg="" then
							msg=getUserInput(request.querystring("msg"),0)
						end if
						If msg<>"" then	%>
                            <tr>
                              <td colspan="2" align="center">
                                
                                    <div class="pcErrorMessage">
                                    <%=msg%>
                                    </div>
                              </td>
                            </tr>
                        <%  end if %>
						<tr>
							<td colspan="2" align="center">
								<img src="<%=rsIconObj("requiredicon")%>">&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_21")%>							</td>
						</tr>
						<tr>
							<td width="30%" nowrap align="right">
								<p><%=dictLanguage.Item(Session("language")&"_prv_11")%></p>							</td>
							<td><input type="text" size="45" name="Field1" value="<%=session("pcStrCustName")%>">
							<%For i=0 to FCount-1
								if Fi(i)="1" then
									if FRe(i)="1" then%>
										<img src="<%=rsIconObj("requiredicon")%>">
									<%end if
									exit for
								end if
							Next%>							</td>
						</tr>
						<tr>
							<td nowrap align="right">
								<p><%=dictLanguage.Item(Session("language")&"_prv_12")%></p>							</td>
							<td><input type="text" size="45" name="Field2" value="">
							<%For i=0 to FCount-1
								if Fi(i)="2" then
									if FRe(i)="1" then%>
										<img src="<%=rsIconObj("requiredicon")%>">
									<%end if
									exit for
								end if
							Next%></td>
						</tr>
						<%if pcv_RatingType="0" then%>
							<tr>
								<td nowrap align="right">
									<p><%=dictLanguage.Item(Session("language")&"_prv_5")%></p>								</td>
								<td>
									<input name="feel" type="hidden" value="<%=pcv_feel%>"><img src="catalog/<%=pcv_Img1%>" align="absbottom"><input name="feel1" value="2" type="radio" onclick="document.rating.feel.value='2';" <%if pcv_feel="2" then%>checked<%end if%> class="clearBorder"><%=pcv_MainRateTxt2%> <img src="catalog/<%=pcv_Img2%>" align="absbottom"><input name="feel1" value="1" type="radio" onclick="document.rating.feel.value='1';" <%if pcv_feel="1" then%>checked<%end if%> class="clearBorder"><%=pcv_MainRateTxt3%>
								<img src="<%=rsIconObj("requiredicon")%>">								</td>
							</tr>
						<%else
						    if pcv_CalMain="1" then
						    %>
							<tr>
								<td nowrap align="right">
									<p><%=dictLanguage.Item(Session("language")&"_prv_5")%></p>									
								</td>
								<td>
									<input name="rate" type="hidden" value="<%=pcv_rate%>"><%pcv_showNote=1%><%for k=1 to pcv_MaxRating%><input name="rate1" value="<%=k%>" type="radio" onclick="document.rating.rate.value='<%=k%>';" <%if pcv_rate<>"" then%><%if clng(k)=clng(pcv_rate) then%>checked<%end if%><%end if%> class="clearBorder"><span class="pcSmallText"><%=k%></span>&nbsp;<%next%> <%=dictLanguage.Item(Session("language")&"_prv_13")%><%=pcv_MaxRating%><%=dictLanguage.Item(Session("language")&"_prv_13a")%>
								    <img src="<%=rsIconObj("requiredicon")%>">									
								</td>
							</tr>
						    <%end if
						
						end if%>
						<%For i=0 to FCount-1
							if (Fi(i)<>"1") and (Fi(i)<>"2") then
								pcv_test=1
								IF FType(i)="2" THEN
								query="SELECT pcRL_Name,pcRL_Value FROM pcRevLists WHERE pcRL_IDField=" & Fi(i)
								set rs=connTemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
										set rs=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
								if rs.eof then
								pcv_test=0
								end if
								END IF
								if pcv_test=1 then%>
								<tr>
									<td align="right" valign="top">
										<p><%=FName(i)%>:</p>									</td>
									<td>
										<%IF FType(i)="0" THEN%>
											<input type="text" size="45" name="Field<%=Fi(i)%>" value="">
										<%END IF%>
										<%IF FType(i)="1" THEN%>
											<textarea cols="40" rows="5" name="Field<%=Fi(i)%>" <%if pcv_RewardForReviewMinLength>"0" then%>onkeyup="javascript:testchars(this,'1',<%=pcv_RewardForReviewMinLength%>);"<%end if%>></textarea>
											<% ' PRV41 begin
											if pcv_RewardForReviewMinLength>"0" Then
											%>
											<script>
											function testchars(tmpfield,idx,maxlen)
											{
												var tmp1=tmpfield.value;
												if (tmp1.length>=maxlen) {
												   document.getElementById('charcount').style.display='none';}
												else {
												   document.getElementById('charcount').style.display='';
												}
												document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
											}
											</script>

											<br><span id="charcount">
											<%response.write dictLanguage.Item(Session("language")&"_prv_34")%><span id="countchar1" name="countchar1" style="font-weight: bold"><%=pcv_RewardForReviewMinLength%></span> <%response.write dictLanguage.Item(Session("language")&"_prv_35")%>
											</span>
											<%end If
											' PRV41 end %>
										<%END IF%>
										<%IF FType(i)="2" THEN%>
										<select name="Field<%=Fi(i)%>">
										<%if FRe(i)<>"1" then%>
											<option value=""></option>
										<%end if%>
										<%
										pcArray=rs.getRows()
										intCount=ubound(pcArray,2)
										For j=0 to intCount
										%>
										<option value="<%=pcArray(1,j)%>"><%=pcArray(0,j)%></option>
									<%Next%>
									</select>
								<%END IF%>
								<%IF FType(i)="3" THEN%>
									<input name="Field<%=Fi(i)%>" type="hidden" value="0"><img src="catalog/<%=pcv_Img1%>" align="absbottom"><input name="Field<%=Fi(i)%>a" value="2" type="radio" onclick="document.rating.Field<%=Fi(i)%>.value='2';" class="clearBorder"><%=pcv_SubRateTxt1%> <img src="catalog/<%=pcv_Img2%>" align="absbottom"><input name="Field<%=Fi(i)%>a" value="1" type="radio" onclick="document.rating.Field<%=Fi(i)%>.value='1';" class="clearBorder"><%=pcv_SubRateTxt2%>
								<%END IF%>
								<%IF FType(i)="4" THEN%>
									<input name="Field<%=Fi(i)%>" type="hidden" value="0"><%for k=1 to pcv_MaxRating%><input name="Field<%=Fi(i)%>a" value="<%=k%>" type="radio" onclick="document.rating.Field<%=Fi(i)%>.value='<%=k%>';" class="clearBorder"><span class="pcSmallText"><%=k%></span>&nbsp;<%next%>
									<%if pcv_showNote=0 then
										pcv_showNote=1%>
										<%=dictLanguage.Item(Session("language")&"_prv_13")%><%=pcv_MaxRating%><%=dictLanguage.Item(Session("language")&"_prv_13a")%>
									<%end if%>
								<%END IF%>
								<%if FRe(i)="1" then%>
									<img src="<%=rsIconObj("requiredicon")%>">
								<%end if%>								</td>
								</tr>
							<%else
							IF FType(i)="2" THEN%>
								<input type="hidden" name="Field<%=Fi(i)%>" value="">
							<%end if
							end if
							end if
						Next%>
						<%
						Session("store_ReviewReg")="1"
						Session("store_ReviewRegpostnum")=""
						session("store_ReviewRegnum")="      "
						%>
						<% if (scSecurity=1) and (scReview=1) then%>
						<tr>
							<td nowrap align="right"><p><%=dictLanguage.Item(Session("language")&"_security_1")%></p></td>
							<td valign="top"><!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" --></td>
						</tr>
						<%end if%>
					    <tr>
						<td align="right" valign="top">&nbsp;</td>
						<td>
						<br>
						<input type="submit" name="submit" value="<%=dictLanguage.Item(Session("language")&"_prv_9")%>" class="submit2">&nbsp;<input type="reset" name="Reset" value="Reset"></td>
						</tr>
					</table>				</td>
			</tr>
		</table>
		</form>
	</div>
	</body>
</html>
<%

conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing

END IF
set rs=nothing
call closedb()
%>