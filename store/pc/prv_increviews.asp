<%
'Review Text Length Limitation
pcv_RevLenLimit=500
tmp_strMore=dictLanguage.Item(Session("language")&"_viewPrd_21")
tmp_strMore=replace(tmp_strMore,"...","")

IF pcv_Active="1" THEN
	query="SELECT pcRE_IDProduct FROM pcRevExc WHERE pcRE_IDProduct=" & pcv_IDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if rs.eof then
		Prv_Accept=1
	else
		Prv_Accept=0
	end if
	set rs=nothing
	
	IF Prv_Accept=1 THEN
		Call CreateList() %>
		<script>
		function openbrowser(url) {
						self.name = "productPageWin";
						popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
						if (navigator.appName == 'Netscape') {
										popUpWin.focus();
						}
		}
		</script>
		<form name="rating" class="pcForms">
			<table class="pcShowContent">
				<tr>
					<td colspan="2" class="pcSectionTitle">
					<p><a name="productReviews"></a><%=dictLanguage.Item(Session("language")&"_prv_1")%></p>
					</td>
				</tr>
				<tr>
					<td colspan="2">
				
			<%
				if pcv_ShowRatSum="1" then
				pcv_SaveRating=CalRating()
					IF pcv_RatingType="0" then
						query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pIDProduct
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)

						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if

						pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)

						query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pIDProduct & " AND pcRev_Active=1 AND pcRev_MainRate>0"
						set rs=connTemp.execute(query)

						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if

						intCount = clng(rs("ct"))

						set rs=Nothing
						%>
						<p>
						<%if pcv_tmpRating>"0" then%><%=dictLanguage.Item(Session("language")&"_prv_2")%><img src="catalog/<%=pcv_Img1%>" align="absbottom"><%=pcv_tmpRating%>% <%=pcv_MainRateTxt1%> (<%=intCount%>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)<%end if%>
						</p>
					<%
					ELSE
						if pcv_CalMain="1" then     ' Can be set independently of sub-ratings 
							query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pIDProduct
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)

							if err.number<>0 then
								call LogErrorToDatabase()
								set rs=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if

							pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)

							set rs=nothing
							if pcv_tmpRating>"0" then%>
							<p><%=dictLanguage.Item(Session("language")&"_prv_39")%>
							<% Call WriteStar(pcv_tmpRating,1) %></p>
							<%end if%>
						<% 
						else    ' Will be calculated automatically by averaging sub-ratings
						pcv_tmpRating=pcv_SaveRating
							if pcv_tmpRating>"0" then%>
							<p><%=dictLanguage.Item(Session("language")&"_prv_2")%>
							<% Call WriteStar(pcv_tmpRating,1) %></p>
							<%end if%>
						<% 
						end if
					END IF 'Main Rating %>
				</td>
			</tr>
				
				<% '******** Display Sub-Rating
				if FCount>"0" then
					For m=0 to FCount-1
						if FType(m)>"2" then%>
							
							<%
							IF FType(m)="3" then
								if FRecord(m)>"0" then
									if FValue(m)="0" then
										pcv_Rpercent=0
									else
										pcv_Rpercent=Round((Fvalue(m)/FRecord(m))*100)
									end if
									if pcv_Rpercent<>"0" then%>
									<tr>
									<td width="40%">
										<p><%=FName(m)%>:</p>
									</td>
									<td width="60%">
										<img src="catalog/<%=pcv_Img1%>" align="absbottom" alt="<%=pcv_SubRateTxt1%>"> <%=pcv_Rpercent%>%&nbsp;<img src="catalog/<%=pcv_Img2%>" align="absbottom" alt="<%=pcv_SubRateTxt2%>"> <%=100-pcv_Rpercent%>%
									</td>
									</tr>
								<% end if
								end if
							ELSE
								if FRecord(m)>"0" then
									if FValue(m)="0" then
										Rev_Rating=0
									else
										Rev_Rating=Fvalue(m)/FRecord(m)
									end if
								else
									Rev_Rating=0
								end if
								if Rev_Rating<>"0" then%>
								<tr>
									<td width="40%">
										<p><%=FName(m)%>:</p>
									</td>
									<td width="60%">
										<%Call WriteStar(Rev_Rating,0)%>
									</td>
								</tr>
								<%end if%>
							<%END IF%>
							
						<%end if
					Next%>
			<% end if
            Else

					IF pcv_RatingType="0" then
						query="SELECT count(*) FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 and pcRev_MainRate>0"
                    else
						query="SELECT count(*) FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 and pcRev_MainDRate>0"
                    End if
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
					
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end If
                        intCount = CLng(rs(0))

                        rs.close
                        Set rs = Nothing
                                    
			END IF 'Show Rating Sumary%>	
			<tr>
				<td colspan="2">
					<p>
					<%
					if pcv_RevCount="0" then 
						tmppcv_RevCount=1
					else
						tmppcv_RevCount=pcv_RevCount
					end if
					query="SELECT top " & tmppcv_RevCount & " pcRev_IDReview,pcRev_Date,pcRev_MainRate,pcRev_MainDRate FROM pcReviews where pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 ORDER BY pcRev_Date DESC"
					set rs=connTemp.execute(query)
					if not rs.eof Then
                       If intCount>0 then%>
					      <a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><strong>&gt;&gt;&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_3")%></strong> [<%=intCount%>]</a>
					      &nbsp;|&nbsp;
					<%
                       End If
                    end if
					set rs=nothing%><a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>
                    </p>
					<p>&nbsp;</p>
					<%if pcv_RatingType="0" then%>
						<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="feel" type="hidden" value=""><img src="catalog/<%=pcv_Img1%>" align="absbottom"><input name="feel1" value="2" type="radio" onclick="document.rating.feel.value='2';" class="clearBorder"> <%=pcv_MainRateTxt2%>  <img src="catalog/<%=pcv_Img2%>" align="absbottom"><input name="feel1" value="1" type="radio" onclick="document.rating.feel.value='1';" class="clearBorder"> <%=pcv_MainRateTxt3%> <input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&feel=' + document.rating.feel.value);" class="submit2"></p>
					<%else%>
						<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="rate" type="hidden" value=""><%if pcv_CalMain="1" then%><%for k=1 to pcv_MaxRating%><input name="rate1" value="<%=k%>" type="radio" onclick="document.rating.rate.value='<%=k%>';" class="clearBorder"><span class="pcSmallText"><%=k%></span>&nbsp;<%next%><%end if%><input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&rate=' + document.rating.rate.value);" class="submit2"></p>
                    <%end if%>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<hr>
					<%IF pcv_RevCount>"0" then
						pcv_CShow=pcv_RevCount
						iPageCurrent=1%>
						<!--#include file="prv_incshow.asp"-->
					<%END IF%>
				</td>
			</tr>
			</table>
		</form>

		<%if RCount>="5" then%>
			<form name="rating1" class="pcForms">
			<p><%
				query="SELECT pcRev_IDReview,pcRev_Date,pcRev_MainRate,pcRev_MainDRate FROM pcReviews where pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 ORDER BY pcRev_Date DESC"
				set rs=connTemp.execute(query)
				if not rs.eof then%>
				<a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><strong>&gt;&gt;&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_3")%></strong> [<%=intCount%>]</a>&nbsp;|&nbsp;
			<%end if
			set rs=nothing%>
			<a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>
            </p>
			<p>&nbsp;</p>
				<%if pcv_RatingType="0" then%>
					<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="feel" type="hidden" value=""><img src="catalog/<%=pcv_Img1%>" align="absbottom"><input name="feel1" value="2" type="radio" onclick="document.rating1.feel.value='2';" class="clearBorder"> <%=pcv_MainRateTxt2%>  <img src="catalog/<%=pcv_Img2%>" align="absbottom"><input name="feel1" value="1" type="radio" onclick="document.rating1.feel.value='1';" class="clearBorder"> <%=pcv_MainRateTxt3%> <input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&feel=' + document.rating1.feel.value);" class="submit2"></p>
				<%else%>
					<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="rate" type="hidden" value=""><%if pcv_CalMain="1" then%><%for k=1 to pcv_MaxRating%><input name="rate1" value="<%=k%>" type="radio" onclick="document.rating1.rate.value='<%=k%>';" class="clearBorder"><span class="pcSmallText"><%=k%></span>&nbsp;<%next%><%end if%><input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&rate=' + document.rating1.rate.value);" class="submit2"></p>
				<%end if%>
			</form>
		<%end if%>
	<% END IF
END IF%>