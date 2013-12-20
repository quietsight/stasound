<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<% 
'*******************************
' START: Check store on/off, start PC session, check affiliate ID
'*******************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*******************************
' END: Check store on/off, start PC session, check affiliate ID
'*******************************

' Check to see if the user is updating the product after adding it to the shopping cart
tIndex=0
tUpdPrd=request.QueryString("imode")
if tUpdPrd="updOrd" then
	tIndex=request.QueryString("index")
end if

dim query, conntemp, rs, pcv_IDProduct, pIDProduct, pcv_IDCategory

pcv_IDProduct=trim(request("IDProduct"))
	if not validNum(pcv_IDProduct) then
		response.redirect "msg.asp?message=85"
	end if
	pIDProduct=pcv_IDProduct
	
pcv_IDCategory=trim(request("IDCategory"))
	if not validNum(pcv_IDCategory) then
		pcv_IDCategory=0
	end if
	
call opendb()
%>
<!--#include file="prv_getsettings.asp"-->
<!--#include file="prv_incfunctions.asp"-->
<%
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
		set rs=nothing
		call closedb()
		response.redirect "viewPrd.asp?IDProduct=" & pcv_IDProduct & "&IDCategory=" & pcv_IDCategory
	end if
	
	IF Prv_Accept=1 THEN
		Call CreateList()
		query="SELECT description, active FROM products WHERE idproduct=" & pcv_IDProduct
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

		pcv_PrdName=rs("description") 
		pcIntProductStatus = rs("active")
			if pcIntProductStatus=0 or isNull(pcIntProductStatus) or pcIntProductStatus="" then
				set rs = nothing
				call closeDb()
				response.redirect "msg.asp?message=95"
			end if
			
		set rs=nothing
		%>
		<script>
		function openbrowser(url) {
						self.name = "productPageWin";
						popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
						if (navigator.appName == 'Netscape') {
										popUpWin.focus();
						}
		}
		</script>
		
			<div id="pcMain">
				<table class="pcMainTable">
					<tr>
						<td>
							<h1><%=dictLanguage.Item(Session("language")&"_prv_1")%></h1>
							<h2><span style="font-weight: normal;"><%=dictLanguage.Item(Session("language")&"_prv_10")%></span> <%=pcv_PrdName%></h2>
							<p align="right"><a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>&nbsp;|&nbsp;<a href="viewPrd.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><%=dictLanguage.Item(Session("language")&"_prv_30")%></a></p>
						</td>
					</tr>
					<tr>
						<td>
								<table class="pcShowContent">
									<tr>
										<td colspan="2">
											<%
											if pcv_ShowRatSum="1" then
												pcv_SaveRating=CalRating()
												if pcv_RatingType="0" then
													call opendb()
													' PRV41 Start
													' Why not just use "SELECT (AVG(CONVERT(decimal(3, 2), pcRev_MainRate)) - 1) * 100 AS AverageRating from pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 and pcRev_MainRate>0;"  and be done?
													' PRV41 End


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
														
														
														
														call closedb()
														%>
														<%if pcv_tmpRating>"0" then%><p><%=dictLanguage.Item(Session("language")&"_prv_2")%><img src="catalog/<%=pcv_Img1%>" align="absbottom"><%=pcv_tmpRating%>% <%=pcv_MainRateTxt1%> (<%=intCount%>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)<%end if%></p>
												<%ELSE
													if pcv_CalMain="1" then
														call opendb()

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

															call closedb() %>
															<p>
															<%=dictLanguage.Item(Session("language")&"_prv_2")%>
															<%Call WriteStar(pcv_tmpRating,1)	%>
															</p>
														<%else
															pcv_tmpRating=pcv_SaveRating
															%>
															<p>
															<%=dictLanguage.Item(Session("language")&"_prv_2")%>
															<%Call WriteStar(pcv_tmpRating,1)%>
															</p>
														<% end if
													END IF 'Main Rating%>
											</td>
										</tr>
					
										<% '******** Display Sub-Rating
										if FCount>"0" then%>
												<%For m=0 to FCount-1
													if FType(m)>"2" then%>
													<tr>
														<td width="40%">
															<p><%=FName(m)%>:</p>
														</td>
														<td width="60%">
															<p>
																<%IF FType(m)="3" then
																	if FRecord(m)>"0" then
																		if FValue(m)="0" then
																			pcv_Rpercent=0
																		else
																			pcv_Rpercent=Round((FValue(m)/FRecord(m))*100)
																		end if %>
																		<img src="catalog/<%=pcv_Img1%>" align="absbottom" alt="<%=pcv_SubRateTxt1%>"> <%=pcv_Rpercent%>%&nbsp;<img src="catalog/<%=pcv_Img2%>" align="absbottom" alt="<%=pcv_SubRateTxt2%>"> <%=100-pcv_Rpercent%>%
																	<%else%>
																		<%=dictLanguage.Item(Session("language")&"_prv_15")%>
																	<% end if
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
																	Call WriteStar(Rev_Rating,0)
																END IF%>
															</p>
														</td>
													</tr>
									<%end if
								Next%>
						<%end if
					END IF 'Show Rating Sumary%>
					<tr>
						<td colspan="2"><hr></td>
					</tr>
					<tr>
						<td colspan="2">
						<form name="rating" class="pcForms">
							<%if pcv_RatingType="0" then%>
								<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="feel" type="hidden" value=""><img src="catalog/<%=pcv_Img1%>" align="absbottom"><input name="feel1" value="2" type="radio" onclick="document.rating.feel.value='2';" class="clearBorder"> <%=pcv_MainRateTxt2%>  <img src="catalog/<%=pcv_Img2%>" align="absbottom"><input name="feel1" value="1" type="radio" onclick="document.rating.feel.value='1';" class="clearBorder"> <%=pcv_MainRateTxt3%> <input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&feel=' + document.rating.feel.value);" class="submit2"></p>
							<%else%>
								<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="rate" type="hidden" value=""><%if pcv_CalMain="1" then%><%for k=1 to pcv_MaxRating%><input name="rate1" value="<%=k%>" type="radio" onclick="document.rating.rate.value='<%=k%>';" class="clearBorder"><span class="pcSmallText"><%=k%></span>&nbsp;<%next%><%end if%> <input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&rate=' + document.rating.rate.value);" class="submit2"></p>
						<%end if%>
						</form>
						</td>
					</tr>
					<tr>
						<td colspan="2"><hr></td>
					</tr>
					<tr>
						<td colspan="2">
							<%pcv_CShow=15
							iPageCurrent=request("page")
							if (iPageCurrent="") then
								iPageCurrent=1
							end if
							if not validNum(iPageCurrent) then
								iPageCurrent=1
							end if
							call opendb()%>
							<!--#include file="prv_incshow.asp"-->
							<% If iPageCount>1 then
							iPageCurrent=clng(iPageCurrent) %>
							<!-- If Page count is more then 1 show page navigation -->
									<p><b> 
									<% If iPageCurrent > 1 Then %>
										<a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>&page=<%=iPageCurrent -1 %>"><img src="<%=rsIconObj("previousicon")%>"></a> 
									<% End If
									For I=1 To iPageCount
										If I=iPageCurrent Then %>
											<%= I %> 
										<% Else %>
											<a class=privacy href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>&page=<%=I%>"><%=I%></a> 
										<% End If 
									Next
								
									If iPageCurrent < iPageCount Then %>
										<a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>&page=<%=iPageCurrent + 1%>"><img src="<%=rsIconObj("nexticon")%>"></a> 
									<% End If %>
									</b>
									</p>
							<!-- end of page navigation -->
							<% end if%>
							</td>
						</tr>
					</table>
			</td>
		</tr>
		<tr>
			<td>
				<%if RCount>="5" then%>
					<p align="right"><a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a></p>
			</td>
		</tr>
		<tr>
			<td><hr></td>
		</tr>
		<tr>
			<td>
			<form name="rating1" class="pcForms">
			<%if pcv_RatingType="0" then%>
				<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="feel" type="hidden" value=""><img src="catalog/<%=pcv_Img1%>" align="absbottom"><input name="feel1" value="2" type="radio" onclick="document.rating1.feel.value='2';" class="clearBorder"> <%=pcv_MainRateTxt2%>  <img src="catalog/<%=pcv_Img2%>" align="absbottom"><input name="feel1" value="1" type="radio" onclick="document.rating1.feel.value='1';" class="clearBorder"> <%=pcv_MainRateTxt3%> <input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&feel=' + document.rating1.feel.value);" class="submit2"></p>
			<%else%>
				<p><b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="rate" type="hidden" value=""><%if pcv_CalMain="1" then%><%for k=1 to pcv_MaxRating%><input name="rate1" value="<%=k%>" type="radio" onclick="document.rating1.rate.value='<%=k%>';" class="clearBorder"><span class="pcSmallText"><%=k%></span>&nbsp;<%next%><%end if%> <input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&rate=' + document.rating1.rate.value);" class="submit2"></p>
		<%end if%>
			</form>
			</td>
		</tr>
		<%end if%>
		<tr>
			<td>
			<p>&nbsp;</p>
			<p>
			<form class="pcForms">
			<input type="button" name="back" value="<%=dictLanguage.Item(Session("language")&"_prv_30")%>" onclick="location='viewPrd.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>';">
			</form>
			</p>
		</td>
	</tr>
	</table>
</div>
	<%
	call closeDb()
	END IF
END IF%>
<!--#include file="footer.asp"-->