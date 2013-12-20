<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!-- #Include File="checkdate.asp" -->
<%
Dim rstemp, connTemp, mySQL
Dim iPageCurrent 

iPageCurrent=getUserInput(Request("iPageCurrent"),10)
if validNum(iPageCurrent) then
	session("PHiPageCurrent")=iPageCurrent
else	
	if validNum(session("PHiPageCurrent")) and (request("Order")="") and (request("sort")="") then
		iPageCurrent=session("PHiPageCurrent")
	else
		iPageCurrent=1 
		session("PHiPageCurrent")=iPageCurrent
	end if
end If

%>
<!-- #Include File="header.asp" -->
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<h1><%response.write dictLanguage.Item(Session("language")&"_viewPostings_3")%></h1>
			</td>
		</tr>
		<tr>
			<td>
					<% Dim lngIDOrder,TempStr
					lngIDOrder=getUserInput(request("IDOrder"),50)
					session("IDOrder")=lngIDOrder
					lngIDOrder=Clng(lngIDOrder)-clng(scpre)
			
					call openDb()
			
					TempStr=" and IDCustomer=" & session("IDCustomer")
					MySQL="Select IDCustomer from Orders where IDOrder=" & lngIDOrder & TempStr
					set rstemp=connTemp.execute(mySQL)
					' Check to see if the order exists and the customer has permissions to view it
					IF rstemp.eof then
					' Order-related information cannot be accessed
					%>
						<div class="pcErrorMessage">
							<%response.write dictLanguage.Item(Session("language")&"_viewPostings_a")%>
						</div>
					<%
					ELSE
					' Customer can access order-related information

						dim A(30,2),Count
			
						MySQL="Select pcFStat_IDStatus,pcFStat_Name from pcFStatus"
						set rstemp=connTemp.execute(mySQL)
			
						Count=0
						do while not rstemp.eof
							Count=Count+1
							A(Count-1,0)=rstemp("pcFStat_IDStatus")
							A(Count-1,1)=rstemp("pcFStat_Name")
							rstemp.movenext
						loop
						redim B(Count-1)

						MySQL="Select pcComm_FStatus from pcComments where pcComm_IDParent=0 and pcComm_IDOrder=" & lngIDOrder
						
						set rstemp5=connTemp.execute(mySQL)
						FCount=0
						do while not rstemp5.eof
						FCount=FCount+1
						For k=0 to Count-1
						if cint(rstemp5("pcComm_FStatus"))=cint(A(k,0)) then
						B(k)=B(k)+1
						end if
						Next
						rstemp5.Movenext
						loop
					%>

						<p><%response.write dictLanguage.Item(Session("language")&"_viewPostings_b")%><a href="CustviewPastD.asp?idOrder=<%=(scpre+lngIDOrder)%>"><strong><%=clng(scpre)+clng(lngIDOrder)%></strong></a></p>
						<p><%response.write dictLanguage.Item(Session("language")&"_viewPostings_m")%></p>
						
						<table class="pcShowContent">
							<tr>
								<th nowrap>
								<%response.write dictLanguage.Item(Session("language")&"_viewPostings_o")%><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_Priority&Sort=Desc"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_Priority&Sort=Asc"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a>
								</th>
								<th nowrap>
								<%response.write dictLanguage.Item(Session("language")&"_viewPostings_p")%><%if scDB="Access" then%><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_Description&Sort=Desc"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_Description&Sort=Asc"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a><%end if%>
								</th>
								<th nowrap>
								<%response.write dictLanguage.Item(Session("language")&"_viewPostings_q")%><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_FType&Sort=Desc"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_FType&Sort=Asc"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a>
								</th>
								<th nowrap>
								<%response.write dictLanguage.Item(Session("language")&"_viewPostings_s")%><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_EditedDate&Sort=Desc"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_EditedDate&Sort=Asc"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a>
								</th>
								<th nowrap>
								<%response.write dictLanguage.Item(Session("language")&"_viewPostings_t")%><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_IDUser&Sort=Desc"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_IDUser&Sort=Asc"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a>
								</th>
								<th nowrap>
								<%response.write dictLanguage.Item(Session("language")&"_viewPostings_u")%><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_FStatus&Sort=Desc"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&Order=pcComm_FStatus&Sort=Asc"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a>
								</th>
							</tr>
							<tr>
								<td class="pcSpacer" colspan="6"></td>
							</tr>
							<%
							
							Dim SOrder,SSort,APageCount,strsortOrder,iPageCount
							
							if request("order")<>"" then
								SOrder=getUserInput(request("order"),0)
								session("PHorder")=SOrder
							else
								if session("PHorder")<>"" then
									SOrder=session("PHorder")
								else
									SOrder="pcComm_EditedDate"
									session("PHorder")=SOrder
								end if	
							end if
							
							if request("sort")<>"" then
								SSort=getUserInput(request("sort"),0)
								session("PHsort")=SSort
							else
								if session("PHsort")<>"" then
									SSort=session("PHsort")
								else	
									SSort="Desc"
									session("PHsort")=SSort
								end if	
							end if
							
							APageCount=getUserInput(request("APageCount"),50)
							if not validNum(APageCount) then
								APageCount=10
							end if
							
							strsortOrder=" Order by " & SOrder & " " & SSort
							
							MySQL="Select * from pcComments where pcComm_IDOrder=" & lngIDOrder & " and pcComm_IDParent=0" & strsortOrder
							Set rstemp=Server.CreateObject("ADODB.Recordset")
							
							if APageCount<>"" then
								rstemp.CacheSize=APageCount
								rstemp.PageSize=APageCount
							end if
							
							rstemp.Open mySQL, connTemp, 3, 1
							
							IF rstemp.eof THEN
							%>
							<tr>
								<td colspan="6">
								<p><%response.write dictLanguage.Item(Session("language")&"_viewPostings_v")%></p>
								</td>
							</tr>
							<%
							ELSE
							rstemp.MoveFirst
							iPageCount=rstemp.PageCount
							If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
							If iPageCurrent < 1 Then iPageCurrent=1
							rstemp.AbsolutePage=iPageCurrent
							
							Count=0
							
							DO While not rstemp.eof and Count < rstemp.PageSize
							
							Dim lngIDfeedback,lngIDUser,dtcreatedDate,dteditedDate,intFType,intFStatus,intPriority,strFDesc
						
							lngIDfeedback=rstemp("pcComm_idfeedback")
							lngIDUser=rstemp("pcComm_iduser")
							dtcreatedDate=rstemp("pcComm_createdDate")
							dteditedDate=rstemp("pcComm_editedDate")
							intFType=rstemp("pcComm_FType")
							intFStatus=rstemp("pcComm_FStatus")
							intPriority=rstemp("pcComm_Priority")
							strFDesc=rstemp("pcComm_Description")
						
							Dim rstemp1,strFBgColor,intshowbgcolor
							
							MySQL="Select * from pcFStatus where pcFStat_IDStatus=" & intFStatus
								set rstemp1=connTemp.execute(mySQL)
						
								strFBgColor=""
								if not rstemp1.eof then
								strFBgColor=rstemp1("pcFStat_BgColor")
								intshowbgcolor=1
								end if
								%>    
								<tr <%if intshowbgcolor="1" then
								if strfbgcolor<>"" then%>bgcolor="<%=strFBgColor%>"<%end if
								end if%>>
									<td>
										<p align="center">
									<%
									Dim strPName,strPImg,intPriorityImage
									
									MySQL="Select * from pcPriority where pcPri_IDPri=" & intPriority
									set rstemp1=connTemp.execute(mySQL)
									if not rstemp1.eof then
									strPName=rstemp1("pcPri_Name")
									strPImg=rstemp1("pcPri_Img")
									intPriorityImage=rstemp1("pcPri_ShowImg")
										if intPriorityImage="1" then
										if strPImg<>"" then%>
										<img src="images/<%=strPImg%>" alt="<%=strPName%>" border="0">
										<%end if
										else%>
										<%=strPName%>
										<%end if
									end if%>
									</p>
									</td>
									<td>
									<a href="userviewfeedback.asp?IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&IDFeedback=<%=lngIDfeedback%>"><%=strFDesc%></a>
									</td>
									<td>
									<%
									
									Dim intTypeImage
									
									MySQL="Select * from pcFTypes where pcFType_IDType=" & intFType
									set rstemp1=connTemp.execute(mySQL)
									if not rstemp1.eof then
									strPName=rstemp1("pcFType_Name")
									strPImg=rstemp1("pcFType_Img")
									intTypeImage=rstemp1("pcFType_ShowImg")
										if intTypeImage="1" then
										if strPImg<>"" then%>
										<img src="images/<%=strPImg%>" alt="<%=strPName%>" border="0">
										<%end if
										else%>
										<%=strPName%>
										<%end if
									end if%>
									</td>
									<td><%=CheckDate(dteditedDate)%></td>
									<td>
									<%
									if (lngIDUser<>"") and (lngIDUser<>"0") then
									MySQL="Select email,name,lastname from Customers where IDCustomer=" & lngIDUser
									set rstemp1=connTemp.execute(mySQL)
									if not rstemp1.eof then%>
									<%=rstemp1("Name") & " " & rstemp1("LastName")%>
									<%end if
									else	%>
									<%response.write dictLanguage.Item(Session("language")&"_viewPostings_2")%>
									<%end if%>
									</td>
									<td>
									<%
									
									Dim intStatusImage
									
									MySQL="Select * from pcFStatus where pcFStat_IDStatus=" & intFStatus
									set rstemp1=connTemp.execute(mySQL)
									if not rstemp1.eof then
									strPName=rstemp1("pcFStat_Name")
									strPImg=rstemp1("pcFStat_Img")
									intStatusImage=rstemp1("pcFStat_ShowImg")
										if intStatusImage="1" then
										if strPImg<>"" then%>
										<img src="images/<%=strPImg%>" alt="<%=strPName%>" border="0">
										<%end if
										else%>
										<%=strPName%>
										<%end if
									end if%>
									</td>
								</tr>
								<%
								Count=Count+1
								rstemp.MoveNext
							
								LOOP
								
							END IF%>
							</table>
							<%
							If iPageCount>1 Then %>
							<br><br><%response.write dictLanguage.Item(Session("language")&"_viewPostings_x")%> 
							<%' display Next / Prev links
							For I=1 To iPageCount
							If Cint(I)=Cint(iPageCurrent) Then %>
							<b><%=I%></b>
							<% Else %>
							<a href="userviewallposts.asp?iPageCurrent=<%=I%>&order=<%=SOrder%>&sort=<%=SSort%>&IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>"><%=I%></a> 
							<% End If %>
							<% Next
							if APageCount<>"" then
							else %>
							&nbsp;|&nbsp;
							<a href="userviewallposts.asp?order=<%=SOrder%>&sort=<%=SSort%>&IDOrder=<%=clng(scpre)+clng(lngIDOrder)%>&APageCount=<%=FBperPage*(iPageCount+1)%>"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_z")%><</a>
							<%end if
							End If %>
							<p>&nbsp;</p>
							<p align="center">
								<a href="useraddfeedback.asp"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_1")%></a> : <a href="CustViewPast.asp"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_4")%></a>
							</p>
							<%END IF%>
			</td>
		</tr>
	</table>
</div>
<%call closedb()%>
<!-- #Include File="footer.asp" -->