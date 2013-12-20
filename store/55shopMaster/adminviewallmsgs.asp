<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!-- #Include file="../pc/checkdate.asp" -->
<% Dim pageTitle, Section
pageTitle="Manage Help Desk - View All Messages"
pageIcon="pcv4_icon_helpDesk.png"
Section="orders" %>
<%

Dim rstemp, connTemp, query, iPageCurrent 

if request("iPageCurrent")<>"" then
	iPageCurrent=Request("iPageCurrent")
	session("PHiPageCurrent")=iPageCurrent
else	
	if (session("PHiPageCurrent")<>"") and (request("Order")="") and (request("sort")="") then
		iPageCurrent=session("PHiPageCurrent")
	else
	    iPageCurrent=1 
	    session("PHiPageCurrent")=iPageCurrent
	end if
end If

%>
<!-- #Include File="Adminheader.asp" -->
<%
call openDb() 

'// Loading order-specific messages
Dim lngIDOrder
lngIDOrder=getUserInput(request("IDOrder"),0)
tmpIDOrder="0"
if validNum(lngIDOrder) then
	tmpIDOrder=lngIDOrder
	session("admin_IDOrder")=lngIDOrder
	query="SELECT idorder FROM Orders WHERE IDOrder=" & lngIDOrder
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	IF rstemp.eof Then
		set rstemp = nothing
		call closedb()
		response.Redirect "msgb.asp?message="& Server.Urlencode(dictLanguage.Item(Session("language")&"_viewPostings_a"))
	End If
	query1 = " AND pcComm_IDOrder=" & lngIDOrder
	addQueryString = "&idOrder=" & lngIDOrder
else
	lngIDOrder=""
end if

'// Loading customer-specific messages
Dim pcIntCustomerId
pcIntCustomerId=getUserInput(request("idCustomer"),0)
if validNum(pcIntCustomerId) then
	query4 = " AND pcComm_IDOrder IN (SELECT orders.idorder FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND customers.idcustomer="& pcIntCustomerId &")"
	addQueryString = "&idCustomer=" & pcIntCustomerId
else
	pcIntCustomerId=""
	query4=""
end if

dim A(30,2),Count,FCount,k

query="Select pcFStat_IDStatus,pcFStat_Name from pcFStatus"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)

Count=0
do while not rstemp.eof
	Count=Count+1
	A(Count-1,0)=rstemp("pcFStat_IDStatus")
	A(Count-1,1)=rstemp("pcFStat_Name")
	rstemp.movenext
loop

redim B(Count-1)

query="SELECT pcComm_FStatus FROM pcComments WHERE pcComm_IDParent=0" & query1
set rstemp=connTemp.execute(query)

FCount=0
do while not rstemp.eof
	FCount=FCount+1
	For k=0 to Count-1
		if cint(rstemp("pcComm_FStatus"))=cint(A(k,0)) then
			B(k)=B(k)+1
		end if
	Next
	rstemp.Movenext
loop

set rstemp = nothing

%>
<div style="padding: 8px;">A list of the Help Desk messages received on all orders within the specified date range.</div>
<table class="pcCPcontent">
	<tr>
		<th nowrap valign="middle"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_n")%><a href="adminviewallmsgs.asp?Order=pcComm_IDFeedback&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_IDFeedback&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
		<th nowrap valign="middle"><img src="images/pcv3_infoIcon.gif" alt="Order" title="Order Number"><a href="adminviewallmsgs.asp?Order=pcComm_IDOrder&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_IDOrder&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
		<th nowrap><img src="images/pcv3_infoIcon.gif" alt="Priority" title="Message Priority"><a href="adminviewallmsgs.asp?Order=pcComm_Priority&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_Priority&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
		<th nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_p")%></th>
		<th nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_q")%><a href="adminviewallmsgs.asp?Order=pcComm_FType&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_FType&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
		<th nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_r")%><a href="adminviewallmsgs.asp?Order=pcComm_CreatedDate&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_CreatedDate&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
		<th nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_s")%><a href="adminviewallmsgs.asp?Order=pcComm_EditedDate&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_EditedDate&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
		<th nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_t")%><a href="adminviewallmsgs.asp?Order=pcComm_IDUser&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_IDUser&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
		<th nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_u")%><a href="adminviewallmsgs.asp?Order=pcComm_FStatus&Sort=Desc<%=addQueryString%>"><img src="images/sortdesc.gif" width="14" height="14" alt="Sort Descending" border="0"></a><a href="adminviewallmsgs.asp?Order=pcComm_FStatus&Sort=Asc<%=addQueryString%>"><img src="images/sortasc.gif" width="14" height="14" alt="Sort Ascending" border="0"></a></th>
	</tr>
	<tr class="main">
		<td colspan="7" class="pcCPspacer"></td>
	</tr>
<%

Dim SOrder,SSort,APageCount,strsortOrder

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

APageCount=request("APageCount")
FBPerpage=request("FBPerpage")

if FBPerpage<>"" then
session("adminFBperPage")=FBPerpage
iPageCurrent=0
else
if (session("adminFBperPage")<>0) and (request("Type")="") then
FBperPage=session("adminFBperPage")
else
FBperPage=50
session("adminFBperPage")=FBperPage
iPageCurrent=0
end if
end if

strsortOrder=" ORDER BY pcComments." & SOrder & " " & SSort

FromDate=request("fromdate")
ToDate=request("todate")

if (FromDate<>"") and (IsDate(FromDate)) then
	session("adminFBFromDate")=FromDate
else
	if (session("adminFBFromDate")<>"") and (request("Type")="") then
	FromDate=session("adminFBFromDate")
	else
	session("adminFBFromDate")=""
	end if
end if

if (ToDate<>"") and (IsDate(ToDate)) then
	session("adminFBToDate")=ToDate
else
	if (session("adminFBToDate")<>"") and (request("Type")="") then
	ToDate=session("adminFBToDate")
	else
	session("adminFBToDate")=""
	end if
end if

if (FromDate<>"") and (IsDate(FromDate)) then
	if SQL_Format="1" then
		FromDate=Day(FromDate)&"/"&Month(FromDate)&"/"&Year(FromDate)
	else
		FromDate=Month(FromDate)&"/"&Day(FromDate)&"/"&Year(FromDate)
	end if
	if scDB="Access" then
		query1= " AND pcComments.pcComm_CreatedDate>=#" & FromDate & "#"
	else
		query1= " AND pcComments.pcComm_CreatedDate>='" & FromDate & "'"
	end if
end if

if (ToDate<>"") and (IsDate(ToDate)) then
	ToDate1=ToDate
	if scDB="Access" then
		query2= " AND pcComments.pcComm_EditedDate<=#" & ToDate1 & "#"
	else
		query2= " AND pcComments.pcComm_EditedDate<='" & ToDate1 & "'"
	end if
end if

if validNum(lngIDOrder) then
	query3 = " AND pcComm_IDOrder=" & lngIDOrder
end if

query="SELECT * FROM pcComments WHERE pcComments.pcComm_IDParent=0 " & query1 & query2 & query3 & query4 & strsortOrder
Set rstemp=Server.CreateObject("ADODB.Recordset")

Dim iPageCount,lngIDfeedback,lngIDUser,dtcreatedDate,dteditedDate,intFType,intFStatus,intPriority,strFDesc

if APageCount<>"" then
	rstemp.CacheSize=APageCount
	rstemp.PageSize=APageCount
else
	rstemp.CacheSize=FBperPage
	rstemp.PageSize=FBperPage
end if

rstemp.Open query, connTemp, 3, 1

if rstemp.eof then
%>
	<tr>
        <td colspan="7">
            <div class="pcCPmessage"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_v")%></div>
        </td>
	</tr>
<%else
	rstemp.MoveFirst
	iPageCount=rstemp.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
	rstemp.AbsolutePage=iPageCurrent
	Count=0
	DO While not rstemp.eof and Count < rstemp.PageSize
	lngIDOrder=rstemp("pcComm_IDOrder")
	lngIDfeedback=rstemp("pcComm_idfeedback")
	lngIDUser=rstemp("pcComm_iduser")
	dtcreatedDate=rstemp("pcComm_createdDate")
	dteditedDate=rstemp("pcComm_editedDate")
	intFType=rstemp("pcComm_FType")
	intFStatus=rstemp("pcComm_FStatus")
	intPriority=rstemp("pcComm_Priority")
	strFDesc=rstemp("pcComm_Description")
	
	Dim rstemp1,strFBgColor,intshowbgcolor

	query="Select * from pcFStatus where pcFStat_IDStatus=" & intFStatus
	set retemp1=Server.CreateObject("ADODB.Recordset")
    set rstemp1=connTemp.execute(query)
    FBgColor=""
    if not rstemp1.eof then
    strFBgColor=rstemp1("pcFStat_BgColor")
    intshowbgcolor=1
    end if
    %>    
	<tr class="main" <%if intshowbgcolor="1" then
	if strFBgColor<>"" then%>bgcolor="<%=strFBgColor%>"<%end if
	end if%>>
  	<td style="border-bottom: 1px solid #FFF;" nowrap><a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><%=lngIDfeedback%></a></td>
	<td style="border-bottom: 1px solid #FFF;" nowrap><a href="orddetails.asp?id=<%=lngIDOrder%>"><%=clng(scpre)+clng(lngIDOrder)%></a></td>    
    <td style="border-bottom: 1px solid #FFF;">
    <%
	Dim strPName,strPImg,intPriorityImage
	query="Select * from pcPriority where pcPri_IDPri=" & intPriority
	set rstemp1=connTemp.execute(query)
	if not rstemp1.eof then
		strPName=rstemp1("pcPri_Name")
		strPImg=rstemp1("pcPri_Img")
		intPriorityImage=rstemp1("pcPri_ShowImg")
		if intPriorityImage="1" then
			if strPImg<>"" then%>
				<img src="../pc/images/<%=strPImg%>" alt="<%=strPName%>" border="0">
			<%end if
		else%>
			<%=strPName%>
		<%end if
	end if
	set rstemp1=nothing
	%>
    </td>
    <td style="border-bottom: 1px solid #FFF;"><a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><%=strFDesc%></a></td>
    <td style="border-bottom: 1px solid #FFF;">
    <%
	Dim intTypeImage

	query="Select * from pcFTypes where pcFType_IDType=" & intFType
	set rstemp1=connTemp.execute(query)
	if not rstemp1.eof then
		strPName=rstemp1("pcFType_Name")
		strPImg=rstemp1("pcFType_Img")
		intTypeImage=rstemp1("pcFType_ShowImg")
		if intTypeImage="1" then
			if strPImg<>"" then%>
				<img src="../pc/images/<%=strPImg%>" alt="<%=strPName%>" border="0">
			<%end if
		else%>
			<%=strPName%>
		<%end if
	end if%>
    </td>
    <td style="border-bottom: 1px solid #FFF; width: 10%;"><%=CheckDate(dtcreatedDate)%></td>
    <td style="border-bottom: 1px solid #FFF; width: 10%;"><%=CheckDate(dteditedDate)%></td>
    <td style="border-bottom: 1px solid #FFF;">
    <%
    if validNum(lngIDUser) and lngIDUser<>0 then
		query="Select email,name,lastname from Customers where IDCustomer=" & lngIDUser
		set rstemp1=connTemp.execute(query)
		if not rstemp1.eof then%>
			<a href="modcusta.asp?idcustomer=<%=lngIDUser%>" target="_blank"><%=rstemp1("Name") & " " & rstemp1("LastName")%></a>
		<%else%>
		Customer has been deleted
		<%		
		end if
		else
		%>
			<%response.write dictLanguage.Item(Session("language")&"_viewPostings_2")%>
		<%end if%>
	</td>
    <td style="border-bottom: 1px solid #FFF;">
<%

	Dim intStatusImage

	query="Select * from pcFStatus where pcFStat_IDStatus=" & intFStatus
	set rstemp1=connTemp.execute(query)
	if not rstemp1.eof then
		strPName=rstemp1("pcFStat_Name")
		strPImg=rstemp1("pcFStat_Img")
		intStatusImage=rstemp1("pcFStat_ShowImg")
		if intStatusImage="1" then
	    	if strPImg<>"" then%>
		    	<a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><img src="../pc/images/<%=strPImg%>" alt="<%=strPName%>" border="0"></a>
	    	<%end if
    	else%>
			<a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><%=strPName%></a>
		<%end if
	end if
	set retemp1=nothing
	%>
    </td>
	</tr>
	<%
	Count=Count+1
	rstemp.MoveNext
	Loop
	set rstemp=nothing
end if
call closeDb()
%>
</table>
<%
If iPageCount>1 Then
%>
<div style="margin-top: 20px; margin-left: 5px;">
	<%response.write dictLanguage.Item(Session("language")&"_viewPostings_x")%>
	<%' display Next / Prev links
	For I=1 To iPageCount
		If Cint(I)=Cint(iPageCurrent) Then %>
			<b><%=I%></b>
		<% Else %>
			<a href="adminviewallmsgs.asp?iPageCurrent=<%=I%>&order=<%=SOrder%>&sort=<%=SSort%>"><%=I%></a> 
		<% End If %>
	<% Next
	if APageCount<>"" then
	else %>
	&nbsp;|&nbsp;<a href="adminviewallmsgs.asp?order=<%=SOrder%>&sort=<%=SSort%>&APageCount=<%=FBperPage*(iPageCount+1)%>"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_z")%></a>
	<%
	end if
	%>
</div>
<%
End If
%>

<div style="margin-top: 20px;">
    <form method="post" name="filter" action="adminviewallmsgs.asp" class="pcForms">
    <input type="hidden" name="order" value="<%=SOrder%>">
    <input type="hidden" name="sort" value="<%=SSort%>">
    <input type="hidden" name="Type" value="1">
    <table class="pcCPcontent">
        <tr>
            <th colspan="2">Update Date Range</th>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>        
        <tr>
            <td width="10%">Date:</td>
            <td width="90%">From: <input type=text name="FromDate" size="10" value="<%=FromDate%>"> To: <input type=text name="ToDate" size="10" value="<%=ToDate%>"></td>
        </tr>
        <tr>
            <td nowrap>Messages shown:</td>
            <td><input type=text name="FBPerpage" size="5" value="<%=FBPerpage%>"> per page</td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"><hr></td>
        </tr>     
        <tr>
            <td colspan="2">
            <input type="submit" name="submit" value="Update" class="submit2">&nbsp;
            <input type="button" value="Write New Message" onClick="location.href='adminaddfeedback.asp?type=1&IDOrder=<%=tmpIDOrder%>'">&nbsp;
            <input type="button" value="Manage Help Desk" onClick="location.href='adminFBsettings.asp'">
            </td>
        </tr>
    </table>
    </form>
</div>
<!-- #Include File="Adminfooter.asp" -->