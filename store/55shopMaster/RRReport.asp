<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt" -->
<html>
<head>
<title>Product Reviews Notifications Report</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px;">
<div id="pcCPmain" style="background-image: none;">
<table class="pcCPcontent">
        
<%pcPageName="RRReport.asp"
Dim connTemp,rs, RMAStatusCond

' Choose the records to display
err.clear
Dim strTDateVar, strTDateVar2, DateVar, DateVar2
strTDateVar=Request("FromDate")
DateVar=strTDateVar
if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
end if
strTDateVar2=Request("ToDate")
DateVar2=strTDateVar2
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	if err.number<>0 then
		DateVar=Request("FromDate")
		DateVar2=Request("ToDate")
	end if
end if
err.clear

tmpD="pcReviewNotifications.pcRN_DateSent"

if (DateVar<>"") and IsDate(DateVar) then
    if SQL_Format="1" then DateVar=day(DateVar) & "/" & month(DateVar) & "/" & year(DateVar)
	if scDB="Access" then
		TempSQL1=tmpD & " >=#" & DateVar & "# "
	else
		TempSQL1=tmpD & " >='" & DateVar & "' "
	end if
else
	TempSQL1=""
end if

if (DateVar2<>"") and IsDate(DateVar2) then
    if SQL_Format="1" then DateVar2=day(DateVar2) & "/" & month(DateVar2) & "/" & year(DateVar2)
	if TempSQL1<>"" then
		TempSQL2=" AND "
	else
		TempSQL2=""
	end if
	if scDB="Access" then
		TempSQL2=TempSQL2 & tmpD & " <=#" & DateVar2 & "# "
	else
		TempSQL2=TempSQL2 & tmpD & " <='" & DateVar2 & "' "
	end if
else
	TempSQL2=""
end if

iPageCurrent=request("iPageCurrent")
if iPageCurrent="" then
	iPageCurrent=1
end if

call opendb()
query = "SELECT Customers.[name],Customers.lastName,Customers.[email],pcReviewNotifications.pcRN_idOrder,pcReviewNotifications.pcRN_UniqueID,pcReviewNotifications.pcRN_DateSent,pcReviewNotifications.pcRN_DateLastViewed FROM Customers INNER JOIN pcReviewNotifications ON Customers.idCustomer=pcReviewNotifications.pcRN_idCustomer WHERE " 
query = query & TempSQL1 
If NOT (DateVar=DateVar2) Then
	query = query & TempSQL2 
End If
query = query & " Order by " & tmpD & " DESC"
Set rs=Server.CreateObject("ADODB.Recordset")
iPageSize=50
rs.CacheSize=iPageSize
rs.PageSize=iPageSize
rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

' if there are no records in recordset
If rs.EOF Then %>
	<tr> 
		<td colspan="6">No records match your query</td>
	</tr>
<% Else 
	iPageCount=rs.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=1
	
	rs.AbsolutePage=Cint(iPageCurrent)
	pcArr = rs.getRows(rs.PageSize)
	set rs=nothing
	intCount = UBound(pcArr,2)
	%>
	<tr> 
		<td colspan="6"><h2>Product Reviews notifications were sent from <%=strTDateVar%> to <%=strTDateVar2%></h2>
	</tr>
	<tr valign="top"> 
		<th nowrap>Sent Date</th>
		<th nowrap>Sent To</th>
		<th nowrap>E-mail</th>
		<th nowrap>Order #</th>
		<th nowrap>Review Code</th>
		<th nowrap>Last Viewed<br>by Customer</th>
	</tr>
	<tr>
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<%
	For i=0 to intCount
		pcv_CustName=pcArr(0,i) & " " & pcArr(1,i)
		pcv_CustEmail="<a href='mailto:" & pcArr(2,i) & "'>" & pcArr(2,i) & "</a>"
		pcv_IdOrder="<a target='_blank' href='Orddetails.asp?id=" & pcArr(3,i) & "'>" & clng(pcArr(3,i))+scpre & "</a>"
		pcv_ReviewCode=pcArr(4,i)
		pcv_DateSent=pcArr(5,i)
		if scDateFrmt="DD/MM/YY" then
			pcv_DateSent=(day(pcv_DateSent)&"/"&month(pcv_DateSent)&"/"&year(pcv_DateSent))
		end if
		pcv_DateViewed=pcArr(6,i)
		if pcv_DateViewed<>"" then
		if scDateFrmt="DD/MM/YY" then
			pcv_DateViewed=(day(pcv_DateViewed)&"/"&month(pcv_DateViewed)&"/"&year(pcv_DateViewed))
		end if
		else
		pcv_DateViewed="Not Viewed Yet"
		end if
	%>
	<tr>  
		<td nowrap valign="top"><%=pcv_DateSent%></td>
		<td nowrap valign="top"><%=pcv_CustName%></td>
		<td nowrap valign="top"><%=pcv_CustEmail%></td>
		<td nowrap valign="top"><%=pcv_IdOrder%></td>
		<td nowrap valign="top"><%=pcv_ReviewCode%></td>
		<td nowrap valign="top"><%=pcv_DateViewed%></td>
	</tr>
	<%Next%>
	<tr>
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<%If iPageCount>1 Then%>                     
		<tr> 
			<td colspan="6" class="cpLinksList">
			<%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%>
            <br><br>
			<%' display Next / Prev buttons
			if iPageCurrent > 1 then %>
			<a href="<%=pcPageName%>?FromDate=<%=request("FromDate")%>&ToDate=<%=request("ToDate")%>&iPageCurrent=<%=iPageCurrent-1%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
			<%
			end If
			For I=1 To iPageCount
			If Cint(I)=Cint(iPageCurrent) Then %>
				<b><%=I%></b> 
			<%
			Else
			%>
				<a href="<%=pcPageName%>?FromDate=<%=request("FromDate")%>&ToDate=<%=request("ToDate")%>&iPageCurrent=<%=I%>"><%=I%></a> 
			<%
			End If
			Next
			if CInt(iPageCurrent) < CInt(iPageCount) then %>
					<a href="<%=pcPageName%>?FromDate=<%=request("FromDate")%>&ToDate=<%=request("ToDate")%>&iPageCurrent=<%=iPageCurrent+1%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
			<%
			end If
			%>
		</td>
		</tr>
	<%End If%>
	<tr>
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
<%END IF%>     
</table>
<%	' Done. Now release Objects
	Set rs=Nothing
	call closedb()
%>
</div>
</body>
</html>