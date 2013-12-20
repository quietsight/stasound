<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim rstemp, conntemp, query
call opendb()

'Start SDBA
pcv_pageType=request("pagetype")
'End SDBA
pageTitle="List Sent Messages"
section="mngAcc"
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>
        <th nowrap>#ID</th>
        <th nowrap>Message Title</th>
        <th nowrap>Date</th>
        <th nowrap>Recipients</th>
        <th nowrap></th>
    </tr>
    <tr>
		<td colspan="5" class="pcCPspacer"></td>
    </tr>
		<%query="select idnews, FromDate, fromname, fromemail, title, msgbody, msgtype, CustTotal, CustFile FROM News order by IDNews desc"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
  
		if rstemp.eof then %>
		<tr>
			<td colspan="5"><div class="pcCPmessage">No sent messages found.</div></td>
		</tr>
		<%else
		do while not rstemp.eof
			intIdNews=rstemp("idnews")
			dtFromDate=rstemp("FromDate")
			strFromName=rstemp("fromname")
			strFromEmail=rstemp("fromemail")
			strTitle=rstemp("title")
			strMsgBody=rstemp("msgbody")
			intMsgType=rstemp("msgtype")
			strCustTotal=rstemp("CustTotal")
			strCustFile=rstemp("CustFile")%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td><%=intIdNews%></td>
				<td width="80%"><%=strTitle%></td>
				<td><%=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)%></td>
				<td align="center"><%=strCustTotal%></td>
				<td nowrap class="cpLinksList"><a href="viewmsg.asp?pagetype=<%=pcv_pageType%>&idnews=<%=intIdNews%>">View Details</a> | <a href="germsg.asp?pagetype=<%=pcv_pageType%>&idnews=<%=intIdNews%>">Send to new list</a> | <a href="javascript:if (confirm('You are about to remove this message from your database. Are you sure you want to complete this action?')) location='delmsg.asp?idnews=<%=intIdNews%>'">Delete</a></td>
			</tr>
			<%rstemp.movenext
		loop
		end if
		set rstemp=nothing
		call closedb()
		%>

<tr>
	<td colspan="5" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="5">
		<input type="button" name="buttonnew" value="Start Newsletter Wizard" onClick="location='<%if pcv_pageType<>"" then%>sds_newsWizStep1.asp?pagetype=<%=pcv_pageType%><%else%>newsWizStep1.asp<%end if%>'" class="ibtnGrey">
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->