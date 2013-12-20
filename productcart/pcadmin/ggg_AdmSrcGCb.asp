<% 'GGG Add-on Only File %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin="1*2*3*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<% pageTitle = "Generated Gift Certificates - Search Results" %>
<% section = "products" %>
<!--#include file="Adminheader.asp"--> 
<% dim iPageSize
dim connTemp
dim iPageCurrent

IF (request("action")="search") then
	session("adm_gen_gcs_viewall")=""
	submit2=""
Else
	if request("submit2")<>"" then
	submit2=request("submit2")
	session("adm_gen_gcs_viewall")=submit2
	else
	submit2=session("adm_gen_gcs_viewall")
	end if
End if

IF (request("action")<>"search") and (submit2="") then

	if request.queryString("iPageSize")<>"" then
		iPageSize=server.HTMLEncode(request.querystring("iPageSize"))
		session("adm_gen_gcs_iPageSize")=iPageSize
	else
		iPageSize=session("adm_gen_gcs_iPageSize")
	end if
	
	if request.queryString("iPageCurrent")<>"" then
		iPageCurrent=server.HTMLEncode(request.querystring("iPageCurrent"))
		session("adm_gen_gcs_iPageCurrent")=iPageCurrent
	else
		iPageCurrent=session("adm_gen_gcs_iPageCurrent")
	end if

	pIDProduct=session("adm_gen_gcs_IDProduct")
	pGiftCode=session("adm_gen_gcs_GiftCode")
	pExpDate=session("adm_gen_gcs_ExpDate")

ELSE

	iPageSize=getUserInput(request("resultCnt"),10)
	if iPageSize="" then
		iPageSize=request("iPageSize")
	end if
	if request("iPageCurrent")="" then
		iPageCurrent=1 
	else
		iPageCurrent=server.HTMLEncode(request("iPageCurrent"))
	end if
	
	pIDProduct=request("IDGC")
	session("adm_gen_gcs_IDProduct")=pIDProduct
	pGiftCode=request("GCCode")
	session("adm_gen_gcs_GiftCode")=pGiftCode
	pExpDate=request("ExpDate")
	session("adm_gen_gcs_ExpDate")=pExpDate
	
	session("adm_gen_gcs_iPageSize")=iPageSize
	session("adm_gen_gcs_iPageCurrent")=iPageCurrent
END IF

call opendb()

Dim DefaultBoolean
DefaultBoolean="OR"

' create sql statement
Dim query

query="Select products.IDProduct,products.Description,pcGCOrdered.pcGO_IDOrder,pcGCOrdered.pcGO_GcCode,pcGCOrdered.pcGO_ExpDate,pcGCOrdered.pcGO_Amount,pcGCOrdered.pcGO_Status from Products,pcGCOrdered where pcGCOrdered.pcGO_IDProduct=products.IDProduct and pcGCOrdered.pcGO_idOrder=0 and products.pcprod_GC=1 and pcGCOrdered.pcGO_Status<=1"

IF submit2="" then 'not View All

	if (pIdProduct<>"0") and (pIdProduct<>"") then
		query=query & " AND Products.IDProduct=" & pIdProduct
	end if

	if (pGiftCode<>"") then
		query=query & " AND pcGCOrdered.pcGO_GcCode like '%" & pGiftCode & "%'"
	end if

	if (pExpDate<>"") then
		if SQL_Format="1" then
			pExpDate=(day(pExpDate)&"/"&month(pExpDate)&"/"&year(pExpDate))
		else
			pExpDate=(month(pExpDate)&"/"&day(pExpDate)&"/"&year(pExpDate))
		end if
		if scDB="Access" then
			query=query & " AND pcGCOrdered.pcGO_ExpDate=#" & pExpDate & "#"
		else
			query=query & " AND pcGCOrdered.pcGO_ExpDate='" & pExpDate & "'"
		end if
	end if

END IF 'Not View All

query=query & " ORDER by pcGCOrdered.pcGO_ExpDate DESC"
Set rsTemp=Server.CreateObject("ADODB.Recordset")     

if submit2<>"" then
	iPageSize=25
end if

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize

'response.end
rsTemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
  	response.redirect "techErr.asp?error="&Server.UrlEncode("Error in page advSrcb. Error: "&err.description)
end If

if not rsTemp.eof then 
	dim iPageCount
	iPageCount=rstemp.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then Cint(iPageCurrent)=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
	rstemp.AbsolutePage=iPageCurrent
end if	

%>
<table class="pcCPcontent">
<tr> 
	<td colspan="6" class="pcCPspacer"></td>
</tr>
<%IF rstemp.eof THEN %>
<tr> 
	<td colspan="6">
		<div class="pcCPmessage">No Generated Gift Certificate Found</div>
	</td>
</tr>
<%ELSE%>
	<tr>
		<th nowrap>Product Name</th>
		<th nowrap>GC Code</th>
		<th nowrap>Expiring on</th>
		<th nowrap>Available</th>
		<th nowrap colspan="2">Status</th>
	</tr>
	<tr> 
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<%
	Dim Count
	Count = 0	
	do while (not rsTemp.eof) and (count < rsTemp.pageSize)
								
		gIDProduct=rstemp("idProduct")
		gName=rstemp("Description")
		gIDOrder=rstemp("pcGO_IDOrder")
		gCode=rstemp("pcGO_GcCode")
		gExpDate=rstemp("pcGO_ExpDate")
		if year(gExpDate)="1900" then
			gExpDate=""
		end if
		if gExpDate<>"" then
			if scDateFrmt="DD/MM/YY" then
				gExpDate=(day(gExpDate)&"/"&month(gExpDate)&"/"&year(gExpDate))
			else
				gExpDate=(month(gExpDate)&"/"&day(gExpDate)&"/"&year(gExpDate))
			end if
		end if
		gAmount=rstemp("pcGO_Amount")
		gStatus=rstemp("pcGO_Status")
		%>
				 
		<!-- start of display -->
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td><%= gName %></td>
			<td><a href="ggg_modGc.asp?idproduct=<%=gIDProduct%>&GcCode=<%=gCode%>&gen=1"><%= gCode %></a></td>
			<td><%= gExpDate %></td>
			<td><%= scCurSign & money(gAmount)%></td>
			<td class="cpLinksList"><%if gStatus="1" then%>Active<%else%>Inactive<%end if%></td>
			<td nowrap align="right"><a href="ggg_modGc.asp?idproduct=<%=gIDProduct%>&GcCode=<%=gCode%>&gen=1">Edit</a>&nbsp;|&nbsp;<a href="ggg_AdmSendGCs.asp?GcCode=<%=gCode%>">Send</a></td>
		</tr>
		<%
		count=count + 1
		rsTemp.MoveNext
	loop
END IF
set rsTemp=nothing
call closeDb()
%>
</table>
<br>
<!-- end of display -->
<table class="pcCPcontent">
<tr>
	<td>
		<% If iPageCount>1 Then %>
            <%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & "<br>")%>
			<p class="pcPageNav">
				<%if iPageCurrent > 1 then %>
					<a href="ggg_AdmSrcGCb.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=iPageCurrent - 1%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a>
				<% end If
				For I = 1 To iPageCount
					If Cint(I) = Cint(iPageCurrent) Then %>
						<b><%=I%></b>
					<% Else %>
						<a href="ggg_AdmSrcGCb.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=I%>"><%=I%></a>
					<% End If %>
				<%Next %>
				<%if CInt(iPageCurrent) < CInt(iPageCount) then %>
					<a href="ggg_AdmSrcGCb.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=iPageCurrent + 1%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
				<%end If %>
			</p>
		<% End If %>
	</td>
</tr>
</table>
<table class="pcCPcontent">
<tr>
	<td align="center" valign="top">
		<form class="pcForms">
			<input TYPE="button" VALUE="New Search" onClick="location.href='ggg_AdmManageGCs.asp'">
		<%if submit2="" then%>
			&nbsp;<input TYPE="button" VALUE="View All" onClick="location.href='ggg_AdmSrcGCb.asp?iPageSize=99999&submit2=viewall'">
		<%end if%>
		</form>
	</td>
</tr>
</table>
<!--#include file="Adminfooter.asp"-->