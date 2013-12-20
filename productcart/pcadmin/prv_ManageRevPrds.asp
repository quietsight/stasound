<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pcv_RevName=request("nav")
if pcv_RevName="" then
	pcv_RevName="2"
end if
if pcv_RevName="1" then
	pageTitle="Manage Pending Reviews"
else
	pageTitle="Manage Product Reviews"
end if
pageIcon="pcv4_icon_reviews.png"
section="reviews" 
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim rs, connTemp, query

call openDb()

query="SELECT pcRS_NeedCheck FROM pcRevSettings;"
set rs=server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

pcRS_NeedCheck=0
if not rs.eof then
	pcRS_NeedCheck=rs("pcRS_NeedCheck")
	if IsNull(pcRS_NeedCheck) or pcRS_NeedCheck="" then
		pcRS_NeedCheck=0
	end if
end if
set rs=nothing
	
query="SELECT pcReviews.pcRev_IDProduct,products.sku,products.description FROM pcReviews,products WHERE products.idproduct=pcReviews.pcRev_IDProduct group by pcReviews.pcRev_IDProduct,products.description,products.sku order by products.description asc"
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
<form method="POST" action="prv_ManageRevPrds.asp?action=update" name="checkboxform" class="pcForms">
	<table class="pcCPcontent">
		<tr>
            <td colspan="5"><%if pcv_RevName="1" then%>The following products have reviews that are awaiting approval<%else%>The following products have live reviews<%end if%>.</td>
		</tr>
		<tr>
            <td colspan="5" class="pcCPspacer"></td>
		</tr>
		<tr>
            <th width="55%" nowrap>Product Name</th>
            <th width="5%" nowrap>SKU</th>
            <th width="5%" nowrap><%if pcv_RevName="1" then%>Pending Reviews<%else%>Live Reviews<%end if%></th>
            <th width="5%" nowrap>Last Posted</th>
            <th width="30%">&nbsp;</th>
		</tr>
		<tr>
            <td colspan="5" class="pcCPspacer"></td>
		</tr>
        <% If DataEmpty=1 Then %>
			<tr> 
				<td colspan="5">
					<div class="pcCPmessage">
						No products have <%if pcv_RevName="1" then%>pending <% else %>live <%end if%>reviews.
					</div>
				</td>
			</tr>
		<% Else 
			Dim Count
			Count=0
			HaveCount=0
			For k=0 to intCount
			Count=Count+1
				pcv_ID=pcArray(0,k)
	
				if pcv_RevName="1" then
					query1=" and pcRev_Active=0 "
					query2=" and pcRev_Active=1 "
				else
					query1=" and pcRev_Active=1 "
					query2=" and pcRev_Active=0 "
				end if
	
				query="SELECT pcRev_IDReview,pcRev_Date FROM pcReviews WHERE pcRev_IDProduct=" & pcv_ID & query1 & " order by pcRev_Date desc"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
	
				pcv_RCount=0
				pcv_lastposted=""
	
				if not rs.eof then
					pcArray1=rs.getRows()
					pcv_RCount=ubound(pcArray1,2)+1
					pcv_lastposted=pcArray1(1,0)
				end if
				
				query="SELECT pcRev_IDReview,pcRev_Date FROM pcReviews WHERE pcRev_IDProduct=" & pcv_ID & query2 & " order by pcRev_Date desc"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
	
				pcv_LCount=0
				
				if not rs.eof then
					pcArray1=rs.getRows()
					pcv_LCount=ubound(pcArray1,2)+1
					if pcv_lastposted="" then
					pcv_lastposted=pcArray1(1,0)
					end if
				end if
	
				pcv_sku=pcArray(1,k)
				pcv_Name=pcArray(2,k)
				set rs=nothing
				
				if (pcv_RCount>"0") then
				HaveCount=1%>
				
				<% If scDateFrmt="DD/MM/YY" then 
                    pcv_lastposted = day(pcv_lastposted) & "/" & month(pcv_lastposted) & "/" & year(pcv_lastposted)
                Else
                    pcv_lastposted = month(pcv_lastposted) & "/" & day(pcv_lastposted) & "/" & year(pcv_lastposted)
                End If %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td valign="top"><%=pcv_Name%></td>
					<td valign="top" nowrap><%=pcv_sku%></td>
					<td valign="top"><%=pcv_RCount%> post(s)</td>
					<td valign="top" nowrap><%=pcv_lastposted%></td>
					<td valign="top" nowrap class="cpLinksList" align="right"><%if (pcv_RCount>"0") then%><a href="prv_ManageReviews.asp?IDProduct=<%=pcv_ID%>&nav=<%=pcv_RevName%>"><%if (pcv_RevName="1") then%>View/Edit/Approve/Delete<%else%>View/Edit/Delete<%end if%></a><br><%end if%>
					<%if (pcv_RCount="0") and (pcv_LCount>"0") then%><a href="prv_ManageReviews.asp?IDProduct=<%=pcv_ID%>&nav=<%if pcv_RevName="1" then%>0<%else%>1<%end if%>"><%if (pcv_RevName="1") then%>View Live Reviews<%else%>View Pending Reviews<%end if%></a><%end if%></td>
				</tr>
				<%end if%>
			<% Next
			if HaveCount=0 then%>
				<tr> 
					<td colspan="5">
						<div class="pcCPmessage">
							No products have <%if pcv_RevName="1" then%>pending <% else %>live <%end if%>reviews.
						</div>
					</td>
				</tr>
			<%
			end if
		End If
		%>
		<tr>
            <td colspan="5" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td colspan="5">
		<%if pcv_RevName="1" then%>
		<input type="button" value="View Live Reviews" onClick="location.href='prv_ManageRevPrds.asp?nav=2'" class="submit2">
		&nbsp;
		<% else %>
		<input type="button" value="View Pending Reviews" onClick="location.href='prv_ManageRevPrds.asp?nav=1'" class="submit2">
		&nbsp;
		<% end if %>
		<input type="button" value=" Back " onClick="javascript:history.back()">
		</td>
		</tr>
	</table>
</form>
<%
call closeDb()
%>
<!--#include file="AdminFooter.asp"-->