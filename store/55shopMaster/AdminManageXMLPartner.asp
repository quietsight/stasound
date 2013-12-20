<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="XML Tools - Manage Partners" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim rs, connTemp, query
call openDB()
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td>
		<p>For security reasons, you need to explicitly allow a third party application to communicate XML data feeds with your ProductCart-powered store. Use this feature to create/manage these applications, which we call &quot;XML Partners&quot;.</p>
		<ul class="pcListIcon">
			<li><a href="instXMLPartner.asp">Add New Partner</a></li>
            <li><a href="XMLToolsManager.asp">XML Tools Manager</a></li>
		<%
		query="SELECT pcXL_ID FROM pcXMLLogs WHERE pcXP_ID=0;"
		set rs=connTemp.execute(query)
		if not rs.eof then%>
		<li><a href="viewXMLPartnerLogs.asp?idPartner=0">Unauthorized Transactions Log</a></li>
		<%end if
		set rs=nothing%>
        </ul>
		</td>
	</tr>
	
	<% If Request.QueryString("msg")<>"" Then %>
	<tr> 
		<td align="center">
			<div class="pcCPmessageSuccess">
			<%if Request.QueryString("msg")="added" then%>
				The XML Partner has been added successfully!
			<%end if%>
			<%if Request.QueryString("msg")="deleted" then%>
				The XML Partner was removed successfully!
			<%end if%>
			</div>
		</td>
	</tr>
	<% End If %>
	
	<tr>
		<td align="center">
			<table class="pcCPcontent">
				<tr>
					<th nowrap>Partner ID</th>
					<th nowrap>Partner Key</th>
					<th nowrap>Partner Name</th>
					<th nowrap>Actions</th>
				</tr>
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<%
				query="select pcXP_ID,pcXP_PartnerID,pcXP_Key,pcXP_Name,pcXP_Email,pcXP_Status,pcXP_ExportAdmin FROM pcXMLPartners WHERE pcXP_Removed=0 ORDER BY pcXP_PartnerID ASC;"
				set rstemp=Server.CreateObject("ADODB.Recordset")
				set rstemp=connTemp.execute(query)
				If rstemp.eof Then
				%>             
					<tr> 
						<td colspan="4"><p>No XML Partners Found.</p></td>
					</tr>            
				<%
				Else
					pcArr=rstemp.getRows()
					set rstemp=nothing
					intCount=ubound(pcArr,2)
					For i=0 to intCount
					pa_ID=pcArr(0,i)
					pa_UserID=pcArr(1,i)
					pa_Key=pcArr(2,i)
					pa_Name=pcArr(3,i)
					pa_Email=trim(pcArr(4,i))
					pa_Status=pcArr(5,i)
					pa_ExportAdmin=pcArr(6,i)
					%>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td nowrap><%=pa_UserID%><%if clng(pa_Status)<>1 then%>&nbsp;<img src="images/notactive.gif" border="0"><%end if%><%if pa_ExportAdmin="1" then%>&nbsp(<b>Export Admin</b>)<%end if%></td>
						<td nowrap><%=pa_Key%></td>
						<td nowrap><%=pa_Name%></td>
						<%query="SELECT pcXL_id FROM pcXMLLogs WHERE pcXP_id=" & pa_ID & ";"
						set rs=connTemp.execute(query)
						pa_HaveTrans=0
						if not rs.eof then
							pa_HaveTrans=1
						end if
						set rs=nothing%>
						<td nowrap>
							<a href="modXMLPartner.asp?idPartner=<%=pa_ID%>">View/Edit</a>
							<%if pa_Email<>"" then%>&nbsp;|&nbsp;<a href="mailto:<%=pa_Email%>">E-mail</a><%end if%>
							<%if pa_HaveTrans=1 then%>&nbsp;|&nbsp;<a href="viewXMLPartnerLogs.asp?idPartner=<%=pa_ID%>">Transactions</a><%end if%>
							&nbsp;|&nbsp;<a href="javascript:if (confirm('You are about to remove this XML Partner from your database. Are you sure you want to complete this action?')) location='delXMLPartner.asp?action=del&idPartner=<%=pa_ID%>'">Delete</a>
						</td>
					</tr>              
					<%
					Next
				End If
				%>
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="4"><a href="instXMLPartner.asp">Add New Partner</a> | <a href="XMLToolsManager.asp">XML Tools Manager</a></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->