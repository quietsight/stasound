<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% Dim pageTitle, Section
pageTitle="Manage Help Desk - Message Status"
Section="layout" %>
<!-- #Include File="Adminheader.asp" -->
<%

Dim rs, connTemp, query, intSShowImg
call openDB()

if request("action")="update" then
	intSShowImg=request("ShowImg")
	if intSShowImg<>"1" then
		intSShowImg="0"
	end if
	query="update pcFStatus set pcFStat_ShowImg=" & intSShowImg
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
end if
%>

<table class="pcCPcontent">
    <tr>
        <td colspan="4" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
		<td colspan="4">When you review a message posted by a customer and reply to it, you can then change its status using one of the options shown below.
        <ul class="pcListIcon">
        	<li><a href="adminCreateFBStatus.asp">Add New</a></li>
            <li><a href="adminFBsettings.asp">Help Desk Settings</a></li>
        </ul>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
    <tr>
        <th>Name</th>
        <th>Background Color</th>
        <th nowrap="nowrap" colspan="2">Image file</th>
    </tr>
			<%
			query="select pcFStat_ShowImg,pcFStat_IDStatus,pcFStat_name,pcFStat_Img,pcFStat_BgColor from pcFStatus"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			IF rs.eof THEN
			%>
			<tr>
				<td colspan="3">No Feedback Status found.</td>
			</tr>
			<%
			ELSE
				Dim intFStatShowImg,strTName,strTImg,lngIDStatus
				intFStatShowImg=rs("pcFStat_ShowImg")
				do while not rs.eof
					lngIDStatus=rs("pcFStat_IDStatus")
					strTName=rs("pcFStat_name")
					strTImg=rs("pcFStat_Img")
					strPBgColor=rs("pcFStat_BgColor")
			%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td nowrap="nowrap"><%=strTname%></td>
				<td nowrap="nowrap"><%=strPBgColor%></td>
				<td nowrap="nowrap"><%=strTImg%></td>
				<td align="right"><a href="adminEditFBStatus.asp?IDPro=<%=lngIDStatus%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit"></a>&nbsp;<a href="javascript:if (confirm('You are about to remove this feedback status from ProductCart. Are you sure you want to complete this action?')) location='adminDelFBStatus.asp?IDPro=<%=lngIDStatus%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete"></a></td>
			</tr>
			<%
			rs.MoveNext
			Loop
			set rs = nothing
			call closeDb()
			%>
            <tr>
            	<td colspan="4" class="pcCPspacer"></td>
            </tr>
			<tr>
				<td colspan="4">
					<form name=show action="adminFBStatusManager.asp?action=update" method=post>
                    	<input type="checkbox" name="ShowImg" value="1" <%if intFStatShowImg<>"1" then%><%else%>checked<%end if%> class="clearBorder">
						Show Message Status Images
                        &nbsp;<input type="submit" name="submit" value="Update Message Status Setting" class="submit2">
					</form>
				</td>
			</tr>
			<%END IF%>
    	</table>
<!-- #Include File="Adminfooter.asp" -->