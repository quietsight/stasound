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
pageTitle="Manage Help Desk - Message Priority Levels"
Section="layout" %>
<!-- #Include File="Adminheader.asp" -->
<%

Dim rs, connTemp, query, intPriShowImg

call openDB()

if request("action")="update" then
	intPriShowImg=request("ShowImg")
	if intPriShowImg<>"1" then
		intPriShowImg="0"
	end if
	query="update pcPriority set pcPri_ShowImg=" & intPriShowImg
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
end if
%>

<table class="pcCPcontent">
    <tr>
        <td colspan="3" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>
	<td colspan="3">When customers post a message related to an order, they can specify how urgent the matter is by selecting from a drop-down menu listing the options shown below.
    	<ul class="pcListIcon">
        	<li><a href="adminCreateFBPriority.asp">Add New</a></li>
            <li><a href="adminFBsettings.asp">Help Desk Settings</a></li>
        </ul>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="3"></td>
	</tr>
    <tr>
        <th><strong>Name</strong></th>
        <th nowrap colspan="2"><strong>Image file</strong></th>
    </tr>
		<%
		query="select pcPri_ShowImg,pcPri_IDPri,pcPri_name,pcPri_Img from pcPriority order by pcPri_IDPri asc"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		IF rs.eof THEN
		%>
			<tr>
				<td colspan="3">No Feedback Priority found.</td>
			</tr>
		<%ELSE
		
			intPriShowImg=rs("pcPri_ShowImg")
			
			Dim lngIDPri,strTName,strTImg
			Do while not rs.eof
				lngIDPri=rs("pcPri_IDPri")
				strTName=rs("pcPri_name")
				strTImg=rs("pcPri_Img")
		%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td nowrap="nowrap"><%=strTname%></td>
				<td nowrap="nowrap"><%=strTImg%></td>
				<td align="right"><a href="adminEditFBPriority.asp?IDPro=<%=lngIDPri%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit"></a>&nbsp;<a href="javascript:if (confirm('You are about to remove this feedback priority from ProductCart. Are you sure you want to complete this action?')) location='adminDelFBPriority.asp?IDPro=<%=lngIDPri%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete"></a></td>
			</tr>
		<%
			rs.MoveNext
			Loop
			set rs=nothing
			call closeDb()
		%>
            <tr>
                <td class="pcCPspacer" colspan="3"></td>
            </tr>
			<tr>
				<td colspan="3">
					<form name=show action="adminFBPriorityManager.asp?action=update" method=post>
                        <input type="checkbox" name="ShowImg" value="1" <%if intPriShowImg<>"1" then%><%else%>checked<%end if%> class="clearBorder">
                        Show Message Priority Images
                        &nbsp;<input type="submit" name="submit" value="Update Message Priority Settings" class="submit2">
					</form>
				</td>
			</tr>
		<%END IF%>
</table>
<!-- #Include File="Adminfooter.asp" -->