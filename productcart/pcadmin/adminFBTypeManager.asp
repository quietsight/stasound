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
<% 
Dim pageTitle, Section
pageTitle="Manage Help Desk - Message Types"
Section="layout" 
%>
<!-- #Include File="Adminheader.asp" -->
<%
Dim rs, connTemp, query, intTShowImg

if request("action")="update" then
	intTShowImg=request("ShowImg")
	if not validNum(intTShowImg) then
		intTShowImg="0"
	end if
	call openDB()
	query="update pcFTypes set pcFType_ShowImg=" & intTShowImg
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
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
		<td align="left" valign="top" colspan="3">When customers post a message related to an order, they can select the type of message that they are posting from a drop-down menu showing the message types listed below (e.g. problem, comment, suggestion, etc.).
        <ul class="pcListIcon">
        	<li><a href="adminCreateFBType.asp">Add New</a></li>
            <li><a href="adminFBsettings.asp">Help Desk Settings</a></li>
        </ul>
        </td>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Name</th>
		<th nowrap="nowrap" colspan="2">Image file</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
		<%
            call openDb()
            query="select pcFType_ShowImg,pcFType_IDType,pcFType_name,pcFType_Img from pcFTypes"
            set rs=Server.CreateObject("ADODB.Recordset")
            set rs=connTemp.execute(query)
            IF rs.eof THEN
        %>
                <tr>
                    <td colspan="3">No Feedback Types found.</td>
                </tr>
		<%
            ELSE
        
            Dim intFTypeShowImg,lngIDType,strTName,strTImg
            do while not rs.eof
            intFTypeShowImg=rs("pcFType_ShowImg")
            lngIDType=rs("pcFType_IDType")
            strTName=rs("pcFType_name")
            strTImg=rs("pcFType_Img")
        %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td nowrap="nowrap"><%=strTname%></td>
                    <td nowrap="nowrap"><%=strTImg%></td>
                    <td align="right"><a href="adminEditFBType.asp?IDPro=<%=lngIDType%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit"></a>&nbsp;<a href="javascript:if (confirm('You are about to remove this feedback type from ProductCart. Are you sure you want to complete this action?')) location='adminDelFBType.asp?IDPro=<%=lngIDType%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete"></a></td>
                </tr>
		
		<%
            rs.MoveNext
            Loop
            set rs=nothing
            call closeDb()
        %>
            <tr>
                <td colspan="3" class="pcCPspacer"></td>
            </tr>
            <tr>
                <td colspan="3" valign="top">
                    <form name="show" action="adminFBTypeManager.asp?action=update" method="post" class="pcForms">
                    	<input type="checkbox" name="ShowImg" value="1" <%if intFTypeShowImg=1 then%>checked<%end if%> class="clearBorder"> Show Message Type Images &nbsp; <input type="submit" name="submit" value="Update Message Type Setting" class="submit2">
                    </form>
                </td>
            </tr>
		<%END IF%>
        <tr>
            <td colspan="3" class="pcCPspacer"></td>
        </tr>
</table>
<!-- #Include File="Adminfooter.asp" -->