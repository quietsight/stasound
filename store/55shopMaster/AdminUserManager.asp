<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Users" %>
<% section="layout" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
		<tr>
			<td colspan="3" class="pcCPspacer">
			<!--#include file="pcv4_showMessage.asp"-->
            </td>
		</tr>
        <tr>
            <td colspan="3">
            User this feature to <a href="AdminAddUser.asp">create new store managers</a> that have limited access to the areas of the Control Panel that you select. <a href="AdminAddUser.asp">Add New</a>.
            </td>
        </tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
        <tr>
            <th width="10%" align="left" nowrap>User ID</th>
            <th width="20%" align="left" nowrap>Contact Name</th>
            <th width="60%" align="left">Permissions</td>
            <th width="10%" align="right" nowrap></th>
        </tr>
        <tr>
            <td colspan="3" class="pcCPspacer"></td>
        </tr>
        <%
            Dim rs, connTemp, query
            call openDB()
            query="SELECT * FROM admins ORDER BY id ASC"
            set rs=Server.CreateObject("ADODB.Recordset")
            set rs=connTemp.execute(query)
            rs.MoveFirst
            rs.MoveNext
            If rs.eof Then
        %>             
                <tr> 
                    <td colspan="3">No additional Control Panel Users found.</td>
                </tr>            
        <%
            Else 
                Do While NOT rs.EOF
                IDAdmin=rs("ID")
                AdminUser=rs("IDAdmin")
				AdminLevel=rs("adminLevel")
				AdminName=rs("adm_ContactName")
				AdminEmail=rs("adm_ContactEmail")
					
					Dim myArr, pcv_intCount, permissionDetails, mycount
					permissionDetails=""
					myArr=split(AdminLevel,"*")
					pcv_intCount=ubound(myArr)-1
						mycount=0
						for i=0 to pcv_intCount
							permissionDetails=permissionDetails & ", "
							query="SELECT pmName FROM permissions WHERE idpm =" & myArr(i)
							set rstemp=Server.CreateObject("ADODB.Recordset")
							set rstemp=connTemp.execute(query)
							permissionDetails=permissionDetails & rstemp("pmName")
						next
						set rstemp=nothing
						x=len(permissionDetails)-2
						permissionDetails=right(permissionDetails,x)
            %>
                                        
            <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                <td align="left" valign="top"><%=AdminUser%></td>
                <td align="left" valign="top"><a href="AdminEditUser.asp?id=<%=IDAdmin%>"><%=AdminName%></a></td>
                <td align="left" valign="top"><%=permissionDetails%></td>
                <td align="right" valign="top" nowrap="nowrap" class="cpLinksList"><a href="AdminEditUser.asp?id=<%=IDAdmin%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="View/Edit" title="View/Edit"></a><% if adminEmail<>"" then%>&nbsp;<a href="mailto:<%=adminEmail%>"><img src="images/pcIconList.jpg" width="12" height="12" alt="Email" title="Email"></a><% end if %>&nbsp;<a href="javascript:if (confirm('You are about to remove this Control Panel user from your database. Are you sure you want to complete this action?')) location='AdminDelUser.asp?id=<%=IDAdmin%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete" title="Delete"></a>
                </td>
            </tr>              
            <%
                    rs.MoveNext
                    Loop
                End If
                set rs=nothing
                call closeDb()
            %>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->