<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Manage Blackout Dates"
pageIcon="pcv4_icon_calendar.png"
section="layout"
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<%
Dim rs, connTemp, query
call openDB()
%>
<!--#include file="AdminHeader.asp"-->
	<table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr>
            <td colspan="3">This feature works <u>in conjunction</u> with the <a href="checkoutOptions.asp#datetime">Delivery Date &amp; Time</a> Checkout Option to allow you to notify the customer of dates that may not be selected during checkout. For example, a catering company may not provide its services on certain holidays. Therefore the customer should not be able to select those dates during checkout.</td>
        </tr>
        <tr>
            <td colspan="3" class="pcCPspacer"></td>
        </tr>
        <tr> 
            <th nowrap width="20%">Blackout Date</th>
            <th nowrap width="80%" colspan="2">Message</th>
        </tr>
        <tr>
            <td colspan="3" class="pcCPspacer"></td>
        </tr>
              
		<%
            query="select * from Blackout order by Blackout_Date asc"
            set rs=Server.CreateObject("ADODB.Recordset")
            set rs=connTemp.execute(query)
            If rs.eof Then
        %>
        <tr> 
            <td colspan="3">No Blackout Dates Found.</td>
        </tr>
              
		<% Else 
                Dim strCol
                strCol="#E1E1E1"
                Do While NOT rs.EOF
                    Blackout_Date=rs("Blackout_Date")
                    Blackout_Message=rs("Blackout_Message")
        %>

                        <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                            <td><a href="Blackout_edit.asp?Blackout_Date=<%=Blackout_Date%>"><%=showdateFrmt(Blackout_Date)%></a></td>
                            <td><a href="Blackout_edit.asp?Blackout_Date=<%=Blackout_Date%>"><%=Blackout_Message%></a></td>
                            <td align="right" nowrap><a href="Blackout_edit.asp?Blackout_Date=<%=Blackout_Date%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit" border="0"></a>&nbsp;<a href="javascript:if (confirm('You are about to remove this Blackout Date from your database. Are you sure you want to complete this action?')) location='Blackout_delete.asp?Blackout_Date=<%=Blackout_Date%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete" border="0"></a>
                            </td>
                        </tr>
              
<%
				rs.MoveNext
				Loop
			End If
%>
            <tr>
                <td colspan="3" class="pcCPspacer"></td>
            </tr>
            <tr>
                <td colspan="3" align="center">
                    <form class="pcForms">
                    <input type="button" value="Add New Blackout Date" name="btnCreate" onClick="location='Blackout_add.asp';" class="submit2">
                    &nbsp;
                    <input type="button" name="Button" value="Back" onClick="javascript:history.back()">
                    </form>
                </td>
            </tr>
	</table>
<!--#include file="AdminFooter.asp"-->