<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
'Start SDBA
if request("pagetype")="1" then
	pcv_PageType="1"
	pcv_Title="Drop-Shippers"
else
	if request("pagetype")="0" then
		pcv_PageType="0"
		pcv_Title="Suppliers"
	else
		pcv_PageType=""
		pcv_Title="Customers"
	end if
end if
'End SDBA
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<% pageTitle="Newsletter Wizard: Message Sent and Saved to Your Archive" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->

<form name="d1">
    <table class="pcCPcontent">
        <tr>
            <td colspan="2">
                <table width="100%">
                <tr>
                    <td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
                    <td width="95%"><font color="#A8A8A8">Select <%=pcv_Title%></font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step2.gif"></td>
                    <td><font color="#A8A8A8">Verify <%=pcv_Title%></font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step3.gif"></td>
                    <td><font color="#A8A8A8">Enter message</font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step4.gif"></td>
                    <td><font color="#A8A8A8">Test message</font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step5a.gif"></td>
                    <td><b>Message sent</b></td>
                </tr>
                </table>
            <p>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td>
            <div class="pcCPmessageSuccess"><b>Results:</b><br><b><%=request("sent")%></b> e-mails of <%=request("total")%> have been sent successfully!<br><br>
            <a href="manageNewsWiz.asp">Back to the Newsletter Wizard start page</a>.
            </div></td>
        </tr>
    </table>
</form>
<p>&nbsp;</p>
<!--#include file="AdminFooter.asp"-->