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
<% pageTitle="Newsletter Wizard: Review Test and Send Message" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->

<form name="hForm" method="post" action="newsWizStep5a.asp?action=send&pagetype=<%=pcv_PageType%>" class="pcForms">
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
                    <td><b>Send message</b></td>
                </tr>
                </table>
            <p>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td>
                <%if request("from")="4" then%>
                	<h2>Review the Test Message</h2>
                    <b>Did you receive the TEST message?</b> Check the address to which you have sent the test message. If everything looks in order, click on 'Send Message' to send the message to the <%=pcv_Title%> list that you have built using the Wizard.
                <%else%>
                    <b>Your are ready to go</b>! Click on 'Send Message' to send the message to the <%=pcv_Title%> list that you have built using the Wizard.
                <%end if%>
            </td>
        </tr>
        <tr>
            <td><hr noshade size="1" color="#e1e1e1"></td>
        </tr>
        <tr>
            <td><strong>NOTE</strong>: The ProductCart Newsletter Wizard is not intended to handle large email lists. Messages are sent one by one, to avoid exceeding limitations to the number of concurrent receipients that may be in place on your Web server's mail server. In our tests, sending a message typically took between 1 and 2 seconds each. Therefore, sending a newsletter to 100 <%=pcv_Title%> should take about 3 minutes.</td>
        </tr>
        <tr>
            <td><strong>NO SPAM</strong>: DO NOT USE this feature to send SPAM email. Regardless of whether or not spam is considered illegal in your State and/or Country, sending unsolicited messages is not what this feature is meant for. It is also not a good marketing practice and it will harm your business in the long run.</td>
        </tr>
        <tr>
            <td align="center">&nbsp;</td>
        </tr>
        <tr>
            <td align="center">
                <input type="submit" name="submit" value="Send message" class="submit2">&nbsp;&nbsp;
                <input type="button" name="back" value="Back" onClick="location='newsWizStep4.asp?pagetype=<%=pcv_PageType%>'">
            </td>
        </tr>
        <tr>
            <td align="center">&nbsp;</td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->