<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Newsletter Wizard" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<%'Start SDBA
pcv_pageType=request("pagetype")
'End SDBA
call opendb()
query="SELECT pcMailUpSett_RegSuccess FROM pcMailUpSettings WHERE pcMailUpSett_RegSuccess=1;"
set rs=connTemp.execute(query)
if not rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "mu_manageNewsWiz.asp"
end if
set rs=nothing
call closedb()
%>
<table class="pcCPcontent">
	<tr>
		<td>
		<p>The Newsletter Wizard allows you obtain a list of <%if pcv_pageType="0" then%>suppliers<%else%><%if pcv_pageType="1" then%>drop-shippers<%else%>customers<%end if%><%end if%> using a number of filters, and then export the list or send a message within ProductCart. You can also use a previously sent message to send a new message to the same list.
		<ul>
		<li><a href="<%if pcv_pageType<>"" then%>sds_newsWizStep1.asp?pagetype=<%=pcv_pageType%><%else%>newsWizStep1.asp<%end if%>">Start the Wizard</a> to create a new message</li>
		<li>View <a href="manageNews.asp?pagetype=<%=pcv_pageType%>">previously sent messages</a></li>
		</ul>
		</td>
	</tr>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
    <th>NO SPAM</th>
	</tr>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
    <p>DO NOT USE this feature to send SPAM e-mail. Regardless of whether or not SPAM is considered illegal in your State or Country, sending unsolicited messages is not what this feature is meant for. It is also not a good marketing practice and it will harm your business in the long run.</p>
    <p><strong>US STORES: You must comply with  <a href="http://www.ftc.gov/spam/" target="_blank">CAN-SPAM law</a></strong></p>
    <p>Make sure that you comply with the CAN-SPAM.  <a href="http://www.ftc.gov/spam/" target="_blank">Click here for more details</a>. Failure to comply could result in fines and possible imprisonment. In a nutshell, all commercial e-mail messages:</p>
    <ul>
        <li>Must not present misleading information in the From field or header information.</li>
        <li>Must include a link for and honor unsubscribe requests.</li>
        <li>Must conspicuously state that all commercial, promotional mail is an advertisement, unless all recipients have opted in </li>
        <li>Must Include a valid, physical mailing address (postal address) in all email campaigns.</li>
    </ul>
    </td>
  </tr>
	<%
	'MailUp-S
	Dim query,rs,connTemp
	call opendb()
	pcIncMailUp=0
	query="SELECT pcMailUpSett_RegSuccess FROM pcMailUpSettings WHERE pcMailUpSett_RegSuccess=1;"
	set rs=connTemp.execute(query)
		if err.number<>0 then
			set rs = nothing
			call closedb()
			response.Redirect("upddb_MailUp.asp")
		end if
	if not rs.eof then
		pcIncMailUp=1
	end if
	set rs=nothing
	call closedb()
	IF pcIncMailUp=0 THEN
	''MailUp-E%>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
    <th>Not a professional e-mail marketing system</th>
	</tr>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td><p><img src="images/pc2008-mailup-box-big.jpg" alt="MailUp professional e-mail marketing system integrated with ProductCart" width="258" height="191" align="right" style="margin-left: 15px;">Please keep in mind that the ProductCart Newsletter Wizard is not a professional e-mail marketing system and it not intended to handle large e-mail lists. There are a number of reasons why it might make sense for you to upgrade to a more robust and feature-rich e-mail marketing system</p>
    <ul>
    	<li>Double-opt in mechanism (subscription pending until confirmed by e-mail)</li>
      <li>Message tracking (reads, opens, clicks, etc.)</li>
      <li>Separate subscription management for different multiple lists (e.g. 'Product Updates' vs. 'Specials and Promotions')</li>
      <li>List-specific, one-click unsubscribe and bounced messages management</li>
      <li>Robust infrastructure to send a message to a high number of recipients</li>
     </ul>
     <p>There are many providers of e-mail marketing management services. At NetSource Commerce we chose to work with a long-time partner of ours - NWEB - to integrate their <a href="http://www.earlyimpact.com/productcart/mailup/" target="_blank">MailUp</a> service - a professional <a href="http://www.earlyimpact.com/productcart/mailup/" target="_blank">e-mail newsletter management system</a> - with ProductCart.</p>
     <p>&nbsp;</p>
    </td>
	</tr>
	<%
	'MailUp-S
	ELSE%>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
    <th>MailUp Newsletter Management System</th>
	</tr>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<a href="mu_manageNewsWiz.asp"><img src="images/MailUp_200.jpg" alt="MailUp Newsletter Management System" align="right" style="margin: 30px;" /></a>MailUp is a professional e-mail &amp; SMS message management system. It allows you to efficiently broadcast and track messages sent to e-mail or mobile phone recipients. The integration with ProductCart connects your customer database - and your customer newsletter subscription preferences - with your MailUp console.</p>
			<p><a href="mu_manageNewsWiz.asp">Click here</a> to use ProductCart - MailUp Intergration.</p>
		</td>
	</tr>
	<%END IF
	'MailUp-E%>

</table>
<!--#include file="AdminFooter.asp"-->