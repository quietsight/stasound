<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage your e-mail marketing professionally with MailUp" %>
<% Section="genRpts" %>
<%PmAdmin="10*7*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim rs, conntemp
pcPageName="mailup_home.asp"

'// START - Check for MailUp and redirect to Add-on Home page

	tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess,pcMailUpSett_TurnOff FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		session("SF_MU_APIUser")=pcMailUpSett_APIUser
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		session("SF_MU_APIPassword")=pcMailUpSett_APIPassword
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		session("SF_MU_URL")=pcMailUpSett_URL
		tmp_Auto=rs("pcMailUpSett_AutoReg")
		if IsNull(tmp_Auto) or tmp_Auto="" then
			tmp_Auto=0
		end if
		session("SF_MU_Auto")=tmp_Auto
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
		session("SF_MU_Setup")=tmp_setup
		tmp_TurnOff=rs("pcMailUpSett_TurnOff")
		if IsNull(tmp_TurnOff) OR tmp_TurnOff="" then
			tmp_TurnOff=0
		end if
	end if
	set rs=nothing
	call closedb()

if (session("SF_MU_Setup")="1") OR (tmp_TurnOff="1") then
	response.Redirect("mu_manageNewsWiz.asp")
end if
'// END
%>
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 		
		<td>
    	<p><a href="http://www.earlyimpact.com/productcart/mailup/" target="_blank"><img src="images/pc2008-mailup-box-big.jpg" alt="ProductCart Recurring Billing Add-on" width="258" height="191" align="right" style="margin-left: 15px;"></a>Create, send, and track e-mail newsletters reliably and professionally.</p>
		  <p style="padding-top: 6px;">E-mail marketing is an integral part of running a successful e-commerce store. That's why we decided to integrate MailUp - a proven e-mail newsletter management system - with our shopping cart software.</p>
		  <ul>
      	<li>Unlimited lists, unlimited contacts, unlimited messages</li>
      	<li>Over 100 professionally designed e-mail templates</li>
        <li>Reads, clicks, opens... detailed statistics down to the user level</li>
        <li>Dynamic, two-way integration with ProductCart</li>
        <li><a href="http://www.earlyimpact.com/productcart/mailup/" target="_blank">Learn more</a></li>
        <li><a href="mu_manageNewsWiz.asp">Activate MailUp</a></li>
          </ul>
    </td>
	</tr>	
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 	
</table>
<!--#include file="AdminFooter.asp"-->