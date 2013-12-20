<%
response.Buffer=true
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"

'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Choose a way to accept credit card payments"
pageIcon="pcv4_icon_pg.png"
section="paymntOpt" 

pcSpryCP="PP" %>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
dim query, conntemp, rstemp
call openDb()

if request("mode")="disable9" then
	query= "DELETE FROM payTypes WHERE gwCode=9"
	set rs = server.CreateObject("ADODB.RecordSet")
	set rs = conntemp.execute(query)
end if

dim pcCode2, pcCode3, pcCode99, pcCode80, pcCode53, pcCode999999, pcCode46
pcCode2 = 0
pcCode3 = 0
pcCode99 = 0
pcCode80 = 0
pcCode53 = 0
pcCode999999 = 0
pcCode46 = 0
paymentActive = 0
paymentCSS2 = "CollapsiblePanelContentDisabled"
paymentCSS3 = "CollapsiblePanelContentDisabled"
paymentCSS99 = "CollapsiblePanelContentDisabled"
paymentCSS80 = "CollapsiblePanelContentDisabled"
paymentCSS53 = "CollapsiblePanelContentDisabled"
paymentCSS46 = "CollapsiblePanelContentDisabled"
paymentCSS999999 = "CollapsiblePanelContentDisabled"


query = "SELECT idPayment, gwCode, paymentDesc FROM payTypes;"
set rs = server.CreateObject("ADODB.RecordSet")
set rs = conntemp.execute(query)
do until rs.eof
	idPayment = rs("idPayment")
	gwCode = rs("gwCode")
	paymentDesc = rs("paymentDesc")
	
	Select Case gwCode
	case "2"
		pcCode2 = 1
		idPayment2 = idPayment
		paymentDesc2 = paymentDesc
		paymentActive = 1
		paymentCSS2 = "CollapsiblePanelContentEnabled"
	case "3"
		pcCode3 = 1
		idPayment3 = idPayment
		paymentDesc3 = paymentDesc
		paymentActive = 1
		paymentCSS3 = "CollapsiblePanelContentEnabled"
	case "99"
		pcCode99 = 1
		idPayment99 = idPayment
		paymentDesc99 = paymentDesc
		paymentActive = 1
		paymentCSS99 = "CollapsiblePanelContentEnabled"
	case "80"
		pcCode80 = 1
		idPayment80 = idPayment
		paymentDesc80 = paymentDesc
		paymentActive = 1
		paymentCSS80 = "CollapsiblePanelContentEnabled"
	case "53"
		pcCode53 = 1
		idPayment53 = idPayment
		paymentDesc53 = paymentDesc
		paymentActive = 1
		paymentCSS53 = "CollapsiblePanelContentEnabled"
	case "999999"
		pcCode999999 = 1
		idPayment999999 = idPayment
		paymentDesc999999 = paymentDesc
		paymentActive = 1
		paymentCSS999999 = "CollapsiblePanelContentEnabled"
	case "46"
		pcCode46 = 1
		idPayment46 = idPayment
		paymentDesc46 = paymentDesc
		paymentActive = 1
		paymentCSS46 = "CollapsiblePanelContentEnabled"
	end select
	rs.moveNext
loop
set rs = nothing

dim myCountry
myCountry=request("SelectPaymentCountry")

call opendb()

if myCountry&""="" then
	'pull from database
	query = "SELECT PP_Country FROM paypal;"
	set rs = server.CreateObject("ADODB.RecordSet")
	set rs = conntemp.execute(query)
	myCountry = rs("PP_Country")
else
	'save to database
	query = "UPDATE paypal SET PP_Country='"&myCountry&"';"
	set rs = server.CreateObject("ADODB.RecordSet")
	set rs = conntemp.execute(query)
end if
set rs = nothing
call closedb()

dim jumptogw1, jumptogw2, jumptogw3

if request("BTN1") = "Enable" then
	tmpRedirect = request("ALLINONE")
	'check if gateway is active
	call opendb()
	query = "SELECT idPayment, gwCode, paymentDesc FROM payTypes WHERE gwCode = "&tmpRedirect&";"
set rs = server.CreateObject("ADODB.RecordSet")
set rs = conntemp.execute(query)
	if rs.eof then
		call closedb()
		response.redirect "pcConfigurePayment.asp?gwchoice="&tmpRedirect
		response.end
	else
		tmpIdPayment = rs("idPayment")
		call closedb()
		response.redirect "pcConfigurePayment.asp?mode=Edit&id="&tmpIdPayment&"&gwchoice="&tmpRedirect
		response.end
	end if
end if

if request("BTN2") = "Enable" then
	tmpRedirect = request("BANK")
	'check if gateway is active
	call opendb()
	query = "SELECT idPayment, gwCode, paymentDesc FROM payTypes WHERE gwCode = "&tmpRedirect&";"
set rs = server.CreateObject("ADODB.RecordSet")
set rs = conntemp.execute(query)
	if rs.eof then
		call closedb()
		response.redirect "pcConfigurePayment.asp?gwchoice="&tmpRedirect
	else
		tmpIdPayment = rs("idPayment")
		call closedb()
		response.redirect "pcConfigurePayment.asp?mode=Edit&id="&tmpIdPayment&"&gwchoice="&tmpRedirect
		response.end
	end if
end if
if request("BTN3") = "Enable" then
	'Check if Offline methods were selected<br>
	If request("ALTERNATIVE") = "OL" or request("ALTERNATIVE")="CP" Then
		call closedb()
		response.redirect "AddCCPaymentOpt.asp"
	End If
	'check if gateway is active
	call opendb()
	query = "SELECT active FROM payTypes where gwCode=50"
	set rs = server.CreateObject("ADODB.RecordSet")
	set rs = conntemp.execute(query)
	tmpActive = rs("active")
	set rs = nothing
	call closedb()
	if tmpActive="0" then
		response.redirect "ConfigureGoogleCheckout.asp"
	else
		response.redirect "ConfigureGoogleCheckout2.asp"
	end if
end if

%>
<form name="formname" method="post" action="<%=pcStrPageName%>" class="pcForms">  
    <% if myCountry&""="" then %>
        <table class="pcCPcontent">
            <tr>
                <td>
                    <table width="25%">
                        <tr>
                          <td colspan="3" width="35%" align="left" style="font-size:18px;">Payment Setup<hr /></td>
                        </tr>
                        <tr>
                            <td width="35%" align="left" style="font-size:14px;" nowrap>Select your country</td>
                            <td width="60%" align="right">
                                <select name="SelectPaymentCountry">
                                    <option value="">Select your country</option>
                                    <option value="US">United States</option>
                                    <option value="CA">Canada</option>
                                    <option value="UK">United Kingdom</option>
                                    <option value="Global">Global</option>
                                </select></td>
                            <td width="5%" align="right" style="font-size:15px;"><input name="GO" type="submit" id="GO" value="GO"></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    <% else %>
        <table class="pcCPcontent">
            <tr>
                <td>
                    <table width="100%">
                        <tr>
                            <td width="35%" align="left" style="font-size:18px;">Payment Setup</td>
                            <td width="60%" align="right">
                                <select name="SelectPaymentCountry">
                                    <option value="">Select your country</option>
                                    <option value="US" <% if myCountry="US" then%>selected<% end if %>>United States</option>
                                    <option value="CA" <% if myCountry="CA" then%>selected<% end if %>>Canada</option>
                                    <option value="UK" <% if myCountry="UK" then%>selected<% end if %>>United Kingdom</option>
                                    <option value="Global" <% if myCountry="Global" then%>selected<% end if %>>Global</option>
                                </select></td>
                            <td width="5%" align="right" style="font-size:15px;"><input name="GO" type="submit" id="GO" value="GO"></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
	<% end if %>
</form>
<% if myCountry&""<>"" then %>
<form name="formname" method="post" action="<%=pcStrPageName%>" class="pcForms">  
<table class="pcCPcontent"><tr><td>
	<% if paymentActive = 0 then %>
        <table style="background-color:#EEE;border: solid 1px #CCC;">
            <tr>
                <td><p><strong>Understanding Online Payments</strong></p></td>
            </tr>
            <tr>
                <td><p>Watch this short video and learn the basics of how online payment processing works. You'll also find out what to look for when selecting the best payment processing solution for your business.<br>
                <br>
                <INPUT TYPE='BUTTON' VALUE='Watch Video' onClick="open('http://www.youtube.com/embed/d3QG_1R3hI0', 'Sample',   'location=yes,scrollbars=no,width=640,height=380')"></p>
            </tr>
        </table>
        <br>
    <% end if %>
    <table width="100%" style="background-color:#EEE;border: solid 1px #CCC;"><tr><td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                    <tr>
                        <td colspan="2" class="pcPanelTitle1"><table><tr><td><img src="images/expand.gif" width="19" height="19" hspace="2" vspace="2"></td><td><span class="pcSubmenuHeader">All-in-One Payment Solutions</span></td></tr></table></td>
                    </tr>
                    <tr class="pcPanelDesc"><td colspan="2"><hr style="background:#ccc;border:0;" /></td></tr>
                    <tr class="pcPanelDesc">
                        <td width="24" rowspan="2" valign="top"><img src="Gateways/logos/paypal_logo.png" width="228" height="55" alt="PayPal Payments"></td>
                        <td width="580" class="pcPanelItalic">Everything you need.</td>
                    </tr>
                    <tr class="pcPanelDesc">
                        <td>Get a merchant account and a payment gateway for a quick, easy way to accept all types of payment.</td>
                    </tr>
                    <tr class="pcPanelDesc">
                        <td valign="top">&nbsp;</td>
                        <td>&nbsp;</td>
                    </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
				<% 
                If pcCode3 = 1 Then
                    pcLinkString = "mode=Edit&id="&idPayment3&"&gwchoice=3"
                    pcButtonString = "Edit"
                    pcTD1 = "&nbsp;"
                    pcTD2 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                    pcTD3 = "<input type='button' name='Disable' value='Disable' onClick=""javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id="&idPayment3&"&gwchoice=3&page=pcPaymentSelection.asp'"">"
                Else
                    pcLinkString = "gwchoice=3"
                    pcButtonString = "Enable"
                    pcTD1 = "<INPUT TYPE='BUTTON' VALUE='Demo' onClick=""open('https://merchant.paypal.com/us/cgi-bin/?cmd=_render-content&amp;content_ID=merchant/demo_WPS', 'Sample',   'location=yes,scrollbars=no,width=640,height=380')"">"
                    pcTD2 = "<INPUT TYPE='BUTTON' VALUE='Learn More' onClick=""open('https://www.paypal.com/webapps/mpp/paypal-payments-standard', 'Sample',   'location=yes,scrollbars=yes,resizable=yes')"">"
                    pcTD3 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                End If
                %>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            <table class="<%=paymentCSS3%>" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td width="65%">
                                        <span class="pcSubmenuHeader">PayPal Payments Standard</span>
                                        <br />
                                        <span class="pcSubmenuContent">Accept credit cards quickly and securely. Buys are sent to PayPal to pay, and then return to your site when finished. Setup is easy, there are no monthly charges, and buyers don't need a PayPal account.</span>
                                    </td>
                                    <td width="9%" class="pcSubmenuContent"><%=pcTD1%></td>
                                    <td width="13%" class="pcSubmenuContent"><%=pcTD2%></td>
                                    <td width="13%" class="pcSubmenuContent"><%=pcTD3%></td>
            
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <% If myCountry="US" Then
                    If pcCode80 = 1 Then
                        pcLinkString = "mode=Edit&id="&idPayment80&"&gwchoice=80"
                        pcButtonString = "Edit"
                        pcTD1 = "&nbsp;"
                        pcTD2 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                        pcTD3 = "<input type='button' name='Disable' value='Disable' onClick=""javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id="&idPayment80&"&gwchoice=80&page=pcPaymentSelection.asp'"">"
                    Else
                        pcLinkString = "gwchoice=80"
                        pcButtonString = "Enable"
                        pcTD1 = "<INPUT TYPE='BUTTON' VALUE='Demo' onClick=""open('PPAVideo.html', 'Sample',   'location=yes,scrollbars=no,width=580,height=325')"">"
                        pcTD2 = "<INPUT TYPE='BUTTON' VALUE='Learn More' onClick=""open('https://www.paypal.com/webapps/mpp/paypal-payments-advanced', 'Sample',   'location=yes,scrollbars=yes,resizable=yes')"">"
                        pcTD3 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                    End If
                    %>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td>
                                <table class="<%=paymentCSS80%>" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td width="65%"><span class="pcSubmenuHeader">PayPal Payments Advanced</span><br /><span class="pcSubmenuContent">The easy way to create a professional checkout experience that lets buyers pay without leaving your site and PayPal processes credit cards behind the scenes, healping you simplify PCI compliance.</span></td>
                                        <td width="9%" class="pcSubmenuContent"><%=pcTD1%></td>
                                        <td width="13%" class="pcSubmenuContent"><%=pcTD2%></td>
                                        <td width="13%" class="pcSubmenuContent"><%=pcTD3%></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                <% End If %>
                <% If myCountry="US" OR myCountry="UK" Then
                    If myCountry="US" Then
                        tempPPPname = "PayPal Payments Pro"
                        tempPPPlink = "46"
                        tempPaymentCSS = paymentCSS46
                    End If %>
                    <% If myCountry="UK" Then
                        tempPPPname = "Website Payments Pro"
                        tempPPPlink = "53"
                        tempPaymentCSS = paymentCSS53
                    End If %>
                    <% 
                    If pcCode46 = 1 Then
                        pcLinkString = "mode=Edit&id="&idPayment46&"&gwchoice="&tempPPPlink
                        pcButtonString = "Edit"
                        pcTD1 = "&nbsp;"
                        pcTD2 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                        pcTD3 = "<input type='button' name='Disable' value='Disable' onClick=""javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id="&idPayment46&"&gwchoice=46&page=pcPaymentSelection.asp'"">"
                    Else
                        pcLinkString = "gwchoice="&tempPPPlink
                        pcButtonString = "Enable"
                        pcTD1 = "&nbsp;" '<INPUT TYPE='BUTTON' VALUE='Demo' onClick=""open('https://merchant.paypal.com/cms_content/US/en_US/files/merchant/pp_msr_player.swf?config=https://merchant.paypal.com/cms_content/US/en_US/files/merchant/pp_ms_paymentsadvanced_config.xml&debug=false', 'Sample',   'location=yes,menubar=no,scrollbars=no,titlebar=no,toolbar=no,width=640,height=380')"">"
                        pcTD2 = "<INPUT TYPE='BUTTON' VALUE='Learn More' onClick=""open('https://www.paypal.com/webapps/mpp/paypal-payments-pro', 'Sample',   'location=yes,scrollbars=yes,resizable=yes')"">"
                        pcTD3 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                    End If
                    %>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td>
                        <table class="<%=tempPaymentCSS%>" cellspacing="0" cellpadding="0">
                            <tr>
                                <td width="65%"><span class="pcSubmenuHeader"><%=tempPPPname%></span><br /> 
                                    <span class="pcSubmenuContent"> Fully customize your checkout pages and accept credit cards directly on your site. PayPal simplifies PCI compliance for you, plus you get Virtual Terminal at no added cost.</span>
                                </td>
                                <td width="9%" class="pcSubmenuContent"><%=pcTD1%></td>
                                <td width="13%" class="pcSubmenuContent"><%=pcTD2%></td>
                                <td width="13%" class="pcSubmenuContent"><%=pcTD3%></td>
                            </tr>
                        </table>
                    </td></tr></table>
                <% End If %>
                <table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td>
                    <table width="100%" class="CollapsiblePanelContentDisabled">
                        <tr>
                            <td><strong>Other All-in-One Payment Solutions</strong></td>
                            <td class="pcSubmenuContent">&nbsp;</td>
                        </tr>
                        <tr>
                            <td><span class="pcSubmenuHeader">
                            <select name="ALLINONE" id="ALLINONE">
                            <option value="13">2Checkout (2CO)</option>
                            <option value="29">BluePay</option>
                            <option value="57">Beanstream</option>
                            <option value="60">Dow Commerce</option>
                            <option value="65">EC Suite - Transaction Gateway System</option>
                           	<option value="31">eWay (AU)</option>
                            <option value="64">Pay Junction - QuickLink</option>
                            </select>
                            </span></td>
                            <td width="13%" class="pcSubmenuContent"><input type="submit" name="BTN1" id="BTN1" value="Enable"></td>
                        </tr>
                    </table>
                </td></tr></table>
            </div>
        </div>
        <div id="CollapsiblePanel2" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                <tr>
                <td colspan="2" class="pcPanelTitle1">
                  <table>
                    <tr>
                      <td><img src="images/expand.gif" alt="" width="19" height="19" hspace="2" vspace="2"></td>
                      <td><span class="pcSubmenuHeader">Gateway Solutions</span></td>
                    </tr>
                  </table></td>
                </tr>
                  <tr class="pcPanelDesc"><td colspan="2"><hr style="background:#ccc;border:0;" /></td></tr>
                <tr class="pcPanelDesc">
                <td width="24" rowspan="2" valign="top"><img src="Gateways/logos/payflow_logo.jpg" width="150" height="68" alt="PayPal Payments"></td>
                <td width="580" class="pcPanelItalic">Join forces with your bank.</td>
                </tr>
                <tr class="pcPanelDesc">
                <td>Use a merchant account from your financial institution to acceopt online payments.</td>
                </tr>
                <tr class="pcPanelDesc">
                  <td valign="top">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <% If myCountry="US" OR myCountry="CA" Then
                    If pcCode99 = 1 Then
                        pcLinkString = "mode=Edit&id="&idPayment99&"&gwchoice=99"
                        pcButtonString = "Edit"
                        pcTD1 = "&nbsp;"
                        pcTD2 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                        pcTD3 = "<input type='button' name='Disable' value='Disable' onClick=""javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id="&idPayment46&"&gwchoice=99&page=pcPaymentSelection.asp'"">"
                    Else
                        pcLinkString = "gwchoice=99"
                        pcButtonString = "Enable"
                        pcTD1 = "&nbsp;"
                        pcTD2 = "<INPUT TYPE='BUTTON' VALUE='Learn More' onClick=""open('https://merchant.paypal.com/us/cgi-bin/?&cmd=_render-content&content_ID=merchant/payment_gateway&nav=2.1.2&nav=2.08', 'Sample',   'location=yes,scrollbars=yes,resizable=yes')"">"
                        pcTD3 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                    End If
                    %>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td>
                                <table class="<%=paymentCSS99%>" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td width="65%"><span class="pcSubmenuHeader">PayPal Payflow Link</span><br /> 
                                        <span class="pcSubmenuContent">connect your merchant account with a PCI-compliant gateway. Setup is quick and customers pay without leaving your site. </span>
                                        </td>
                                        <td width="9%" class="pcSubmenuContent"><%=pcTD1%></td>
                                        <td width="13%" class="pcSubmenuContent"><%=pcTD2%></td>
                                        <td width="13%" class="pcSubmenuContent"><%=pcTD3%></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                <% End If %>
                <% 
                If pcCode2 = 1 Then
                    pcLinkString = "mode=Edit&id="&idPayment2&"&gwchoice=2"
                    pcButtonString = "Edit"
                    pcTD1 = "&nbsp;"
                    pcTD2 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                    pcTD3 = "<input type='button' name='Disable' value='Disable' onClick=""javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id="&idPayment46&"&gwchoice=2&page=pcPaymentSelection.asp'"">"
                Else
                    pcLinkString = "gwchoice=2"
                    pcButtonString = "Enable"
                    pcTD1 = "&nbsp;"
                    pcTD2 = "<INPUT TYPE='BUTTON' VALUE='Learn More' onClick=""open('https://merchant.paypal.com/us/cgi-bin/?&amp;cmd=_render-content&amp;content_ID=merchant/payment_gateway&amp;nav=2.1.2&amp;nav=2.08', 'Sample',   'location=yes,scrollbars=yes,resizable=yes')"">"
                    pcTD3 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                End If
                %>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            <table class="<%=paymentCSS2%>" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td width="65%"><span class="pcSubmenuHeader">PayPal Payflow Pro</span><br /> 
                                   <span class="pcSubmenuContent"> Use your own merchant account and stay in control of your checkout pages with this fully customizable gateway solution. PayPal simplifies PCI compliance for you, if needed.</span>
                                    </td>
                                    <td width="9%" class="pcSubmenuContent"><%=pcTD1%></td>
                                    <td width="13%" class="pcSubmenuContent"><%=pcTD2%></td>
                                    <td width="13%" class="pcSubmenuContent"><%=pcTD3%></td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            <table width="100%" class="CollapsiblePanelContentDisabled">
                                <tr>
                                    <td><strong>Other Gateway Solutions</strong></td>
                                    <td class="pcSubmenuContent">&nbsp;</td>
                                </tr>
                                <tr>
                                    <td><span class="pcSubmenuHeader">
                                    <select name="BANK" id="BANK">
                                    <option value="67">NetSource Commerce Gateway</option>
                                    <option value="39">ACH Direct, Inc</option>
                                    <option value="1">AuthorizeNet</option>
                                    <option value="52">ChronoPay</option>
                                    <option value="32">CyberSource</option>
                                    <option value="42">eProcessing Network, LLC</option>
                                    <option value="54">ETS - EMoney</option>
                                    <option value="37">Fastcharge</option>
                                    <option value="8">FirstData</option>
                                    <option value="58">Global Pay</option>
                                    <option value="30">InternetSecure</option>
                                    <option value="5">iTransact, Inc.</option>
                                    <option value="8">LinkPoint (FirstData)</option>
                                    <option value="11">Moneris - eSelect Plus Direct Post</option>
                                    <option value="27">NETbilling</option>
                                    <option value="55">Ogone</option>
                                    <option value="59">Omega</option>
                                    <option value="12">Payment Express &reg; PX Pay</option>
                                    <option value="47">Payment Express &reg; PX Post</option>
                                    <option value="48">PayPoint.Net (formerly SECPay)</option>
                                    <option value="4">PSiGate</option>
                                    <option value="26">Sage Pay (Protx)</option>
                                    <option value="49">Skipjack</option>
                                    <option value="63">TotalWeb Solutions</option>
                                    <option value="70">Transaction Express™ - TransFirst</option>
                                    <option value="24">TrustCommerce - TCLink</option>
                                    <option value="35">USA ePay</option>
                                    <option value="56">Virtual Merchant</option>
                                    <option value="10">WorldPay - Select Junior</option>
                                    </select>
                                    </span></td>
                                    <td width="13%" class="pcSubmenuContent"><input type="submit" name="BTN2" id="BTN2" value="Enable"></td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        <div id="CollapsiblePanel3" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td colspan="2" class="pcPanelTitle1">
                      <table>
                        <tr>
                          <td><img src="images/expand.gif" alt="" width="19" height="19" hspace="2" vspace="2"></td>
                          <td><span class="pcSubmenuHeader">Add Alternative Payment Methods</span></td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr class="pcPanelDesc"><td colspan="2"><hr style="background:#ccc;border:0;" /></td></tr>
                  <tr class="pcPanelDesc">
                    <td width="24" rowspan="2" valign="top"><img src="Gateways/logos/paypal_express_logo.gif" width="145" height="42" alt="PayPal Payments Express"></td>
                    <td width="580" class="pcPanelItalic">Quick and easy setup.</td>
                  </tr>
                  <tr class="pcPanelDesc">
                    <td>Give buyers another way to pay by adding an alternative payment method. Including Offline Credit Card Processing or Customized Payments.</td>
                  </tr>
                  <tr class="pcPanelDesc">
                    <td valign="top">&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <% 
                If pcCode999999 = 1 Then
                    pcLinkString = "mode=Edit&id="&idPayment999999&"&gwchoice=999999"
                    pcButtonString = "Edit"
                    pcTD1 = "&nbsp;"
                    pcTD2 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                    pcTD3 = "<input type='button' name='Disable' value='Disable' onClick=""javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id="&idPayment3&"&gwchoice=999999&page=pcPaymentSelection.asp'"">"
                Else
				    pcLinkString = "gwchoice=999999"
                    pcButtonString = "Enable"
                    pcTD1 = "<INPUT TYPE='BUTTON' VALUE='Demo' onClick=""open('https://cms.paypal.com/us/cgi-bin/?&cmd=_render-content&content_ID=merchant/demo_express_checkout', 'Sample', 'location=yes,scrollbars=no,width=640,height=380')"">"
                    pcTD2 = "<INPUT TYPE='BUTTON' VALUE='Learn More' onClick=""open('https://www.paypal.com/webapps/mpp/express-checkout', 'Sample',   'location=yes,scrollbars=yes,resizable=yes')"">"
                    pcTD3 = "<input type='button' name='"&pcButtonString&"' value='"&pcButtonString&"' onClick=""location.href='pcConfigurePayment.asp?"&pcLinkString&"';"">"
                End If
                %>
                <table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td>
                    <table class="<%=paymentCSS999999%>" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="65%"><span class="pcSubmenuHeader">PayPal Express Checkout</span><br /> 
                            <span class="pcSubmenuContent"> If you already accept credit cards online, add PayPal as an alternative way to pay. Tapping into millions of shoppers who prefer paying with PayPal is a quick and easy way to lift your sales. </span>
                            </td>
                            <td width="9%" class="pcSubmenuContent"><%=pcTD1%></td>
                            <td width="13%" class="pcSubmenuContent"><%=pcTD2%></td>
                            <td width="13%" class="pcSubmenuContent"><%=pcTD3%></td>
                        </tr>
                    </table>
                </td></tr></table>
                <table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td>
                    <table width="100%" class="CollapsiblePanelContentDisabled">
                        <tr>
                            <td><strong>Other Alternative Payment Solutions</strong></td>
                            <td class="pcSubmenuContent">&nbsp;</td>
                        </tr>
                        <tr>
                            <td><span class="pcSubmenuHeader">
                            <select name="ALTERNATIVE" id="ALTERNATIVE">
                                <option value="GC">Google Checkout</option>
                                <option value="OL">Offline Credit Card Processing</option>
                                <option value="CP">Custom Payment Options</option>
                            </select>
                            </span></td>
                            <td width="13%" class="pcSubmenuContent"><input type="submit" name="BTN3" id="BTN3" value="Enable"></td>
                        </tr>
                    </table>
                </td></tr></table>
        	</div>
    	</div>
		<script type="text/javascript">
        <!--
        var CollapsiblePanel1 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel1", {contentIsOpen:false});
        var CollapsiblePanel2 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel2", {contentIsOpen:false});;
        var CollapsiblePanel3 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel3", {contentIsOpen:false});
        //-->
        </script>
    </td></tr></table>
</td></tr></table>
</form>
<% End If %>
<!--#include file="AdminFooter.asp"-->