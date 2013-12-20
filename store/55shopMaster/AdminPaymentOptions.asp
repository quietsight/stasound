<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

pageTitle="Add New Payment Option"
pageIcon="pcv4_icon_pg.png"
section="paymntOpt" 
%>
<%PmAdmin=5%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<style>
	ul li {
		padding-top: 4px;
	}
</style>

<%
'// Page view modes
'// viewall = customer selected to View All Payment Options
'// step2 = customer selected one of the radio buttons on step 1
Dim pcvPageMode, pcvPaymentWizard
pcvPageMode=request("mode")
pcvPaymentWizard1=request("paymentWizard1")

'// Get next step after redirecting
	if pcvPageMode="nextStep" then
		select case pcvPaymentWizard1
			case "1" ' Accept credit cards &amp; PayPal
				response.redirect "AdminPaymentOptionsWizard.asp?type=1"
			case "2" ' Accept credit cards
		    	response.redirect "AdminPaymentOptionsWizard.asp?type=2"
			case "3" ' Setup a custom, offline payment method
				response.redirect "AdminPaymentOptionsWizard.asp?type=3"
		end select
	end if
	
Dim rs, conntemp
pcStrPageName = "AdminPaymentOptions.asp"

'// Check Payment Types Exist
call opendb()
query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
pcv_strPayTypeExists=0
if NOT rs.eof then
	pcv_strPayTypeExists=-1
end if
set rs=nothing
call closedb()

%>

<!--#include file="inc_PayPalExpressCheck.asp"-->
<form name="form1" method="post" action="<%=pcStrPageName%>" class="pcForms">
<table class="pcCPcontent">  
	<% 
	'// MODE = VIEWALL
	'// Customer opted to view all payment options on first step
	'// OR customer has already setup one or more payment options and click on Add New
	
	If pcv_strPayTypeExists=-1 OR Request("mode")="viewall" Then %>	
	<tr> 		
		<td> 
                
        <h2>PayPal Payment Options</h2>
        <ul>
          <li><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=0">Website Payments Standard</a></li>
          <li><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=2">Website Payments Pro (US)</a></li>
          <li><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=4">Website Payments Pro (UK)</a></li>  
        </ul>

		<% ' Check if EIG is active
        call opendb()
		Dim pcv_strEigExists
        query="SELECT active, idPayment, gwCode, paymentDesc, paymentNickName FROM paytypes WHERE gwCode=67;"
        set rs=Server.CreateObject("ADODB.Recordset")     
        set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		pcv_strEigExists=0
		if NOT rs.eof then
			pcv_strEigExists=1
		end if
		set rs=nothing
		call closedb()
		%>

        <div style="float: right; margin: 40px 20px 20px 20px;"><a href="http://www.earlyimpact.com/gateway/" target="_blank"><img src="images/ei_logo_gradient_payment_gateway_175.jpg" alt="NetSource Commerce Payment Gateway" style="border: none;" /></a></div>
        <h2>Eary Impact Payment Gateway</h2>

        <% if pcv_strEigExists=0 then %>
        <p>Includes advanced payment features and may reduce the scope of PCI compliance.</p>
        <ul>
        	<li><a href="http://www.earlyimpact.com/gateway/" target="_blank">Learn more about the NetSource Commerce Payment Gateway</a></li>
            <li><a href="AddModRTPayment.asp?gwchoice=67">Activate now</a></li>
        </ul>
        <% else %>
        <p>The NetSource Commerce Payment Gateway is active.
        <ul>
            <li><a href="AddModRTPayment.asp?mode=Edit&id=3&gwchoice=67">Edit settings</a></li>
            <li><a href="https://www.earlyimpact.com/gateway/" target="_blank">Login</a></li>
        	<li><a href="http://wiki.earlyimpact.com/productcart/early_impact_payment_gateway" target="_blank">Documentation</a></li>
        </ul>
        <% end if %>

        <h2>Other Real-Time Payment Options</h2>
        <ul>
          <li><a href="AddPayPalPaymentOpt.asp?gwchoice=VeriSignPP">PayPal Payflow Pro</a></li>
          <li><a href="AddPayPalPaymentOpt.asp?gwchoice=VeriSignLk">PayPal Payflow Link</a></li>
          <li><a href="AddRTPaymentOpt.asp">Authorize.Net, First Data, etc.</a></li>
        </ul>

        <h2>Alternative Checkout Processes</h2>			
        <ul>
          <li><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=1">PayPal Express Checkout</a></li>
          <li><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=3">PayPal Express Checkout (UK)</a></li>
          <li><a href="ConfigureGoogleCheckout.asp">Google Checkout</a></li>
        </ul>

        <h2>Non Real-Time Payment Options</h2>
        <ul>
            <li><a href="AddCCPaymentOpt.asp">Offline credit cards, check, Net 30, etc.</a></li>
            <li><a href="AddCustomCardPaymentOpt.asp">Debit cards, store cards, etc.</a></li>
        </ul>
        
        <h2>Current Payment Options</h2>
        <div>
        	<a href="PaymentOptions.asp">View/Edit</a> payment options that are currently active.
        </div>
       </td>
	</tr>
	<% Else %>
	<tr> 		
		<td>
		  <div>We have detected that no payment options are active on your store. A quick Wizard will take you through a few simple steps to select the payment option(s) that will work best for you.</div></td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 		
		<td> 
			<h2>What would you like to do?</h2>
            <div style="float: right; margin: -35px 100px 0 0;">
            	<img src="https://www.paypal.com/en_US/i/bnr/horizontal_solution_PPeCheck.gif" border="0" alt="Solution Graphics">
            </div>
            <input type="radio" name="paymentWizard1" value="1" checked class="clearBorder"> Accept credit cards &amp; PayPal <br />
            <input type="radio" name="paymentWizard1" value="2" class="clearBorder"> Accept credit cards <br />
            <input type="radio" name="paymentWizard1" value="3" class="clearBorder"> Setup a custom, offline payment method (e.g. "Net 30")
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 		
		<td>
        <hr>
        <input type="Submit" value="Next Step" class="submit2">
        <input type="hidden" name="mode" value="nextStep">
        </td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<% End If %>
</table>
</form>
<!--#include file="AdminFooter.asp"-->