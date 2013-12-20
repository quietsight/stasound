<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Payment Option" %>
<% Section="paymntOpt" %>
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
'// Get page type for querystring
	pcvPaymentWizard=request("type")
	select case pcvPaymentWizard
		case "1" ' Accept credit cards &amp; PayPal
		%>
		
			<table class="pcCPcontent"> 
            	<tr>
                	<td> 
                    	<div style="float: right; margin: -10px 10px 0 0;">
                        <img src="https://www.paypal.com/en_US/i/bnr/horizontal_solution_PPeCheck.gif" border="0" alt="Solution Graphics">
                        </div>
                      <h2>Accept Credit Cards &amp; PayPal</h2>
                      <div style="margin-bottom: 20px;">Choose a solution to accept credit cards and PayPal.</div>
                      
                      <div style="width: 340px; margin: 10px; padding: 10px; border: 1px solid #CCC; float: left;">
                      	<h2><strong>PayPal Standard</strong><span class="pcSmallText"> - <a href="https://www.paypal.com/us/cgi-bin/webscr?cmd=_wp-standard-overview-outside" target="_blank">More information</a></span></h2>
                        <div style="padding: 5px; background-color:#FC0;">
                        Easy to get started. No monthly fees.</div>
                        <div style="margin-top: 15px;">
                            <strong>Features</strong>
                            <ul>
                                <li>Accept VISA, MC, Amex, Discover, PayPal and more at one low rate</li>
                                <li>Buyers enter credit card information on secure PayPal pages and immediately return to your site. Your buyers do not need a PayPal account.</li>
                                <li>Start selling as soon as you sign up.</li>
                            </ul>
                            <strong>Pricing</strong>
                            <ul>
                                <li>No monthly fees</li>
                                <li>No set up or cancellation fees</li>
                                <li>Transaction fees: 1.9% - 2.9% +$0.30 USD (based on sales volume)</li>
                            </ul>
                        </div>
                        <div style="text-align: center;"><input type="button" value="Select" class="submit2" onClick="document.location.href='AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=0'"></div>
                      </div>
                      
                      <div style="width: 340px; margin: 10px; padding: 10px; border: 1px solid #CCC; float: right;">
                      	<h2><strong>PayPal Website Payments Pro</strong><span class="pcSmallText"> - <a href="https://www.paypal.com/us/cgi-bin/webscr?cmd=_wp-pro-overview-outside" target="_blank">More information</a></span></h2>
                        <div style="padding: 5px; background-color:#FC0;">
                        Advanced ecommerce solutions for established businesses.</div>
                        <div style="margin-top: 15px;">
                            <strong>Features</strong>
                            <ul>
                                <li>Accept VISA, MC, Amex, Discover, PayPal and more at one low rate</li>
                                <li>Buyers enter credit card information directly on your site, and do not need a PayPal account.</li>
                                <li>Business credit application required to start selling, decision usually comes within 24 hours.</li>
                                <li>Includes Virtual Terminal - accept payments for orders taken via phone, fax, and mail.</li>
                            </ul>
                            <strong>Pricing</strong>
                            <ul>
                                <li>$30 per month</li>
                                <li>No set up or cancellation fees</li>
                                <li>Transaction fees: 2.2% - 2.9% +$0.30 USD (based on sales volume)</li>
                            </ul>
                        </div>
                        <div style="text-align: center;"><input type="button" value="Select (US Merchants)" class="submit2" onClick="document.location.href='AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=2'">&nbsp;<input type="button" value="Select (UK Merchants)" class="submit2" onClick="document.location.href='AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=4'"></div>
                      </div>
                      
                      <div style="clear: both;" class="pcForms">
                      <hr>
                      	<form>
                        	<input type="button" value="Back" onClick="document.location.href='AdminPaymentOptions.asp'">
                        </form>
                      </div>
					</td>
				</tr>
			</table>
		
		<%
		case "2" ' Accept credit cards
		%>
        
			<table class="pcCPcontent"> 
            	<tr>
                	<td>
                      <h2>Accept Credit Cards</h2>
                      <h3 style="margin-bottom: 0;">All-in-one Solutions</h3>
                      <div>These payment systems typically <u>do not</u> require a separate Internet merchant account with a bank.</div>
						<ul>
                        <% if scCompanyCountry="US" then %>
            			<li><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=2">PayPal Website Payments Pro</a> (US only) - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_wp-pro-overview-outside" class="pcSmallText" style="color:#999;" target="_blank">Learn more</a></li>
                        <% else %>
                        <li><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=4">PayPal Website Payments Pro</a> (UK only) - <a href="https://www.paypal.com/uk/cgi-bin/webscr?cmd=_wp-pro-overview-outside" class="pcSmallText" style="color:#999;" target="_blank">Learn more</a></li>
                        <% end if %>
                        <li><a href="AddModRTPayment.asp?gwchoice=13">2Checkout</a> - <a href="http://www.2checkout.com/" class="pcSmallText" style="color:#999;" target="_blank">Learn more</a></li>
                        <li><a href="ConfigureGoogleCheckout.asp">Google Checkout</a> - <a href="http://www.earlyimpact.com/productcart/googleCheckout/" class="pcSmallText" style="color:#999;" target="_blank">Learn more</a></li>
                        <li><a href="AddModRTPayment.asp?gwchoice=10">WorldPay</a> - <a href="http://www.rbsworldpay.com/products/index.php?page=ecom&c=UK" class="pcSmallText" style="color:#999;" target="_blank">Learn more</a></li>
                        <li><a href="AddRTPaymentOpt.asp">Other payment systems</a></li>
                      </ul>
                      
                      <h3 style="margin-bottom: 0;">Payment gateways</h3>
						<% ' Highlight NetSource Commerce Payment Gateway %>
                        <div style="position: relative;">
                            <div style="float: right; margin: 30px 0 0 0;"><a href="AddModRTPayment.asp?gwchoice=67"><img src="images/ei_logo_gradient_payment_gateway_175.jpg" alt="NetSource Commerce Payment Gateway" border="0" /></a></div>
                        </div>
                      <div>These payment systems require an <a href="AdminPaymentMerchantAccount.asp">Internet merchant account</a> account with a bank [<a href="AdminPaymentMerchantAccount.asp">get one</a>].</div>
						<ul>
            			<li><a href="AddPayPalPaymentOpt.asp?gwchoice=VeriSignPP">PayPal Payflow Pro</a> - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_payflow-pro-overview-outside" class="pcSmallText" style="color:#999;" target="_blank">Learn more</a></li>
                        <li><a href="AddPayPalPaymentOpt.asp?gwchoice=VeriSignLk">PayPal Payflow Link</a> - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_payflow-link-overview-outside" class="pcSmallText" style="color:#999;" target="_blank">Learn more</a></li>
                        <li><a href="AddModRTPayment.asp?gwchoice=67"><strong>NetSource Commerce Payment Gateway</strong></a> - <a href="http://www.earlyimpact.com/gateway/" class="pcSmallText" style="color:#999;">Learn more</a></li>
                        <li><a href="AddRTPaymentOpt.asp">Other payment gateways</a></li>
                      </ul>
                      
                      <div style="clear: both;" class="pcForms">
                      <hr>
                      	<form style="margin: 20px 0;">
                            <input type="button" value="I need a merchant account" onClick="document.location.href='AdminPaymentMerchantAccount.asp'">
                            <input type="button" value="View all payment options" onClick="document.location.href='AdminPaymentOptions.asp?mode=viewall'">
                          <input type="button" value="Back" onClick="document.location.href='AdminPaymentOptions.asp'">
                        </form>
                      </div>
				  </td>
				</tr>
			</table>        
        
        
        <%
		case "3" ' Setup a custom, offline payment method
		%>
		
			<table class="pcCPcontent"> 
            	<tr>
                	<td> 
                      <h2>Non Real-Time Payment Options</h2>
                      Your store will not connect to an outside payment system. You can create multiple, custom payment options using the links below.<a href="http://wiki.earlyimpact.com/productcart/productcart#chapter_7payment_options" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="Learn more about this feature" width="16" height="16" border="0"></a>
                      <ul>
            			<li><a href="AddCCPaymentOpt.asp">Offline credit cards, check, Net 30, etc.</a></li>
                        <li><a href="AddCustomCardPaymentOpt.asp">Debit cards, store cards, etc.</a></li>
                      </ul>
					</td>
				</tr>
			</table>
		
		<%
	end select
%>
<!--#include file="AdminFooter.asp"-->