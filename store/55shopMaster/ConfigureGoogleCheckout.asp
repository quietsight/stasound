<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Google Checkout Activation Wizard" %>
<% Section="paymntOpt" %>
<% pcPageName = "ConfigureGoogleCheckout.asp" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% 
Dim query, rs, conntemp


'**************************************************************************
' START: Check if customer checked box
'**************************************************************************
if request.form("submit")<>"" then
	
	'// Set error count
	pcv_intErr=0	
	
	'// generic error for page
	pcv_strGenericPageError = "Please check the checkbox that confirms that you have properly setup your Google Checkout account. If you have any questions, make sure to review the Guide to using Google Checkout with ProductCart."
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ValidateTextField	"FedExMode", true, 0	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	If pcv_intErr>0 Then
		response.redirect pcPageName & "?msg=" & pcv_strGenericPageError
	Else
		response.redirect "ConfigureGoogleCheckout2.asp"
	End If
end if
'**************************************************************************
' END: If registration request was submitted, process request
'**************************************************************************

msg=request.querystring("msg")
if msg<>"" then 
%>
<div class="pcCPmessage">
	<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
</div>
<% end if %>
	
<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Getting your Google Checkout Account ready</th>
		</tr>
		<tr>
		  <td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2"><p>Before you get started <a href="http://wiki.earlyimpact.com/widgets/integrations/googlecheckout" target="_blank">carefully review the documentation on this topic</a>.</p>
			  <p>To activate Google Checkout, follow these steps: </p></td>
		</tr>
		<tr>
			<td width="5%" valign="top" align="right"><img border="0" src="images/step1a.gif"></td>
			<td width="95%"><p>If you have not yet done so, <a href="http://checkout.google.com/sell?promo=seei" target="_blank">sign up for Google Checkout</a></td>
		</tr>
		<tr>
			<td valign="top" align="right"><img border="0" src="images/step2a.gif"></td>
			<td><p>Obtain your <span style="font-style: italic">&quot;<a href="http://checkout.google.com/support/sell/bin/answer.py?answer=42963&amp;topic=8670" target="_blank">Google merchant ID</a>&quot;</span> and <span style="font-style: italic">&quot;<a href="http://checkout.google.com/support/sell/bin/answer.py?answer=42963&amp;topic=8670" target="_blank">Google merchant key</a>&quot;</span> from the <span style="font-style: italic">&quot;Settings &gt; Integrations&quot;</span> page inside your Google Account.</p></td>
		</tr>
		<%
		if scSSLUrl="" then
			psvCallBackURL = "NO"
		else
			psvCallBackURL=replace((scSslURL&"/"&scPcFolder&"/pc/pcPay_GoogleCheckout_Callback.asp"),"//","/")
			psvCallBackURL=replace(psvCallBackURL,"https:/","https://")
		end if
		%>
		<tr>
			<td valign="top" align="right"><img border="0" src="images/step3a.gif"></td>
			<td>
			<p>Also on the <span style="font-style: italic">&quot;Settings &gt; Integrations&quot;</span> page, correctly enter the <span style="font-style: italic">&quot;<a href="http://checkout.google.com/support/sell/bin/answer.py?hl=en&answer=134463" target="_blank">Call Back URL</a>&quot;</span>. Your Call Back URL is:<br /><br />
		<% if psvCallBackURL = "NO" then %>
		<div class="pcCPmessage">The Call Back URL cannot be properly configured as you are not using an SSL certificate. <a href="AdminSettings.asp">Configure the SSL settings now</a>.</div>
		<% else %>
		<input type="text" value="<%=psvCallBackURL%>" size="120">
		<% end if %>
		</p></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Activate Google Checkout in ProductCart</th>
		</tr>
		<tr>
		  <td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">
		  <p>After checking the box below to confirm you have completed these instructions, you may click the "Continue" button.  ProductCart will then ask you for your Google Checkout Credentials and guide your though the Checkout options. If you have any problems creating your Google Account, or obtaining your <span style="font-style: italic">&quot;Google merchant ID&quot;</span> and <span style="font-style: italic">&quot;Google merchant key&quot;</span>, please review the following <a href="http://checkout.google.com/support/sell/bin/topic.py?topic=8668" target="_blank">FAQs.</a></p>
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="right"><input name="FedExMode" type="checkbox" value="YES" class="clearBorder"></td>
		  <td><strong>I have a Google Checkout account, a Merchant ID  and Merchant Key, and I saved my API Callback URL</strong></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
	
		<tr> 
			<td colspan="2">
			<% 
			if scXML<>".3.0" then
				psvXMLParser = "NO"
			end if
			if psvCallBackURL = "NO" OR psvXMLParser = "NO" then %>
				<% if  psvCallBackURL = "NO" then %>
					<div class="pcCPmessage">
						You cannot activate Google Checkout because your store is not using an SSL certificate. 
						<a href="AdminSettings.asp">Configure the SSL settings now</a>.
					</div>
				<% end if %>
				<% if  psvXMLParser = "NO" then %>
					<div class="pcCPmessage">
						You cannot activate Google Checkout because your store is not using a supported XML Parser. Google Checkout currently supports XML Parser v3. <a href="pcTSUtility.asp">Change the XML parser</a> (<a href="http://www.earlyimpact.com/faqs/afmviewfaq.asp?faqid=377" target="_blank">Help on this topic</a>).
				  </div>
				<% end if %>
				<br />
			<% else %>
				<input type="submit" name="Submit" value="Continue" class="submit2">
			<% end if %>
			&nbsp;
			<input type="button" name="back" value="Back" onClick="JavaScript:history.back();">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->