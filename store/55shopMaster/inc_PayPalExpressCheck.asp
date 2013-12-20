<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>

<style>
	
	#showPayPalExpress {
		padding: 10px;
		margin: 10px;
		border: 1px solid #CCC;
		background-image: url(images/paypal_29794_screenshot2.gif);
		background-position: right;
		background-repeat: no-repeat;
		min-height: 150px;
	}
	
	#showPayPalExpressTitle {
		font-size: 15px;
		font-weight: bold;
		margin-bottom: 10px;
	}
	
	#showPayPalExpressText {
		color: #666;
		padding-right: 330px;
	}
	
	#showPayPalExpressTextSmall {
		color: #999;
		font-size: 9px;
		padding-right: 330px;
	}
</style>

<%
Dim pcIntShowPPExpressBanner, rsPPE
pcIntShowPPExpressBanner=1


'// 1.) Check to see if PayPal Express is active
call opendb()
query="SELECT active,idPayment,gwCode,paymentDesc FROM paytypes WHERE gwCode=999999 OR gwCode=46 OR gwCode=53 OR gwCode=9 OR gwCode=99 OR gwCode=80"
set rsPPE=Server.CreateObject("ADODB.Recordset")     
set rsPPE=conntemp.execute(query)
if err.number <> 0 then
	set rsPPE=nothing
	call closedb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
end If
If not rsPPE.eof then
	pcIntShowPPExpressBanner=0 '// PayPal Express is active = no need to show the banner
else
	pcIntShowPPExpressBanner=1 '// PayPal Express is not active -> show banner
end if
set rsPPE=nothing
call closedb()



'// 2.) Check for No Show Cookie

'// TEST CODE / Remove Cookie
'Response.Cookies("pcHideExpressSignUp")=""
'Response.Cookies("pcHideExpressSignUp").Expires=Date() - 365
'MyCookiePath=Request.ServerVariables("PATH_INFO")
'do while not (right(MyCookiePath,1)="/")
'	MyCookiePath=mid(MyCookiePath,1,len(MyCookiePath)-1)
'loop
'Response.Cookies("pcHideExpressSignUp").Path=MyCookiePath

'// Check Payment Types Exist
pcv_strHideAlert=0
CookieVar=Request.Cookies("pcHideExpressSignUp")
If CookieVar="Agreed" then
	pcv_strHideAlert=-1
End If

if pcIntShowPPExpressBanner=1 AND pcv_strHideAlert=0 AND session("pcPayPalExpressCookie")="" then
%>
<form>
<div id="showPayPalExpress">
	<div id="showPayPalExpressTitle">
		<input type="checkbox" name="PayPalExpressActive" id="PayPalExpressActive" value="1" class="clearBorder" checked>
		PayPal Express Checkout 
	</div>
	<div id="showPayPalExpressText">According to Jupiter Research, 23% of online shoppers consider PayPal one of their favorite ways to pay online<sup>1</sup>. Accepting PayPal in addition to credit cards is proven to increase your sales<sup>2</sup>. <a href="https://www.paypal.com/us/cgi-bin/?&cmd=_additional-payment-overview-outside" target="_blank">See Quick Demo</a>.</div>
	<div id="showPayPalExpressTextSmall">(1) Payment Preferences Online, Jupiter Research, September 2007. <br />
	(2) Applies to online businesses doing up to $10 million/year in online sales. Based on a Q4 2007 survey of PayPal shoppers conducted by Northstar Research, and PayPal internal data on Express Checkout transactions.
	</div>
    <a href="pcConfigurePayment.asp?gwchoice=999999">Click here to activate Express Checkout now.</a>
</div>
</form>
<script>
	$(document).ready(function()
	{
		$('#PayPalExpressActive').click(function(){
			PayPalExpressSession();
		});		
		function PayPalExpressSession() {
				var isChecked = 0;
				if ($("#PayPalExpressActive").is(':checked')) 
				{
					isChecked = 1;
				}
				$.ajax({
					type: "POST",
					url: "inc_PayPalExpressSession.asp",
					data: "checked=" + isChecked,
					timeout: 5000,
					global: false,
					success: function(data, textStatus){
						if (data=="SECURITY")
						{
							window.location="login_1.asp";
							
						} else {
							
							if (data=="OK")
							{
								
								// no action
								
							} else {
								
								// no action
								
							}
						}
					}
				});
		}
		PayPalExpressSession();
	});	
</script>
<%
end if

'// PAYPAL EXPRESS CHECKOUT - END


'// PAYPAL EXPRESS CHECKOUT - START
%>