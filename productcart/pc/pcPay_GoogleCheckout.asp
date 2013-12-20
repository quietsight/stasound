<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="pcPay_GoogleCheckout_Global.asp"--> 
<!--#include file="pcPay_GoogleCheckout_Checkout.asp"--> 
<% 
'// Google Checkout button implementation	
Dim buttonW
Dim buttonH
Dim buttonStyle
Dim buttonVariant
Dim buttonLoc
Dim buttonSrc	

'// Button Size
pcv_strGoogleBtnSize = GOOGLEBTNSIZE

select case pcv_strGoogleBtnSize
	case "small"
		buttonW = "160"
		buttonH = "43"
	case "medium"
		buttonW = "168"
		buttonH = "44"
	case "large"
		buttonW = "180"
		buttonH = "46"
	case else
		buttonW = "168"
		buttonH = "44"
end select

'// Disable for Gift Registry
if Session("Cust_BuyGift")<>"" then
	pcv_strGoogleBtnVar = 0
else
	pcv_strGoogleBtnVar = 1
end if
				
select case pcv_strGoogleBtnVar
	case 1
		buttonVariant = "text"
		buttonAction = "pcPay_GoogleCheckout_Start.asp?action=checkout"
	case 0
		buttonVariant = "disabled"
		buttonAction = ""
	case else
		buttonVariant = "disabled"
		buttonAction = ""
end select

buttonStyle = "white" '// "trans"
buttonLoc = "en_US"
if GOOGLETESTMODE="YES" then
	buttonSrc = _
		"https://sandbox.google.com/checkout/buttons/checkout.gif" & _
		"?merchant_id=" & strMerchantId & _
		"&w=" & buttonW & _
		"&h=" & buttonH & _
		"&style=" & buttonStyle & _
		"&variant=" & buttonVariant & _
		"&loc=" & buttonLoc
else
	buttonSrc = _
		"https://checkout.google.com/buttons/checkout.gif" & _
		"?merchant_id=" & strMerchantId & _
		"&w=" & buttonW & _
		"&h=" & buttonH & _
		"&style=" & buttonStyle & _
		"&variant=" & buttonVariant & _
		"&loc=" & buttonLoc
end if
%>
<% 
'SB S
pcIsSubscription = findSubscription(Session("pcCartSession"), Session("pcCartIndex"))
If pcIsSubscription then		
	strAndSub = "AND (pcPayTypes_Subscription = 1)"
Else		
	strAndSub = ""		
End if
'SB E

'SB S
query="SELECT idPayment FROM paytypes WHERE active=-1 AND gwCode=50 " & strAndSub & ";"
'SB E
set rsGoogle=Server.CreateObject("ADODB.Recordset")     
set rsGoogle=connTemp.execute(query)		
if rsGoogle.eof then
	pcv_intGoogleActive=0
else
	pcv_intGoogleActive=-1
end if
if pcv_intGoogleActive=-1 then 
%>
	<div style="padding-top: 12px;" class="bottomCell">
	<%
	IF trim(scGoogleAnalytics)<>"" AND NOT IsNull(scGoogleAnalytics) THEN
	%>
	<a href="javascript:setGoogleCheckout();"><img src="<%=buttonSrc%>" id="GoogleCheckout" name="GoogleCheckout" height="<%=buttonH%>" width="<%=buttonW%>" border="0"></a>
	<%
	ELSE
	%>
	<a href="<%=buttonAction%>"><img src="<%=buttonSrc%>" id="GoogleCheckout" name="GoogleCheckout" height="<%=buttonH%>" width="<%=buttonW%>" border="0"></a>
	<%
	END IF
	%>
	</div>
<% end if %>