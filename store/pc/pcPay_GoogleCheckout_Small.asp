<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<% 
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

'// Open Private Connection String
Set conGoogle=Server.CreateObject("ADODB.Connection")
conGoogle.Open scDSN
		
'SB S
IF trim(scGoogleAnalytics)<>"" AND NOT IsNull(scGoogleAnalytics) THEN
	sbCartArr=Session("pcCartSession")
	If (sbCartArr(1,38)>0) then
		pcIsSubscription = True		
		strAndSub = "AND (pcPayTypes_Subscription = 1)"
	Else		
		pcIsSubscription = False		
		strAndSub = ""		
	End if 
ELSE
	pcIsSubscription = findSubscription(Session("pcCartSession"), Session("pcCartIndex"))
	If pcIsSubscription then	
		strAndSub = "AND (pcPayTypes_Subscription = 1)"
	Else		
		strAndSub = ""		
	End if 
END IF
'SB E		
		
'SB S
query="SELECT idPayment FROM paytypes WHERE active=-1 AND gwCode=50 " & strAndSub & ";"
'SB E
set rsGoogle=Server.CreateObject("ADODB.Recordset")     
set rsGoogle=conGoogle.execute(query)	
	
if rsGoogle.eof then
	pcv_intGoogleActive=0
else
	pcv_intGoogleActive=-1
end if
if pcv_intGoogleActive=-1 then 

	IF trim(scGoogleAnalytics)<>"" AND NOT IsNull(scGoogleAnalytics) THEN
	%>
	<div style="padding-top: 10px;"><a href="javascript:setGoogleCheckout();"><img src="<%=buttonSrc%>" id="GoogleCheckout" name="GoogleCheckout" height="<%=buttonH%>" width="<%=buttonW%>" border="0"></a></div>
	<% 
	ELSE
	%>
	<div style="padding-top: 10px;"><a href="<%=buttonAction%>"><img src="<%=buttonSrc%>" id="GoogleCheckout" name="GoogleCheckout" height="<%=buttonH%>" width="<%=buttonW%>" border="0"></a></div>
	<% 
	END IF
end if 

'// Close Private Connection String
conGoogle.Close
Set conGoogle=nothing
%>