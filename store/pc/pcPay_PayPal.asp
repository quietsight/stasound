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
'SB S
pcIsSubscription = findSubscription(Session("pcCartSession"), Session("pcCartIndex"))
If pcIsSubscription then		
	strAndSub = "AND (pcPayTypes_Subscription = 1)"
Else		
	strAndSub = ""		
End if 
'SB E

'SB S
if session("customerType")=1 then
	query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 AND (gwCode=999999 OR gwCode=46 OR gwCode=53 OR gwCode=80 OR gwCode=99) " & strAndSub & " ORDER by paymentPriority;"
else
	query="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, gwcode, type, paymentNickName FROM paytypes WHERE active=-1 and Cbtob=0 AND (gwCode=999999 OR gwCode=46 OR gwCode=53 OR gwCode=80 OR gwCode=99) " & strAndSub & " ORDER by paymentPriority;"
end if
'SB E

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
If NOT rs.eof Then
	If rs("gwCode")="80" OR rs("gwCode")="99" Then
		if rs("gwCode")="80" then %>
			<div style="padding-top: 12px;" class="bottomCell">
                <p>&nbsp;OR</p>
                <br />
                <p><a href="pcPay_ExpressPayPPA_Start.asp?refer=<%=pcStrPageName%>"><img  src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" border="0" alt="Acceptance Mark"></a></p>
</div>
		<% else %>
			<div style="padding-top: 12px;" class="bottomCell">
                <p>&nbsp;OR</p>
                <br />
                <p><a href="pcPay_ExpressPayPPL_Start.asp?refer=<%=pcStrPageName%>"><img  src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" border="0" alt="Acceptance Mark"></a></p>
</div>
        <% end if
	Else
		'// Determine which API to use (US or UK)
		query="SELECT pcPay_PayPal.pcPay_PayPal_Partner, pcPay_PayPal.pcPay_PayPal_Vendor FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
		set rsPayPalType=Server.CreateObject("ADODB.Recordset")
		set rsPayPalType=conntemp.execute(query)
		pcPay_PayPal_Partner=rsPayPalType("pcPay_PayPal_Partner")
		pcPay_PayPal_Vendor=rsPayPalType("pcPay_PayPal_Vendor")
		if pcPay_PayPal_Partner<>"" AND pcPay_PayPal_Vendor<>"" then  
			pcPay_PayPal_Version = "UK"			
		else
			pcPay_PayPal_Version = "US"						
		end if
		set rsPayPalType=nothing
		'// Display the API Button Code
		if pcPay_PayPal_Version = "US" then %>
			<div style="padding-top: 12px;" class="bottomCell">
                <p>&nbsp;OR</p>
                <br />
				<p><a href="pcPay_ExpressPay_Start.asp?refer=<%=pcStrPageName%>"><img  src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" border="0" alt="Acceptance Mark"></a></p>
</div>
		<% else %>
			<div style="padding-top: 12px;" class="bottomCell">
			<p>&nbsp;OR</p>
            <br />
            <p><a href="pcPay_ExpressPayUK_Start.asp?refer=<%=pcStrPageName%>"><img  src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" border="0" alt="Acceptance Mark"></a></p>
</div>
		<%
		end if
	End If

End If
set rs=nothing
%>