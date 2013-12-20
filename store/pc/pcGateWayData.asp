
<%
call opendb()

pcGatewayDataIdOrder=int(pcGatewayDataIdOrder)

'// Customer Hit "Back Button" or "Session Idle"
If pcGatewayDataIdOrder=0 Then
	call closedb()
	response.redirect "msg.asp?message=38"
End If

pcGatewayDataIdOrder=cLng(pcGatewayDataIdOrder)-cLng(scPre)

pcv_strIdPayment=Request("idPayment")

'SB S
pcIsSubscription = Request("pcIsSubscription")
If len(pcIsSubscription)=0 Then
	pcIsSubscription = session("pcIsSubscription")
End If
pcIsSubTrial = Request("pcIsSubTrial")
If len(pcIsSubTrial)=0 Then
	pcIsSubTrial = session("pcIsSubTrial")
End If
'SB E

if len(pcv_strIdPayment)>0 then
	session("pcSFIdPayment")=Request("idPayment")
end if
%>
<!--#include file="pcCheckReferer.asp"-->
<%
'SB S
query="SELECT orders.idCustomer, orders.total, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.taxAmount, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.pcOrd_ShippingEmail, orders.pcOrd_ShippingFax, orders.pcOrd_shippingPhone, orders.pcOrd_SubTax, orders.pcOrd_SubTrialTax, orders.pcOrd_SubShipping, orders.pcOrd_SubTrialShipping, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pcGatewayDataIdOrder&"));"
'SB E

set rs=server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcIdCustomer=rs("idCustomer")
pcBillingTotal=rs("total")
pcBillingAddress=rs("address")
pcBillingPostalCode=rs("zip")
pcBillingStateCode=rs("stateCode")
pcBillingProvince=rs("state")
pcBillingCity=rs("city")
pcBillingCountryCode=rs("countryCode")
pcBillingTaxAmount=rs("taxAmount")
pcShippingAddress=rs("shippingAddress")
pcShippingStateCode=rs("shippingStateCode")
pcShippingProvince=rs("shippingState")
pcShippingCity=rs("shippingCity")
pcShippingCountryCode=rs("shippingCountryCode")
pcShippingPostalCode=rs("shippingZip")
pcShippingFullName=rs("shippingFullName")
pcBillingAddress2=rs("address2")
pcShippingCompany=rs("shippingCompany")
pcShippingAddress2=rs("shippingAddress2")
pcShippingEmail=rs("pcOrd_ShippingEmail")
pcShippingFax=rs("pcOrd_ShippingFax")
pcShippingPhone=rs("pcOrd_shippingPhone")
pcBillingFirstName=rs("name")
pcBillingLastName=rs("lastName")
pcBillingCompany=rs("customerCompany")
pcBillingPhone=rs("phone")
'SB S
pcSubTax=rs("pcOrd_SubTax")
pcSubTrialTax=rs("pcOrd_SubTrialTax")
pcSubShipping=rs("pcOrd_SubShipping")
pcSubTrialShipping=rs("pcOrd_SubTrialShipping")
'SB E
pcCustomerEmail=rs("email")



set rs=nothing

'SB S
pcBillingSubScriptionTotal = 0.00
if pcIsSubscription Then

	query="SELECT quantity, pcPO_SubAmount, pcPO_SubTrialAmount, pcPO_IsTrial, pcPO_LinkID FROM productsOrdered WHERE idOrder=" & pcGatewayDataIdOrder & " AND pcSubscription_id > 0"
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)

	do while not rs.eof

		'// Get the total subscription amount thats not going to be billed through first gateway pass
		pcv_SubQty = rs("quantity")
		pcv_SubAmount = rs("pcPO_SubAmount")
		pcv_TrialAmount = rs("pcPO_SubTrialAmount")
		pcv_LinkID = rs("pcPO_LinkID")
		pcv_strLinkID = pcv_LinkID
		pcv_intIsTrial = rs("pcPO_IsTrial")

		if pcv_intIsTrial then
			pcBillingSubScriptionTotal = pcBillingSubScriptionTotal + pcv_TrialAmount
		else
			pcBillingSubScriptionTotal = pcBillingSubScriptionTotal + pcv_SubAmount
		end if

		'// Amount
		pcv_TotalSubAmount = pcv_TotalSubAmount + (pcv_SubAmount * pcv_SubQty)

		'// Trial Amount
		pcv_TotalTrialAmount = pcv_TotalTrialAmount + (pcv_TrialAmount * pcv_SubQty)

	rs.movenext
	loop

	'// Amount
	'response.Write("Amount:  " & pcv_TotalSubAmount & "<br />")

	'// Trial Amount
	'response.Write("Trial Amount:  " & pcv_TotalTrialAmount & "<br />")

	'// Tax
	'response.Write("Tax:  " & pcSubTax & "<br />")

	'// Trial Tax
	'response.Write("Trial Tax:  " & pcSubTrialTax & "<br />")

	'// Shipping
	'response.Write("Shipping:  " & pcSubShipping & "<br />")

	'// Trial Shipping
	'response.Write("Trial Shipping:  " & pcSubTrialShipping & "<br />")


	if isnull(pcBillingSubScriptionTotal) or pcBillingSubScriptionTotal = "" Then
		pcBillingSubScriptionTotal = 0.00
	end if

	'// Tax and Shipping Charges
	if pcv_intIsTrial then
		pcv_TaxAndShipping = pcSubTrialTax + pcSubTrialShipping
	else
		pcv_TaxAndShipping = pcSubTax + pcSubShipping
	end if
	pcBillingSubScriptionTotal = pcBillingSubScriptionTotal + pcv_TaxAndShipping

	'// Subtract it from the order total
	pcBillingTotal = (pcBillingTotal - pcBillingSubScriptionTotal)

	'// Pay Now
	'response.Write("Pay Now:  " & pcBillingTotal & ".<br />")
	'response.Write("Pay Later:  " & pcBillingSubScriptionTotal & ".<br />")
	'response.Write("Agree:  " & session("pcAgreeAll") & "<br />")
	'response.Write("<br />")

End if
'SB E

pcBillingState=pcBillingStateCode
if len(pcBillingStateCode)<1 then
	pcBillingState=pcBillingProvince
end if

pcShippingState=pcShippingStateCode
if len(pcShippingStateCode)<1 then
	pcShippingState=pcShippingProvince
end if

'SAVE customer IP to order
'save only the first 15 characters in case this is returned as a list of IP addresses
pcCustIpAddress = left(pcCustIpAddress,15)

query="UPDATE orders SET pcOrd_CustomerIP='"&pcCustIpAddress&"' WHERE (((orders.idOrder)="&pcGatewayDataIdOrder&"));"

set rs=server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if


set rs=nothing

call closedb()
%>
