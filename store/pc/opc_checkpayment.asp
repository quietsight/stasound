<%
call opendb()

pcIdPayment=""
query="SELECT pcCustomerSessions.pcCustSession_IdPayment FROM pcCustomerSessions INNER JOIN customers ON pcCustomerSessions.idCustomer = customers.idcustomer WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&")) ORDER BY pcCustomerSessions.idDbSession DESC;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	pcIdPayment=rsQ("pcCustSession_IdPayment")
end if
set rsQ=nothing

if session("pcSFIdPayment")<>"" AND session("pcSFIdPayment")<>"0" then
pidPayment=session("pcSFIdPayment")
else
pidPayment=pcIdPayment
end if
if request("idpayment")<>"" then
	pidPayment=getUserInput(request("idpayment"),0)
end if
if pidPayment="" then
	pidPayment=pcIdPayment
else
	if not IsNumeric(pidPayment) then
		pidPayment=pcIdPayment
	end if
end if
if pidPayment="" then
	pidPayment=0
end if
pcIdPayment=pidPayment
session("pcSFIdPayment")=pidPayment

if pidPayment<>0 and pidPayment<>"" and pidPayment<>999999 then
	query="SELECT paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE idPayment=" &pidPayment
	set rsQ=server.CreateObject("ADODB.RecordSet")
	set rsQ=connTemp.execute(query)

	if rsQ.eof then
		pidPayment=0
		pcIdPayment=pidPayment
		session("pcSFIdPayment")=pidPayment
	end if
end if

'CHECK PAYMENT DATA
if pidPayment<>0 and pidPayment<>"" and pidPayment<>999999 then
	query="SELECT paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE idPayment=" &pidPayment
	set rsQ=server.CreateObject("ADODB.RecordSet")
	set rsQ=connTemp.execute(query)
	
	if rsQ.eof then
		set rsQ=nothing
		call closeDb() 
		response.redirect "msg.asp?message=200"
	end if
	
	pPaymentDesc=rsQ("paymentDesc")
	pPaymentPriceToAdd=rsQ("priceToAdd")
	pPaymentpercentageToAdd=rsQ("percentageToAdd")
	
	set rsQ=nothing
elseif pidPayment=0 then

	'SB S
	strAndSub = ""
	if pcIsSubscription = True Then
	   strAndSub = " AND pcPayTypes_Subscription = 1 ORDER by pcPayTypes_Subscription, paymentPriority"
	else
	   strAndSub = " ORDER by paymentPriority"
	End if 
	'SB E

	if session("customerType")=1 then
		query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE active=-1 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
	else
		query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE active=-1 AND Cbtob=0 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
	end if 
	set rsQ=server.CreateObject("ADODB.RecordSet")
	set rsQ=connTemp.execute(query)
	
	if rsQ.eof then
		set rsQ=nothing
		call closeDb() 
		response.redirect "msg.asp?message=200"
	end if
	
	pidPayment=rsQ("idPayment")
	pPaymentDesc=rsQ("paymentDesc")
	pPaymentPriceToAdd=rsQ("priceToAdd")
	pPaymentpercentageToAdd=rsQ("percentageToAdd")
	
	set rsQ=nothing
	
	pcIdPayment=pidPayment
else
	if pidPayment=999999 then
		query="SELECT paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE gwcode=46 OR gwcode=53 OR gwcode=999999;"
		set rsQ=server.CreateObject("ADODB.RecordSet")
		set rsQ=connTemp.execute(query)
		if rsQ.eof then
			pPaymentDesc="Paypal Express Checkout"
			pPaymentPriceToAdd=0
			pPaymentpercentageToAdd=0
		else
			pPaymentDesc=rsQ("paymentDesc")
			pPaymentPriceToAdd=rsQ("priceToAdd")
			pPaymentpercentageToAdd=rsQ("percentageToAdd")
		end if
		set rsQ=nothing
	end if
end if
'END CHECK PAYMENT DATA
call closeDb()
%>