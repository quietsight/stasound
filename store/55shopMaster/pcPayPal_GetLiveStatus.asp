<%Function GetLivePaymentStatus(TransactionID)
	set objPayPalClass=nothing
	Set resArray=nothing
	Set Session("nvpResArray")=nothing

	set objPayPalClass = New pcPayPalClass
	
	'// Query our PayPal Settings
	objPayPalClass.pcs_SetAllVariables()

	'// Add the required NVP’s
	nvpstr="" '// clear
	objPayPalClass.AddNVP "STARTDATE", "2005-01-01T00:00:00Z"
	objPayPalClass.AddNVP "TRANSACTIONID", TransactionID
	objPayPalClass.AddNVP "TRXTYPE", "Q"
	
	'// Post to PayPal by calling .hash_call
	Set resArray = objPayPalClass.hash_call("TransactionSearch",nvpstr)

	'// Set Response
	Set Session("nvpResArray")=resArray

	'// Check for success
	ack = UCase(resArray("ACK"))
	
	'// Check for code errors
	if err.number <> 0 then 
		'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
		pcv_PayPalErrMessage=""
		%><!--#include file="../includes/pcPayPalErrors.asp"--><%                                             
	end if

	If ack="SUCCESS" Then
		For resindex = 0 To resArray.Count - 1
			TrxnID="L_TRANSACTIONID"&resindex
			If resArray(TrxnID) = TransactionID AND resArray("L_STATUS"&resindex)<>"..2e" Then
				PaymentStatus=resArray("L_STATUS"&resindex)
				exit for
			End if
		Next
	End if
	GetLivePaymentStatus=PaymentStatus
End Function%>
