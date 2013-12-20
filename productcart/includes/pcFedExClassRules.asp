<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
'// Enable Logging
pcv_strEnableClientLogs = -1

'/////////////////////////////////////////////////////////////
'// LOGGING
'/////////////////////////////////////////////////////////////
If pcv_strEnableClientLogs = -1 Then
	'// Log our Transaction
	call objFedExClass.pcs_LogTransaction(fedex_postdata, "ErrLog__"&pcv_strMethodName&".in", true)
	'// Log our Error Response
	call objFedExClass.pcs_LogTransaction(FEDEX_result, "ErrLog__"&pcv_strMethodName&".out", true)
End If

'// These Rules will display "Soft" Error messages. We will add to the list as needed.
Select Case pcv_strErrorCodeReturn
case "FE1E"
	response.redirect ErrPageName & "?msg=There was an error processing your request. Please verify all the information you want to submit is correct, then try again."
case "2106"
	if instr(pcv_strErrorMsg,"AccountNumber")>0 then
		response.redirect ErrPageName & "?msg=There was an error processing your request. Please retype your Account Number with no dashes or spaces."
	else
		response.redirect ErrPageName & "?msg=There was an error processing your request. Please verify all the information you want to submit is correct, then try again."
	end if
case "5012"
	response.redirect ErrPageName & "?msg=There was an error processing your request.  Your Account Number was not added to the FedEx database. Please contact FedEx at the email below. "
case else	
	response.redirect ErrPageName & "?msg=There was an error processing your request. " & pcv_strErrorMsg
End Select
%>