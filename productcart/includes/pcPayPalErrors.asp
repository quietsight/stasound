<% 
' NOTE: As Specific Errors are reported or discovered during testing, place a user friendly version of that error in this file.

'////////////////////////////////////////////
'// START: User Friendly PayPal Errors
'////////////////////////////////////////////
pcv_strErrNumber=Err.Number
pcv_strErrDescription=Err.Description

Select Case pcv_strErrNumber

	'// 1.) Object Cant be created.
	Case "-2147221005": pcv_PayPalErrMessage="Server can't create object. Your server must be running &quot;Msxml2.serverXmlHttp&quot; to use PayPal Services.<hr/>"


	'// 2.) Object Cant be created. Type Mismatch.
	Case "13": pcv_PayPalErrMessage="There was a problem communicating with PayPal. Please check your server is running &quot;Msxml2.serverXmlHttp&quot;. Also verify that your API Credentials are correct, you’re using the correct currency, and that PayPal is in the correct mode (e.g. Sandbox –OR- Live mode).<hr/>"
	
	'// 3.) Object Cant be created. Type Mismatch.
	Case "-2147012867": pcv_PayPalErrMessage="We can not connect to PayPal servers. If you are in Test Mode please make sure the Sandbox is not offline.<hr/>"	

	'// Else) General Error.
	Case Else: pcv_PayPalErrMessage="There was an error communicating with PayPal.<hr/>"	
	
End Select	

If pcv_strErrNumber<>0 Then
	Err.Number=0
End If
'////////////////////////////////////////////
'// END: User Friendly PayPal Errors
'////////////////////////////////////////////
%>
