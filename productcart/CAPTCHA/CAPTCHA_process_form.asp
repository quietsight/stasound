<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz CAPTCHA(TM)
'**  http://www.webwizCAPTCHA.com
'**                                                              
'**  Copyright (C)2005-2008 Web Wiz(TM). All rights reserved.   
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM 'WEB WIZ'.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN 'WEB WIZ' IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwizguide.com/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwizguide.com
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************          

'Initialise variables
Dim blnCAPTCHAcodeCorrect		'Set to true if the CAPTCHA code entered is correct
Dim strCAPTCHAenteredCode		'Holds the code entered by the user

'Initilise the security code bulleon to false
blnCAPTCHAcodeCorrect = false


'Run the CAPTCHA processing code if this is a postback
If Request("CAPTCHA_Postback") OR (len(CAPTCHA_Postback)>0)  Then

	'Check CAPTCHA ocde is correct (case sensitive)
	If blnCAPTCHAcaseSensitive Then
	 	
	 If Session("strSecurityCode") = Request("securityCode") AND Session("strSecurityCode") <> "" Then blnCAPTCHAcodeCorrect = true    
	 
	 
	'Check CAPTCHA ocde is correct (non-case sensitive)
	Else
		If LCase(Session("strSecurityCode")) = LCase(Request("securityCode")) AND Session("strSecurityCode") <> "" Then blnCAPTCHAcodeCorrect = true    
	End If
	
	
	'Reset the security code session variable so it can not be reused
	'Clear session variable
	Session.Contents.Remove("strSecurityCode")
	
	
	'If a redirect has been setup for incorrect CAPTCHA code redirect to it
	If strIncorrectCAPTCHApage <> "" AND blnCAPTCHAcodeCorrect = False Then Response.Redirect(strIncorrectCAPTCHApage)
End If

%>