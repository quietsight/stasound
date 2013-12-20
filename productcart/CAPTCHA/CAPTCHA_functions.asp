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

'Get about status
Private Function about()

	about = 1 

	'Calulate the lentgh
	If LEN(strLicense) > 47 Then
		
		Dim intAbout
		intAbout = 1
		
		'Read in the data
		If isNumeric(Mid(strLicense, len(strLicense)-3, 4)) Then intAbout = CInt(Mid(strLicense, len(strLicense)-3, 4))
		strDisplayLicense = Mid(strLicense, 4, 40)
		about = CBool(intAbout MOD 18) 
	End If
End Function
 

'Get CAPTCHA info
Private Sub captchaInfo()

	Response.Write("" & _
	vbCrLf & "<pre>" & _
	vbCrLf & "*********************************************************" & _
	vbCrLf & "Software: Web Wiz CAPTCHA(TM)" & _
	vbCrLf & "Version: " & strCAPTCHAversion & _
	vbCrLf & "License: " & strDisplayLicense & _
	vbCrLf & "Author: Web Wiz(TM)." & _
	vbCrLf & "Address: Unit 10E, Dawkins Raod Ind Est, Poole, Dorset, UK" & _
	vbCrLf & "Info: http://www.webwiznewspad.com" & _
	vbCrLf & "Copyright: (C)2001-2008 Web Wiz(TM). All rights reserved" & _
	vbCrLf & "*********************************************************" & _
	vbCrLf & "</pre")
	
	Response.Flush
	Response.End
End Sub

%>