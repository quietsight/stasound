<% @ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="CAPTCHA_configuration.asp" --> 
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

'Create random CAPTCHA bitmap of 50 x 120 pixels

'Set the buffer to true
Response.Buffer = True

'Set browser headers
Response.Clear 
Response.ContentType = "image/bmp"
Response.AddHeader "Content-Disposition", "inline; filename=captcha.bmp"
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.CacheControl = "No-Store"
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"





'Declare variables
Dim strHexImage
Dim intRandomNumber
Dim strBlankCanvas
Dim strCanvas
Dim sarryCanvas(49,119)
Dim intRowLoop
Dim intColumnLoop
Dim intLoop
Dim sarryCharacter(17,32)
Dim strCAPTCHAcode
Dim sarrySkewCharcater(4)
Dim intCharPixelSpacing
Dim blnSkewSet





'Reverse the HEX colours
strCanvasColour = bmpHexColourSwicth(strCanvasColour)
strBorderColour = bmpHexColourSwicth(strBorderColour)
strCharacterColour = bmpHexColourSwicth(strCharacterColour)
strNoiseColour1 = bmpHexColourSwicth(strNoiseColour1)
strNoiseLinesColour1 = bmpHexColourSwicth(strNoiseLinesColour1)
strNoiseColour2 = bmpHexColourSwicth(strNoiseColour2)
strNoiseLinesColour2 = bmpHexColourSwicth(strNoiseLinesColour2)

'Randomise the system timer for genretaing random numbers later
Randomize Timer

'Initilise varaibles
Session("strSecurityCode") = ""
blnSkewSet = False
intCharPixelSpacing = 0






'Character 2
sarryCharacter(0,24) = "6,7,8,9,10"
sarryCharacter(0,23) = "4,5,6,10,11,12"
sarryCharacter(0,22) = "3,4,5,6,7,8,10,13"
sarryCharacter(0,21) = "3,4,9,10,12,13,14"
sarryCharacter(0,20) = "2,3,10,13,14"
sarryCharacter(0,19) = "2,11,13,14"
sarryCharacter(0,18) = "11,13,14"
sarryCharacter(0,17) = "11,13,14"
sarryCharacter(0,16) = "11,13"
sarryCharacter(0,14) = "11,12,13"
sarryCharacter(0,14) = "10,12"
sarryCharacter(0,13) = "10,11,12"
sarryCharacter(0,12) = "9,11"
sarryCharacter(0,11) = "9,10"
sarryCharacter(0,10) = "8,9"
sarryCharacter(0,9) = "8"
sarryCharacter(0,8) = "7"
sarryCharacter(0,7) = "6"
sarryCharacter(0,6) = "5,15"
sarryCharacter(0,5) = "4,13,14,15"
sarryCharacter(0,4) = "3,4,5,6,7,8,9,10,11,12,13,14"
sarryCharacter(0,3) = "2,3,4,5,6,7,8,9,10,11,12,14"
sarryCharacter(0,2) = "1,2,13,14"
sarryCharacter(0,1) = "1,2,3,4,5,6,7,8,9,10,11,12,13"


'Character 3
sarryCharacter(1,24) = "4,5,6,7,8,9,10,11,12,13,14,15,16,17,18"
sarryCharacter(1,23) = ""
sarryCharacter(1,22) = ""
sarryCharacter(1,21) = "4,5,6,7,8,9,10,11,12,13,14,15"
sarryCharacter(1,20) = "4,13,14"
sarryCharacter(1,19) = "3,12,13"
sarryCharacter(1,18) = "11,12"
sarryCharacter(1,17) = "10,11"
sarryCharacter(1,16) = "9,10"
sarryCharacter(1,15) = "9,10"
sarryCharacter(1,14) = "8,9,10,11"
sarryCharacter(1,13) = "7,11,12,13,15"
sarryCharacter(1,12) = "13,14,16"
sarryCharacter(1,11) = "13,14,17"
sarryCharacter(1,10) = "14,15,17,18"
sarryCharacter(1,9) = "14,15,17,18"
sarryCharacter(1,8) = "14,15,17,18"
sarryCharacter(1,7) = "14,15,17,18"
sarryCharacter(1,6) = "14,15,17,18"
sarryCharacter(1,5) = "2,14,17,18"
sarryCharacter(1,4) = "3,13,14,17"
sarryCharacter(1,3) = "3,12,13,16"
sarryCharacter(1,2) = "4,5,11,12,15"
sarryCharacter(1,1) = "5,6,7,8,9,10"




'Character 5
sarryCharacter(2,26) = "17"
sarryCharacter(2,25) = "8,9,10,11,12,13,14,15,16,17"
sarryCharacter(2,24) = "8,9,10,11,12,13,14,15,16,17"
sarryCharacter(2,23) = ""
sarryCharacter(2,22) = "7,8,9,10"
sarryCharacter(2,21) = "7,9,10"
sarryCharacter(2,20) = "7,9,10"
sarryCharacter(2,19) = "6,7,8,9"
sarryCharacter(2,18) = "6,8,9"
sarryCharacter(2,17) = "6,8,9"
sarryCharacter(2,16) = "5,8"
sarryCharacter(2,15) = "5,6,7,8,9,10,11,12,13"
sarryCharacter(2,14) = "5,6,7,8,9,10,11,12,13"
sarryCharacter(2,13) = "10,13,14"
sarryCharacter(2,12) = "11,12,14,15"
sarryCharacter(2,11) = "13,15,16"
sarryCharacter(2,10) = "14,16,17"
sarryCharacter(2,9) = "14,16,17"
sarryCharacter(2,8) = "14,16,17"
sarryCharacter(2,7) = "14,16,17"
sarryCharacter(2,6) = "14,16,17"
sarryCharacter(2,5) = "14,16"
sarryCharacter(2,4) = "13,15,16"
sarryCharacter(2,3) = "12,13,15"
sarryCharacter(2,2) = "11,12,14"
sarryCharacter(2,1) = "9,10,12,13"
sarryCharacter(2,0) = "3,4,5,6,7,8,9,10,11"



'Character 6
sarryCharacter(3,21) = "13,14,15"
sarryCharacter(3,20) = "11,12"
sarryCharacter(3,19) = "9,10,11"
sarryCharacter(3,18) = "8,9,10"
sarryCharacter(3,17) = "8,9"
sarryCharacter(3,16) = "7,8"
sarryCharacter(3,15) = "5,7"
sarryCharacter(3,14) = "4,6,7"
sarryCharacter(3,13) = "4,6,7,8,9,10,11,12,13"
sarryCharacter(3,12) = "3,4,6,7,13,14,16"
sarryCharacter(3,11) = "3,4,6,14,15,17"
sarryCharacter(3,10) = "3,6,15,17"
sarryCharacter(3,9) = "3,6,15,17,18"
sarryCharacter(3,8) = "3,6,15,18"
sarryCharacter(3,7) = "3,6,15,18"
sarryCharacter(3,6) = "3,6,15,18"
sarryCharacter(3,5) = "4,6,15,18"
sarryCharacter(3,4) = "4,6,15,17,18"
sarryCharacter(3,3) = "4,6,7,15,17"
sarryCharacter(3,2) = "4,7,14,15,16"
sarryCharacter(3,1) = "7,8,13,14"
sarryCharacter(3,0) = "9,10,11,12"




'Character 7
sarryCharacter(4,29) = "6"
sarryCharacter(4,28) = "5,6"
sarryCharacter(4,27) = "5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"
sarryCharacter(4,26) = "5,6,7,8,9,10,11,12,13,14,15,16,17,18,19"
sarryCharacter(4,25) = "4,5,6,7,19"
sarryCharacter(4,24) = "4,5,18,19"
sarryCharacter(4,23) = "4,5,6,7,8,9,10,11,12,13,14,15,16,17,18"
sarryCharacter(4,22) = "3,4,5,18"
sarryCharacter(4,21) = "3,17"
sarryCharacter(4,20) = "3,17"
sarryCharacter(4,19) = "2,16,17"
sarryCharacter(4,18) = "16"
sarryCharacter(4,17) = "15,16"
sarryCharacter(4,16) = "15,16"
sarryCharacter(4,15) = "14,15"
sarryCharacter(4,14) = "14,15"
sarryCharacter(4,13) = "13,14"
sarryCharacter(4,12) = "13,14"
sarryCharacter(4,11) = "12,13,14"
sarryCharacter(4,10) = "12,13"
sarryCharacter(4,9) = "11,12,13"
sarryCharacter(4,8) = "11,12,13"
sarryCharacter(4,7) = "10,11,12"
sarryCharacter(4,6) = "10,11,12"
sarryCharacter(4,5) = "9,10,11,12"
sarryCharacter(4,4) = "9,10,11"
sarryCharacter(4,3) = "8,9,10,11"
sarryCharacter(4,2) = "8,9,10"
sarryCharacter(4,1) = "8,9,10"
sarryCharacter(4,0) = "8,9"




'Character 8
sarryCharacter(5,25) = "8,9,10,11,12,13,14"
sarryCharacter(5,24) = "6,7,8,14,15,16"
sarryCharacter(5,23) = "5,7,15,17"
sarryCharacter(5,22) = "5,7,16,17,18"
sarryCharacter(5,21) = "4,6,7,16,18"
sarryCharacter(5,20) = "4,6,7,16,18,19"
sarryCharacter(5,19) = "4,6,7,16,18,19"
sarryCharacter(5,18) = "4,6,7,8,16,18"
sarryCharacter(5,17) = "4,7,8,16,18"
sarryCharacter(5,16) = "5,7,8,9,15,17,18"
sarryCharacter(5,15) = "6,8,9,10,11,14,16,17"
sarryCharacter(5,14) = "7,10,11,12,13,15"
sarryCharacter(5,13) = "8,12,13,14"
sarryCharacter(5,12) = "6,7,8,9,14,15,16"
sarryCharacter(5,11) = "4,5,6,11,12,16,17,18"
sarryCharacter(5,10) = "3,5,6,13,14,17,18,19"
sarryCharacter(5,9) = "3,5,15,16,18,19,20"
sarryCharacter(5,8) = "2,4,5,17,19,20"
sarryCharacter(5,7) = "2,4,5,17,19,20,21"
sarryCharacter(5,6) = "2,4,5,17,20,21"
sarryCharacter(5,5) = "2,4,5,17,20"
sarryCharacter(5,4) = "2,4,5,17,19,20"
sarryCharacter(5,3) = "3,5,6,17,19,20"
sarryCharacter(5,2) = "4,6,7,16,19"
sarryCharacter(5,1) = "5,7,8,15,16,18"
sarryCharacter(5,0) = "6,7,8,9,10,11,12,13,14,15,16,17"







'Character B
sarryCharacter(6,28) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16"
sarryCharacter(6,27) = "4,6,7,8,9,16,18,19"
sarryCharacter(6,26) = "5,6,7,8,17,19,20"
sarryCharacter(6,25) = "5,6,7,8,17,20,21"
sarryCharacter(6,24) = "5,6,7,8,18,20,21,22"
sarryCharacter(6,23) = "5,6,7,8,18,20,21,22"
sarryCharacter(6,22) = "5,6,7,8,18,21,22"
sarryCharacter(6,21) = "5,6,7,8,18,21,22"
sarryCharacter(6,20) = "5,6,7,8,18,21,22"
sarryCharacter(6,19) = "5,6,7,8,18,20,21,22"
sarryCharacter(6,18) = "5,6,7,8,18,20,21,22"
sarryCharacter(6,17) = "5,6,7,8,17,19,20,21"
sarryCharacter(6,16) = "5,6,7,8,16,18,19,20"
sarryCharacter(6,15) = "5,6,7,8,8,9,10,11,12,13,14,15,16,17,18"
sarryCharacter(6,14) = "5,6,7,8,8,9,10,11,12,13,14,15,16,17"
sarryCharacter(6,13) = "5,6,7,8,16,18,19,20"
sarryCharacter(6,12) = "5,6,7,8,17,18,20,21"
sarryCharacter(6,11) = "5,6,7,8,18,20,21,22"
sarryCharacter(6,10) = "5,6,7,8,19,21,22"
sarryCharacter(6,9) = "5,6,7,8,19,21,22,23"
sarryCharacter(6,8) = "5,6,7,8,19,21,22,23"
sarryCharacter(6,7) = "5,6,7,8,19,21,22,23"
sarryCharacter(6,6) = "5,6,7,8,19,21,22,23"
sarryCharacter(6,5) = "5,6,7,8,19,21,22,23"
sarryCharacter(6,4) = "5,6,7,8,19,21,22"
sarryCharacter(6,3) = "5,6,7,8,18,20,21,22"
sarryCharacter(6,2) = "5,6,7,8,17,19,20,21"
sarryCharacter(6,1) = "4,6,7,8,9,16,18,19,20"
sarryCharacter(6,0) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"




'Character C
sarryCharacter(7,23) = "9,10,11,12,13,14,15,18"
sarryCharacter(7,22) = "7,8,9,14,15,16,17,18"
sarryCharacter(7,21) = "5,6,7,16,17,18"
sarryCharacter(7,20) = "4,5,6,16,17,18"
sarryCharacter(7,19) = "4,5,6,17,18"
sarryCharacter(7,18) = "3,4,5,17,18"
sarryCharacter(7,17) = "3,4,5,18"
sarryCharacter(7,16) = "3,4"
sarryCharacter(7,15) = "2,3,4"
sarryCharacter(7,14) = "2,3,4"
sarryCharacter(7,13) = "2,3,4"
sarryCharacter(7,12) = "2,3,4"
sarryCharacter(7,11) = "2,3,4"
sarryCharacter(7,10) = "2,3,4"
sarryCharacter(7,9) = "2,3,4"
sarryCharacter(7,8) = "2,3,4"
sarryCharacter(7,7) = "2,3,4"
sarryCharacter(7,6) = "3,4,5,18"
sarryCharacter(7,5) = "3,4,5,18"
sarryCharacter(7,4) = "3,4,5,17"
sarryCharacter(7,3) = "4,5,6,17"
sarryCharacter(7,2) = "5,6,7,16"
sarryCharacter(7,1) = "6,7,8,15,16"
sarryCharacter(7,0) = "8,9,10,11,12,13,14"


'Character D
sarryCharacter(8,24) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15"
sarryCharacter(8,23) = "3,5,6,7,8,14,15,16,17"
sarryCharacter(8,22) = "4,6,7,15,17,18"
sarryCharacter(8,21) = "4,6,7,16,18,19"
sarryCharacter(8,20) = "4,6,7,17,19,20"
sarryCharacter(8,19) = "4,6,7,18,20"
sarryCharacter(8,18) = "4,6,7,18,20,21"
sarryCharacter(8,17) = "4,6,7,18,20,21"
sarryCharacter(8,16) = "4,6,7,19,21,22"
sarryCharacter(8,15) = "4,6,7,19,21,22"
sarryCharacter(8,14) = "4,6,7,19,21,22"
sarryCharacter(8,13) = "4,6,7,19,21,22"
sarryCharacter(8,12) = "4,6,7,19,21,22"
sarryCharacter(8,11) = "4,6,7,19,21,22"
sarryCharacter(8,10) = "4,6,7,19,21,22"
sarryCharacter(8,9) = "4,6,7,19,21,22"
sarryCharacter(8,8) = "4,6,7,19,21,22"
sarryCharacter(8,7) = "4,6,7,18,20,21"
sarryCharacter(8,6) = "4,6,7,18,20,21"
sarryCharacter(8,5) = "4,6,7,18,20"
sarryCharacter(8,4) = "4,6,7,17,19,20"
sarryCharacter(8,3) = "4,6,7,16,17,19"
sarryCharacter(8,2) = "4,6,7,15,16,18"
sarryCharacter(8,1) = "3,5,6,7,8,14,16,17"
sarryCharacter(8,0) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15"


'Character E
sarryCharacter(9,26) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"
sarryCharacter(9,25) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"
sarryCharacter(9,24) = "3,5,6,7,8,9,16,19,20,21"
sarryCharacter(9,23) = "4,6,7,8,17,18,20,21"
sarryCharacter(9,22) = "4,6,7,8,18,19,20,21"
sarryCharacter(9,21) = "4,6,7,8,20,21"
sarryCharacter(9,20) = "4,6,7,8,21"
sarryCharacter(9,19) = "4,6,7,8"
sarryCharacter(9,18) = "4,6,7,8"
sarryCharacter(9,17) = "4,6,7,8,15,16"
sarryCharacter(9,16) = "4,6,7,8,15,16"
sarryCharacter(9,15) = "4,6,7,8,15,16"
sarryCharacter(9,14) = "4,6,7,8,13,14,15,16"
sarryCharacter(9,13) = "4,5,6,7,8,8,9,10,11,12,15,16"
sarryCharacter(9,12) = "4,6,7,8,13,14,15,16"
sarryCharacter(9,11) = "4,6,7,8,15,16"
sarryCharacter(9,10) = "4,6,7,8,15,16"
sarryCharacter(9,9) = "4,6,7,8,15,16"
sarryCharacter(9,8) = "4,6,7,8"
sarryCharacter(9,7) = "4,6,7,8"
sarryCharacter(9,6) = "4,6,7,8,22"
sarryCharacter(9,5) = "4,6,7,8,21,22"
sarryCharacter(9,4) = "4,6,7,8,20,21,22"
sarryCharacter(9,3) = "4,6,7,8,18,19,20,21"
sarryCharacter(9,2) = "4,6,7,8,17,19,20,21"
sarryCharacter(9,1) = "3,5,6,7,8,9,16,19,20,21"
sarryCharacter(9,0) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"


'Character F
sarryCharacter(10,24) = "2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18"
sarryCharacter(10,23) = "4,6,7,8,9,15,18"
sarryCharacter(10,22) = "5,7,8,16,18,19"
sarryCharacter(10,21) = "5,7,8,17,18,19"
sarryCharacter(10,20) = "5,7,8,18,19"
sarryCharacter(10,19) = "5,7,8,19"
sarryCharacter(10,18) = "5,7,8,20"
sarryCharacter(10,17) = "5,7,8,15,16"
sarryCharacter(10,16) = "5,7,8,15,16"
sarryCharacter(10,15) = "5,7,8,15,16"
sarryCharacter(10,14) = "5,7,8,14,15,16"
sarryCharacter(10,13) = "5,7,8,13,15,16"
sarryCharacter(10,12) = "5,7,8,9,10,11,12,15,16"
sarryCharacter(10,11) = "5,7,8,13,14,15,16"
sarryCharacter(10,10) = "5,7,8,15,16"
sarryCharacter(10,9) = "5,7,8,15,16"
sarryCharacter(10,8) = "5,7,8,15,16"
sarryCharacter(10,7) = "5,7,8"
sarryCharacter(10,6) = "5,7,8"
sarryCharacter(10,5) = "5,7,8"
sarryCharacter(10,4) = "5,7,8"
sarryCharacter(10,3) = "5,7,8"
sarryCharacter(10,2) = "5,7,8"
sarryCharacter(10,1) = "4,6,7,8,9"
sarryCharacter(10,0) = "2,3,4,5,6,7,8,9,10,11"



'Character H
sarryCharacter(11,22) = "1,2,3,4,5,6,7,8,9,15,16,17,18,19,20,21,22,23"
sarryCharacter(11,21) = "3,4,5,6,7,18,19,20,21"
sarryCharacter(11,20) = "4,5,6,18,19,20"
sarryCharacter(11,19) = "4,5,6,18,19,20"
sarryCharacter(11,18) = "4,5,6,18,19,20"
sarryCharacter(11,17) = "4,5,6,18,19,20"
sarryCharacter(11,16) = "4,5,6,18,19,20"
sarryCharacter(11,15) = "4,5,6,18,19,20"
sarryCharacter(11,14) = "4,5,6,18,19,20"
sarryCharacter(11,13) = "4,5,6,18,19,20"
sarryCharacter(11,12) = "4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"
sarryCharacter(11,11) = "4,5,6,18,19,20"
sarryCharacter(11,10) = "4,5,6,18,19,20"
sarryCharacter(11,9) = "4,5,6,18,19,20"
sarryCharacter(11,8) = "4,5,6,18,19,20"
sarryCharacter(11,7) = "4,5,6,18,19,20"
sarryCharacter(11,6) = "4,5,6,18,19,20"
sarryCharacter(11,5) = "4,5,6,18,19,20"
sarryCharacter(11,4) = "4,5,6,18,19,20"
sarryCharacter(11,3) = "4,5,6,18,19,20"
sarryCharacter(11,2) = "4,5,6,7,17,18,19,20"
sarryCharacter(11,1) = "3,4,5,6,7,17,18,19,20,21"
sarryCharacter(11,0) = "1,2,3,4,5,6,7,8,9,15,16,17,18,19,20,21,22,23"





'Character K
sarryCharacter(12,27) = "1,2,3,4,5,6,7,8,9,10,13,14,15,16,17,18,19,20"
sarryCharacter(12,26) = "3,5,6,7,8,14,18,19"
sarryCharacter(12,25) = "4,6,7,15,17"
sarryCharacter(12,24) = "4,6,7,15,16"
sarryCharacter(12,23) = "4,6,7,15"
sarryCharacter(12,22) = "4,6,7,14,15"
sarryCharacter(12,21) = "4,6,7,14"
sarryCharacter(12,20) = "4,6,7,13"
sarryCharacter(12,19) = "4,6,7,12"
sarryCharacter(12,18) = "4,6,7,11"
sarryCharacter(12,17) = "4,6,7,11"
sarryCharacter(12,16) = "4,6,7,10,11,12"
sarryCharacter(12,15) = "4,6,7,9,11,12"
sarryCharacter(12,14) = "4,6,7,8,9,11,12,13"
sarryCharacter(12,13) = "4,6,7,8,10,12,13"
sarryCharacter(12,12) = "4,6,7,10,13,14"
sarryCharacter(12,11) = "4,6,7,11,13,14"
sarryCharacter(12,10) = "4,6,7,11,14,15"
sarryCharacter(12,9) = "4,6,7,12,14,15"
sarryCharacter(12,8) = "4,6,7,12,15,16"
sarryCharacter(12,7) = "4,6,7,13,16,17"
sarryCharacter(12,6) = "4,6,7,13,16,17"
sarryCharacter(12,5) = "4,6,7,14,17,18"
sarryCharacter(12,4) = "4,6,7,15,17,18"
sarryCharacter(12,3) = "4,6,7,16,18,19"
sarryCharacter(12,2) = "4,6,7,16,18,19"
sarryCharacter(12,1) = "3,5,6,7,8,16,17,18,19,20"
sarryCharacter(12,0) = "1,2,3,4,5,6,7,8,9,10,13,14,15,16,17,18,19,20,21"





'Character N
sarryCharacter(13,22) = "2,3,4,5,6,7,8,15,16,17,18,19,20,21,22"
sarryCharacter(13,21) = "4,5,6,7,8,17,18,19,20"
sarryCharacter(13,20) = "5,6,7,8,9,18,19"
sarryCharacter(13,19) = "5,6,7,8,9,18,19"
sarryCharacter(13,18) = "5,6,7,8,9,10,18,19"
sarryCharacter(13,17) = "5,6,7,8,9,10,11,18,19"
sarryCharacter(13,16) = "5,6,8,9,10,11,18,19"
sarryCharacter(13,15) = "5,6,9,10,11,12,18,19"
sarryCharacter(13,14) = "5,6,9,10,11,12,13,18,19"
sarryCharacter(13,13) = "5,6,10,11,12,13,18,19"
sarryCharacter(13,12) = "5,6,10,11,12,13,14,18,19"
sarryCharacter(13,11) = "5,6,11,12,13,14,18,19"
sarryCharacter(13,10) = "5,6,12,13,14,15,18,19"
sarryCharacter(13,9) = "5,6,12,13,14,15,16,18,19"
sarryCharacter(13,8) = "5,6,13,14,15,16,18,19"
sarryCharacter(13,7) = "5,6,14,15,16,17,18,19"
sarryCharacter(13,6) = "5,6,14,15,16,17,18,19"
sarryCharacter(13,5) = "5,6,15,16,17,18,19"
sarryCharacter(13,4) = "5,6,16,17,18,19"
sarryCharacter(13,3) = "5,6,16,17,18,19"
sarryCharacter(13,2) = "5,6,17,18,19"
sarryCharacter(13,1) = "4,5,6,7,17,18,19"
sarryCharacter(13,0) = "2,3,4,5,6,7,8,9,18,19"




'Character P
sarryCharacter(14,27) = "2,3,4,5,6,7,8,9,10,11,12,13,14,15"
sarryCharacter(14,26) = "4,6,7,8,9,14,15,16,17"
sarryCharacter(14,25) = "5,7,8,15,17,18"
sarryCharacter(14,24) = "5,7,8,16,17,18,19"
sarryCharacter(14,23) = "5,7,8,17,18,19"
sarryCharacter(14,22) = "5,7,8,17,18,19,20"
sarryCharacter(14,21) = "5,7,8,17,18,19,20"
sarryCharacter(14,20) = "5,7,8,17,18,19,20"
sarryCharacter(14,19) = "5,7,8,17,18,19,20"
sarryCharacter(14,18) = "5,7,8,17,18,19,20"
sarryCharacter(14,17) = "5,7,8,17,18,19,20"
sarryCharacter(14,16) = "5,7,8,16,17,18,19"
sarryCharacter(14,15) = "5,7,8,16,17,18"
sarryCharacter(14,14) = "5,7,8,14,15,16,17"
sarryCharacter(14,13) = "5,7,8,9,10,11,12,13,14,15"
sarryCharacter(14,12) = "5,7,8"
sarryCharacter(14,11) = "5,7,8"
sarryCharacter(14,10) = "5,7,8"
sarryCharacter(14,9) = "5,7,8"
sarryCharacter(14,8) = "5,7,8"
sarryCharacter(14,7) = "5,7,8"
sarryCharacter(14,6) = "5,7,8"
sarryCharacter(14,5) = "5,7,8"
sarryCharacter(14,4) = "5,7,8"
sarryCharacter(14,3) = "5,7,8"
sarryCharacter(14,2) = "5,7,8"
sarryCharacter(14,1) = "4,6,7,8,9"
sarryCharacter(14,0) = "2,3,4,5,6,7,8,9,10,11,12"




'Character R
sarryCharacter(15,23) = "1,2,3,4,5,6,7,8,9,10,11,12"
sarryCharacter(15,22) = "4,6,7,8,9,10,13,14"
sarryCharacter(15,21) = "4,6,7,11,14,15"
sarryCharacter(15,20) = "4,6,7,12,15,16"
sarryCharacter(15,19) = "4,6,7,13,15,16"
sarryCharacter(15,18) = "4,6,7,13,15,16"
sarryCharacter(15,17) = "4,6,7,13,15,16"
sarryCharacter(15,16) = "4,6,7,13,15,16"
sarryCharacter(15,15) = "4,6,7,13,15,16"
sarryCharacter(15,14) = "4,6,7,13,15,16"
sarryCharacter(15,13) = "4,6,7,12,14,15"
sarryCharacter(15,12) = "4,6,7,11,13"
sarryCharacter(15,11) = "4,6,7,8,9,10,11,12"
sarryCharacter(15,10) = "4,6,7,11,12,13"
sarryCharacter(15,9) = "4,6,7,11,13,14"
sarryCharacter(15,8) = "4,6,7,12,14"
sarryCharacter(15,7) = "4,6,7,12,14,15"
sarryCharacter(15,6) = "4,6,7,13,15"
sarryCharacter(15,5) = "4,6,7,13,15,16"
sarryCharacter(15,4) = "4,6,7,14,16,17"
sarryCharacter(15,3) = "4,6,7,14,17"
sarryCharacter(15,2) = "4,6,7,15,17,18"
sarryCharacter(15,1) = "4,6,7,16,18"
sarryCharacter(15,0) = "1,2,3,4,5,6,7,8,9,10,17,18,19,20"





'Character S
sarryCharacter(16,25) = "6,7,8,9,10,14,15"
sarryCharacter(16,24) = "4,5,11,12,13,14,15"
sarryCharacter(16,23) = "3,4,12,14,15"
sarryCharacter(16,22) = "3,13,15"
sarryCharacter(16,21) = "2,3,13,15,16"
sarryCharacter(16,20) = "2,3,14,15,16"
sarryCharacter(16,19) = "2,3,15,16"
sarryCharacter(16,18) = "2,3,15,16"
sarryCharacter(16,17) = "2,3,4"
sarryCharacter(16,16) = "2,4,5,6,7"
sarryCharacter(16,15) = "3,4,5,6,7,8,9,10,11"
sarryCharacter(16,14) = "3,6,7,8,9,10,11,12,13,14"
sarryCharacter(16,13) = "4,8,9,10,11,12,13,14,15"
sarryCharacter(16,12) = "5,6,12,13,14,15,16"
sarryCharacter(16,11) = "7,8,9,10,11,15,16,17"
sarryCharacter(16,10) = "12,13,14,16,17,18"
sarryCharacter(16,9) = "15,17,18"
sarryCharacter(16,8) = "2,3,15,17,18"
sarryCharacter(16,7) = "2,3,16,17,18"
sarryCharacter(16,6) = "2,3,4,16,17,18"
sarryCharacter(16,5) = "2,3,4,16,17,18"
sarryCharacter(16,4) = "2,4,5,16,17"
sarryCharacter(16,3) = "2,4,5,15,16,17"
sarryCharacter(16,2) = "2,5,6,15,16"
sarryCharacter(16,1) = "2,3,4,6,7,8,13,14,15"
sarryCharacter(16,0) = "2,3,9,10,11,12,13"




'Character T
sarryCharacter(17,23) = "2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21"
sarryCharacter(17,22) = "2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21"
sarryCharacter(17,21) = "2,9,3,10,12,13,14,20,21"
sarryCharacter(17,20) = "2,10,12,13,14,21"
sarryCharacter(17,19) = "2,10,12,13,14,21"
sarryCharacter(17,18) = "2,10,12,13,14,21"
sarryCharacter(17,17) = "10,12,13,14"
sarryCharacter(17,16) = "10,12,13,14"
sarryCharacter(17,15) = "10,12,13,14"
sarryCharacter(17,14) = "10,12,13,14"
sarryCharacter(17,13) = "10,12,13,14"
sarryCharacter(17,12) = "10,12,13,14"
sarryCharacter(17,11) = "10,12,13,14"
sarryCharacter(17,10) = "10,12,13,14"
sarryCharacter(17,9) = "10,12,13,14"
sarryCharacter(17,8) = "10,12,13,14"
sarryCharacter(17,7) = "10,12,13,14"
sarryCharacter(17,6) = "10,12,13,14"
sarryCharacter(17,5) = "10,12,13,14"
sarryCharacter(17,4) = "10,12,13,14"
sarryCharacter(17,3) = "10,12,13,14"
sarryCharacter(17,2) = "10,12,13,14"
sarryCharacter(17,1) = "10,12,13,14"
sarryCharacter(17,0) = "7,8,9,10,11,12,13,14,15,16,17"



'Initlaise the blank canvas and write binary output
Call writeBinary("424D864600000000000036000000280000007800000032000000010018000000000050460000120B0000120B00000000000000000000")



'Key code for canvas pixels
'0 - background
'1 - border
'2 - character
'3 - Noise #1
'4 - Line Noise #1
'5 - Noise #2
'6 - Line Noise #2


'If skewing is enabled then workout which charecters are going to be italic
If blnSkewing Then

	'Loop through selecting random charcater for skewing
	For intLoop = 0 To 4
		If CInt(Rnd * 5) > 3 Then 
			sarrySkewCharcater(intLoop) = True
			blnSkewSet = True
		End If
	Next
	
	'If none are set choice just one random character
	If blnSkewSet = False Then sarrySkewCharcater(CInt(Rnd * 4)) = True
	
End If





'Genreate random CAPTCHA code
For intLoop = 0 To 4

	'Set the number of pixels between characters
	intCharPixelSpacing = 24
	
	'Calculate current charcter pixel position
	intCharPixelSpacing = intLoop * intCharPixelSpacing

	'Create random number 
	intRandomNumber = CInt(Rnd * UBound(sarryCharacter))
	
	'Get the number of character
	Select Case intRandomNumber
		Case 0
			strCAPTCHAcode = strCAPTCHAcode & "2"
		Case 1
			strCAPTCHAcode = strCAPTCHAcode & "3"
		Case 2
			strCAPTCHAcode = strCAPTCHAcode & "5"
		Case 3
			strCAPTCHAcode = strCAPTCHAcode & "6"
		Case 4
			strCAPTCHAcode = strCAPTCHAcode & "7"
		Case 5
			strCAPTCHAcode = strCAPTCHAcode & "8"
		Case 6
			strCAPTCHAcode = strCAPTCHAcode & "B"
		Case 7
			strCAPTCHAcode = strCAPTCHAcode & "C"
		Case 8
			strCAPTCHAcode = strCAPTCHAcode & "D"
		Case 9
			strCAPTCHAcode = strCAPTCHAcode & "E"
		Case 10
			strCAPTCHAcode = strCAPTCHAcode & "F"
		Case 11
			strCAPTCHAcode = strCAPTCHAcode & "H"
		Case 12
			strCAPTCHAcode = strCAPTCHAcode & "K"
		Case 13
			strCAPTCHAcode = strCAPTCHAcode & "N"
		Case 14
			strCAPTCHAcode = strCAPTCHAcode & "P"
		Case 15
			strCAPTCHAcode = strCAPTCHAcode & "R"
		Case 16
			strCAPTCHAcode = strCAPTCHAcode & "S"
		Case 17
			strCAPTCHAcode = strCAPTCHAcode & "T"
	End Select

	'Set the session varaible with the CAPTCHA code
	Session("strSecurityCode") = strCAPTCHAcode
	
	'If skewing is enabled see if this character needs to be skewed
	If blnSkewing Then
		
		'If this charcater is to be in skewed then change the line spacing (depending on canvas placemnet)
		If sarrySkewCharcater(intLoop) Then
			Select Case intLoop
				Case 0
					intCharPixelSpacing = 0
				Case 1
					intCharPixelSpacing = intCharPixelSpacing - 8
				Case 2
					intCharPixelSpacing = intCharPixelSpacing - 8
				Case 3
					intCharPixelSpacing = intCharPixelSpacing - 8
				Case 4
					intCharPixelSpacing = intCharPixelSpacing - 12
			End Select
		End If
	End If
	
	
	'Map character to canvas
	Call mapCharacter(intRandomNumber, intCharPixelSpacing, sarrySkewCharcater(intLoop))
	
Next


'0 - background
'1 - border
'2 - character
'3 - Noise #1
'4 - Line Noise #1
'5 - Noise #2
'6 - Line Noise #2


'Add noise #1 lines to canvas
Call  mapCanvasLineNoise(intNoiseLines1, "4")

'Add noise #2 lines to canvas
Call  mapCanvasLineNoise(intNoiseLines2, "6")

'Add noise #1 to canvas
Call  mapCanvasNoise(intNoiseLevel1, "3")

'Add noise #2 to canvas
Call  mapCanvasNoise(intNoiseLevel2, "5")

'Map the broder last so that it overlays anything else on the canvas
Call mapBorder()






'Build the flat 50 x 120 canvas string from the array
'Loop through the rows (50 rows in all)
For intRowLoop = 0 to 49

	'Loop throuh the coulmns (120 in all)
	For intColumnLoop = 0 TO 119
		
		'If border (1) then set this pixel as the border colour
		If sarryCanvas(intRowLoop, intColumnLoop) = "1" Then
			writeBinary(strBorderColour)
			
		'If character (2) then set this pixel as the character colour
		ElseIf sarryCanvas(intRowLoop, intColumnLoop) = "2" Then
			writeBinary(strCharacterColour)
		
		'If noise #1 (3) then set this pixel as the noise #1 colour
		ElseIf sarryCanvas(intRowLoop, intColumnLoop) = "3" Then
			writeBinary(strNoiseColour1)
		
		'If noise line #1 (4) then set this pixel as the line noise colour
		ElseIf sarryCanvas(intRowLoop, intColumnLoop) = "4" Then
			writeBinary(strNoiseLinesColour1)
			
		'If noise #2 (5) then set this pixel as the noise #2 colour
		ElseIf sarryCanvas(intRowLoop, intColumnLoop) = "5" Then
			writeBinary(strNoiseColour2)
		
		'If noise line #2 (5) then set this pixel as the line noise #2 colour
		ElseIf sarryCanvas(intRowLoop, intColumnLoop) = "6" Then
			writeBinary(strNoiseLinesColour2)
			
		'Else set as the canvas background colour
		Else
			writeBinary(strCanvasColour)
		
		End If
		
		
	Next
Next



'Send buffered client response to the client
Response.Flush




'******************************************
'***  Converet HEX to Binary Output   *****
'******************************************

'convert hex to binary output for browser
Private Sub writeBinary(strHex)

	Dim lngHexLoop
	
	'Remove any spaces (Error handling (there should not be any spaces anyway))
	strHex = Replace(strHex, " ", "")

	'Loop through the in steps of 2 for each 2 character hex
	For lngHexLoop = 1 to Len(strHex) step 2
		'Write the binary output
		Response.BinaryWrite ChrB(CByte("&H" & Mid(strHex, lngHexLoop, 2)))
	Next
End Sub



'******************************************
'***  	Map Character to Canvas	      *****
'******************************************

'Map Charcter Sub, builds the chacater in the canvas
Private Sub mapCharacter(intCharacter, intCharacterPostion, blnItalicChar)
	
	Dim sarryPixels
	Dim intPixelPosition
	Dim intPixelLoop
	Dim intRandomLinePlace
	Dim intCharcterHeight
	
	'Max charcter height
	'First lets see if this is s short charceter
	If sarryCharacter(intCharacter, 25) = "" Then
		intCharcterHeight = 24 'Short chracter so it can be placed higher on the canvas
	Else
		intCharcterHeight = 29 'Tallest charcter
	End If
	
	'Generate a random number for line placement
	If blnRandomLinePlacement Then
		intRandomLinePlace = CInt(Rnd * (49 - intCharcterHeight))
	Else
		intRandomLinePlace = 12
	End If
	
	'Loop through each row of the character
	For intRowLoop = 0 to intCharcterHeight
	
		'If an skew charcter move pixels diagonal
		If blnItalicChar Then
			If intRowLoop MOD 2 Then intCharacterPostion = intCharacterPostion + 1
		End If
	
		'Split the pixels to be used into an array
		sarryPixels = Split(sarryCharacter(intCharacter, intRowLoop), ",")
		
		'Loop through each pixel
		For intPixelLoop = 0 to Ubound(sarryPixels)
		
			'Calculate the pixel postion
			intPixelPosition = intCharacterPostion + sarryPixels(intPixelLoop)
			
			'Place the pixel into the canvas array
			If intPixelPosition < 120 Then sarryCanvas(intRowLoop + intRandomLinePlace, intPixelPosition) = "2"
		Next
	Next
End Sub





'******************************************
'***  	Add Noise to Canvas	      *****
'******************************************

'Adds noise pixels to canvas
Private Sub mapCanvasNoise(intNoiseLevel, strCanvasColour)
	
	Dim intRow
	Dim intColumn
	
	'Exit subroutine if Noise Level is set to 0 (0 = off)
	If intNoiseLevel = 0 Then Exit Sub
	
	'Times the noise level by 50 to get the number of random pixels to set
	intNoiseLevel = intNoiseLevel * 25
	
	'Loop through each row of the character
	For intRowLoop = 0 to intNoiseLevel
	
		'Get a random row and coloumn to change the pixel on
		intRow = CInt(Rnd * 49)
		intColumn = CInt(Rnd * 119)

		'Place the noise pixel into the canvas array
		sarryCanvas(intRow, intColumn) = strCanvasColour
	Next
End Sub






'******************************************
'***  	Add Line Noise to Canvas	  *****
'******************************************

'Adds line noise to canvas
Private Sub mapCanvasLineNoise(intNoiseLines, strCanvasColour)
	
	
	Dim intLinePixels
	Dim blnTravelDirection
	Dim intPixelsPerRow
	Dim intPixelLoop
	Dim intMapLineLoop
	Dim intRow
	Dim intColumn
	Dim intRowPixels
	
	'Exit subroutine if Noise Level Lines is set to 0 (0 = off)
	If intNoiseLines = 0 Then Exit Sub
	
	'Loop through the number of lines we are placing on the canvas
	For intMapLineLoop = 1 to intNoiseLines
	
		'Initialise variables
		intRowPixels = 1
		
		'Generate random start position
		intRow = CInt(Rnd * 49)
		intColumn = CInt(Rnd * 80) 'Set as 80 as don't want the start column to far over
		
		'Generate random number for how many pixles the line will have
		intLinePixels = CInt(Rnd * 120)
		
		'Generate how many pixels we will have per row (this dictates what angle the line travels at)
		intPixelsPerRow = CInt(Rnd * 5)
		
		'This is the direction of travel (if the line will go down or up through the rows)
		blnTravelDirection = CBool(CInt(Rnd * 1))
		
		
		'Loop through each pixel in the line maping it to the canvas
		For intPixelLoop = 0 to intLinePixels
		
			'Exit loop if out of canvas space
			If intColumn <  1 OR intColumn > 118 Then Exit For
			If intRow < 1 OR intRow > 48 Then Exit For
			
			'Increament column count
			intColumn = intColumn + 1
			
			'See if we need to increament/decrement row count
			'If the number of pixels chnaged in this row = the number of pixels per row increament/decrement row count
			If intRowPixels = intPixelsPerRow Then
				
				'The direction of travel sets if we are incrementing (moving up) or decrementing (moving down) rows
				If blnTravelDirection Then
					intRow = intRow - 1
				Else
					intRow = intRow + 1
				End If
				
				'Reset the pixels set on this row to 1
				intRowPixels = 1
			
			'Increament the number of pixels set on this row
			Else
				intRowPixels = intRowPixels + 1
			End If
			
			
			'Place the noise line pixel into the canvas array
			sarryCanvas(intRow, intColumn) = strCanvasColour
			
		Next
	Next
End Sub





'******************************************
'***  	Hex Colour Reverser 	      *****
'******************************************

'The HEX colours in bitmaps are in reverse, so we need to switch them for easier setup
Private Function bmpHexColourSwicth(strHexColour)

	Dim sarryHexColour
	Dim strBmpHexColour

	'If HEX colour has spaces
	If InStr(strHexColour, " ") AND Len(strHexColour) = 8 Then
		
		'Split at the space
		sarryHexColour = Split(strHexColour, " ")
		
		'Make sure we have 3 HEX values
		If UBound(sarryHexColour) = 2 Then
			'Reverse the 3 HEX values
			strBmpHexColour = sarryHexColour(2)
			strBmpHexColour = strBmpHexColour & sarryHexColour(1)
			strBmpHexColour = strBmpHexColour & sarryHexColour(0)
		
		'Else it's not in correct format so just display as white
		Else
			strBmpHexColour = "CCCCCC"
		End If
		
	'If the HEX colour has no spaces make sure is 6 chars long
	ElseIf Len(strHexColour) = 6 Then
		
		'Reverse the 3 HEX values
		strBmpHexColour = Mid(strHexColour, 5, 2)
		strBmpHexColour = strBmpHexColour & Mid(strHexColour, 3, 2)
		strBmpHexColour = strBmpHexColour & Mid(strHexColour, 1, 2)
	
	'Else the value is not correct so set as white
	Else
		strBmpHexColour = "CCCCCC"
	End If
	
	'Make sure it is in uppper case
	strBmpHexColour = LCase(strBmpHexColour)
	
	'Return functions
	bmpHexColourSwicth = strBmpHexColour
	
End Function






'******************************************
'***  	Map Border to Canvas	      *****
'******************************************

'Map border sub, builds up the border in the canvas array
Private Sub mapBorder()

	'Create border sides (do this last so it doesn't overwrite the characters)
	For intRowLoop = 0 to 49
		sarryCanvas(intRowLoop, 0) = "1"
		sarryCanvas(intRowLoop, 119) = "1"
		
	Next
	
	'Create border top and bottom
	For intColumnLoop = 0 to 119
		sarryCanvas(49, intColumnLoop) = "1"
		sarryCanvas(0, intColumnLoop) = "1"
	Next
End Sub

%>