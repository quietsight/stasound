<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'=====================================================================================
'= Cardinal Commerce (http://www.cardinalcommerce.com)
'= pcPay_Cent_Utility.asp
'= General Utilities used for Thin Client Integrations
'= Version 4.2.2 01/25/2005
'=====================================================================================

Function determineCardType(Card_Number)
     
	Dim cardType

	cardType = "UNKNOWN"   ' VISA, MASTERCARD, JCB, AMEX, UNKNOWN

	If (Len(Card_Number) = "16" AND Left(Card_Number, 1) = "4") Then
		cardType = "VISA"
	ElseIf (Len(Card_Number) = "13" AND Left(Card_Number, 1) = "5") Then
		cardType = "MASTERCARD"
	ElseIf (Len(Card_Number) = "16" AND Left(Card_Number, 1) = "5") Then
		cardType = "MASTERCARD"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 4) = "2131") Then
		cardType = "JCB"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 4) = "1800") Then
		cardType = "JCB"
	ElseIf (Len(Card_Number) = "16" AND Left(Card_Number, 1) = "3") Then
		cardType = "JCB"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 2) = "34") Then
		cardType = "AMEX"
	ElseIf (Len(Card_Number) = "15" AND Left(Card_Number, 2) = "37") Then
		cardType = "AMEX"
	End If

	determineCardType = cardType   
	 
End Function

%>
