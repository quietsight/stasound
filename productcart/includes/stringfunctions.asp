<%
'// Main function used to sanitize strings
function getUserInput(input,stringLength)
 dim tempStr,i,known_bad

 known_bad= array("*","--","expression(")
 if stringLength>0 then
  tempStr	= left(trim(input),stringLength) 
 else
  tempStr	= trim(input)
 end if
 for i=lbound(known_bad) to ubound(known_bad)
 	if (instr(1,tempStr,known_bad(i),vbTextCompare)<>0) then
		tempStr	= replace(tempStr,known_bad(i),"",1,-1,1)
 	end if
 next
 if tempStr<>"" then
	 tempStr	= replace(tempStr,"'","''")
	 tempStr	= replace(tempStr,"<","&lt;")
	 tempStr	= replace(tempStr,">","&gt;")
	 tempStr	= replace(tempStr,"%0d","")
	 tempStr	= replace(tempStr,"%0D","")
	 tempStr	= replace(tempStr,"%0a","")
	 tempStr	= replace(tempStr,"%0A","")
	 tempStr	= replace(tempStr,"\r\n","")
	 tempStr	= replace(tempStr,"\r","")
	 tempStr	= replace(tempStr,"\n","")
	 tempStr	= replace(tempStr,"\R\N","")
	 tempStr	= replace(tempStr,"\R","")
	 tempStr	= replace(tempStr,"%28","")
	 tempStr	= replace(tempStr,"%29","")
	 tempStr	= replace(tempStr,"\N","")
	 tempStr	= replace(tempStr,"EXEC(","",1,-1,1) 
	 tempStr	= replace(tempStr,"NOT IN (SELECT","",1,-1,1)
	 tempStr	= replace(tempStr,"(SELECT","",1,-1,1)
	 tempStr	= replace(tempStr,"FROMCHARCODE","",1,-1,1)
	 tempStr	= replace(tempStr,"ALERT","",1,-1,1)
	 tempStr	= replace(tempStr,"ResponseSplitting","",1,-1,1)
	 tempStr	= replace(tempStr,"Content-Type","",1,-1,1)
 end if
	
	if tempStr<>"" then
	 	if IsNumeric(tempStr) then
	 		if InStr(Cstr(10/3),",")>0 then
				if Instr(tempStr,".")>0 then
					tempStr=FormatNumber(tempStr,,,,0)
	 				tempStr=replace(tempStr,".",",")
				end if
	 		end if
	 	end if
	end if
 
 getUserInput	= tempStr 
end function


'// Slightly different function used for "Recently viewed products" feature
function getUserInput2(input,stringLength)
 dim tempStr,i

 if stringLength>0 then
  tempStr	= left(trim(input),stringLength) 
 else
  tempStr	= trim(input)
 end if
 if tempStr<>"" then
	 tempStr	= replace(tempStr,"'","''")
	 tempStr	= replace(tempStr,"<","&lt;")
	 tempStr	= replace(tempStr,">","&gt;")
	 tempStr	= replace(tempStr,"%0d","")
	 tempStr	= replace(tempStr,"%0D","")
	 tempStr	= replace(tempStr,"%0a","")
	 tempStr	= replace(tempStr,"%0A","")
	 tempStr	= replace(tempStr,"\r\n","")
	 tempStr	= replace(tempStr,"\r","")
	 tempStr	= replace(tempStr,"\n","")
	 tempStr	= replace(tempStr,"\R\N","")
	 tempStr	= replace(tempStr,"\R","")
	 tempStr	= replace(tempStr,"%28","")
	 tempStr	= replace(tempStr,"%29","")
	 tempStr	= replace(tempStr,"\N","")
	 tempStr	= replace(tempStr,"EXEC(","",1,-1,1)
	 tempStr	= replace(tempStr,"NOT IN (SELECT","",1,-1,1)
	 tempStr	= replace(tempStr,"(SELECT","",1,-1,1)
	 tempStr	= replace(tempStr,"FROMCHARCODE","",1,-1,1)
	 tempStr	= replace(tempStr,"ALERT","",1,-1,1)
	 tempStr	= replace(tempStr,"ResponseSplitting","",1,-1,1)
	 tempStr	= replace(tempStr,"Content-Type","",1,-1,1)
 end if
	
	if tempStr<>"" then
	 	if IsNumeric(tempStr) then
	 		if InStr(Cstr(10/3),",")>0 then
				if Instr(tempStr,".")>0 then
					tempStr=FormatNumber(tempStr,,,,0)
	 				tempStr=replace(tempStr,".",",")
				end if
	 		end if
	 	end if
	end if
 
 getUserInput2	= tempStr 
end function


'[ClearHTMLTags2]

'Coded by Jóhann Haukur Gunnarsson
'joi@innn.is

'	Purpose: This function clears all HTML tags from a string using Regular Expressions.
'	 Inputs: strHTML2;	A string to be cleared of HTML TAGS
'		 intWorkFlow2;	An integer that if equals to 0 runs only the regEx2p filter
'							  .. 1 runs only the HTML source render filter
'							  .. 2 runs both the regEx2p and the HTML source render
'							  .. >2 defaults to 0
'	Returns: A string that has been filtered by the function


function ClearHTMLTags2(strHTML2, intWorkFlow2)

	'Variables used in the function
	
	dim regEx2, strTagLess2
	
	'---------------------------------------
	strTagLess2 = strHTML2
	'Move the string into a private variable
	'within the function
	'---------------------------------------
	
	'---------------------------------------
	'NetSource Commerce codes
	IF strTagLess2<>"" THEN
		strTagLess2=replace(strTagLess2,"<br>"," ")
		strTagLess2=replace(strTagLess2,"<BR>"," ")
		strTagLess2=replace(strTagLess2,"<p>"," ")
		strTagLess2=replace(strTagLess2,"<P>"," ")
		strTagLess2=replace(strTagLess2,"</p>"," ")
		strTagLess2=replace(strTagLess2,"</P>"," ")
		strTagLess2=replace(strTagLess2,vbcrlf," ")
		strTagLess2=replace(strTagLess2,"™","&trade;")
		strTagLess2=replace(strTagLess2,"©","&copy;")
		strTagLess2=replace(strTagLess2,"®","&reg;")
		strTagLess2=trim(strTagLess2)
		do while instr(strTagLess2,"  ")>0
			strTagLess2=replace(strTagLess2,"  "," ")
		loop
	END IF
	'Modify the string to a friendly ONLY 1 LINE string
	'---------------------------------------
	
	IF strTagLess2<>"" THEN

	'regEx2 initialization

		'---------------------------------------
		set regEx2 = New regExp 
		'Creates a regEx2p object		
		regEx2.IgnoreCase = True
		'Don't give frat about case sensitivity
		regEx2.Global = True
		'Global applicability
		'---------------------------------------
		'Phase I
		'	"bye bye html tags"
		if intWorkFlow2 <> 1 then
			'---------------------------------------
			regEx2.Pattern = "<[^>]*>"
			'this pattern mathces any html tag
			strTagLess2 = regEx2.Replace(strTagLess2, "")
			'all html tags are stripped
			'---------------------------------------
		end if

		'Phase II
		'	"bye bye rouge leftovers"
		'	"or, I want to render the source"
		'	"as html."

		'---------------------------------------
		'We *might* still have rouge < and > 
		'let's be positive that those that remain
		'are changed into html characters
		'---------------------------------------	
		if intWorkFlow2 > 0 and intWorkFlow2 < 3 then
			regEx2.Pattern = "[<]"
			'matches a single <
			strTagLess2 = regEx2.Replace(strTagLess2, "&lt;")

			regEx2.Pattern = "[>]"
			'matches a single >
			strTagLess2 = regEx2.Replace(strTagLess2, "&gt;")
			'---------------------------------------
		end if
		
		'Clean up
		'---------------------------------------
		set regEx2 = nothing
		'Destroys the regEx2p object
		'---------------------------------------	
	END IF 'vefiry strTagLess2 (null strings)
	
	'---------------------------------------
	ClearHTMLTags2 = strTagLess2
	'The results are passed back
	'---------------------------------------
end function
function ClearHTMLTags3(strHTML2, intWorkFlow2)

	'Variables used in the function
	
	dim regEx2, strTagLess2
	
	'---------------------------------------
	strTagLess2 = strHTML2
	'Move the string into a private variable
	'within the function
	'---------------------------------------
	
	'---------------------------------------
	'NetSource Commerce codes
	IF strTagLess2<>"" THEN
		strTagLess2=replace(strTagLess2,"<br>"," ")
		strTagLess2=replace(strTagLess2,"<BR>"," ")
		strTagLess2=replace(strTagLess2,"<p>"," ")
		strTagLess2=replace(strTagLess2,"<P>"," ")
		strTagLess2=replace(strTagLess2,"</p>"," ")
		strTagLess2=replace(strTagLess2,"</P>"," ")
		strTagLess2=replace(strTagLess2,vbcrlf," ")
		strTagLess2=replace(strTagLess2,"™","&trade;")
		strTagLess2=replace(strTagLess2,"©","&copy;")
		strTagLess2=replace(strTagLess2,"®","&reg;")
		strTagLess2=trim(strTagLess2)
		do while instr(strTagLess2,"  ")>0
			strTagLess2=replace(strTagLess2,"  "," ")
		loop
	END IF
	'Modify the string to a friendly ONLY 1 LINE string
	'---------------------------------------
	
	IF strTagLess2<>"" THEN

	'regEx2 initialization

		'---------------------------------------
		set regEx2 = New regExp 
		'Creates a regEx2p object		
		regEx2.IgnoreCase = True
		'Don't give frat about case sensitivity
		regEx2.Global = True
		'Global applicability
		'---------------------------------------
		'Phase I
		'	"bye bye html tags"
		if intWorkFlow2 <> 1 then
			'---------------------------------------
			regEx2.Pattern = "<[^>]*>"
			'this pattern mathces any html tag
			strTagLess2 = regEx2.Replace(strTagLess2, "")
			'all html tags are stripped
			'---------------------------------------
		end if

		'Phase II
		'	"bye bye rouge leftovers"
		'	"or, I want to render the source"
		'	"as html."

		'---------------------------------------
		'We *might* still have rouge < and > 
		'let's be positive that those that remain
		'are changed into html characters
		'---------------------------------------	
		if intWorkFlow2 > 0 and intWorkFlow2 < 3 then
			regEx2.Pattern = "[<]"
			'matches a single <
			strTagLess2 = regEx2.Replace(strTagLess2, "&lt;")

			regEx2.Pattern = "[>]"
			'matches a single >
			strTagLess2 = regEx2.Replace(strTagLess2, "&gt;")
			'---------------------------------------
		end if
		
		'Clean up
		'---------------------------------------
		set regEx2 = nothing
		'Destroys the regEx2p object
		'---------------------------------------	
	END IF 'vefiry strTagLess2 (null strings)
	
	'---------------------------------------
	ClearHTMLTags3 = strTagLess2
	'The results are passed back
	'---------------------------------------
end function

'check for real integers
Function validNum(strInput)
	DIM iposition		' Current position of the character or cursor
	validNum =  true 
	if isNULL(strInput) OR trim(strInput)="" then
		validNum = false
	else
		'loop through each character in the string and validate that it is a number or integer
		For iposition=1 To Len(trim(strInput))
			if InStr(1, "12345676890", mid(strInput,iposition,1), 1) = 0 then
				validNum =  false
				Exit For
			end if
		Next
	end if
end Function

'// This function is used to wrap text in e-mails (e.g. product name, product options)
'// Assumption is that Width < len(Text)
Public Function WrapString(ByVal Width, ByVal Text)
	wrapPos=0
	iwrap = Width
	do while mid(Text,iwrap,1) <> " " AND iwrap > 0
		iwrap=iwrap-1
	loop
	if iwrap > 0 then
		WrapString=left(Text,iwrap)
		wrapPos = iwrap
	else
		WrapString=left(Text,Width)
		wrapPos = Width
	end if
End Function


'// Strip out Currency and Percent Characters From a String
Public Function pcf_ReplaceChars(pricenumber)
	on error resume next
	pcf_ReplaceChars=replace(pricenumber,"%","")
	pcf_ReplaceChars=replace(pcf_ReplaceChars,scCurSign,"")
	'// If they still did not type a valid number set equal to zero.
	if not isNumeric(pcf_ReplaceChars) then
		pcf_ReplaceChars=0
	end if
End Function


Public Function pcf_PayPalExpressOnly()

		'// Open Private Connection String
		Set conPayPalExpressOnly=Server.CreateObject("ADODB.Connection")
		conPayPalExpressOnly.Open scDSN

		pcf_PayPalExpressOnly=1		
		if session("customerType")=1 then
			query="SELECT idPayment FROM paytypes WHERE active=-1 AND gwCode<>999999 AND gwCode<>50;"
		else
			query="SELECT idPayment FROM paytypes WHERE active=-1 AND Cbtob=0 AND gwCode<>999999 AND gwCode<>50;"
		end if		

		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conPayPalExpressOnly.execute(query)		
		if err.number<>0 then
			set rs=nothing
			exit function
		end if	
				
		PayTypesIndex=Cint(0)	
		if NOT rs.eof then
			pcv_strTempPaymentOptionValue=rs("idPayment")
			PayTypesIndex=PayTypesIndex + 1	
		else
			set rs=nothing
			pcf_PayPalExpressOnly=0			
		end if
		set rs=nothing		
		
		if int(PayTypesIndex)=0 then
			pcf_PayPalExpressOnly=0
		end if
		

		
		'// Close Private Connection String
		conPayPalExpressOnly.Close
		Set conPayPalExpressOnly=nothing
		
End Function


Public Function pcf_PaymentTypes(GatewayName)

		'// Open Private Connection String
		Set conPaymentTypes=Server.CreateObject("ADODB.Connection")
		conPaymentTypes.Open scDSN

		'SB S
		sbCartArr=Session("pcCartSession")
		If (sbCartArr(1,38)>0) then
			pcIsSubscription = True		
			strAndSub = "AND (pcPayTypes_Subscription = 1)"
		Else
			pcIsSubscription = False		
			strAndSub = ""		
		End if 
		'SB E
		
		pcf_PaymentTypes=1		
		'SB S	
		if session("customerType")=1 then
			query="SELECT idPayment FROM paytypes WHERE active=-1 AND gwCode<>50 AND gwCode<>999999 " & strAndSub & ";"
		else
			query="SELECT idPayment FROM paytypes WHERE active=-1 AND Cbtob=0 AND gwCode<>50 AND gwCode<>999999 " & strAndSub & ";"
		end if			
		'SB E		
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conPaymentTypes.execute(query)		
		if err.number<>0 then
			set rs=nothing
			exit function
		end if	
				
		PayTypesIndex=Cint(0)	
		if NOT rs.eof then
			pcv_strTempPaymentOptionValue=rs("idPayment")
			PayTypesIndex=PayTypesIndex + 1	
		else
			set rs=nothing
			pcf_PaymentTypes=0			
		end if
		set rs=nothing		
		
		if int(PayTypesIndex)=0 then
			pcf_PaymentTypes=0
		end if
		
		query="SELECT idPayment FROM paytypes WHERE active=-1 and gwCode=50 " & strAndSub & ";"
		set rsGoogle=Server.CreateObject("ADODB.Recordset")     
		set rsGoogle=conPaymentTypes.execute(query)		
		if rsGoogle.eof then
			pcv_intGoogleActive=0
		else
			pcv_intGoogleActive=-1
		end if
		
		query="SELECT idPayment FROM paytypes WHERE active=-1 and gwCode=999999 " & strAndSub & ";"
		set rsPayPal=Server.CreateObject("ADODB.Recordset")     
		set rsPayPal=conPaymentTypes.execute(query)	
		if rsPayPal.eof then
			pcv_intPayPalExpressActive=0
		else
			pcv_intPayPalExpressActive=-1
		end if

		Select Case GatewayName
			Case "GoogleCheckout":pcv_strPayTypeFilter=(pcv_intGoogleActive=-1)
			Case "PayPalExp":pcv_strPayTypeFilter=(pcv_intPayPalExpressActive=-1)
			Case "":pcv_strPayTypeFilter=(pcv_intPayPalExpressActive=-1 OR pcv_intGoogleActive=-1)
			Case Else:pcv_strPayTypeFilter=(pcv_intPayPalExpressActive=-1 OR pcv_intGoogleActive=-1)
		End Select

		if (pcv_strPayTypeFilter) AND pcf_PaymentTypes=0 then
			pcf_PaymentTypes=0
		else
			pcf_PaymentTypes=1
		end if			
		
		'// Close Private Connection String
		conPaymentTypes.Close
		Set conPaymentTypes=nothing
		
End Function


Public Function pcf_DiscountedOptions(OriginalOptionsTotal, ProductID, Quantity, CustomerType)
	
	'// Dim pcv_intDiscountPerUnit  ///// Notes: Dim this variable before this function call to use the value throughout your page.
	pDiscountPerUnit=0
	pDiscountPerWUnit=0
	pPercentage=0
	pbaseproductonly=0
	pcv_intDiscountPerUnit=0
	
	query="SELECT discountPerUnit, percentage, discountPerWUnit, baseproductonly FROM discountsPerQuantity WHERE idProduct="&ProductID&" AND quantityFrom<="&Quantity&" AND quantityUntil>="&Quantity
	Dim rsDiscountedOptions
	set rsDiscountedOptions=server.CreateObject("ADODB.RecordSet")
	set rsDiscountedOptions=conntemp.execute(query)
	if not rsDiscountedOptions.eof and err.number<>9 then								
		'// There are quantity discounts defined for that quantity and product
		pDiscountPerUnit=rsDiscountedOptions("discountPerUnit")
		pDiscountPerWUnit=rsDiscountedOptions("discountPerWUnit")
		pPercentage=rsDiscountedOptions("percentage")
		pbaseproductonly=rsDiscountedOptions("baseproductonly")										
		'// )nly factor if the options were included									
		if pbaseproductonly<>"-1" then																			
			if pPercentage<>"0" then 	
				if CustomerType=1 Then								
					pcv_intDiscountPerUnit=pDiscountPerWUnit
				else
					pcv_intDiscountPerUnit=pDiscountPerUnit
				end if
				OriginalOptionsTotal= OriginalOptionsTotal - ((pcv_intDiscountPerUnit/100) * OriginalOptionsTotal)								
			end if										
		end if '// if pbaseproductonly<>"-1" then
	end if		
	if NOT isNumeric(OriginalOptionsTotal) then
		OriginalOptionsTotal=0
	end if
	pcf_DiscountedOptions=OriginalOptionsTotal
	'// rsDiscountedOptions=nothing /// Keep this line commented. Close the Obj after the loop.
End Function


Function pcf_Round(numericValue,decimals) 
	'// Validate Numeric Value
	If isNULL(numericValue)=True OR numericValue="" Then
		pcf_Round=0 '// Set the Value to Zero if Validation Fails
		Exit Function '// Do Not Round... Exit
	End If 
	If isnumeric(numericValue)=False Then
		pcf_Round=numericValue '// Restore the Original Value if Validation Fails
		Exit Function '// Do Not Round... Exit
	End If 
	'// Perform the Round	
	pcf_Round=round(numericValue, decimals)
End Function 

'Print last Four Digits of Credit Card in the Invoice

Function ShowLastFour(CreditCardNumber)
	
	'Check for Special Characters(Hyphen and Space) in the Credit Card Number
	
	CreditCardNumberTemp = CreditCardNumber
	CreditCardNumber = Replace(CreditCardNumber,"-","")
	CreditCardNumber = Replace(CreditCardNumber," ","")
	
	if IsNumeric(CreditCardNumber) and len(CreditCardNumber)>10 then
		CreditCardNumber="****" & Right(CreditCardNumber,4)
		else
		CreditCardNumber=CreditCardNumberTemp
	end if	
	ShowLastFour=CreditCardNumber
End Function

'// Determine Credit Card Type, if possible
Function ShowCardType(CreditCardNumber)
	
	CreditCardType="Not Available"

	'Check for Special Characters(Hyphen and Space) in the Credit Card Number
	CreditCardNumberTemp = CreditCardNumber
	CreditCardNumber = Replace(CreditCardNumber,"-","")
	CreditCardNumber = Replace(CreditCardNumber," ","")
	
	if IsNumeric(CreditCardNumber) and len(CreditCardNumber)>10 then
		
		'Check the first two digits first
		Select Case CInt(Left(CreditCardNumber, 2)) ' 1
		   Case 34, 37
			  CreditCardType = "American Express"
		   Case 36
			  CreditCardType = "Diners Club"
		   Case 38
			  CreditCardType = "Carte Blanche"
		   Case 51, 52, 53, 54, 55
			  CreditCardType = "Master Card"
		   Case Else
			  'None of the above - so check the
			  'first four digits collectively
			  Select Case CInt(Left(CreditCardNumber, 4)) ' 2
				 Case 2014, 2149
					CreditCardType = "EnRoute"
				 Case 2131, 1800
					CreditCardType = "JCB"
				 Case 6011
					CreditCardType = "Discover"
				 Case Else
					'None of the above - so check the
					'first three digits collectively
					Select Case CInt(Left(CreditCardNumber, 3)) ' 3
					   Case 300, 301, 302, 303, 304, 305
						  CreditCardType = "American Diners Club"
					   Case Else
					   'None of the above -
					   'so simply check the first digit
					   Select Case CInt(Left(CreditCardNumber, 1)) ' 4
						  Case 3
							 CreditCardType = "JCB"
						  Case 4
							CreditCardType = "Visa"
					   End Select '4 
					End Select ' 3
				End Select '2 
			End Select '1 
	end if
	ShowCardType=CreditCardType
	
End Function


'// To DB #1
Function pcf_SanitizeApostrophe(UserInput)
	If UserInput<>"" AND NOT isNULL(UserInput) Then
		pcf_SanitizeApostrophe=replace(UserInput,"'","''")
	Else
		pcf_SanitizeApostrophe=UserInput
	End If
End Function

'// To DB #2
function removeReplaceSQ(myString)
	if isNULL(myString) then
		removeReplaceSQ=""
	else
		removeReplaceSQ=replace(myString,"''","'") '// normalize
		removeReplaceSQ=replace(removeReplaceSQ,"'","''")
	end if
end function

'// To Text File
function removeSQ(myString)
	if isNULL(myString) then
		removeSQ=""
	else
		myString=replace(myString,"''","'") '// normalize
		removeSQ=replace(myString,"""","&quot;")
	end if
end function

'// To Print
Function pcf_PrintCharacters(UserInput)
	Dim tempStr
	tempStr=UserInput
	If tempStr<>"" AND NOT isNULL(tempStr) Then
		tempStr=replace(tempStr,"'","''") '// normalize
		tempStr=replace(tempStr,"''","'")
		tempStr=replace(tempStr,"&quot;","""") 
		if Instr(tempStr,Vbcrlf)>0 then
			do while Instr(tempStr,Vbcrlf&Vbcrlf)>0
			tempStr=replace(tempStr, Vbcrlf&Vbcrlf,Vbcrlf)
			loop
		tempStr=replace(tempStr, Vbcrlf," ")
		end if
		tempStr=replace(tempStr, Chr(8),"")
		tempStr=replace(tempStr, Chr(9),"")
		tempStr=replace(tempStr, Chr(10),"")
		tempStr=replace(tempStr, Chr(13),"")
		pcf_PrintCharacters=tempStr
	Else
		pcf_PrintCharacters=tempStr
	End If
End Function

'// To Email
Function pcf_RestoreCharacters(UserInput)
	Dim tempStr
	tempStr=UserInput
	If tempStr<>"" AND NOT isNULL(tempStr) Then
		tempStr=replace(tempStr,"&trade;","™")
		tempStr=replace(tempStr,"&copy;","©")
		tempStr=replace(tempStr,"&reg;","®")
		pcf_RestoreCharacters=tempStr
	Else
		pcf_RestoreCharacters=tempStr
	End If
End Function

'// Convert Special Characters To HTML
Function pcf_ReplaceCharacters(UserInput)
	' NOTES:
	' DO NOT add a replace on " (double quotes) because this function is used when saving
	' HTML code too (e.g. long product description) and that would break the HTML
	Dim tempStr
	tempStr=removeReplaceSQ(UserInput) '// Sanitize UserInput
	If tempStr<>"" And Not isNull(tempStr) Then		
		tempStr=replace(tempStr,"™","&trade;")
		tempStr=replace(tempStr,"©","&copy;")
		tempStr=replace(tempStr,"®","&reg;")
		pcf_ReplaceCharacters=tempStr
	Else
		pcf_ReplaceCharacters=tempStr
	End If
End Function

Function pcf_ReplaceQuotes(UserInput)
	Dim tempStr
	tempStr=UserInput
	If tempStr<>"" And Not isNull(tempStr) Then
		tempStr=replace(tempStr,"""","&quot;")
		pcf_ReplaceQuotes=tempStr
	Else
		pcf_ReplaceQuotes=tempStr
	End If
End Function

Function pcf_AddToCart(IdProduct)
	IF scQuickBuy = "1" THEN '// Feature is inactive
			pcf_AddToCart=False
			Exit Function
			
	ELSE '// Feature is active
	
		If IdProduct<>"" AND NOT isNULL(IdProduct) Then
	
			'// Show Add To Cart Button
			pcf_AddToCart=True
	
			'// 1.) PRODUCT FOR SALE:  Verify level is 0 OR is 1 with a custmer type of 1	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// NOTES	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Check for order level permission "scorderlevel".
			' scorderlevel = 0 // everybody
			' scorderlevel = 1 // wholesale only
			' scorderlevel = 2 // catalog only
			
			' Also check what level the current customer is classified.
			' session("customerType") = "" // not logged in
			' session("customerType") = 1  // wholesale
			' session("customerType") = 0  // retail
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If NOT (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) Then '// if [everyone] OR [wholesale w/ wholesale only turned on]
				pcf_AddToCart=False
				Exit Function
			End If
	
	
			'// 2.) STOCK CONSIDERATIONS:  If out of stock AND out of stock purchase is allowed show button.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// NOTES	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' The Following Variable Must Be Defined:
			' pStock
			' pserviceSpec
			' pNoStock
			' pcv_intBackOrder
			' iBTOOutofstockpurchase
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If NOT pcf_OutStockPurchaseAllow Then  
				pcf_AddToCart=False
				Exit Function
			End If
	
	
			'// 3.) Product is a BTO product
			If (pserviceSpec<>0) Then  
				pcf_AddToCart=False
				Exit Function
			End If
	
	
			'// 4.) Product has required options (if options are not required, then visiting the product details page is not mandatory) 
			If pcf_CheckForReqOptions(IdProduct)=1 Then
				pcf_AddToCart=False
				Exit Function
			End If
	
	
			'// 5.) Product has required input fields 
			If pcf_CheckForReqInputFields(IdProduct)=1 Then
				pcf_AddToCart=False
				Exit Function
			End If		
	
	
			'// 6.) Product has required accessories 
			If pcf_CheckForReqAccessories(IdProduct)=1 Then
				pcf_AddToCart=False
				Exit Function
			End If		
	
	
			'// 7.) Product is an Apparel Product 	
			'//			Does not apply to stores not running the APP Add-on
	
	
	
	
	
	
	
	
	
	
			
			'// 8.) Product has Qty Minimums	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// NOTES	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Use "pcv_SkipCheckMinQty" to skip the Qty Minimums validation section
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			If pcf_CheckMinQty(IdProduct)=1 AND pcv_SkipCheckMinQty<>-1 Then
				pcf_AddToCart=False
				Exit Function
			End If	
	
	
			'// 9.) Not for Sale
			If pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0 Then
				pcf_AddToCart=False
				Exit Function
			End If			
			
		Else  '// If IdProduct<>"" AND NOT isNULL(IdProduct) Then
			
			pcf_AddToCart=False
			Exit Function
			
		End If  '// If IdProduct<>"" AND NOT isNULL(IdProduct) Then
		
	END IF '// Feature is active
	
End Function

Public Function pcf_InitializePrototype()  
	If Not (pcv_strPrototype) Then
		pcv_strPrototype=True
		pcv_NewLine=CHR(10)
		pcv_strModal=""
		pcv_strModal=pcv_strModal&"<link href='screen.css' rel='stylesheet' type='text/css' />"& pcv_NewLine
		pcv_strModal=pcv_strModal&"<script type='text/javascript' src='../includes/javascripts/highslide.html.packed.js'></script>"& pcv_NewLine	
		pcf_InitializePrototype = pcv_strModal
	End If	
End Function



Public Function pcf_ModalWindow(Message, ID, Width)  
	Dim pcv_strModal, pcv_NewLine
	pcv_NewLine=CHR(10)
	pcv_strModal=""
	pcv_strModal=pcv_strModal&"<!-- Start: Modal Window - "& ID &" -->"& pcv_NewLine
	pcv_strModal="<a href=""javascript:;"" id="""& ID &""" onclick=""return hs.htmlExpand(this, { contentId: 'modal_"& ID &"', align: 'center', width: "& Width &" } )"" class=""highslide""></a>"& pcv_NewLine
	pcv_strModal=pcv_strModal&"<div id='modal_"& ID &"' class=""highslide-maincontent"">"& pcv_NewLine
	pcv_strModal=pcv_strModal&"		<div align='center'>"& Message &"</div>"& pcv_NewLine
	pcv_strModal=pcv_strModal&"</div>"& pcv_NewLine
	pcv_strModal=pcv_strModal&"<!-- End: Modal Window - "& ID &" -->"& pcv_NewLine
	pcv_strModal=pcv_strModal&"<!-- Start: Modal Functions - "& ID &" -->"& pcv_NewLine
	pcv_strModal=pcv_strModal&"<script language=""JavaScript"">"& pcv_NewLine
	pcv_strModal=pcv_strModal&"<!--"& pcv_NewLine
	pcv_strModal=pcv_strModal&"function pcf_Close_"& ID &"() { var t=setTimeout(""hs.close('"& ID &"')"",50) }"& pcv_NewLine
	pcv_strModal=pcv_strModal&"function pcf_Open_"& ID &"() { document.getElementById('"& ID &"').onclick() }"& pcv_NewLine
	pcv_strModal=pcv_strModal&"//-->"& pcv_NewLine
	pcv_strModal=pcv_strModal&"</script>"& pcv_NewLine
	pcv_strModal=pcv_strModal&"<!-- End: Modal Functions - "& ID &" -->"& pcv_NewLine
	pcf_ModalWindow = pcv_strModal
End Function


Sub pcs_GetSubCats(tmpIDParent)
	Dim intCount,pcArrSubCat
	querySubCat="SELECT idcategory FROM categories WHERE idParentCategory=" & tmpIDParent & ";"
	set rsSubCat=connTemp.execute(querySubCat)
	if not rsSubCat.eof then
		pcArrSubCat=rsSubCat.getRows()
		set rsSubCats=nothing
		For intCount=0 to ubound(pcArrSubCat,2)
			TmpCatList=TmpCatList & "," & pcArrSubCat(0,intCount)
			call pcs_GetSubCats(pcArrSubCat(0,intCount))
		Next
	end if
	set rsSubCat=nothing
End Sub



Function pcf_ColumnToArray(pMDA, iCol)
	Dim iLoop, Result(), max
	max = UBound(pMDA,2)
	ReDim Result(max)
	For iLoop = 0 To max
		Result(iLoop) = pMDA(iCol, iLoop)
	Next
	pcf_ColumnToArray = Result
End Function

'*****************************************************************************************************
'// START: Add function for ZIP code
'*****************************************************************************************************
Public Function pcf_PostCodes(code)
	if len(code)>0 then
		code = trim(code)
		code = ucase(code)
	end if
	pcf_PostCodes=code
End Function
'*****************************************************************************************************
'// END: Add function for ZIP code
'*****************************************************************************************************

'*****************************************************************************************************
'// START: Function used to look for admin user in user permission array
'*****************************************************************************************************
Function findUser(ByRef arr, ByVal val, pcArrCount)
	findUser=Null
	For i=0 To pcArrCount
			If CLng(val) = CLng(arr(i)) Then
				findUser=i
				Exit Function
			End If
	Next
End Function
'*****************************************************************************************************
'// END
'*****************************************************************************************************

'*****************************************************************************************************
'Get Active Security Key
'*****************************************************************************************************
Function pcs_GetSecureKey()
	Dim querygsk,rsGSKObj
	querygsk = "SELECT pcSecurityKey FROM pcSecurityKeys WHERE pcActiveKey = 1;"
	set rsGSKObj = server.CreateObject("ADODB.RecordSet")
	set rsGSKObj = connTemp.execute(querygsk)
	
	if rsGSKObj.eof then
		pcv_SecurityPass = scCrypPass
	else
		pcv_SecurityPass = rsGSKObj("pcSecurityKey")
	end if
	pcs_GetSecureKey = pcv_SecurityPass
	set rsGSKObj=nothing
End Function

Function pcs_GetKeyID()
	Dim querygsk,rsGSKObj
	querygsk = "SELECT pcSecurityKeyID FROM pcSecurityKeys WHERE pcActiveKey = 1;"
	set rsGSKObj = server.CreateObject("ADODB.RecordSet")
	set rsGSKObj = connTemp.execute(querygsk)
	
	if rsGSKObj.eof then
		pcv_SecurityKeyID = 0
	else
		pcv_SecurityKeyID = rsGSKObj("pcSecurityKeyID")
	end if
	pcs_GetKeyID = pcv_SecurityKeyID
	set rsGSKObj=nothing
End Function
'*****************************************************************************************************
'Get Active Security Key
'*****************************************************************************************************

'*****************************************************************************************************
'//Get the pcSecurityKey 
'*****************************************************************************************************
Function pcs_GetKeyUsed(keyid)
	if isNULL(keyid) or keyid&""="" then
		keyid = 0
	end if
	if keyid = 0 then
		pcs_GetKeyUsed = scCrypPass
	else
		query = "SELECT pcSecurityKey FROM pcSecurityKeys WHERE pcSecurityKeyID = "&keyid&";"
		set rsGSKObj = server.CreateObject("ADODB.RecordSet") 
		set rsGSKObj = connTemp.execute(query)
		if rsGSKObj.eof then
			'redirect
		else
			pcs_GetKeyUsed = rsGSKObj("pcSecurityKey")
		end if
		set rsGSKObj = nothing
	end if
End Function
'*****************************************************************************************************
'//Get the pcSecurityKey 
'*****************************************************************************************************
'*****************************************************************************************************
Function pcf_PurgeCardNumber(CreditCardNumber,keyid)
	if isNULL(keyid) or keyid&""="" then
		keyid = 0
	end if
	if keyid = 0 then
		sSecurePass = scCrypPass
	else
		query = "SELECT pcSecurityKey FROM pcSecurityKeys WHERE pcSecurityKeyID = "&keyid&";"
		set rsGSKObj = server.CreateObject("ADODB.RecordSet") 
		set rsGSKObj = connTemp.execute(query)
		if rsGSKObj.eof then
			'redirect
		else
			sSecurePass = rsGSKObj("pcSecurityKey")
		end if
		set rsGSKObj = nothing
	end if
	tempCC=CreditCardNumber
	tempCC=enDeCrypt(tempCC, sSecurePass)
	tempfourR=right(tempCC,4)
	tempfour="************"&tempfourR
	pcf_PurgeCardNumber=enDeCrypt(tempfour, sSecurePass)
End Function
'*****************************************************************************************************
'*****************************************************************************************************
'// Get the Customer Status
'*****************************************************************************************************
Function pcf_GetCustType(CustID)
	if CustID="" then CustID=0
	query = "SELECT pcCust_Guest FROM customers WHERE idcustomer=" & CustID
	set rsCustCType = Server.CreateObject("ADODB.Recordset")
	set rsCustCType = conntemp.execute(query)
	if NOT rsCustCType.EOF then
		pcf_GetCustType = rsCustCType("pcCust_Guest")
	else
		pcf_GetCustType = ""
	end if
	set rsCustCType = nothing
End Function
'*****************************************************************************************************
'*****************************************************************************************************
'// Get the Total Discount from Promotions
'*****************************************************************************************************
Function pcf_GetPromoTotal(PromoArr1,PromoIndex)
	TotalPromotions=0
	For m=1 to PromoIndex
		if PromoArr1(m,2)>0 then
			TotalPromotions=TotalPromotions+cdbl(PromoArr1(m,2))
		end if
	Next
	pcf_GetPromoTotal = TotalPromotions
End Function

' begin PRV41


'*****************************************************************************************************
'// Format a date for the currently defined database
'*****************************************************************************************************
   Function formatDateForDB(varDate)

      If IsDate(varDate) then
         if scDB="Access" Then
            formatDateForDB = "#" & varDate & "#"
	     Else
            formatDateForDB = "'" & varDate & "'"
	     End If
      Else
         formatDateForDB = "null"
      End If
   End function
'*****************************************************************************************************
   

'*****************************************************************************************************
'// Read a complete text file
'*****************************************************************************************************
   Function strReadAll(strFilename)

      Dim objFSO, objFile
	  Set objFSO = server.CreateObject("Scripting.FileSystemObject")
	  Err.number=0
	  Set objFile = objFSO.OpenTextFile(strFileName, 1)
	  strReadAll = objFile.ReadAll()
	  objFile.close
	  Set objFile = Nothing
	  Set objFSO = nothing

   End Function
'*****************************************************************************************************


'*****************************************************************************************************
'// simple proper case. if a string is all caps or all lower, forcer to Proper Case
'// (note: does not handle the McGillicuddies and St. Jeans of the world)
'*****************************************************************************************************
    Function ProperCase(varString)

	   If Len(varString)>0 then
		  If LCase(varString)=varstring Or UCase(varstring)=varstring Then
			 ProperCase = UCase(Left(varString,1)) & LCase(Mid(varString,2))
		  Else
			 ProperCase = varString
		  End If
	   End If
	   
	End Function
'*****************************************************************************************************


'*****************************************************************************************************
'// Generate a GUID
'*****************************************************************************************************
	Function genGUID()

		 Dim newGuid

		 newGuid = server.createobject("scriptlet.typelib").guid

		 ' Take the GUID (stripping out the brackets and any extraneous data at the end
		 newGuid=mid(Left(newGuid,instr(newGuid,"}")-1),2)

		 genGUID=newGuid
		 Set newGuid=nothing

	end Function
	

'*****************************************************************************************************
'// Passed what 'should' be a number, return zero if it's null or non-numeric, else return the number
'*****************************************************************************************************
	Function fnZeroIfNull(varAny)

	   fnZeroIfNull = 0

	   If IsNull(varAny) Then Exit Function
	   If Len(varAny)=0 Then Exit Function
	   If IsNumeric(varAny) = False Then Exit Function

	   
	   fnZeroIfNull = varAny

	End function

' end PRV41



'// Replace Comma
'*****************************************************************************************************
function replacecomma(pricenumber)
  If IsNumeric(pricenumber) then
	  if scDecSign="," then
		  replacecomma=replace(pricenumber,".","")
		  replacecomma=replace(replacecomma,",",".")
	  else
		  replacecomma=replace(pricenumber,",","")
	  end if
  End If
end function

'//Fix Javascript string issues
'*****************************************************************************************************
function FixLang(str)
Dim tmp1
	tmp1=str
	if tmp1<>"" then
		tmp1=replace(tmp1,"\""","""")
		tmp1=replace(tmp1,"\'","'")
		tmp1=replace(tmp1,"""","\""")
		tmp1=replace(tmp1,"'","\'")
	end if
	FixLang=tmp1
end function

%>