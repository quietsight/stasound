<%  
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.


Dim pcv_strGenericPageError, pcv_intErr, msg, pcv_strRequiredIcon, pcv_strErrorIcon, pcv_strSessionPrefix

'// set error to zero
pcv_intErr=0 
				
'// determine the path for the icons
pcv_strRequiredIcon = pcf_GenerateIconURL(rsIconObj("requiredicon"))
pcv_strErrorIcon = pcf_GenerateIconURL(rsIconObj("errorfieldicon"))

'// determine the prefix for the Sessions
if len(pcv_strAdminPrefix)>0 then
	pcv_strSessionPrefix = "pcAdmin"
else
	pcv_strSessionPrefix = "pcSF"
end if

'// If the Session is present then we dont need to re-fill from the database
Public Function pcf_ResetFormField(SV,LV)
	If len(SV)>0 Then
		pcf_ResetFormField=SV
	Else
		pcf_ResetFormField=LV
	End If 
End Function

'// This Validates any Required Field has been filled, Sets the session for a Non-Required Field
Public Sub pcs_ValidateTextField(FieldName,isRequiredField,MaxLength)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	pcv_strFieldName=getUserInput(trim(request.form(FieldName)),MaxLength)
	session(pcv_strSessionPrefix&FieldName)=pcv_strFieldName
	if ( (isRequiredField="True" OR isRequiredField=true) AND pcv_strFieldName="" ) then
		pcv_intErr=pcv_intErr+1
	end if
End Sub


'// Use this in place of Text Field when basic html formatting is acceptable.
Public Sub pcs_ValidateHTMLField(FieldName,isRequiredField,MaxLength)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	pcv_strFieldName=request.form(FieldName)	 
	known_bad= array("*","--")
	if stringLength>0 then
		pcv_strFieldName = left(trim(pcv_strFieldName),0) 
	else
		pcv_strFieldName = trim(pcv_strFieldName)
	end if
	for i=lbound(known_bad) to ubound(known_bad)
	if (instr(1,pcv_strFieldName,known_bad(i),vbTextCompare)<>0) then
		pcv_strFieldName	= replace(pcv_strFieldName,known_bad(i),"")
	end if
	next	
	pcv_strFieldName	= replace(pcv_strFieldName,"%0d","")
	pcv_strFieldName	= replace(pcv_strFieldName,"%0D","")
	pcv_strFieldName	= replace(pcv_strFieldName,"%0a","")
	pcv_strFieldName	= replace(pcv_strFieldName,"%0A","")
	pcv_strFieldName	= replace(pcv_strFieldName,"\r\n","")
	pcv_strFieldName	= replace(pcv_strFieldName,"\r","")
	pcv_strFieldName	= replace(pcv_strFieldName,"\n","")
	pcv_strFieldName	= replace(pcv_strFieldName,"\R\N","")
	pcv_strFieldName	= replace(pcv_strFieldName,"\R","")
	pcv_strFieldName	= replace(pcv_strFieldName,"\N","")		
	session(pcv_strSessionPrefix&FieldName)=pcv_strFieldName
	if ( (isRequiredField="True" OR isRequiredField=true) AND pcv_strFieldName="" ) then
		pcv_intErr=pcv_intErr+1
	end if
End Sub


'// This Validates State or Province has been filled, Sets the session for a Non-Required Field
Public Sub pcs_ValidateStateProvField(FieldName,isRequiredField,MaxLength)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	Session("Err"&FieldName)=""
	pcv_strFieldName=getUserInput(trim(request.form(FieldName)),MaxLength)
	session(pcv_strSessionPrefix&FieldName)=pcv_strFieldName
	if isRequiredField="True" OR isRequiredField=true then
		if pcv_strFieldName="" then
			pcv_intErr=pcv_intErr+1
		else
			'if custom then 
			'	pcv_intErr=pcv_intErr+1
			'	Session("Err"&FieldName)=1
			'end if 
		end if
	end if
End Sub

'// This Validates Email Addresses  
Public Sub pcs_ValidateEmailField(FieldName,isRequiredField,MaxLength)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	Session("Err"&FieldName)=""
	pcv_strFieldName=getUserInput(trim(request.form(FieldName)),MaxLength)
	pcv_strFieldName=replace(pcv_strFieldName," ","")
	session(pcv_strSessionPrefix&FieldName)=pcv_strFieldName
	if isRequiredField="True" OR isRequiredField=true then
		if pcv_strFieldName="" then
			pcv_intErr=pcv_intErr+1
		else
			if instr(pcv_strFieldName,"@")=0 OR instr(pcv_strFieldName,".")=0 then 
				pcv_intErr=pcv_intErr+1
				Session("Err"&FieldName)=1
			end if 
		end if
	end if
End Sub




'// This Validates Phone Numbers    
Public Sub pcs_ValidatePhoneNumber(FieldName,isRequiredField,MaxLength)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	Session("Err"&FieldName)=""
	pcv_strFieldName=getUserInput(trim(request.form(FieldName)),MaxLength)
	pcv_strFieldName=replace(pcv_strFieldName," ","")
	session(pcv_strSessionPrefix&FieldName)=pcv_strFieldName
	if isRequiredField="True" OR isRequiredField=true then
		if pcv_strFieldName="" then
			pcv_intErr=pcv_intErr+1
		else
			pcv_strFieldName=replace(pcv_strFieldName," ","")
			pcv_strFieldName=replace(pcv_strFieldName,"-","")
			pcv_strFieldName=replace(pcv_strFieldName,".","")
			pcv_strFieldName=replace(pcv_strFieldName,"(","")
			pcv_strFieldName=replace(pcv_strFieldName,")","")
			if Not isNumeric(pcv_strFieldName) then 
				pcv_intErr=pcv_intErr+1
				Session("Err"&FieldName)=1
			end if 
		end if
	end if
End Sub




'// This Controls the Validation Display Icons
Public Sub 	pcs_RequiredImageTag(FieldName,isRequiredField)
	if isRequiredField="True" OR isRequiredField=true then
		if msg="" then 
		%>
		<img src="<%=pcv_strRequiredIcon%>"> 
		<% 
		else
			if session(pcv_strSessionPrefix&FieldName)="" then 
			%>
			<img src="<%=pcv_strErrorIcon%>"> 
			<% 
			end if
		end if 
	end if
	' closes the session if not closed
	if len(session(pcv_strSessionPrefix&FieldName))>0 then
		session(pcv_strSessionPrefix&FieldName)=""
	end if
	' closes the special error session if not closed
	if Session("Err"&FieldName)<>"" then
		Session("Err"&FieldName)=""
	end if
End Sub

'// This Controls the Validation Display Icons
Public Sub 	pcs_UPSRequiredImageTag(FieldName,isRequiredField)
	if isRequiredField="True" OR isRequiredField=true then
		if msg<>"" AND session(pcv_strSessionPrefix&FieldName)="" then 
			%>
			<img src="<%=pcv_strErrorIcon%>"> 
			<% 
		else
			%>
			<img src="<%=pcv_strRequiredIcon%>"> 
			<% 
		end if
	end if
	' closes the session if not closed
	if len(session(pcv_strSessionPrefix&FieldName))>0 then
		session(pcv_strSessionPrefix&FieldName)=""
	end if
	' closes the special error session if not closed
	if Session("Err"&FieldName)<>"" then
		Session("Err"&FieldName)=""
	end if
End Sub


'// This Fills the form fields and closes the sessions.
Public Function pcf_FillFormField(FieldName,isRequiredField)
	pcf_FillFormField=Session(pcv_strSessionPrefix&FieldName)
	if isNULL(pcf_FillFormField)=True then
		pcf_FillFormField=""
	end if
	if len(pcf_FillFormField)>0 then
		pcf_FillFormField=replace(pcf_FillFormField,"''","'")
	end if
	if (Not isRequiredField="True") OR (Not isRequiredField=true) then
		' closes the session if not closed
		if len(session(pcv_strSessionPrefix&FieldName))>0 then
			session(pcv_strSessionPrefix&FieldName)=""
		end if
		' closes the special error session if not closed
		if Session("Err"&FieldName)<>"" then
			Session("Err"&FieldName)=""
		end if
	end if
End Function


'// This removes the double apostrophes that were added by "getUserInput" for database insert.
Public Function pcf_ReverseGetUserInput(FieldValue)
	if isNULL(FieldValue)=True then
		FieldValue=""
	end if
	if len(FieldValue)>0 then
		FieldValue=replace(FieldValue,"''","'")
	end if
	pcf_ReverseGetUserInput=FieldValue
End Function

'// This Selects a dropdown if there is a match
Public Function pcf_SelectOption(FieldName, OptionValue)
	pcf_SelectOption=Session(pcv_strSessionPrefix&FieldName)
	if pcf_SelectOption = OptionValue then
		pcf_SelectOption = "selected"
	else
		pcf_SelectOption = ""
	end if
End Function

'// This Checks a box if there is a match
Public Function pcf_CheckOption(FieldName, OptionValue)
	pcf_CheckOption=Session(pcv_strSessionPrefix&FieldName)
	if pcf_CheckOption = OptionValue then
		pcf_CheckOption = "checked"
	else
		pcf_CheckOption = ""
	end if
End Function


'// This Takes an Action based off the Results  
Public Sub pcs_ValidateResults  
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&redirectUrl=" & pRedirectURL
	Else
		'// Sheri is recoding the session database part.
		'// For now clear the sessions and redirect back so we can test more.
		pcs_ClearAllSessions '(NOTE: If the page validation fails in anyway the sessions auto clear, but not until the field value is refilled)
		response.redirect pcStrPageName
	End If
End Sub


'// This Clears All Session Before we leave the page
Public Sub pcs_ClearAllSessions
	for each field in request.form 
		' obtain field name
		FieldName=field
		' obtain field value
		fieldValue=request.form(FieldName) 
		' closes the session if not closed
		if len(session(pcv_strSessionPrefix&FieldName))>0 then
			session(pcv_strSessionPrefix&FieldName)=""
		end if
		' closes the special error session if not closed
		if Session("Err"&FieldName)<>"" then
			Session("Err"&FieldName)=""
		end if
	next 	
End Sub


' Generates the correct icon path
public function pcf_GenerateIconURL(currentPath)  					
	pcf_GenerateIconURL = currentPath	
end function

	
'// Validate EU VAT ID
Public Sub pcs_ValidateVATIDField(FieldName, FldReq, ISOCountryCode)

	Dim pcv_strFieldName, isValid
	
	isValid = True
	pcv_strFieldName=getUserInput(trim(request.form(FieldName)),50)
		
	session(pcv_strSessionPrefix&FieldName)=pcv_strFieldName

	if (pcv_strFieldName="" and FldReq) then
		pcv_intErr=pcv_intErr+1
		Session("Err"&FieldName)=1
	end if
	if (pcv_strFieldName<>"" and FldReq) then
		ISOCountryCode=UCASE(ISOCountryCode)
		Select Case ISOCountryCode
			Case "IT": isValid = VerificaPIVA(trim(pcv_strFieldName))
			Case Else: isValid = True
		End Select 
		if not (isValid) then
			pcv_intErr=pcv_intErr+1
			Session("Err"&FieldName)=1
		End IF	
	end if
			
End Sub


'// Validate SSN
Public Sub pcs_ValidateSSNField(FieldName, FldReq, ISOCountryCode)

	Dim pcv_strFieldName, isValid
	
	isValid = False
	pcv_strFieldName=getUserInput(trim(request.form(FieldName)),50)
		
	session(pcv_strSessionPrefix&FieldName)=pcv_strFieldName

	if (pcv_strFieldName="" and FldReq) then
		pcv_intErr=pcv_intErr+1
		Session("Err"&FieldName)=1
	end if
	if (pcv_strFieldName<>"" and FldReq) then
		ISOCountryCode=UCASE(ISOCountryCode)				
		Select Case ISOCountryCode
			Case "IT": isValid = VerificaCF(trim(pcv_strFieldName))
			Case Else: isValid = True
		End Select 
		if not (isValid) then
			pcv_intErr=pcv_intErr+1
			Session("Err"&FieldName)=1
		End IF	
	end if
			
End Sub



'// VAT ID: IT (Italy)
function VerificaPIVA(pi)

	'---------------------------------------------------------------------
	if Len(pi) <> 11 then
		VerificaPIVA=false
	else
		'-----------------------------------------------------------
			
		Set objER = New RegExp
		objER.Global = True
		objER.IgnoreCase = True
		objER.Pattern = "^[0-9]+$"
		
		' verifica la corrispondenza con il pattern
		result = objER.Test(pi)
		if result <> true then 
			VerificaPIVA=false
			Set objER = Nothing
		else 
			'------------------------------------------------
			Dim s, s1, s2, c, i, char
			s1 = 0
			for i = 0 to 9
				i = i + 1 
				char = mid(pi , i , 1 )
				s1 = s1 + asc(char) - asc("0")
			next
	
			for i = 1 to 9 
				i = i + 1 
				char = mid(pi , i , 1 )
				c = 2* ( asc(char) - asc("0"))
				if c > 9 then
					c = c - 9
					s2 = s2 + c
				else
					s2 = s2 + c
				end if
			next
			s = s1 + s2
			if( ( 10 - s Mod 10 ) mod 10 <>  asc(Mid(pi, 11, 1)) - asc("0") ) then
				'ControllaPIVA(pi)
				VerificaPIVA=false
			else
				VerificaPIVA=true
			end if
		
		'------------------------------------------------
		end if
	'------------------------------------------------------------
	end if
	'---------------------------------------------------------------------

end function	
	

'// SSN: IT (Italy)
function VerificaCF(CodiceFiscale)
    CodiceFiscale = ucase(CodiceFiscale)
   If len(CodiceFiscale) < 16 Then
		'Check if SSN is a VATID
		VerificaCF = VerificaPIVA(CodiceFiscale)
		Exit Function
   End IF

	Set objER = New RegExp
	objER.Global = True
	objER.IgnoreCase = True
	objER.Pattern = "[^A-Za-z0-9]"
	
	result = objER.Test(CodiceFiscale)
	if result then
		VerificaCF = false
		Set objER = Nothing
		Exit Function
	End If

		
	Dim Lettere(35,2)

	Dim ConfrontoCarattereControllo(25)

	Dim I
	Dim J

	Dim Carattere
	Dim ValorePari
	Dim ValoreDispari
	Dim SommaCaratteri
	Dim PariDispari
	Dim Risultato

	Dim CarattereControllo
	Dim Temp
	Dim Test

	Lettere(0,0) = "A"
	Lettere(0,1) = "0"
	Lettere(0,2) = "1"

	Lettere(1,0) = "B"
	Lettere(1,1) = "1"
	Lettere(1,2) = "0"

	Lettere(2,0) = "C"
	Lettere(2,1) = "2"
	Lettere(2,2) = "5"

	Lettere(3,0) = "D"
	Lettere(3,1) = "3"
	Lettere(3,2) = "7"

	Lettere(4,0) = "E"
	Lettere(4,1) = "4"
	Lettere(4,2) = "9"

	Lettere(5,0) = "F"
	Lettere(5,1) = "5"
	Lettere(5,2) = "13"

	Lettere(6,0) = "G"
	Lettere(6,1) = "6"
	Lettere(6,2) = "15"

	Lettere(7,0) = "H"
	Lettere(7,1) = "7"
	Lettere(7,2) = "17"

	Lettere(8,0) = "I"
	Lettere(8,1) = "8"
	Lettere(8,2) = "19"

	Lettere(9,0) = "J"
	Lettere(9,1) = "9"
	Lettere(9,2) = "21"

	Lettere(10,0) = "K"
	Lettere(10,1) = "10"
	Lettere(10,2) = "2"

	Lettere(11,0) = "L"
	Lettere(11,1) = "11"
	Lettere(11,2) = "4"

	Lettere(12,0) = "M"
	Lettere(12,1) = "12"
	Lettere(12,2) = "18"

	Lettere(13,0) = "N"
	Lettere(13,1) = "13"
	Lettere(13,2) = "20"

	Lettere(14,0) = "O"
	Lettere(14,1) = "14"
	Lettere(14,2) = "11"

	Lettere(15,0) = "P"
	Lettere(15,1) = "15"
	Lettere(15,2) = "3"

	Lettere(16,0) = "Q"
	Lettere(16,1) = "16"
	Lettere(16,2) = "6"

	Lettere(17,0) = "R"
	Lettere(17,1) = "17"
	Lettere(17,2) = "8"

	Lettere(18,0) = "S"
	Lettere(18,1) = "18"
	Lettere(18,2) = "12"

	Lettere(19,0) = "T"
	Lettere(19,1) = "19"
	Lettere(19,2) = "14"

	Lettere(20,0) = "U"
	Lettere(20,1) = "20"
	Lettere(20,2) = "16"

	Lettere(21,0) = "V"
	Lettere(21,1) = "21"
	Lettere(21,2) = "10"

	Lettere(22,0) = "W"
	Lettere(22,1) = "22"
	Lettere(22,2) = "22"

	Lettere(23,0) = "X"
	Lettere(23,1) = "23"
	Lettere(23,2) = "25"

	Lettere(24,0) = "Y"
	Lettere(24,1) = "24"
	Lettere(24,2) = "24"

	Lettere(25,0) = "Z"
	Lettere(25,1) = "25"
	Lettere(25,2) = "23"

	Lettere(26,0) = "0"
	Lettere(26,1) = "0"
	Lettere(26,2) = "1"

	Lettere(27,0) = "1"
	Lettere(27,1) = "1"
	Lettere(27,2) = "0"

	Lettere(28,0) = "2"
	Lettere(28,1) = "2"
	Lettere(28,2) = "5"

	Lettere(29,0) = "3"
	Lettere(29,1) = "3"
	Lettere(29,2) = "7"

	Lettere(30,0) = "4"
	Lettere(30,1) = "4"
	Lettere(30,2) = "9"

	Lettere(31,0) = "5"
	Lettere(31,1) = "5"
	Lettere(31,2) = "13"

	Lettere(32,0) = "6"
	Lettere(32,1) = "6"
	Lettere(32,2) = "15"

	Lettere(33,0) = "7"
	Lettere(33,1) = "7"
	Lettere(33,2) = "17"

	Lettere(34,0) = "8"
	Lettere(34,1) = "8"
	Lettere(34,2) = "19"

	Lettere(35,0) = "9"
	Lettere(35,1) = "9"
	Lettere(35,2) = "21"

	For I = 0 To 25

		ConfrontoCarattereControllo(I) = Chr(65 + I) 'creo in ConfrontoCarattereControllo tutte le lettere maiuscole dalla A (chr(65)) alla Z(chr(90))

	Next

	Carattere=0
	ValorePari=1 'indice della seconda colonna della matrice Lettere
	ValoreDispari=2 'indice della terza colonna della matrice Lettere
	SommaCaratteri=0
	CarattereControllo=Right(CodiceFiscale,1)

	for I=1 to len(CodiceFiscale)-1
		if (I mod 2)=0 then
			PariDispari="P"
		else
			PariDispari="D"
		end if

		Temp =mid(CodiceFiscale,I,1)
		J=0
		do
			Test=Lettere(J,Carattere)
			J=J+1
		loop until Temp=Test 

		J=J-1

		if PariDispari="P" then
			SommaCaratteri=SommaCaratteri + CInt(Lettere(J,ValorePari))
		else
			SommaCaratteri=SommaCaratteri + CInt(Lettere(J,ValoreDispari))
		end if
	Next

	Risultato=SommaCaratteri mod 26

	Risultato=ConfrontoCarattereControllo(Risultato)

	if Risultato<>CarattereControllo then
		VerificaCF=false
	else
		VerificaCF=true
	end if
end function
%>