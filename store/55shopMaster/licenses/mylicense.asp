<%
'--- ProductCart - Downloadable Products Module - License Generator

'--- Set variables for License Key strings that will be returned to ProductCart
'--- A license can contain up to 5 different pieces of information
Dim LicSTR1, LicSTR2, LicSTR3, LicSTR4, LicSTR5
LicSTR1=""
LicSTR2=""
LicSTR3=""
LicSTR4=""
LicSTR5=""

'--- Begin License Generator Function
Sub GenLincense(IdOrder,OrderDate,ProcessDate,IdCustomer,IdProduct,Index,SKU)

'---- Create variables for temporary license keys  
Dim Lic1, Lic2, Lic3, Lic4, Lic5

'--- BEGIN LICENSE GERNERATOR CODE
'--- Replace this code with your license generating script
'--- You do not need to create scripts for all 5 variables (e.g. the product just needs one Serial Number). For instance, if you did not need to assign a value to the 5th variable, you would replace that section of the code with simply Tn5=""
'--- Each Downloadable Product can use a different License Generator
'--- Just save this file with a new name, and enter that name on the Add/Modify product page
'--- The sample code below assigns values to all 5 variables.

'--- Start Generating License Key 1 - This is just SAMPLE CODE
	Tn1=""
	For dd=1 to 6
	Randomize
	Tn1=Tn1 & Cstr(Fix(10*Rnd))
	Next
'--- Finish Generating License Key 1	

'--- Start Generating License Key 2 - This is just SAMPLE CODE	
	Tn2=""
	For dd=1 to 12
	Randomize
	myC=Fix(2*Rnd)
	Select Case myC
	Case 0: Tn2=Tn2 & Cstr(Fix(10*Rnd))
	Case 1: Tn2=Tn2 & Chr(Fix(26*Rnd)+97)		
	End Select
	Next
'--- Finish Generating License Key 2	

'--- Start Generating License Key 3 - This is just SAMPLE CODE	
	Tn3=""
	For dd=1 to 24
	Randomize
	myC=Fix(3*Rnd)
	Select Case myC
	Case 0: Tn3=Tn3 & Chr(Fix(26*Rnd)+65)
	Case 1: Tn3=Tn3 & Cstr(Fix(10*Rnd))
	Case 2: Tn3=Tn3 & Chr(Fix(26*Rnd)+97)		
	End Select
	Next
'--- Finish Generating License Key 3

'--- Start Generating License Key 4 - This is just SAMPLE CODE	
	Tn4=""
	For dd=1 to 24
	Randomize
	myC=Fix(3*Rnd)
	Select Case myC
	Case 0: Tn4=Tn4 & Chr(Fix(26*Rnd)+65)
	Case 1: Tn4=Tn4 & Cstr(Fix(10*Rnd))
	Case 2: Tn4=Tn4 & Chr(Fix(26*Rnd)+97)		
	End Select
	Next
'--- Finish Generating License Key 4

'--- Start Generating License Key 5 - This is just SAMPLE CODE	
	Tn5=""
	For dd=1 to 24
	Randomize
	myC=Fix(3*Rnd)
	Select Case myC
	Case 0: Tn5=Tn5 & Chr(Fix(26*Rnd)+65)
	Case 1: Tn5=Tn5 & Cstr(Fix(10*Rnd))
	Case 2: Tn5=Tn5 & Chr(Fix(26*Rnd)+97)		
	End Select
	Next
'--- Finish Generating License Key 5
	
'--- Assign the values generated above to the temporary license keys	
	Lic1=Tn1
	Lic2=Tn2
	Lic3=Tn3
	Lic4=Tn4
	Lic5=Tn5		

'--- END OF LICENSE GERNERATOR CODE
'--- You should not need to edit the code below this line

'--- Assign values to the 5 License Key strings

LicSTR1=LicSTR1 & Lic1 & "***"
LicSTR2=LicSTR2 & Lic2 & "***"
LicSTR3=LicSTR3 & Lic3 & "***"
LicSTR4=LicSTR4 & Lic4 & "***"
LicSTR5=LicSTR5 & Lic5 & "***"

End Sub
'--- End of License Generator Function


'MAIN PROGRAM

'---- Begin receiving information from the store
pIdOrder=request("IdOrder")
pOrderDate=request("OrderDate")
pProcessDate=request("ProcessDate")
pIdCustomer=request("IdCustomer")
pIdProduct=request("IdProduct")
pQuantity=request("Quantity")
pSKU=request("SKU")
'---- End of receiving information


'---- IF Product Quantity is a number
if IsNumeric(pQuantity) then

'--- then repeat call to License Generator Function N times (N=Product Quantity) ---

For i=1 to CLng(pQuantity)

'---- Call License Generator Function ----

Call GenLincense(pIdOrder,pOrderDate,pProcessDate,pIdCustomer,pIdProduct,i,pSKU)

'---- End Call ----

Next

'--- End of repeat

'--- Return License Keys Strings to the store -----
response.write pIdOrder & "<br>"
response.write pIdProduct & "<br>"
response.write LicSTR1 & "<br>"
response.write LicSTR2 & "<br>"
response.write LicSTR3 & "<br>"
response.write LicSTR4 & "<br>"
response.write LicSTR5 & "<br>"
'--- End of Return ---

end if
'--- END IF Product Quantity is a number -----
%>