<%
'--- ProductCart - GGG Add-on - Gift Certificate Code Generator

'--- Set variables for Gift Code string that will be returned to ProductCart
Dim GiftCode,CodeStrLen
GiftCode=""

'Length of Gift Code String
CodeStrLen=10

'--- Begin Code Generator Function
Sub GenGiftCode(IdOrder,OrderDate,ProcessDate,IdCustomer,IdProduct,SKU)

'--- BEGIN GIFT CERTIFICATE CODE GERNERATOR
'--- Replace this code with yours to generate your own Gift Certificates
'--- Each Gift Certificate can use a different Code Generator
'--- Just save this file with a new name, and enter that name on the Add/Modify product page

'--- Start Generating Gift Code - This is just SAMPLE CODE
    Randomize( )

    dim CharacterSetArray
    CharacterSetArray = Array( _
      Array( 10, "abcdefghijklmnopqrstuvwxyz" ), _
      Array( 5, "0123456789" ) _
    )

    dim i
    dim j
    dim Count
    dim Chars
    dim Index
    dim Temp

    for i = 0 to UBound( CharacterSetArray )

      Count = CharacterSetArray( i )( 0 )
      Chars = CharacterSetArray( i )( 1 )

      for j = 1 to Count

        Index = Int( Rnd( ) * Len( Chars ) ) + 1
        Temp = Temp & Mid( Chars, Index, 1 )

      next

    next

    dim TempCopy

    do until Len( Temp ) = 0

      Index = Int( Rnd( ) * Len( Temp ) ) + 1
      TempCopy = TempCopy & Mid( Temp, Index, 1 )
      Temp = Mid( Temp, 1, Index - 1 ) & Mid( Temp, Index + 1 )

    loop

    RandomString = TempCopy


'--- END OF GIFT CERTIFICATE CODE GERNERATOR

'--- Assign the values generated above to a temporary gift code	
'--- You should not need to edit the code below this line
	GiftCode=RandomString

End Sub
'--- End of License Generator Function


'MAIN PROGRAM

'---- Begin receiving information from the store
pIdOrder=request("IdOrder")
if pIdOrder="" then
pIdOrder=0
end if
if IsNumeric(pIdOrder)=false then
pIdOrder=0
end if
pOrderDate=request("OrderDate")
if pOrderDate="" then
pOrderDate=Date()
end if
if IsDate(pOrderDate)=false then
pOrderDate=Date()
end if
pProcessDate=request("ProcessDate")
if pProcessDate="" then
pProcessDate=Date()
end if
if IsDate(pProcessDate)=false then
pProcessDate=Date()
end if
pIdCustomer=request("IdCustomer")
if pIdCustomer="" then
pIdCustomer=0
end if
if IsNumeric(pIdCustomer)=false then
pIdCustomer=0
end if
pIdProduct=request("IdProduct")
if pIdProduct="" then
pIdProduct=0
end if
if IsNumeric(pIdProduct)=false then
pIdProduct=0
end if
pQuantity=request("Quantity")
if pQuantity="" then
pQuantity=0
end if
if IsNumeric(pQuantity)=false then
pQuantity=0
end if
pSKU=request("SKU")
if pSKU="" then
pSKU="12345"
end if
'---- End of receiving information


'---- Call Code Generator Function ----

Call GenGiftCode(pIdOrder,pOrderDate,pProcessDate,pIdCustomer,pIdProduct,pSKU)

'---- End Call ----


'--- Return Gift Code Strings to the store -----
response.write pIdOrder & "<br>"
response.write pIdProduct & "<br>"
response.write GiftCode & "<br>"
'--- End of Return ---

%>