<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<%
'***********************************************************************************
' START: PROCESS CALLBACK
'***********************************************************************************
Function processMerchantCalculationCallback(domMcCallbackObj)
    Dim xmlMcResults	
	
	'// Process <merchant-calculation-callback> and create <merchant-calculation-results>
    xmlMcResults = createMerchantCalculationResults(domMcCallbackObj)	
    
	'// Respond with <merchant-calculation-results> XML
    Response.write xmlMcResults	
End Function
'***********************************************************************************
' END: PROCESS CALLBACK
'***********************************************************************************


'***********************************************************************************
' START: RECIEVE THE CALLBACK
'***********************************************************************************
Function createMerchantCalculationResults(domMcCallbackObj)
	on error resume next
    '// Define the objects used to create the <merchant-calculation-callback> 
    Dim domMcResultsObj
    Dim domMcResults
    Dim domMerchantCodeResults
    Dim domMerchantCodeResultsRoot
    Dim domResults
    Dim domResult
    Dim domResponse
    Dim domMcCallbackObjRoot
    Dim domTaxList
    Dim calcTax
    Dim domMethodList
    Dim domMethod
    Dim attrShippingName
    Dim domAnonymousAddressList
    Dim domAnonymousAddress
    Dim attrAddressId
    Dim domMerchantCodeList
    Dim totalTax
    Dim domTotalTax
    Dim domShippingRate
    Dim shippingRate
    Dim domShippable
    Dim shippable
	Dim attrAvailableList
	Dim attrPriceList
	Dim Nodes
	Dim Node
	Dim taxLoc
	Dim	taxPrdAmount
	Dim pTaxableTotal, pSubTotal, pCartTotalWeight, pCartQuantity, pEryPassword
	Dim pshippingStateCode, pshippingCountryCode
	Dim taxCalcAmt
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Get XML Results
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Set domMcResultsObj = Server.CreateObject(strMsxmlDomDocument)
    domMcResultsObj.async = False
    domMcResultsObj.appendChild( domMcResultsObj.createProcessingInstruction("xml", strXmlVersionEncoding))
	
    '// Create root tag for <merchant-calculation-results> response and set xmlns attribute
    Set domMcResults = domMcResultsObj.appendChild( domMcResultsObj.createElement("merchant-calculation-results"))
    domMcResults.setAttribute "xmlns", strXmlns
    
	'// Create child element <results>
    Set domResults = domMcResults.appendChild( domMcResultsObj.createElement("results"))
    Set domMcCallbackObjRoot = domMcCallbackObj.documentElement
    
	'// Retrieve Boolean value indicating whether merchant calculates tax for the order.
    Set domTaxList = domMcCallbackObjRoot.getElementsByTagname("tax")
    calcTax = domTaxList(0).text
   
    '// Retrieve the names of the shipping methods available for the order
    Set domMethodList = domMcCallbackObjRoot.getElementsByTagname("method")
   
    '// Retrieve shipping addresses from the <merchant-calculated-callback>
	Set AddressNodes = domMcCallbackObj.selectNodes("//calculate/addresses/anonymous-address")
   
    '// Retrieve a list of coupon and gift certificate codes that should be applied to the order total.
    Set domMerchantCodeList = domMcCallbackObjRoot.getElementsByTagname("merchant-code-string")		
	
	'// Retrieve the Shopping Cart
	Set Nodes = domMcCallbackObj.selectNodes("//shopping-cart/merchant-private-data")
	For Each Node In Nodes
		pcv_strMerchantNote = Node.selectSingleNode("merchant-note").text
	Next	
	
	'// Split the Merchant Private Data to get the Key
	pcArray_MerchantNote = split(pcv_strMerchantNote, chr(124))	
	pcStrCustomerRefKey=pcArray_MerchantNote(0)		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Get XML Results
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Populate the Cart Array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	Dim pcCartArray(100,45)			
	query="SELECT *  FROM pcCartArray  WHERE pcCartArray_Key="& pcStrCustomerRefKey &""	
	set rs2=server.CreateObject("ADODB.RecordSet")
	set rs2=conntemp.execute(query)		
	f=1
	Do while NOT rs2.eof
		for x=0 to 45
			pcCartArray(f,x)=rs2("pcCartArray_"&x)			
		next
		'// Create an item row price for this Cart Index
		itemIndex = f-1
		session("pcUnitPrice"&itemIndex) = (cCur(domMcCallbackObjRoot.getElementsByTagname("unit-price").item(itemIndex).text) * pcCartArray(f,2))
	f=f+1
	rs2.movenext
	loop	
	set rs2=nothing	
	session("pcCartSession")=pcCartArray 
	pcCartIndex=f-1
	session("pcCartIndex")=pcCartIndex
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Populate the Cart Array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



    '// Loop through address IDs to build <result> elements    
	For Each AddressNode In AddressNodes
		
		Session("pcSFcity") = ""
		Session("pcSFStateCode") = ""
		Session("pcSFzip") = ""
		Session("pcSFCountryCode") = ""		
		Session("AddressDenied")=""

        '// Retrieve the address ID
		attrAddressId = AddressNode.getAttribute("id")	
		'// Retrieve the city
		Session("pcSFcity") = AddressNode.selectSingleNode("city").text
		'// Retrieve the state
		Session("pcSFStateCode") = AddressNode.selectSingleNode("region").text	
		'// Retrieve the zip
		Session("pcSFzip") = AddressNode.selectSingleNode("postal-code").text
		'// Retrieve the county code
		Session("pcSFCountryCode") = AddressNode.selectSingleNode("country-code").text
		
		'// Delivery Zip Codes		
		If DeliveryZip = "1" Then			
			query="SELECT zipcodevalidation.zipcode from zipcodevalidation WHERE zipcode='" &Session("pcSFzip")& "'"
			set rsZipCodeObj=server.CreateObject("ADODB.RecordSet")
			set rsZipCodeObj=conntemp.execute(query)
			if rsZipCodeObj.eof then								
				Session("AddressDenied")="1"
			end if	
			set rsZipCodeObj=nothing
		End If
		taxLoc=0
		taxPrdAmount=0
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Tax Rates
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		If ptaxfile=1 then
				
			'// Dynamically retrieve the current directory
			Dim sScriptDir
			sScriptDir = Request.ServerVariables("SCRIPT_NAME")
			sScriptDir = StrReverse(sScriptDir)
			sScriptDir = Mid(sScriptDir, InStr(1, sScriptDir, "/"))
			sScriptDir = StrReverse(sScriptDir)
	
			'// Get the file name
			dim Filename
			Filename=ptaxfilename
			Const ForReading = 1, ForWriting = 2, ForAppending = 8 
	
			dim FSO
			set FSO = Server.CreateObject("scripting.FileSystemObject") 
			
			'// Map the logincal path to the physical system path
			Dim Filepath
			Filepath=Server.MapPath(sScriptDir) & "\tax\" & Filename
			if NOT FSO.FileExists(Filepath) then
				pcv_intError = 1
			end if
			
			TAX_SHIPPING_ALONE="NA"
			TAX_SHIPPING_AND_HANDLING_TOGETHER="NA"
			zipCnt=1
	
			'*******************************************************************************
			' START: TAX FROM FILE
			'*******************************************************************************
			If Session("pcSFCountryCode")="US" OR Session("pcSFCountryCode")="CA" then
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: CHECK STATE IS TAXABLE
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// See if state is a taxable one and then flag
				taxStateArray=split(ptaxRateState,", ")
				taxRateArray=split(ptaxRateDefault,", ")
				taxSNHArray=split(ptaxSNH,", ")
				intTaxableState=0
				for i=0 to ubound(taxStateArray)-1
					if taxStateArray(i)=Session("pcSFStateCode") then
						'// Flag
						intTaxableState=1
						
						if ubound(taxRateArray)>-1 then
							intTaxRateDefault=taxRateArray(i)
						else
							intTaxRateDefault=0
						end if
						
						if ubound(taxSNHArray)>-1 then
							strTaxSNH=taxSNHArray(i)
						else
							strTaxSNH="NN"
						end if
						select case strTaxSNH
							case "YY"
								TAX_SHIPPING_ALONE="Y"
								TAX_SHIPPING_AND_HANDLING_TOGETHER="Y"
							case "YN"
								TAX_SHIPPING_ALONE="Y"
								TAX_SHIPPING_AND_HANDLING_TOGETHER="N"
							case "NN"
								TAX_SHIPPING_ALONE="N"
								TAX_SHIPPING_AND_HANDLING_TOGETHER="N"
						end select
					end if
				next
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: CHECK STATE IS TAXABLE
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				
				if intTaxableState=1 then
				
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: GET AVAILABLE TAX RATES
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					dim f
					set f=FSO.GetFile(Filepath)
					
					Dim TextStream
					set TextStream=f.OpenAsTextStream(ForReading, -2) 
					
					zipCnt=0
					optionStr=""
					do While NOT TextStream.AtEndOfStream
					  '// Line is found, now write it to new string
						Line=TextStream.readline
						'// Ignore First Line
						if instr(ucase(Line), "ZIP") then
							iArray=split(Line,",")
							'// Loop to find correct array for each
							pcv_PostalCodeColumnFlag=False
							for q=0 to ubound(iArray)
								if iArray(q)="ZIP_CODE" then '// identify the zip code column
									ZIP_CODE_NUM=q
									pcv_PostalCodeColumnFlag=True
									'response.write q&"<BR>"
								end if
								if iArray(q)="COUNTY_NAME" then
									COUNTY_NAME_NUM=q								
								end if
								if iArray(q)="CITY_NAME" then
									CITY_NAME_NUM=q								
								end if
								if iArray(q)="TOTAL_SALES_TAX" then
									TOTAL_SALES_TAX_NUM=q								
								end if
								if iArray(q)="TAX_SHIPPING_ALONE" then
									TAX_SHIPPING_ALONE_NUM=q								
								end if
								if iArray(q)="TAX_SHIPPING_AND_HANDLING_TOGETHER" then
									TAX_SHIPPING_AND_HANDLING_TOGETHER_NUM=q								
								end if
							next						
						else
							'// SEE IF MORE THEN ONE ZIP CODE EXIST
							if instr(Line, pcBillingPostalCode) then
								zArray=split(Line,",")						
								pcv_ValidPostalCode=False						
								if pcv_PostalCodeColumnFlag=True then '// If we can identify the zip code column, then check it's valid
									ZIP_CODE_NAME=zArray(ZIP_CODE_NUM)
									if instr(ZIP_CODE_NAME, pcBillingPostalCode) then
										pcv_ValidPostalCode=True
									end if
								end if
							
								if pcv_PostalCodeColumnFlag=False OR pcv_ValidPostalCode=True then
									zipCnt=zipCnt+1								
									COUNTY_NAME=zArray(COUNTY_NAME_NUM)
									CITY_NAME=zArray(CITY_NAME_NUM)
									TOTAL_SALES_TAX=zArray(TOTAL_SALES_TAX_NUM)
									TAX_SHIPPING_ALONE=zArray(TAX_SHIPPING_ALONE_NUM)
									TAX_SHIPPING_AND_HANDLING_TOGETHER=zArray(TAX_SHIPPING_AND_HANDLING_TOGETHER_NUM)
									exit do
								end if
							end if
						end if
					loop			
					TextStream.Close	
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: GET AVAILABLE TAX RATES
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: CALCULATE TAX FINAL RATE
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// One postal code
					If zipCnt=0 AND intTaxableState=1 then
						taxLoc=taxLoc+(intTaxRateDefault/100) 
					End If
					
					'// Many postal codes
					if zipCnt>1 then
						taxLoc=taxLoc+(TOTAL_SALES_TAX)
						zipCnt=1
					else
						taxLoc=taxLoc+(TOTAL_SALES_TAX)
						zipCnt=1
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: CALCULATE TAX FINAL RATE
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
				end if
				
			Else
			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: CALCULATE TAX FINAL RATE (NON "US" OR "CA")
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				taxLoc=taxLoc+(TOTAL_SALES_TAX)
				zipCnt=1
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: CALCULATE TAX FINAL RATE (NON "US" OR "CA")
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
			End if
			'*******************************************************************************
			' END: TAX FROM FILE
			'*******************************************************************************

		
		
		else '// Tax File Split
		
		
			if ptaxVAT="1" then			
			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: VAT
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				taxLoc=0
				TAX_SHIPPING_ALONE="NA"
				TAX_SHIPPING_AND_HANDLING_TOGETHER="NA"
				zipCnt=1
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: VAT
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			
			else			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: Tax Per Place
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					query="SELECT taxLoc.taxLoc, taxLoc.taxDesc FROM taxLoc "
					query=query& "WHERE ((stateCode='" &Session("pcSFStateCode")& "' AND stateCodeEq=-1) "
					query=query& "OR (stateCode IS NULL) OR (stateCode<>'" &Session("pcSFStateCode")& "' AND stateCodeEq=0)) "
					query=query& "AND ((CountryCode='"&Session("pcSFCountryCode")&"' AND CountryCodeEq=-1) "
					query=query& "OR (CountryCode IS NULL) OR (CountryCode<>'" &Session("pcSFCountryCode")& "' AND CountryCodeEq=0)) "
					query=query& "AND ((zip='" &Session("pcSFzip")& "' AND zipEq=-1) "
					query=query& "OR (zip IS NULL) OR (zip<>'" &Session("pcSFzip")& "' AND zipEq=0));"
					set rsTaxObj=server.CreateObject("ADODB.RecordSet")
					set rsTaxObj=conntemp.execute(query)			
					taxCnt=0
					do until rsTaxObj.eof 
						if ptaxseparate="1" then
							taxCnt=taxCnt+1
							session("taxDesc"&taxCnt)=rsTaxObj("taxDesc")
							session("tax"&taxCnt)=rsTaxObj("taxLoc")
						end if     
						taxLoc=taxLoc+rsTaxObj("taxLoc") 
					 rsTaxObj.movenext
					loop
					set rsTaxObj = nothing
					if ptaxseparate="1" then
						session("taxCnt")=taxCnt
					end if 
					TAX_SHIPPING_ALONE="NA"
					TAX_SHIPPING_AND_HANDLING_TOGETHER="NA"
					zipCnt=1
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// END: Tax Per Place
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			end if
		
		End if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Tax Rates
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		attrAvailableList = ""
		attrPriceList = ""

		%>
		<!--#include file="pcPay_GoogleCheckout_Shipping.asp" -->			
		<%

		'// Clear Remaining Sessions
		Session("attrAvailableList")=""
		Session("attrPriceList")=""

		If domMethodList.length > 0 Then

            '// Loop for each merchant-calulated shipping method
            For Each domMethod In domMethodList

                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' CUSTOMER INFORMATION CALLBACK
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Retrieve the name of the shipping method
                attrShippingName = domMethod.getAttribute("name")
                Set domResult = domResults.appendChild( domMcResultsObj.createElement("result") )
                domResult.setAttribute "shipping-name", attrShippingName
                domResult.setAttribute "address-id", attrAddressId

                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' DISCOUNT CALLBACK
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                If domMerchantCodeList.length > 0 Then
                    Set domMerchantCodeResults = createMerchantCodeResults(domMcCallbackObj, domMerchantCodeList, attrAddressId, attrShippingName) 
                    Set domMerchantCodeResultsRoot = domMerchantCodeResults.documentElement
                    domResult.appendChild( domMerchantCodeResultsRoot.cloneNode(true) )
                End If

                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' TAX CALLBACK
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			    If calcTax = "true" Then
                    Set domTotalTax = domResult.appendChild( domMcResultsObj.createElement("total-tax") )
                    domTotalTax.setAttribute "currency", attrCurrency
                    totalTax = getTaxRate(domMcCallbackObj, attrAddressId, attrShippingName)
                    domTotalTax.appendChild( domMcResultsObj.createTextNode(totalTax) )
                End If

                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' SHIPPING CALLBACK
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Call the getShippingRate function to calculate the shipping cost.
                Set domShippingRate = domResult.appendChild( domMcResultsObj.createElement("shipping-rate") )
                domShippingRate.setAttribute "currency", attrCurrency                
				shippingRate = getShippingRate(domMcCallbackObj, attrAddressId, attrShippingName)                
				domShippingRate.appendChild( domMcResultsObj.createTextNode(shippingRate) )
                '// Verify that the order can be shipped to the address
				shippable = verifyShippable(domMcCallbackObj, attrAddressId, attrShippingName)
                Set domShippable = domResult.appendChild( domMcResultsObj.createElement("shippable") )
                domShippable.text = shippable
				
            Next

        '// This block executes if no shipping methods are specified
        Else
           
		   '// Create a <result> element in the response with shipping-name and address-id attributes
            Set domResult = domResults.appendChild( domMcResultsObj.createElement("result"))
            domResult.setAttribute "address-id", attrAddressId

            '// If the <tax> tag in the <merchant-calculation-callback> has a value of "true", call the getTaxRate function.
            If calcTax = "true" Then
				%>
				<!--#include file="pcPay_GoogleCheckout_Tax.asp" -->
				<%	
				totalTax = pcf_CurrencyField(pTaxNoShipAmount)
				Set domTotalTax = domResult.appendChild( domMcResultsObj.createElement("total-tax") )
				domTotalTax.setAttribute "currency", attrCurrency
				domTotalTax.appendChild( domMcResultsObj.createTextNode(totalTax) )					
            End If

            '// If there are coupon or gift certificate codes, call the createMerchantCodeResults function to verify those
            ' codes and to create <coupon-result> or <gift-certificate-result> elements to be included in the <merchant-calculation-response>.
            If domMerchantCodeList.length > 0 Then
                Set domMerchantCodeResults = createMerchantCodeResults(domMcCallbackObj, domMerchantCodeList, attrAddressId, attrShippingName) 
                Set domMerchantCodeResultsRoot = domMerchantCodeResults.documentElement
                domResult.appendChild( domMerchantCodeResultsRoot.cloneNode(true))
            End If
        End If
    Next

    '// Return <merchant-calculation-results> XMLDOM
    createMerchantCalculationResults = domMcResults.xml

    Set domMcResultsObj = Nothing
    Set domMcResults = Nothing
    Set domMerchantCodeResults = Nothing
    Set domMerchantCodeResultsRoot = Nothing
    Set domResults = Nothing
    Set domResult = Nothing
    Set domResponse = Nothing
    Set domMcCallbackObjRoot = Nothing
    Set domTotalTax = Nothing
    Set domShippingRate = Nothing
    Set domShippable = Nothing
    Set domTaxList = Nothing
    Set domMethodList = Nothing
    Set domMethod = Nothing
    Set domAnonymousAddressList = Nothing
    Set domAnonymousAddress = Nothing
    Set domMerchantCodeList = Nothing
	
End Function
'***********************************************************************************
' END: RECIEVE THE CALLBACK
'***********************************************************************************



'***********************************************************************************
' START: COUPONS AND GIFT CERTIFICATES AVAILABLE
'***********************************************************************************
Function getMerchantCodeInfo(domMcCallbackObj, elemCode, addressId, shippingMethod)
	on error resume next
	Dim elemCodeType
    Dim elemCodeValid
    Dim elemCalculatedAmount
    Dim elemMessage
	Dim displayDiscountCode	
	'// The Discount Code
	pDiscountCode=elemCode	
	%>
	<!--#include file="pcPay_GoogleCheckout_Discounts.asp" -->			
	<%	
	if discountTotal="0.00" then
		discountTotal=0
	end if	
	if discountTotal>0 then
		elemCodeType = pcv_strCodeType
		elemCodeValid = "true"
		elemCalculatedAmount = discountTotal
		elemMessage = "You have saved"
	else
		elemCodeType = pcv_strCodeType
		elemCodeValid = "false"
		elemCalculatedAmount = discountTotal
		elemMessage = pcv_strDiscountErrorMsg
	end if
	Set getMerchantCodeInfo = createMerchantCodeResult(elemCodeType, elemCodeValid, elemCode, elemCalculatedAmount, elemMessage)

End Function
'***********************************************************************************
' END: COUPONS AND GIFT CERTIFICATES AVAILABLE 
'***********************************************************************************




'***********************************************************************************
' START: MERCHANT SHIPPING OPTIONS AVAILABLE
'***********************************************************************************
Function verifyShippable(domMcCallbackObj, addressId, shippingMethod)
	on error resume next		
	verifyShippable = "false"		
	if Session(shippingMethod)<>"" then
		verifyShippable = "true"
		Session(shippingMethod)=""
	end if	
	if Session("AddressDenied")="1" then
		verifyShippable = "false"
	end if

End Function
'***********************************************************************************
' END: MERCHANT SHIPPING OPTIONS AVAILABLE
'***********************************************************************************





'***********************************************************************************
' START: MERCHANT SHIPPING PRICES
'***********************************************************************************
Function getShippingRate(domMcCallbackObj, addressId, shippingMethod)
	on error resume next	
	if Session(shippingMethod)<>"" then
		if Session(shippingMethod) = NULL then Session(shippingMethod) = "0.00"
		getShippingRate = Session(shippingMethod)
	else
		getShippingRate = "0.00"
	end if	
		
End Function
'***********************************************************************************
' END: MERCHANT SHIPPING PRICES
'***********************************************************************************





'***********************************************************************************
' START: MERCHANT TAX RATE
'***********************************************************************************
Function getTaxRate(domMcCallbackObj, addressId, shippingMethod)
	on error resume next	
	if Session(shippingMethod)<>"" then
		if Session(shippingMethod & "_tax") = "" then Session(shippingMethod & "_tax") = "0.00"
		if Session(shippingMethod & "_tax2") = "" then Session(shippingMethod & "_tax2") = "0.00"
		
		if Session("FreeShippingFlag")="1" then
			getTaxRate = Session(shippingMethod & "_tax2")
		else
			getTaxRate = Session(shippingMethod & "_tax")	
		end if
		
	else
		getTaxRate = "0.00"
	end if   	

End Function
'***********************************************************************************
' END: MERCHANT TAX RATE
'***********************************************************************************



'***********************************************************************************
' START: COUPONS AND GIFT CERTIFICATES
'***********************************************************************************
Function createMerchantCodeResult(elemCodeType, elemCodeValid, elemCode, elemCalculatedAmount, elemMessage)

    '// Define objects used to create the <coupon-result> or <gift-certificate-result>
    Dim domCodeResultObj
    Dim domMerchantCodeResult
    Dim domValid
    Dim domCode
    Dim domMessage
    Dim domCalculatedAmount

    '// Create an empty XMLDOM
    Set domCodeResultObj = Server.CreateObject(strMsxmlDomDocument)
    domCodeResultObj.async = False

    '// Create root tag for <coupon-result> or <gift-certificate-result>
    Set domMerchantCodeResult = domCodeResultObj.appendChild( domCodeResultObj.createElement(elemCodeType & "-result") )

    '// Create <valid> tag, which will indicate whether the code is valid
    Set domValid = domMerchantCodeResult.appendChild( domCodeResultObj.createElement("valid") )
    domValid.text = elemCodeValid
    
    '// Add the coupon or gift certificate code in a <code> tag
    Set domCode = domMerchantCodeResult.appendChild( domCodeResultObj.createElement("code") )
    domCode.text = elemCode

    '// Add the <calculated-amount> tag if there is a value for the elemCalculatedAmount parameter. You could omit this tag if the code is invalid.
    If elemCalculatedAmount <> "" Then
        Set domCalculatedAmount = domMerchantCodeResult.appendChild( domCodeResultObj.createElement("calculated-amount") )
        domCalculatedAmount.setAttribute "currency", attrCurrency
        domCalculatedAmount.text = elemCalculatedAmount
    End If

    '// Add a <message> tag if the $message parameter has a value
    If elemMessage <> "" Then
        Set domMessage = domMerchantCodeResult.appendChild( domCodeResultObj.createElement("message") )
        domMessage.text = elemMessage
    End If

    Set createMerchantCodeResult = domCodeResultObj

    '// Release objects used to create the <coupon-result> or <gift-certificate-result>
    Set domCodeResultObj = Nothing
    Set domMerchantCodeResult = Nothing
    Set domValid = Nothing
    Set domCode = Nothing
    Set domCalculatedAmount = Nothing
    Set domMessage = Nothing

End Function


Function createMerchantCodeResults(domMcCallbackObj, domMerchantCodeList, addressId, shippingMethod)

    '// Define the objects used to create the <coupon-result> or <gift-certificate-result>
    Dim code
    Dim domMcResultsObj
    Dim merchantCode
    Dim codeType
    Dim calculatedAmount
    Dim message
    Dim domMerchantCodeResults
    Dim domMerchantCodeResultObj
    Dim domMerchantCodeResultRoot

    '// Create an empty XMLDOM
    Set domMcResultsObj = Server.CreateObject(strMsxmlDomDocument)
    domMcResultsObj.async = False
    Set domMerchantCodeResults = domMcResultsObj.appendChild( domMcResultsObj.createElement("merchant-code-results") )
    
    For Each merchantCode In domMerchantCodeList
        code = merchantCode.getAttribute("code")
        Set domMerchantCodeResultObj = getMerchantCodeInfo(domMcCallbackObj, code, addressId, shippingMethod)
        Set domMerchantCodeResultRoot = domMerchantCodeResultObj.documentElement
        domMerchantCodeResults.appendChild( domMerchantCodeResultRoot.cloneNode(true) )        
    Next

    Set createMerchantCodeResults = domMcResultsObj    

    '// Release the objects used to create the <coupon-result> or <gift-certificate-result>
    Set domMcResultsObj = Nothing
    Set domMerchantCodeResultObj = Nothing
    Set domMerchantCodeResultRoot = Nothing

End Function
'***********************************************************************************
' END: COUPONS AND GIFT CERTIFICATES
'***********************************************************************************
%>


