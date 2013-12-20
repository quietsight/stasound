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
Dim domItemsObj
Dim domShoppingCartObj
Dim domDefaultTaxRulesObj
Dim domAltTaxRulesObj
Dim domAltTaxTablesObj
Dim domTaxTablesObj
Dim domShippingRestrictionsObj
Dim domShippingMethodsObj
Dim domMerchantCalculationsObj 
Dim domRoundingPolicyObj
Dim domMerchantCFSObj
Dim domCFSObj 
Dim domCheckoutShoppingCartObj


'***********************************************************************************
' START: CREATES XML SINGLE CART LINE
'***********************************************************************************
Function createItem(elemItemName, elemItemDescription, elemQuantity, elemUnitPrice, elemTaxTableSelector, elemMerchantPrivateItemData)

    Dim strFunctionName
    Dim errorType
    strFunctionName = "createItem()"

    '// Each of these parameters must have a value to create an <item>
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemItemName", elemItemName
    checkForError errorType, strFunctionName, "elemItemDescription", _
        elemItemDescription
    checkForError errorType, strFunctionName, "elemQuantity", elemQuantity
    checkForError errorType, strFunctionName, "elemUnitPrice", elemUnitPrice
    checkForError errorType, strFunctionName, "attrCurrency", attrCurrency

    '// HTML entities need to be escaped properly
    elemItemName = Server.HTMLEncode(elemItemName)
    elemItemDescription = Server.HTMLEncode(elemItemDescription)

    '// Define objects used to create the item
    Dim domItemObj
    Dim domItem
    Dim domItemName
    Dim domItemDescription
    Dim domQuantity
    Dim domUnitPrice
    Dim domTaxTableSelector
    Dim domMerchantPrivateItemDataObj
    Dim domNewMerchantPrivateItemData
    Dim domMerchantPrivateItemDataRoot
    Dim domItemsRoot
    Dim domItemRoot

    '// Create the <items> tag if this is the first item to be created
    If Not(IsObject(domItemsObj)) Then
        Set domItemsObj = Server.CreateObject(strMsxmlDomDocument)
        domItemsObj.async = False
        domItemsObj.appendChild(domItemsObj.createElement("items"))
    End If

    Set domItemObj = Server.CreateObject(strMsxmlDomDocument)
    domItemObj.async = False

    '// Create the <item> tag for the item to be created
    Set domItem = domItemObj.appendChild(domItemObj.createElement("item"))

    ' Add the item name to the XML
    Set domItemName = domItem.appendChild(domItemObj.createElement("item-name"))
    domItemName.Text = elemItemName
    
    '// Add the item description to the XML
    Set domItemDescription = _
        domItem.appendChild(domItemObj.createElement("item-description"))
    domItemDescription.Text = elemItemDescription

    '// Add the quantity to the XML
    Set domQuantity = _
        domItem.appendChild(domItemObj.createElement("quantity"))
    domQuantity.Text = elemQuantity

    '// Add the unit price for the item to the XML
    Set domUnitPrice = _
        domItem.appendChild(domItemObj.createElement("unit-price"))
    domUnitPrice.setAttribute "currency", attrCurrency
    domUnitPrice.Text = elemUnitPrice

    '// If there is an alternate-tax-table associated with this item, specify the table's name using the <tax-table-selector> tag.
    If elemTaxTableSelector <> "" Then
        Set domTaxTableSelector = _
            domItem.appendChild(domItemObj.createElement("tax-table-selector"))
        domTaxTableSelector.Text = elemTaxTableSelector
    End If

    '// If you have provided a value for the elemMerchantPrivateItemData variable, that value will be printed inside the <merchant-private-item-data> tag.
    If elemMerchantPrivateItemData <> "" Then

        Set domMerchantPrivateItemDataObj = _
            Server.CreateObject(strMsxmlDomDocument)
        domMerchantPrivateItemDataObj.async = False
        domMerchantPrivateItemDataObj.loadXml elemMerchantPrivateItemData

        Set domNewMerchantPrivateItemData = domItem.appendChild( _
            domItemObj.createElement("merchant-private-item-data"))

        Set domMerchantPrivateItemDataRoot = _
            domMerchantPrivateItemDataObj.documentElement

        domNewMerchantPrivateItemData.appendChild( _
            domMerchantPrivateItemDataRoot.cloneNode(True))

    End If

    '// The newly created item is added as a child of the <items> tag.
    Set domItemsRoot = domItemsObj.documentElement
    Set domItemRoot = domItemObj.documentElement
    domItemsRoot.appendChild domItemRoot.cloneNode(True)

   Set createItem = domItemObj

    '// Release objects used to create item
    Set domItemObj = Nothing
    Set domItem = Nothing
    Set domItemName = Nothing
    Set domItemDescription = Nothing
    Set domQuantity = Nothing
    Set domUnitPrice = Nothing
    Set domTaxTableSelector = Nothing
    Set domMerchantPrivateItemDataObj = Nothing
    Set domNewMerchantPrivateItemData = Nothing
    Set domMerchantPrivateItemDataRoot = Nothing
    Set domItemsRoot = Nothing
    Set domItemRoot = Nothing

End Function
'***********************************************************************************
' END: CREATES XML SINGLE CART LINE
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML CART ELEMENT
'***********************************************************************************
Function createShoppingCart(dtmCartExpiration, elemMerchantPrivateData)

    Dim strFunctionName
    Dim errorType

    strFunctionName = "createShoppingCart()"
    
    '// There must be at least one item in the shopping cart by the time you call this function or the function will log an error.
    errorType = "MISSING_PARAM"
    If Not(IsObject(domItemsObj)) Then
        errorHandler errorType, strFunctionName, "domItems"
    End If 

    '// Define objects used to create the shopping cart
    Dim domShoppingCart
    Dim domItemsRoot
    Dim domCartExpiration
    Dim domGoodUntilDate
    Dim domNewMerchantPrivateData
    Dim domMerchantPrivateDataObj
    Dim domMerchantPrivateDataRoot

    Set domShoppingCartObj = Server.CreateObject(strMsxmlDomDocument)
    domShoppingCartObj.async = False

    '// Create the <shopping-cart> element
    Set domShoppingCart = domShoppingCartObj.appendChild( _
        domShoppingCartObj.createElement("shopping-cart"))
    Set domItemsRoot = domItemsObj.documentElement
    domShoppingCart.appendChild(domItemsRoot.cloneNode(true))
    
    '// If there is an expiration date ($cart_expiration) for the cart, include it in the <shopping-cart> XML.
    If dtmCartExpiration <> "" Then
        Set domCartExpiration = domShoppingCart.appendChild( _
            domShoppingCartObj.createElement("cart-expiration"))
        Set domGoodUntilDate = domCartExpiration.appendChild( _
            domShoppingCartObj.createElement("good-until-date"))
        domGoodUntilDate.Text = dtmCartExpiration
    End If

    '// If you have provided a value for the $merchant_private_data variable, that value will be printed inside the <merchant-private-data> tag.
    If elemMerchantPrivateData <> "" Then

        Set domMerchantPrivateDataObj = Server.CreateObject(strMsxmlDomDocument)
        domMerchantPrivateDataObj.async = False
        domMerchantPrivateDataObj.loadXml elemMerchantPrivateData

        Set domNewMerchantPrivateData = domShoppingCart.appendChild( _
            domShoppingCartObj.createElement("merchant-private-data"))

        Set domMerchantPrivateDataRoot = _
            domMerchantPrivateDataObj.documentElement

        domNewMerchantPrivateData.appendChild( _
            domMerchantPrivateDataRoot.cloneNode(true))
    End If

    Set createShoppingCart = domShoppingCartObj

    '// Release objects used to create shipping cart
    Set domShoppingCart = Nothing
    Set domItemsObj = Nothing
    Set domItemsRoot = Nothing
    Set domCartExpiration = Nothing
    Set domGoodUntilDate = Nothing
    Set domNewMerchantPrivateData = Nothing
    Set domMerchantPrivateDataObj = Nothing
    Set domMerchantPrivateDataRoot = Nothing

End Function
'***********************************************************************************
' END: CREATES XML CART ELEMENT
'***********************************************************************************



'***********************************************************************************
' START: CREATES XML COUNTRY LIST
'***********************************************************************************
Function createUsCountryArea(areaPlace)
    Set createUsCountryArea = createUsPlaceArea("country", areaPlace)
End Function
'***********************************************************************************
' END: CREATES XML COUNTRY LIST
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML POSTAL AREA
'***********************************************************************************
Function createPostalArea(areaPlace)
    Set createPostalArea = createCountryArea("country-code", areaPlace)
End Function
'***********************************************************************************
' END: CREATES XML POSTAL AREA
'***********************************************************************************



'***********************************************************************************
' START: CREATES XML STATE AREA
'***********************************************************************************
Function createUsStateArea(areaPlace)
    Set createUsStateArea = createUsPlaceArea("state", areaPlace)
End Function
'***********************************************************************************
' END: CREATES XML STATE AREA
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML POSTALCODE AREA
'***********************************************************************************
Function createUsZipArea(areaPlace)
    Set createUsZipArea = createUsPlaceArea("zip", areaPlace)
End Function
'***********************************************************************************
' END: CREATES XML POSTALCODE AREA
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML US PLACE AREA
'***********************************************************************************
Function createUsPlaceArea(areaType, areaPlace)
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createUsPlaceArea()"
    
    '// Both parameters must be specified for the function call to execute.
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "areaType", areaType
    checkForError errorType, strFunctionName, "areaPlace", areaPlace

    '// Define objects used to create the XML block
    Dim domAreaObj
    Dim domArea
    Dim domAreaPlace

    '// Create the parent XML element for the areaType
    Set domAreaObj = Server.CreateObject(strMsxmlDomDocument)
    domAreaObj.async = False
    Set domArea = domAreaObj.appendChild( _
        domAreaObj.createElement("us-" & areaType & "-area"))

    '// Create the element that contains the areaPlace data
    If areaType = "state" Then

        Set domAreaPlace = _
            domArea.appendChild(domAreaObj.createElement("state"))
        domAreaPlace.Text = areaPlace

    ElseIf areaType = "zip" Then

        Set domAreaPlace = _
            domArea.appendChild(domAreaObj.createElement("zip-pattern"))
        domAreaPlace.Text = areaPlace

    ElseIf areaType = "country" Then

        domArea.setAttribute "country-area", areaPlace

    End If

    Set createUsPlaceArea =  domAreaObj

    '// Release objects used to create the XML block
    Set domAreaObj = Nothing
    Set domArea = Nothing
    Set domAreaPlace = Nothing

End Function
'***********************************************************************************
' END: CREATES XML US PLACE AREA
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML COUNTRY AREA
'***********************************************************************************
Function createCountryArea(areaType, areaPlace)
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createCountryArea()"
    
    '// Both parameters must be specified for the function call to execute.
    'errorType = "MISSING_PARAM"
    'checkForError errorType, strFunctionName, "areaType", areaType
    'checkForError errorType, strFunctionName, "areaPlace", areaPlace

    '// Define objects used to create the XML block
    Dim domAreaObj
    Dim domArea
    Dim domAreaPlace

    '// Create the parent XML element for the areaType
    Set domAreaObj = Server.CreateObject(strMsxmlDomDocument)
    domAreaObj.async = False
    Set domArea = domAreaObj.appendChild(domAreaObj.createElement("postal-area"))

    '// Create the element that contains the areaPlace data
    If areaType = "country-code" Then

        Set domAreaPlace = domArea.appendChild(domAreaObj.createElement("country-code"))
        domAreaPlace.Text = areaPlace

    End If

    Set createCountryArea =  domAreaObj

    '// Release objects used to create the XML block
    Set domAreaObj = Nothing
    Set domArea = Nothing
    Set domAreaPlace = Nothing

End Function
'***********************************************************************************
' END: CREATES XML COUNTRY AREA
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML TAX AREA
'***********************************************************************************
Function createTaxArea(taxAreaType, taxAreaPlace) 
    
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createTaxArea()"
    
    '// Both parameters must be specified for the function call to execute.
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "taxAreaType", taxAreaType
    checkForError errorType, strFunctionName, "taxAreaPlace", taxAreaPlace

    '// Define the objects used to create the tax area
    Dim domTaxAreaObj
    Dim domTaxArea
    Dim domArea
    Dim domAreaRoot

    '// Create the <tax-area> element
    Set domTaxAreaObj = Server.CreateObject(strMsxmlDomDocument)
    domTaxAreaObj.async = False
    Set domTaxArea = domTaxAreaObj.appendChild(domTaxAreaObj.createElement("tax-area"))
	
   if GOOGLECURRENCY="GBP" then
  	 	
		'// Create a list of EU Memeber States - Add the <postal-area> element	
		Set domPostalArea = domTaxArea.appendChild( domTaxAreaObj.createElement("postal-area"))
		domPostalArea.Text = ""
		
		'// Create the <merchant-calculations-url> element
		Set domPostalAreaCountryCode = domPostalArea.appendChild( domTaxAreaObj.createElement("country-code"))
		domPostalAreaCountryCode.Text = taxAreaPlace	
		
   else
  	 	'// Add the <world-area> element	
		Set domWorldArea = domTaxArea.appendChild( domTaxAreaObj.createElement("world-area"))
		domWorldArea.Text = ""
	end if  
   
   
    Set createTaxArea = domTaxAreaObj

    '// Release the objects used to create the tax area
    Set domTaxAreaObj = Nothing
    Set domTaxArea = Nothing
    Set domArea = Nothing
    Set domAreaRoot = Nothing

End Function
'***********************************************************************************
' END: CREATES XML TAX AREA
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML TAX RULE
'***********************************************************************************
Function createDefaultTaxRule(elemRate, domTaxArea, elemShippingTaxed) 

    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createDefaultTaxRule()"
    
    '// Check for missing parameters.  You must specify a rate and provide a domTaxArea object for each rule
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemRate", elemRate

    If Not(IsObject(domTaxArea)) Then
        errorHandler errorType, strFunctionName, "domTaxArea"
    End If
    
    '// Define the objects used to create the <default-tax-rule>
    Dim domDefaultTaxRuleObj
    Dim domDefaultTaxRule
    Dim domShippingTaxed
    Dim domRate
    Dim domTaxAreaRoot
    Dim domDefaultTaxRulesRoot
    Dim domDefaultTaxRuleRoot

    '// Create the <default-tax-rule> element
    Set domDefaultTaxRuleObj = Server.CreateObject(strMsxmlDomDocument)
    domDefaultTaxRuleObj.async = False

    Set domDefaultTaxRule = domDefaultTaxRuleObj.appendChild( _
        domDefaultTaxRuleObj.createElement("default-tax-rule"))

    '// Add a <shipping-taxed> element if a elemShippingTaxed value is provided
    Set domShippingTaxed = domDefaultTaxRule.appendChild( _
        domDefaultTaxRuleObj.createElement("shipping-taxed"))

    domShippingTaxed.appendChild( _
        domDefaultTaxRuleObj.createTextNode(elemShippingTaxed))

    '// Add the tax rate for the tax rule
    Set domRate = domDefaultTaxRule.appendChild( _
        domDefaultTaxRuleObj.createElement("rate"))
    domRate.appendChild(domDefaultTaxRuleObj.createTextNode(elemRate))

    Set domTaxAreaRoot = domTaxArea.documentElement
    domDefaultTaxRule.appendChild(domTaxAreaRoot.cloneNode(true))

    '// Create a <tax-rules> element if no other <default-tax-rule>
    ' elements have been created. Append the rule to a list that
    ' will appear under the <tax-rules> element within the <default-tax-table> element
    If Not(IsObject(domDefaultTaxRulesObj)) Then
        Set domDefaultTaxRulesObj = Server.CreateObject(strMsxmlDomDocument)
        domDefaultTaxRulesObj.async = False
        domDefaultTaxRulesObj.appendChild( _
            domDefaultTaxRulesObj.createElement("tax-rules"))
    End If

    '// Add the tax rules to the XML
    Set domDefaultTaxRulesRoot = domDefaultTaxRulesObj.documentElement
    Set domDefaultTaxRuleRoot = domDefaultTaxRuleObj.documentElement
    domDefaultTaxRulesRoot.appendChild(domDefaultTaxRuleRoot.cloneNode(true))

    Set createDefaultTaxRule = domDefaultTaxRuleObj

    '// Release the objects used to create the <default-tax-rule>
    Set domDefaultTaxRuleObj = Nothing
    Set domDefaultTaxRule = Nothing
    Set domShippingTaxed = Nothing
    Set domRate = Nothing
    Set domTaxAreaRoot = Nothing
    Set domDefaultTaxRulesRoot = Nothing
    Set domDefaultTaxRuleRoot = Nothing

End Function
'***********************************************************************************
' END: CREATES XML TAX RULE
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML ALTERNATE TAX RULE
'***********************************************************************************
Function createAlternateTaxRule(elemRate, domTaxArea) 
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createAlternateTaxRule()"
    
    '// You must specify an elemRate and domTaxArea object for each tax rule
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemRate", elemRate

    If Not(IsObject(domTaxArea)) Then
        errorHandler errorType, strFunctionName, "domTaxArea"
    End If

    '// Define the objects used to create the <alternate-tax-rule>
    Dim domAltTaxRuleObj
    Dim domAltTaxRule
    Dim domRate
    Dim domTaxAreaRoot
    Dim domAltTaxRulesRoot
    Dim domAltTaxRuleRoot

    '// Create the <alternate-tax-rule> element
    Set domAltTaxRuleObj = Server.CreateObject(strMsxmlDomDocument)
    domAltTaxRuleObj.async = False
    Set domAltTaxRule = domAltTaxRuleObj.appendChild( _
        domAltTaxRuleObj.createElement("alternate-tax-rule"))

    '// Add the tax rate for the tax rule
    Set domRate = _
        domAltTaxRule.appendChild(domAltTaxRuleObj.createElement("rate"))
    domRate.appendChild(domAltTaxRuleObj.createTextNode(elemRate))

    Set domTaxAreaRoot = domTaxArea.documentElement
    domAltTaxRule.appendChild(domTaxAreaRoot.cloneNode(true))

    '// Create an <alternate-tax-rules> element if this is the first
    ' <alternate-tax-rule> to be created. Append the rule to a list
    ' that will appear under the <alternate-tax-rules> element within an <alternate-tax-table> element
    If Not(IsObject(domAltTaxRulesObj)) Then
        Set domAltTaxRulesObj = Server.CreateObject(strMsxmlDomDocument)
        domAltTaxRulesObj.async = False
        domAltTaxRulesObj.appendChild( _
            domAltTaxRulesObj.createElement("alternate-tax-rules"))
    End If

    '// Add the alternate tax rules to the XML
    Set domAltTaxRulesRoot = domAltTaxRulesObj.documentElement
    Set domAltTaxRuleRoot = domAltTaxRuleObj.documentElement
    domAltTaxRulesRoot.appendChild(domAltTaxRuleRoot.cloneNode(true))

    Set createAlternateTaxRule = domAltTaxRuleObj

    '// Release the objects used to create the <alternate-tax-rule>
    Set domAltTaxRuleObj = Nothing
    Set domAltTaxRule = Nothing
    Set domRate = Nothing
    Set domTaxAreaRoot = Nothing
    Set domAltTaxRulesRoot = Nothing
    Set domAltTaxRuleRoot = Nothing

End Function
'***********************************************************************************
' END: CREATES XML ALTERNATE TAX RULE
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML ALTERNATE TAX TABLE
'***********************************************************************************
Function createAlternateTaxTable(attrStandalone, attrName) 
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createAlternateTaxTable()"
    
    '// You must specify values for the attrStandalone and attrName parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrStandalone", attrStandalone
    checkForError errorType, strFunctionName, "attrName", attrName

    '// There must be at least one alternate tax rule to include
    ' in the <alternate-tax-table>. This tax table will include
    ' any <alternate-tax-rule> elements that were created since
    ' after the last call to the createAlternateTaxTable function.
    If Not(IsObject(domAltTaxRulesObj)) Then
        errorHandler errorType, strFunctionName, "domAlternateTaxRules"
    End If
    
    '// Define the objects used to create the <alternate-tax-rule>
    Dim domAltTaxTableObj
    Dim domAltTaxTable
    Dim domAltTaxRulesRoot
    Dim domAltTaxTablesRoot
    Dim domAltTaxTableRoot

    Set domAltTaxTableObj = Server.CreateObject(strMsxmlDomDocument)
    domAltTaxTableObj.async = False

    '// Create the <alternate-tax-table> element
    Set domAltTaxTable = _
        domAltTaxTableObj.appendChild(domAltTaxTableObj.createElement("alternate-tax-table"))
    domAltTaxTable.setAttribute "standalone", attrStandalone
    domAltTaxTable.setAttribute "name", attrName

    '// Add the <alternate-tax-rules> element as a child element of <alternate-tax-table> elements
    Set domAltTaxRulesRoot = domAltTaxRulesObj.documentElement
    domAltTaxTable.appendChild(domAltTaxRulesRoot.cloneNode(true))

    '// Create an <alternate-tax-tables> element, if one has not yet been created, to contain all <alternate-tax-table> elements
    If Not(IsObject(domAltTaxTablesObj)) Then
        Set domAltTaxTablesObj = Server.CreateObject(strMsxmlDomDocument)
        domAltTaxTablesObj.async = False
        domAltTaxTablesObj.appendChild(domAltTaxTablesObj.createElement("alternate-tax-tables"))
    End If

    '// Add the <alternate-tax-table> element as a child of the <alternate-tax-tables> element
    Set domAltTaxTablesRoot = domAltTaxTablesObj.documentElement
    Set domAltTaxTableRoot = domAltTaxTableObj.documentElement
    domAltTaxTablesRoot.appendChild(domAltTaxTableRoot.cloneNode(true))

    Set domAltTaxRulesObj = Server.CreateObject(strMsxmlDomDocument)
    domAltTaxRulesObj.async = False
    domAltTaxRulesObj.appendChild(domAltTaxRulesObj.createElement("alternate-tax-rules"))

    Set createAlternateTaxTable = domAltTaxTableObj

    '// Release the objects used to create the <alternate-tax-rule>
    Set domAltTaxTableObj = Nothing
    Set domAltTaxTable = Nothing
    Set domAltTaxRulesRoot = Nothing
    Set domAltTaxTablesRoot = Nothing
    Set domAltTaxTableRoot = Nothing

End Function
'***********************************************************************************
' END: CREATES XML ALTERNATE TAX TABLE
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML TAX TABLE ELEMENT
'***********************************************************************************
Function createTaxTables(attrMerchantCalculated) 
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createTaxTables()"
    
    '// Check for missing parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrMerchantCalculated", _
        attrMerchantCalculated

    '// Define the objects used to create the <tax-tables>
    Dim domTaxTables
    Dim domDefaultTaxRulesRoot
    Dim domDefaultTaxTableObj
    Dim domDefaultTaxTable
    Dim domDefaultTaxTableRoot
    Dim domAltTaxTablesRoot

    '// Create the <tax-tables> element
    Set domTaxTablesObj = Server.CreateObject(strMsxmlDomDocument)
    domTaxTablesObj.async = False
    Set domTaxTables = _
        domTaxTablesObj.appendChild(domTaxTablesObj.createElement("tax-tables"))

    '// Set the "merchant-calculated" attribute on the <tax-tables> element
    If attrMerchantCalculated <> "" Then
        domTaxTables.setAttribute "merchant-calculated", attrMerchantCalculated
    End If

    '// Create a <default-tax-table> element and append the default tax rules
    Set domDefaultTaxTableObj = Server.CreateObject(strMsxmlDomDocument)
    domDefaultTaxTableObj.async = False
    Set domDefaultTaxTable = domDefaultTaxTableObj.appendChild( _
        domDefaultTaxTableObj.createElement("default-tax-table"))
    Set domDefaultTaxRulesRoot = domDefaultTaxRulesObj.documentElement
    domDefaultTaxTable.appendChild(domDefaultTaxRulesRoot.cloneNode(true))

    '// Make the <default-tax-table> element a child of <tax-tables> element
    Set domDefaultTaxTableRoot = domDefaultTaxTableObj.documentElement
    domTaxTables.appendChild(domDefaultTaxTableRoot.cloneNode(true))

    '// Add the <alternate-tax-tables> elements as children of <tax-tables>
    If IsObject(domAltTaxTablesObj) Then
        Set domAltTaxTablesRoot = domAltTaxTablesObj.documentElement
        domTaxTables.appendChild(domAltTaxTablesRoot.cloneNode(true))
    End If

    Set createTaxTables = domTaxTablesObj

    '// Release the objects used to create the <tax-tables>
    Set domTaxTables = Nothing
    Set domDefaultTaxRulesObj = Nothing
    Set domDefaultTaxRulesRoot = Nothing
    Set domDefaultTaxTableObj = Nothing
    Set domDefaultTaxTable = Nothing
    Set domDefaultTaxTableRoot = Nothing
    Set domAltTaxRulesObj = Nothing
    Set domAltTaxTablesObj = Nothing
    Set domAltTaxTablesRoot = Nothing

End Function
'***********************************************************************************
' END: CREATES XML TAX TABLE ELEMENT
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML ALLOWED AREAS
'***********************************************************************************
Function addAllowedAreas(attrAllowedCountry, arrayAllowedState, arrayAllowedZip, arrayAllowedCountries, arrayType)

    Set addAllowedAreas =  addAreas(attrAllowedCountry, arrayAllowedState, arrayAllowedZip, arrayAllowedCountries, arrayType)

End Function
'***********************************************************************************
' END: CREATES XML ALLOWED AREAS
'***********************************************************************************




'***********************************************************************************
' START: CREATES XML EXCLUDED AREAS
'***********************************************************************************
Function addExcludedAreas(attrExcludedCountry, arrayExcludedState,  arrayExcludedZip)

    Set addExcludedAreas =  addAreas(attrExcludedCountry, arrayExcludedState, arrayExcludedZip, arrayAllowedCountries,  "excluded")

End Function
'***********************************************************************************
' END: CREATES XML EXCLUDED AREAS
'***********************************************************************************




'***********************************************************************************
' START: ADD AREAS
'***********************************************************************************
Function addAreas(attrCountry, arrayState, arrayZip, arrayCountryCode, allowedOrExcluded)
    
    Dim strFunctionName
    Dim errorType

    strFunctionName = "addAllowedAreas()"

	'//////////////////////////////////////////////////////////////////////////////////////
    '// Verify that arrayState and arrayZip parameters are actually arrays
	'//////////////////////////////////////////////////////////////////////////////////////
    errorType = "INVALID_INPUT_ARRAY"
    If Not(IsArray(arrayAllowedState)) Then
        errorHandler errorType, strFunctionName, "arrayState"
    End If

    If Not(IsArray(arrayAllowedZip)) Then
        errorHandler errorType, strFunctionName, "arrayZip"
    End If
	
    If Not(IsArray(arrayAllowedCountries)) Then
        errorHandler errorType, strFunctionName, "arrayCountryCode"
    End If
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
    '// Verify that at least one region has been specified
	'//////////////////////////////////////////////////////////////////////////////////////
    errorType = "MISSING_PARAM_NONE"
    If attrAllowedCountryArea = "" _
        And UBound(arrayAllowedState) < 0 _
        And UBound(arrayAllowedZip) < 0 _
		And UBound(arrayAllowedCountries) < 0 _
    Then
        errorHandler errorType, strFunctionName, "attrCountry"
    End If
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
    '// Define objects used to create the <allowed-areas> or <excluded-areas> element
	'//////////////////////////////////////////////////////////////////////////////////////
    Dim domAreasObj
    Dim domAreas
    Dim domAreasRoot
    Dim domCountry
    Dim domCountryRoot
	Dim domCountryCode
	Dim domCountryCodeRoot
    Dim domState
    Dim domStateRoot
    Dim domZip
    Dim domZipRoot
    Dim domShippingRestrictionsRoot     
    Dim iUboundState
    Dim iUboundZip
    Dim iState
    Dim iZip
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
    '// Create the <allowed-areas> or <excluded-areas> element
	'//////////////////////////////////////////////////////////////////////////////////////
    Set domAreasObj = Server.CreateObject(strMsxmlDomDocument)
    domAreasObj.async = False
    Set domAreas = domAreasObj.appendChild( domAreasObj.createElement(allowedOrExcluded & "-areas"))
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
	'// Add a <postal-area> element for each item in the arrayState array
	'//////////////////////////////////////////////////////////////////////////////////////
	For iCountryCode = 0 To UBound(arrayCountryCode)
		If arrayCountryCode(iCountryCode) <> "" Then
			Set domCountryCode = createPostalArea(arrayCountryCode(iCountryCode))
			Set domCountryCodeRoot = domCountryCode.documentElement
			domAreas.appendChild(domCountryCodeRoot.cloneNode(true))
			pcv_LocalUse=1
		End If
	Next
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
	'// Add a <us-state-area> element for each item in the arrayState array
	'//////////////////////////////////////////////////////////////////////////////////////
	For iState = 0 To UBound(arrayState)
		If arrayState(iState) <> "" Then
			Set domState = createUsStateArea(arrayState(iState))
			Set domStateRoot = domState.documentElement
			domAreas.appendChild(domStateRoot.cloneNode(true))
			pcv_LocalUse=1
		End If
	Next
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
    '// Add a <us-zip-area> element for each item in the arrayZip array
	'//////////////////////////////////////////////////////////////////////////////////////
    For iZip = 0 To UBound(arrayZip)
        If arrayZip(iZip) <> "" Then
            Set domZip = createUsZipArea(arrayZip(iZip))
            Set domZipRoot = domZip.documentElement
            domAreas.appendChild(domZipRoot.cloneNode(true))
			pcv_LocalUse=1
        End If
    Next
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
	' Create Global elements
	'//////////////////////////////////////////////////////////////////////////////////////
	If pcv_LocalUse<>1 Then
		'// Add the <world-area> element
		If attrCountry <> "" AND attrCountry="AUTO" Then  			
			Set domWorldArea = domAreas.appendChild( domAreasObj.createElement("world-area"))
			domWorldArea.Text = ""		 
		'// Add the <us-country-area> element if an attrCountry is provided
		ElseIf attrCountry <> "" AND attrCountry<>"AUTO" Then  		
			Set domCountry = createUsCountryArea(attrCountry)
			Set domCountryRoot = domCountry.documentElement
			domAreas.appendChild(domCountryRoot.cloneNode(true))		
		End If
	End If
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
    '// Create a <address-filters> parent element if one has not already been created
	'//////////////////////////////////////////////////////////////////////////////////////
    If Not(IsObject(domShippingRestrictionsObj)) Then
        Set domShippingRestrictionsObj = Server.CreateObject(strMsxmlDomDocument)
        domShippingRestrictionsObj.async = False
        domShippingRestrictionsObj.appendChild( domShippingRestrictionsObj.createElement("address-filters"))
    End If
	'//////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////
    '// Add the shipping restrictions to the XML
	'//////////////////////////////////////////////////////////////////////////////////////
    Set domShippingRestrictionsRoot = domShippingRestrictionsObj.documentElement
    Set domAreasRoot = domAreasObj.documentElement
    domShippingRestrictionsRoot.appendChild(domAreasRoot.cloneNode(true))    
	'//////////////////////////////////////////////////////////////////////////////////////
	
    Set addAreas = domShippingRestrictionsObj

    '// Release objects used to create the <allowed-areas> or <excluded-areas> element
    Set domAreasObj = Nothing
    Set domAreas = Nothing
    Set domAreasRoot = Nothing
    Set domCountry = Nothing
    Set domCountryRoot = Nothing
    Set domState = Nothing
    Set domStateRoot = Nothing
    Set domZip = Nothing
    Set domZipRoot = Nothing
    Set domShippingRestrictionsRoot = Nothing

End Function
'***********************************************************************************
' END: ADD AREAS
'***********************************************************************************




'***********************************************************************************
' START: CREATE FLAT RATE
'***********************************************************************************
Function createFlatRateShipping(attrName, elemPrice, domShippingRestrictionsObj)

    Set createFlatRateShipping = createShipping("flat-rate-shipping", attrName, elemPrice, domShippingRestrictionsObj)

End Function
'***********************************************************************************
' END: CREATE FLAT RATE
'***********************************************************************************




'***********************************************************************************
' START: MERCHANT CALCULATED SHIPPING
'***********************************************************************************
Function createMerchantCalculatedShipping(attrName, elemPrice, domShippingRestrictionsObj)

   Set createMerchantCalculatedShipping = createShipping("merchant-calculated-shipping", attrName, elemPrice, domShippingRestrictionsObj)

End Function
'***********************************************************************************
' END: MERCHANT CALCULATED SHIPPING
'***********************************************************************************




'***********************************************************************************
' START: PICKUP
'***********************************************************************************
Function createPickup(attrName, elemPrice)

   Set createPickup = createShipping("pickup", attrName, elemPrice, "")

End Function
'***********************************************************************************
' END: PICKUP
'***********************************************************************************




'***********************************************************************************
' START: CREATE SHIPPING
'***********************************************************************************
Function createShipping(shippingType, attrName, elemPrice, domShippingRestrictionsObj)
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createFlatRateShipping()"

    '// Verify that there are values for all required parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrName", attrName
    checkForError errorType, strFunctionName, "elemPrice", elemPrice
    checkForError errorType, strFunctionName, "attrCurrency", attrCurrency
    
    '// Define the variables used to create the shipping information
    Dim domShippingObj
    Dim domShipping
    Dim domShippingRoot
    Dim domPrice
    Dim domShippingRestrictionsRoot
    Dim domShippingMethodsRoot

    '// Create a new parent element using the shippingType as the element name
    Set domShippingObj = Server.CreateObject(strMsxmlDomDocument)
    domShippingObj.async = False
    Set domShipping = _
        domShippingObj.appendChild(domShippingObj.createElement(shippingType))

    '// Set the name and price for the shipping option
    domShipping.setAttribute "name", attrName
    Set domPrice = _
        domShipping.appendChild(domShippingObj.createElement("price"))
    domPrice.setAttribute "currency", attrCurrency
    domPrice.Text = elemPrice

    '// Add address-filters for <flat-rate-shipping> and <merchant-calculated-shipping>
    If (shippingType = "flat-rate-shipping" _
        Or shippingType = "merchant-calculated-shipping") _
        And IsObject(domShippingRestrictionsObj) _
    Then
        Set domShippingRestrictionsRoot = _
            domShippingRestrictionsObj.documentElement
        domShipping.appendChild(domShippingRestrictionsRoot.cloneNode(true))
    End If

    '// Create a <shipping-methods> element if one has not already been created
    If Not(IsObject(domShippingMethodsObj)) Then
        Set domShippingMethodsObj = Server.CreateObject(strMsxmlDomDocument)
        domShippingMethodsObj.async = False
        domShippingMethodsObj.appendChild( _
            domShippingMethodsObj.createElement("shipping-methods"))
    End If

    '// Add the shipping method to the XML request
    Set domShippingMethodsRoot = domShippingMethodsObj.documentElement
    Set domShippingRoot = domShippingObj.documentElement
    domShippingMethodsRoot.appendChild(domShippingRoot.cloneNode(true))

    Set createShipping = domShippingObj

    '// Release the variables used to create the shipping information
    Set domShippingObj = Nothing
    Set domShipping = Nothing
    Set domShippingRoot = Nothing
    Set domPrice = Nothing
    Set domShippingRestrictionsRoot = Nothing
    Set domShippingMethodsRoot = Nothing

End Function
'***********************************************************************************
' END: CREATE SHIPPING
'***********************************************************************************



'***********************************************************************************
' START: CREATE MERHCANT CALCULATIONS FUNCTION
'***********************************************************************************
Function createMerchantCalculations(elemMerchantCalculationsUrl, elemAcceptMerchantCoupons, elemAcceptGiftCertificates)
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createMerchantCalculations()"

    '// Verify that the elemMerchantCalculationsUrl parameter has a value
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemMerchantCalculationsUrl", _
        elemMerchantCalculationsUrl

    '// Define the variables used to create the <merchant-calculations> element
    Dim domMerchantCalculations
    Dim domMerchantCalculationsUrl
    Dim domAcceptMerchantCoupons
    Dim domAcceptGiftCertificates

    '// Create the <merchant-calculations> element
    Set domMerchantCalculationsObj = Server.CreateObject(strMsxmlDomDocument)
    domMerchantCalculationsObj.async = False
    Set domMerchantCalculations = domMerchantCalculationsObj.appendChild( _
        domMerchantCalculationsObj.createElement("merchant-calculations"))

    '// Create the <merchant-calculations-url> element
    Set domMerchantCalculationsUrl = domMerchantCalculations.appendChild( _
        domMerchantCalculationsObj.createElement("merchant-calculations-url"))
    domMerchantCalculationsUrl.Text = elemMerchantCalculationsUrl

    '// Create the <accepts-merchant-coupons> element
    If elemAcceptMerchantCoupons <> "" Then

        Set domAcceptMerchantCoupons = domMerchantCalculations.appendChild( _
            domMerchantCalculationsObj.createElement( _
                "accept-merchant-coupons"))

        domAcceptMerchantCoupons.Text = elemAcceptMerchantCoupons

    End If

    '// Create the <accepts-gift-certificates> element
    If elemAcceptGiftCertificates <> "" Then

        Set domAcceptGiftCertificates = domMerchantCalculations.appendChild( _
            domMerchantCalculationsObj.createElement( _
                "accept-gift-certificates"))

        domAcceptGiftCertificates.Text = elemAcceptGiftCertificates

    End If
    
    Set createMerchantCalculations = domMerchantCalculationsObj

    '// Release the variables used to create the <merchant-calculations> element
    Set domMerchantCalculations = Nothing
    Set domMerchantCalculationsUrl = Nothing
    Set domAcceptMerchantCoupons = Nothing
    Set domAcceptGiftCertificates = Nothing

End Function
'***********************************************************************************
' END: CREATE MERHCANT CALCULATIONS FUNCTION
'***********************************************************************************




'***********************************************************************************
' START: CREATE ROUNDING POLICY FUNCTION
'***********************************************************************************
Function createRoundingPolicy(elemRoundingPolicyMode, elemRoundingPolicyRule)
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createRoundingPolicy()"
	
    '// Define the variables used to create the <rounding-policy> element
    Dim domRoundingPolicy
    Dim domRoundingPolicyMode
	Dim domRoundingPolicyRule

    '// Create the <rounding-policy> element
    Set domRoundingPolicyObj = Server.CreateObject(strMsxmlDomDocument)
    domRoundingPolicyObj.async = False
    Set domRoundingPolicy = domRoundingPolicyObj.appendChild(domRoundingPolicyObj.createElement("rounding-policy"))

    '// Create the <mode> element
    If elemRoundingPolicyMode <> "" Then
    	Set domRoundingPolicyMode = domRoundingPolicy.appendChild(domRoundingPolicyObj.createElement("mode"))
   	 	domRoundingPolicyMode.Text = elemRoundingPolicyMode
    End If
	
    '// Create the <rule> element
    If elemRoundingPolicyRule <> "" Then
        Set domRoundingPolicyRule = domRoundingPolicy.appendChild(domRoundingPolicyObj.createElement("rule"))
        domRoundingPolicyRule.Text = elemRoundingPolicyRule
    End If
    
    Set createRoundingPolicy = domRoundingPolicyObj

    '// Release the variables used to create the <rounding-policy> element
    Set domRoundingPolicy = Nothing
    Set domRoundingPolicyMode = Nothing
    Set domRoundingPolicyRule = Nothing

End Function
'***********************************************************************************
' END: CREATE ROUNDING POLICY FUNCTION
'***********************************************************************************




'***********************************************************************************
' START: CREATE MERHCANT CHECKOUT FLOW SUPPORT
'***********************************************************************************
Function createMerchantCheckoutFlowSupport(elemEditCartUrl, elemContinueShoppingUrl, elemPlatformID)

    '// Define objects used to create the <merchant-checkout-flow-support> XML
    Dim domMerchantCFS
    Dim domEditCartUrl
	Dim domPlatformID
	Dim domalyticsdata
	Dim domBuyerPhone
    Dim domContinueShoppingUrl
    Dim domShippingMethodsRoot
    Dim domTaxTablesRoot
    Dim domMerchantCalculationsRoot
	Dim domRoundingPolicyRoot

    '// Create the <merchant-checkout-flow-support> element
    Set domMerchantCFSObj = Server.CreateObject(strMsxmlDomDocument)
    domMerchantCFSObj.async = False
    Set domMerchantCFS = domMerchantCFSObj.appendChild( _
        domMerchantCFSObj.createElement("merchant-checkout-flow-support"))

    '// Add the <edit-cart-url> element
    If elemEditCartUrl <> "" Then
        Set domEditCartUrl = domMerchantCFS.appendChild( _
            domMerchantCFSObj.createElement("edit-cart-url"))
        domEditCartUrl.Text = elemEditCartUrl
    End If
    
    '// Add the <continue-shopping-url> element
    If elemContinueShoppingUrl <> "" Then
        Set domContinueShoppingUrl = domMerchantCFS.appendChild( _
            domMerchantCFSObj.createElement("continue-shopping-url"))
        domContinueShoppingUrl.Text = elemContinueShoppingUrl
    End If

    '// Add the <shipping-methods> element
    If IsObject(domShippingMethodsObj) Then
        Set domShippingMethodsRoot = domShippingMethodsObj.documentElement
        domMerchantCFS.appendChild(domShippingMethodsRoot.cloneNode(true))
    End If

    '// Add the <tax-tables> element
    If IsObject(domTaxTablesObj) Then
        Set domTaxTablesRoot = domTaxTablesObj.documentElement
        domMerchantCFS.appendChild(domTaxTablesRoot.cloneNode(true))
    End If

    '// Add the <merchant-calculations> element
    If IsObject(domMerchantCalculationsObj) Then
        Set domMerchantCalculationsRoot = _
            domMerchantCalculationsObj.documentElement
        domMerchantCFS.appendChild(domMerchantCalculationsRoot.cloneNode(true))
    End If
	
    '// Add the <rounding-policy> element
    If IsObject(domRoundingPolicyObj) Then
        Set domRoundingPolicyRoot = _
            domRoundingPolicyObj.documentElement
        domMerchantCFS.appendChild(domRoundingPolicyRoot.cloneNode(true))
    End If
	
    '// Add the <request-buyer-phone-number> element
	Set domBuyerPhone = domMerchantCFS.appendChild( domMerchantCFSObj.createElement("request-buyer-phone-number"))
	domBuyerPhone.Text = "true"

	
    '// Add the <platform-id> element
    If elemPlatformID <> "" Then
        Set domPlatformID = domMerchantCFS.appendChild( domMerchantCFSObj.createElement("platform-id"))
        domPlatformID.Text = elemPlatformID
    End If
	
    '// Add the <analytics-data> element
	If pcv_stranalyticsdata <> "" Then
        Set domalyticsdata = domMerchantCFS.appendChild( domMerchantCFSObj.createElement("analytics-data"))
        domalyticsdata.Text = pcv_stranalyticsdata
    End If

    Set createMerchantCheckoutFlowSupport = domMerchantCFSObj

    '// Release objects used to create the <merchant-checkout-flow-support> XML
    Set domMerchantCFS = Nothing
    Set domEditCartUrl = Nothing
	Set domBuyerPhone = Nothing
	Set domPlatformID = Nothing
    Set domContinueShoppingUrl = Nothing
    Set domShippingMethodsObj = Nothing
    Set domShippingMethodsRoot = Nothing
    Set domTaxTablesObj = Nothing
    Set domTaxTablesRoot = Nothing
    Set domMerchantCalculationsObj = Nothing
    Set domMerchantCalculationsRoot = Nothing
	Set domRoundingPolicyObj = Nothing
	Set domRoundingPolicyRoot = Nothing

End Function
'***********************************************************************************
' END: CREATE MERHCANT CHECKOUT FLOW SUPPORT
'***********************************************************************************




'***********************************************************************************
' START: CREATE MERHCANT CHECKOUT SHOPPING CART
'***********************************************************************************
Function createCheckoutShoppingCart()
    
    '// Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createCheckoutShoppingCart()"

    '// Verify that there is a <shopping-cart> XML DOM and a <merchant-checkout-flow-support> XML DOM
    errorType = "MISSING_PARAM"
    If Not(IsObject(domShoppingCartObj)) Then
        errorHandler errorType, strFunctionName, "domShoppingCartObj", _
            domShoppingCartObj
    End If

    '// Define the variables used to create the <checkout-shopping-cart> element
    Dim domCheckoutShoppingCart
    Dim domShoppingCartRoot
    Dim domCFSRoot
    Dim domCFS
    Dim domMerchantCFSRoot
    
    '// Create the <checkout-flow-support> element and add the <merchant-checkout-flow-support> element as a child element
    Set domCFSObj = Server.CreateObject(strMsxmlDomDocument)
    domCFSObj.async = False
    Set domCFS = _
        domCFSObj.appendChild(domCFSObj.createElement("checkout-flow-support"))

    Set domMerchantCFSRoot = domMerchantCFSObj.documentElement
    domCFS.appendChild(domMerchantCFSRoot.cloneNode(true))

    Set domCheckoutShoppingCartObj = Server.CreateObject(strMsxmlDomDocument)
    domCheckoutShoppingCartObj.async = False

    domCheckoutShoppingCartObj.appendChild( _
        domCheckoutShoppingCartObj.createProcessingInstruction( _
            "xml", strXmlVersionEncoding))

    '// Create the <checkout-shopping-cart> element
    Set domCheckoutShoppingCart = domCheckoutShoppingCartObj.appendChild( _
        domCheckoutShoppingCartObj.createElement("checkout-shopping-cart"))
    domCheckoutShoppingCart.setAttribute "xmlns", strXmlns

    '// Add the <shopping-cart> element as a child element of the <checkout-shopping-cart> element
    Set domShoppingCartRoot = domShoppingCartObj.documentElement
    domCheckoutShoppingCart.appendChild(domShoppingCartRoot.cloneNode(true))

    Set domCFSRoot = domCFSObj.documentElement
    domCheckoutShoppingCart.appendChild(domCFSRoot.cloneNode(true))

    createCheckoutShoppingCart = domCheckoutShoppingCartObj.xml

    '// Release the variables used to create the <checkout-shopping-cart> element
    Set domShoppingCartObj = Nothing
    Set domCheckoutShoppingCart = Nothing
    Set domShoppingCartRoot = Nothing
    Set domCFSObj = Nothing
    Set domCFSRoot = Nothing
    Set domCFS = Nothing
    Set domMerchantCFSObj = Nothing
    Set domMerchantCFSRoot = Nothing

End Function
'***********************************************************************************
' END: CREATE MERHCANT CHECKOUT SHOPPING CART
'***********************************************************************************
%>

