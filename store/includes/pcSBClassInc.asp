<!--#include file="pcSBBase64.asp"-->
<%
Dim gv_EndPoint
Dim gv_EndPointManagement
Dim gv_APIUserName
Dim gv_APIPassword
Dim gv_APISignature

Dim gv_ProxyServer	
Dim gv_ProxyServerPort 
Dim gv_Proxy	

Dim SB_ErrMsg	

private const gv_RootURL 		= "https://www.subscriptionbridge.com/"
gv_EndPoint						= gv_RootURL & "Subscriptions/Service2.svc"
gv_EndPointManagement			= gv_RootURL & "Management/Service3.svc"

gv_APIUserName		= API_USERNAME
gv_APIPassword		= API_PASSWORD
gv_APISignature 	= API_SIGNATURE
gv_Version			= API_VERSION

gv_ProxyServer 		= ""
gv_ProxyServerPort 	= ""
gv_UseProxy 		= ""


Class pcARBClass


	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Properties
	'//////////////////////////////////////////////////////////////////////////////
	
	'' Order Data
	Private sb_IdOrder
	Private sb_IdCustomer
	Private sb_IdProductOrdered
	Private sb_IdSubscription
	
	'Gate Way URLs
	Private sb_DevTestURL
	Private sb_TestURL
	Private sb_AuthPostURL
	Private sb_DevTestMode
	
	'' Payment Data
	Private sb_PayInfoType
	Private sb_PayInfoToken
	Private sb_PayInfoPayerID
	Private sb_PaymentCode
	Private sb_UpdMode

	''' CARD DATA
	Private sb_PayInfoExpYear
	Private sb_PayInfoExpMonth
	Private sb_PayInfoCardNumber
	Private sb_PayInfoCVVNumber
	
	'Check Data
	Private sb_PayInfoDriversLicenseNum
	Private sb_PayInfoDriversLicenseState
	Private sb_PayInfoBankAcctName
	Private sb_PayInfoBankABACode
	Private sb_PayInfoBankAcctNum 
	Private sb_PayInfoAccountNumber
	Private sb_PayInfoCardType
	Private sb_PayInfoBankAcctType
	Private sb_PayInfoInfoBankName
	Private sb_PayInfoBankAcctOrgType 
	Private sb_PayInfoCustomerTaxId 
	
	Private sb_BillingFirstName
	Private sb_BillingLastName
	Private sb_BillingCompany
	Private sb_BillingAddress
	Private sb_BillingAddress2
	Private sb_BillingCity
	Private sb_BillingPostalCode
	Private sb_BillingStateCode
	Private sb_BillingState
	Private sb_BillingProvince
	Private sb_BillingCountryCode
	Private sb_BillingPhone
	Private sb_CustomerEmail
	
	Private sb_ShippingFirstName
	Private sb_ShippingLastName
	Private sb_ShippingCompany
	Private sb_ShippingAddress
	Private sb_ShippingAddress2
	Private sb_ShippingCity
	Private sb_ShippingPostalCode
	Private sb_ShippingStateCode
	Private sb_ShippingState
	Private sb_ShippingProvince
	Private sb_ShippingCountryCode
	Private sb_ShippingPhone
	Private sb_ShippingEmail
	
	Private sb_SubscriptionID
	Private sb_SubUnitPrice 
	Private sb_SubQty 
	Private sb_SubLength 
	Private sb_subUnit 
	Private sb_SubAmount 
	Private sb_SubActive
	Private sb_subTrial
	Private sb_SubTrialOccur 
	Private sb_SubTotalOccur 
	Private sb_subStartDate 
	Private sb_NoShippingFlag
	Private sb_SubUPDStartDate
	Private sb_TrialAmount
	Private sb_SubType 
	Private sb_CustomerPassword
	Private sb_CustomerAccount
	
	Private XML
	Private xmlDoc   
	Private rs
	Private rstemp
	Private sb_IsErrorResults
	Private sb_ResultSubscriptionID
	Private sb_ResultApproval
	Private sb_ResultResponseCode 
	Private	sb_ResultResponseMess
	
	Private dtYear, dtMonth, x_Type, x_Login, x_Password
	Private  x_Curcode, x_AIMType, x_CVV, x_testmode, x_secureSource 
	Private x_TypeArray, Post_Data, strStatus, strRetVal
	
	private sub Class_Initialize() 

		'On Error Resume Next
		'call opendb()
		sb_DevTestMode = 0
		sb_IsErrorResults = 0
		If Err.Number<>0 Then
			Err.Number=0
		End If		
	end sub 
	
	private sub Class_Terminate()		
		'call closedb()
	end sub 
	
	Public Property Let IdOrder(intIdOrder)
		sb_IdOrder = intIdOrder
	End Property
	
	Public Property Get IdOrder
	  IdOrder = sb_IdOrder
	End Property
	
	Public Property Let IdCustomer(intIdCustomer)
		sb_IdCustomer = intIdCustomer
	End Property
	
	Public Property Get IdCustomer
	  IdCustomer = sb_IdCustomer
	End Property
	
	Public Property Let IdProductOrdered(intIdProductOrdered)
		sb_IdProductOrdered = intIdProductOrdered
	End Property
	
	Public Property Get IdProductOrdered
	  IdProductOrdered = sb_IdProductOrdered
	End Property
		
	Public Property Let IdSubscription(intIdSubscription)
		sb_IdSubscription = intIdSubscription
	End Property

	Public Property Get IdSubscription
	  IdSubscription = sb_IdSubscription
	End Property
		
	Public Property Let DevTestURL(strDevTestURL)
		sb_DevTestURL = strDevTestURL
	End Property
	
	Public Property Get DevTestURL
	  DevTestURL = sb_DevTestURL
	End Property
		
	Public Property Let TestURL(strTestURL)
		sb_TestURL = strTestURL
	End Property
	
	Public Property Get TestURL
	  TestURL = sb_TestURL
	End Property

	Public Property Let URL(strURL)
		sb_AuthPostURL = strURL
	End Property
	
	Public Property Get URL
	  URL = sb_AuthPostURL
	End Property
	
	Public Property Let PayInfoType(strPayInfoType)
		sb_PayInfoType = strPayInfoType
	End Property
	
	Public Property Get PayInfoType
	  PayInfoType = sb_PayInfoType
	End Property
	
	Public Property Let PayInfoToken(strPayInfoToken)
		sb_PayInfoToken = strPayInfoToken
	End Property
	
	Public Property Get PayInfoToken
	  PayInfoToken = sb_PayInfoToken
	End Property
	
	Public Property Let PayInfoPayerID(strPayInfoPayerID)
		sb_PayInfoPayerID = strPayInfoPayerID
	End Property
	
	Public Property Get PayInfoPayerID
	  PayInfoPayerID = sb_PayInfoPayerID
	End Property
	
	Public Property Let PayInfoExpMonth(strPayInfoExpMonth)
		sb_PayInfoExpMonth = strPayInfoExpMonth
	End Property
	
	Public Property Get PayInfoExpMonth
	  PayInfoExpMonth = sb_PayInfoExpMonth
	End Property
	
	
	'CC Data
	Public Property Let PayInfoExpYear(strPayInfoExpYear)
		sb_PayInfoExpYear = strPayInfoExpYear
	End Property
	
	Public Property Get PayInfoExpYear
	  PayInfoExpYear = sb_PayInfoExpYear
	End Property
	
	Public Property Let PayInfoCardNumber(strPayInfoCardNumber)
		sb_PayInfoCardNumber = strPayInfoCardNumber
	End Property
	
	Public Property Get PayInfoCardNumber
	  PayInfoCardNumber = sb_PayInfoCardNumber
	End Property

	Public Property Let PayInfoCVVNumber(strPayInfoCVVNumber)
		sb_PayInfoCVVNumber = strPayInfoCVVNumber
	End Property
	
	Public Property Get PayInfoCVVNumber
	  PayInfoCVVNumber = sb_PayInfoCVVNumber
	End Property
	
	
	'Check Data
	Public Property Let PayInfoDriversLicenseNum(strPayInfoDriversLicenseNum)
		sb_PayInfoDriversLicenseNum = strPayInfoDriversLicenseNum
	End Property
	
	Public Property Get PayInfoDriversLicenseNum
	  PayInfoDriversLicenseNum = sb_PayInfoDriversLicenseNum
	End Property
	
	
	Public Property Let PayInfoDriversLicenseState(strPayInfoDriversLicenseState)
		sb_PayInfoDriversLicenseState = strPayInfoDriversLicenseState
	End Property
	
	Public Property Get PayInfoDriversLicenseState
	  PayInfoDriversLicenseState = sb_PayInfoDriversLicenseState
	End Property
	
	
	
	Public Property Let PayInfoDriversLicenseDOB(strPayInfoDriversLicenseDOB)
		sb_PayInfoDriversLicenseDOB   = strPayInfoDriversLicenseDOB 
	End Property
	
	Public Property Get PayInfoDriversLicenseDOB  
		PayInfoDriversLicenseDOB   = sb_PayInfoDriversLicenseDOB 
	End Property
	
	
	Public Property Let PayInfoBankAcctName(strPayInfoBankAcctName)
		sb_PayInfoBankAcctName = strPayInfoBankAcctName
	End Property
	
	Public Property Get PayInfoBankAcctName
		PayInfoBankAcctName = sb_PayInfoBankAcctName
	End Property
	
	Public Property Let PayInfoBankABACode(strPayInfoBankABACode)
		sb_PayInfoBankABACode = strPayInfoBankABACode
	End Property
	
	Public Property Get PayInfoBankABACode
		PayInfoBankABACode = sb_PayInfoBankABACode
	End Property
	
	
	Public Property Let PayInfoBankAcctNum(strPayInfoBankAcctNum)
		sb_PayInfoBankAcctNum  = strPayInfoBankAcctNum 
	End Property
	
	Public Property Get PayInfoBankAcctNum 
		PayInfoBankAcctNum  = sb_PayInfoBankAcctNum 
	End Property
	
	Public Property Let PayInfoAccountNumber(strPayInfoAccountNumber)
		sb_PayInfoAccountNumber  = strPayInfoAccountNumber 
	End Property
	
	Public Property Get PayInfoAccountNumber 
		PayInfoAccountNumber  = sb_PayInfoBankAcctNum 
	End Property
	
	Public Property Let PayInfoCardType(strPayInfoCardType)
		sb_PayInfoCardType  = strPayInfoCardType 
	End Property
	
	Public Property Get PayInfoCardType 
		PayInfoCardType  = sb_PayInfoCardType 
	End Property
	
	Public Property Let PayInfoBankAcctType(strPayInfoBankAcctType)
		sb_PayInfoBankAcctType  = strPayInfoBankAcctType 
	End Property
	
	Public Property Get PayInfoBankAcctType 
		PayInfoBankAcctType  = sb_PayInfoBankAcctType
	End Property
	
	Public Property Let PayInfoInfoBankName(strPayInfoInfoBankName)
		sb_PayInfoInfoBankName  = strPayInfoInfoBankName 
	End Property
	
	Public Property Get PayInfoInfoBankName 
		PayInfoInfoBankName  = sb_PayInfoInfoBankName
	End Property
	
	Public Property Let PayInfoBankAcctOrgType(strPayInfoBankAcctOrgType)
		sb_PayInfoBankAcctOrgType   = strPayInfoBankAcctOrgType  
	End Property
	
	Public Property Get PayInfoBankAcctOrgType  
		PayInfoBankAcctOrgType   = sb_PayInfoBankAcctOrgType 
	End Property
	
	
	Public Property Let PayInfoCustomerTaxId(strPayInfoCustomerTaxId)
		sb_PayInfoCustomerTaxId   = strPayInfoCustomerTaxId  
	End Property
	
	Public Property Get PayInfoCustomerTaxId  
		PayInfoCustomerTaxId   = sb_PayInfoCustomerTaxId 
	End Property
	
	
	' Bill data
	
	Public Property Let BillingFirstName(strBillingFirstName)
		sb_BillingFirstName   = strBillingFirstName  
	End Property
	
	Public Property Get BillingFirstName  
		 BillingFirstName   = sb_BillingFirstName 
	End Property
	
	Public Property Let BillingLastName(strBillingLastName)
		sb_BillingLastName   = strBillingLastName  
	End Property
	
	Public Property Get BillingLastName  
		BillingLastName   = sb_BillingLastName 
	End Property
	
	Public Property Let BillingCompany(strBillingCompany)
		sb_BillingCompany   = strBillingCompany  
	End Property
	
	Public Property Get BillingCompany  
		BillingCompany   = sb_BillingCompany 
	End Property
	
	
	Public Property Let BillingAddress(strBillingAddress)
		sb_BillingAddress   = strBillingAddress  
	End Property
	
	Public Property Get BillingAddress  
		BillingAddress   = sb_BillingAddress 
	End Property
	
	
	Public Property Let BillingAddress2(strBillingAddress2)
		sb_BillingAddress2   = strBillingAddress2  
	End Property
	
	Public Property Get BillingAddress2  
		BillingAddress2   = sb_BillingAddress2 
	End Property
	
	Public Property Let BillingCity(strBillingCity)
		sb_BillingCity   = strBillingCity  
	End Property
	
	Public Property Get BillingCity  
			BillingCity   = sb_BillingCity 
	End Property
	
	Public Property Let BillingPostalCode(strBillingPostalCode)
		sb_BillingPostalCode   = strBillingPostalCode  
	End Property
	
	Public Property Get BillingPostalCode  
			BillingPostalCode   = sb_BillingPostalCode 
	End Property
	
	Public Property Let BillingStateCode(strBillingStateCode)
			sb_BillingStateCode   = strBillingStateCode  
	End Property
	
	Public Property Get BillingStateCode  
			BillingStateCode   = sb_BillingStateCode 
	End Property
	
	Public Property Let BillingProvince(strBillingProvince)
			sb_BillingProvince   = strBillingProvince  
	End Property
	
	Public Property Get BillingProvince  
			BillingProvince   = sb_BillingProvince 
	End Property
	
	Public Property Let BillingCountryCode(strBillingCountryCode)
			sb_BillingCountryCode   = strBillingCountryCode  
	End Property
	
	Public Property Get BillingCountryCode  
			BillingCountryCode   = sb_BillingCountryCode 
	End Property
	
	Public Property Let BillingPhone(strBillingPhone)
			sb_BillingPhone   = strBillingPhone  
	End Property
	
	Public Property Get BillingPhone  
			BillingPhone   = sb_BillingPhone 
	End Property
	
	Public Property Let CustomerEmail(strCustomerEmail)
			sb_CustomerEmail   = strCustomerEmail  
	End Property
	
	Public Property Get CustomerEmail  
			CustomerEmail   = sb_CustomerEmail 
	End Property


	' Shipping data
	
	Public Property Let ShippingFirstName(strShippingFirstName)
			sb_ShippingFirstName   = strShippingFirstName  
	End Property
	
	Public Property Get ShippingFirstName  
			 ShippingFirstName   = sb_ShippingFirstName 
	End Property
	
	Public Property Let ShippingLastName(strShippingLastName)
			sb_ShippingLastName   = strShippingLastName  
	End Property
	
	Public Property Get ShippingLastName  
			ShippingLastName   = sb_ShippingLastName 
	End Property
	
	Public Property Let ShippingCompany(strShippingCompany)
			sb_ShippingCompany   = strShippingCompany  
	End Property
	
	Public Property Get ShippingCompany  
			ShippingCompany   = sb_ShippingCompany 
	End Property
	
	
	Public Property Let ShippingAddress(strShippingAddress)
			sb_ShippingAddress   = strShippingAddress  
	End Property
	
	Public Property Get ShippingAddress  
			ShippingAddress   = sb_ShippingAddress 
	End Property
	
	
	Public Property Let ShippingAddress2(strShippingAddress2)
			sb_ShippingAddress2   = strShippingAddress2  
	End Property
	
	Public Property Get ShippingAddress2  
			ShippingAddress2   = sb_ShippingAddress2 
	End Property
	
	
	Public Property Let ShippingCity(strShippingCity)
			sb_ShippingCity   = strShippingCity  
	End Property
	
	Public Property Get ShippingCity  
			ShippingCity   = sb_ShippingCity 
	End Property
	
	Public Property Let ShippingPostalCode(strShippingPostalCode)
			sb_ShippingPostalCode   = strShippingPostalCode  
	End Property
	
	Public Property Get ShippingPostalCode  
			ShippingPostalCode   = sb_ShippingPostalCode 
	End Property
	
	Public Property Let ShippingStateCode(strShippingStateCode)
			sb_ShippingStateCode   = strShippingStateCode  
	End Property
	
	Public Property Get ShippingStateCode  
			ShippingStateCode   = sb_ShippingStateCode 
	End Property
	
	Public Property Let ShippingProvince(strShippingProvince)
			sb_ShippingProvince   = strShippingProvince  
	End Property
	
	Public Property Get ShippingProvince  
			ShippingProvince   = sb_ShippingProvince 
	End Property
	
	Public Property Let ShippingCountryCode(strShippingCountryCode)
			sb_ShippingCountryCode   = strShippingCountryCode  
	End Property
	
	Public Property Get ShippingCountryCode  
			ShippingCountryCode   = sb_ShippingCountryCode 
	End Property
	
	Public Property Let ShippingPhone(strShippingPhone)
			sb_ShippingPhone   = strShippingPhone  
	End Property
	
	Public Property Get ShippingPhone  
			ShippingPhone 	= sb_ShippingPhone 
	End Property
	
	Public Property Let ShippingEmail(strShippingEmail)
			sb_ShippingEmail   = strShippingEmail  
	End Property
	
	Public Property Get ShippingEmail  
			ShippingEmail   = sb_ShippingEmail 
	End Property
	
	
	' ARB DATA
	
	Public Property Let SubscriptionID(intSubscriptionID)
			sb_SubscriptionID   = intSubscriptionID  
	End Property
	
	Public Property Get SubscriptionID  
			SubscriptionID   = sb_SubscriptionID 
	End Property
	
	
	Public Property Let SubUnitPrice(curSubUnitPrice)
			sb_SubUnitPrice   = curSubUnitPrice  
	End Property
	
	Public Property Get SubUnitPrice  
			SubUnitPrice   = sb_SubUnitPrice 
	End Property
	
	Public Property Let SubQty(intSubQty)
			sb_SubQty   = intSubQty  
	End Property
	
	Public Property Get SubQty  
			SubQty   = sb_SubQty 
	End Property
	
	
	Public Property Let SubLength(intSubLength)
			sb_SubLength   = intSubLength  
	End Property
	
	Public Property Get SubLength  
			SubLength   = sb_SubLength 
	End Property
	
	
	Public Property Let SubUnit(strSubUnit)
			sb_SubUnit   = strSubUnit  
	End Property
	
	Public Property Get SubUnit  
			SubUnit   = sb_SubUnit 
	End Property
	
	 
	Public Property Let SubAmount(curSubAmount)
			sb_SubAmount   = curSubAmount  
	End Property
	
	Public Property Get SubAmount  
			SubAmount   = sb_SubAmount 
	End Property
	
	Public Property Let SubActive(intSubActive)
			sb_SubActive   = intSubActive  
	End Property
	
	Public Property Get SubActive  
			SubActive   = sb_SubActive 
	End Property
	
	
	Public Property Let SubTrial(intSubTrial)
			sb_SubTrial   = intSubTrial  
	End Property
	
	Public Property Get SubTrial  
			SubTrial   = sb_SubTrial 
	End Property
	 
	Public Property Let SubTrialOccur(intSubTrialOccur)
			sb_SubTrialOccur   = intSubTrialOccur  
	End Property
	
	Public Property Get SubTrialOccur  
			SubTrialOccur   = sb_SubTrialOccur 
	End Property
	
	Public Property Let SubTotalOccur(intSubTotalOccur)
			sb_SubTotalOccur   = intSubTotalOccur  
	End Property
	
	Public Property Get SubTotalOccur  
			SubTotalOccur   = sb_SubTotalOccur 
	End Property
	
	
	Public Property Let SubStartDate(dtSubStartDate)
			sb_SubStartDate   = dtSubStartDate  
	End Property
	
	Public Property Get SubStartDate  
			SubStartDate   = sb_SubStartDate 
	End Property
	
	
	Public Property Let NoShippingFlag(intNoShippingFlag)
			sb_NoShippingFlag   = intNoShippingFlag  
	End Property
	
	Public Property Get NoShippingFlag  
			NoShippingFlag   = sb_NoShippingFlag 
	End Property
	
	
	Public Property Let TrialAmount(curTrialAmount)
			sb_TrialAmount   = curTrialAmount  
	End Property
	
	Public Property Get TrialAmount  
			TrialAmount   = sb_TrialAmount 
	End Property
	
	Public Property Let SubType(intSubType)
			sb_SubType   = intSubType  
	End Property
	
	Public Property Get SubType  
			SubType   = sb_SubType 
	End Property 


	Public Property Let CustomerPassword(strCustomerPassword)
			sb_CustomerPassword   = strCustomerPassword  
	End Property
	
	Public Property Get CustomerPassword  
			CustomerPassword   = sb_CustomerPassword 
	End Property 


	Public Property Let CustomerAccount(strCustomerAccount)
			sb_CustomerAccount   = strCustomerAccount  
	End Property
	
	Public Property Get CustomerAccount  
			CustomerAccount   = sb_CustomerAccount 
	End Property 


	Public Property Let SubUPDStartDate(dtSubUPDStartDate)
			sb_SubUPDStartDate   = dtSubUPDStartDate  
	End Property
	
	Public Property Get SubUPDStartDate  
			SubUPDStartDate   = sb_SubUPDStartDate 
	End Property
	
	
	
	
	Public Property Let PaymentCode(strPaymentCode)
			sb_PaymentCode   = strPaymentCode  
	End Property
	
	Public Property Get PaymentCode  
			PaymentCode   = sb_PaymentCode
	End Property
	
	
	Public Property Let UpdMode(strUpdMode)
			sb_UpdMode   = strUpdMode  
	End Property
	
	Public Property Get UpdMode  
			UpdMode   = sb_UpdMode
	End Property
	
	
	Private	sb_GUID
	Public Property Let GUID(strGUID)
			sb_GUID   = strGUID  
	End Property	
	Public Property Get GUID  
			GUID   = sb_GUID
	End Property
	
	Private	sb_CartRegularAmt
	Public Property Let CartRegularAmt(strCartRegularAmt)
			sb_CartRegularAmt   = strCartRegularAmt  
	End Property	
	Public Property Get CartRegularAmt  
			CartRegularAmt   = sb_CartRegularAmt
	End Property
	
	Private	sb_CartTrialAmt
	Public Property Let CartTrialAmt(strCartTrialAmt)
			sb_CartTrialAmt   = strCartTrialAmt  
	End Property	
	Public Property Get CartTrialAmt  
			CartTrialAmt   = sb_CartTrialAmt
	End Property
	
	Private	sb_CartRegularTax
	Public Property Let CartRegularTax(strCartRegularTax)
			sb_CartRegularTax   = strCartRegularTax  
	End Property	
	Public Property Get CartRegularTax  
			CartRegularTax   = sb_CartRegularTax
	End Property
	
	Private	sb_CartTrialTax
	Public Property Let CartTrialTax(strCartTrialTax)
			sb_CartTrialTax   = strCartTrialTax  
	End Property	
	Public Property Get CartTrialTax  
			CartTrialTax   = sb_CartTrialTax
	End Property
	
	Private	sb_CartRegularShipping
	Public Property Let CartRegularShipping(strCartRegularShipping)
			sb_CartRegularShipping   = strCartRegularShipping  
	End Property	
	Public Property Get CartRegularShipping  
			CartRegularShipping   = sb_CartRegularShipping
	End Property
	
	Private	sb_CartTrialShipping
	Public Property Let CartTrialShipping(strCartTrialShipping)
			sb_CartTrialShipping   = strCartTrialShipping  
	End Property	
	Public Property Get CartTrialShipping  
			CartTrialShipping   = sb_CartTrialShipping
	End Property
	
	Private	sb_CartIsShippable
	Public Property Let CartIsShippable(strCartIsShippable)
			sb_CartIsShippable   = strCartIsShippable  
	End Property	
	Public Property Get CartIsShippable  
			CartIsShippable   = sb_CartIsShippable
	End Property
	
	Private	sb_CartShipName
	Public Property Let CartShipName(strCartShipName)
			sb_CartShipName   = strCartShipName  
	End Property	
	Public Property Get CartShipName  
			CartShipName   = sb_CartShipName
	End Property
	
	Private	sb_CartTaxName
	Public Property Let CartTaxName(strCartTaxName)
			sb_CartTaxName   = strCartTaxName  
	End Property	
	Public Property Get CartTaxName  
			CartTaxName   = sb_CartTaxName
	End Property
	
	Private	sb_CartAgreedToTerms
	Public Property Let CartAgreedToTerms(strCartAgreedToTerms)
			sb_CartAgreedToTerms   = strCartAgreedToTerms  
	End Property	
	Public Property Get CartAgreedToTerms  
			CartAgreedToTerms   = sb_CartAgreedToTerms
	End Property
	
	Private	sb_CartLanguageCode
	Public Property Let CartLanguageCode(strCartLanguageCode)
			sb_CartLanguageCode   = strCartLanguageCode  
	End Property	
	Public Property Get CartLanguageCode  
			CartLanguageCode   = sb_CartLanguageCode
	End Property
	
	Private	sb_LinkID
	Public Property Let LinkID(strLinkID)
			sb_LinkID   = strLinkID  
	End Property	
	Public Property Get LinkID  
			LinkID   = sb_LinkID
	End Property
	
	Private	sb_BillingPeriod
	Public Property Let BillingPeriod(strBillingPeriod)
			sb_BillingPeriod   = strBillingPeriod  
	End Property	
	Public Property Get BillingPeriod  
			BillingPeriod   = sb_BillingPeriod
	End Property
	
	Private	sb_BillingFrequency
	Public Property Let BillingFrequency(strBillingFrequency)
			sb_BillingFrequency   = strBillingFrequency  
	End Property	
	Public Property Get BillingFrequency  
			BillingFrequency   = sb_BillingFrequency
	End Property
	
	Private	sb_TotalBillingCycles
	Public Property Let TotalBillingCycles(strTotalBillingCycles)
			sb_TotalBillingCycles   = strTotalBillingCycles  
	End Property	
	Public Property Get TotalBillingCycles  
			TotalBillingCycles   = sb_TotalBillingCycles
	End Property
	
	Private	sb_IsTrial
	Public Property Let IsTrial(strIsTrial)
			sb_IsTrial   = strIsTrial  
	End Property	
	Public Property Get IsTrial  
			IsTrial   = sb_IsTrial
	End Property
	
	Private	sb_TrialBillingPeriod
	Public Property Let TrialBillingPeriod(strTrialBillingPeriod)
			sb_TrialBillingPeriod   = strTrialBillingPeriod  
	End Property	
	Public Property Get TrialBillingPeriod  
			TrialBillingPeriod   = sb_TrialBillingPeriod
	End Property
	
	Private	sb_TrialBillingFrequency
	Public Property Let TrialBillingFrequency(strTrialBillingFrequency)
			sb_TrialBillingFrequency   = strTrialBillingFrequency  
	End Property	
	Public Property Get TrialBillingFrequency  
			TrialBillingFrequency   = sb_TrialBillingFrequency
	End Property
		
	Private	sb_TrialTotalBillingCycles
	Public Property Let TrialTotalBillingCycles(strTrialTotalBillingCycles)
			sb_TrialTotalBillingCycles   = strTrialTotalBillingCycles  
	End Property	
	Public Property Get TrialTotalBillingCycles  
			TrialTotalBillingCycles   = sb_TrialTotalBillingCycles
	End Property
		
	Private	sb_StartDate
	Public Property Let StartDate(strStartDate)
			sb_StartDate   = strStartDate  
	End Property	
	Public Property Get StartDate  
			StartDate   = sb_StartDate
	End Property
		
	Private	sb_IsTrialShipping
	Public Property Let IsTrialShipping(strIsTrialShipping)
			sb_IsTrialShipping   = strIsTrialShipping  
	End Property	
	Public Property Get IsTrialShipping  
			IsTrialShipping   = sb_IsTrialShipping
	End Property
		
	Private	sb_CurrencyCode
	Public Property Let CurrencyCode(strCurrencyCode)
			sb_CurrencyCode   = strCurrencyCode  
	End Property	
	Public Property Get CurrencyCode  
			CurrencyCode   = sb_CurrencyCode
	End Property
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Properties
	'//////////////////////////////////////////////////////////////////////////////


	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Remove Invalid XML
	'//////////////////////////////////////////////////////////////////////////////
	Function removeInvalidXML(element)
		element = replace(element,"&"," and ")
		removeInvalidXML = element
	End Function
	'//////////////////////////////////////////////////////////////////////////////
	'// End:  Remove Invalid XML
	'//////////////////////////////////////////////////////////////////////////////
	
	
	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Call Method
	'//////////////////////////////////////////////////////////////////////////////
	Function methodCall(methodName, xmlStr, endPoint)
		Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP"&scXML)	
		strEndPoint=endPoint&"/"&methodName
		objHttp.open "POST", strEndPoint, False
		objHttp.SetRequestHeader "Content-Type", "text/xml"	
		'objHttp.SetProxy gv_Proxy,  gv_ProxyServer& ":" &gv_ProxyServerPort
		objHttp.Send xmlStr	
		result1 = objHttp.responseText	
		'response.Write(result1)
		'response.End()		
		if err.number<>0 then	
			err.number=0
			err.description=""
			methodCall="ERROR"
			set objHttp=nothing
		else
			tmpStatus=objHttp.Status
			set objHttp=nothing
			if (tmpStatus<>200) then
				'response.Write(result1)
				'response.End()
				if result1<>"" then
					if Instr(result1,"ErrorCode")=0 AND (not IsNumeric(result1)) then
						methodCall="ERROR"
					else
						methodCall=result1
					end if
				else
					methodCall="ERROR"
				end if
			else
				methodCall=result1
			end if
		end if
	
	End Function
	'//////////////////////////////////////////////////////////////////////////////
	'// End:  Call Method
	'//////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Register Account
	'//////////////////////////////////////////////////////////////////////////////
	Dim pcv_token, pcv_username
	
	Public Function RegisterAccXML()
		Dim xmlStr
		xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
		xmlStr = xmlStr & 	"<ActivationRequest>"	
		xmlStr = xmlStr & 		"<Username>" & pcv_username & "</Username>"
		xmlStr = xmlStr & 		"<Token>" & pcv_token & "</Token>"					
		xmlStr = xmlStr & 	"</ActivationRequest>"
		RegisterAccXML=xmlStr
	End Function
	
	Function RegisterAcc(APIUser, APIPass, APIKey)
		Dim to_hash, xmlStr, result, currentTime
		'On Error Resume Next

		'// 1.) Get Timestamp
		result = GetCurrentTime()
		If result = "ERROR" Then
			RegisterAcc="0"
			SB_ErrMsg="We could not contact SubscriptionBridge time server. This means SubscriptionBridge may be down temporarily."
			Exit Function 
		Else
			currentTime = pcf_GetNode(result, "CurrentTime", "//GetTimeResponse")
		End IF 

		'// 2.) Create Request		
		to_hash = APIPass & "|" & cstr(currentTime) '// cstr(ReverseParseDate(4, Now()))
		pcv_token = hex_hmac_sha1(APIKey, to_hash)		
		pcv_username = APIUser
		xmlStr = RegisterAccXML

		result = methodCall("ActivationRequest", xmlStr, gv_EndPoint)

		'// 3.) Process Response
		IF result="ERROR" or result="TIMEOUT" THEN
			RegisterAcc="0"
			SB_ErrMsg="Cannot connect to SubscriptionBridge registration system."
		ELSE
			strStatus = pcf_GetNode(result, "Ack", "//ActivationResponse")	
			SB_ErrMsg=""
			if strStatus="Success" then
				RegisterAcc="1"
			else
				tmpCode = pcf_GetNode(result, "ErrorCode", "//Error")
				RegisterAcc="0"
				Select Case tmpCode
					Case "10002": SB_ErrMsg="Authentication/Authorization Failed"
					Case Else: SB_ErrMsg="Error Code: " & tmpCode
				End Select
			end if			
			Set Node = Nothing
			Set Nodes = Nothing
			Set myXmlDoc = Nothing			
		END IF
		
	End Function
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Register Account
	'//////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Get Packages
	'//////////////////////////////////////////////////////////////////////////////
	Public Function GetPackagesXML()
		Dim xmlStr
		xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
		xmlStr = xmlStr & 	"<GetPackagesRequest>"	
		xmlStr = xmlStr & 		"<Username>" & pcv_username & "</Username>"
		xmlStr = xmlStr & 		"<Token>" & pcv_token & "</Token>"
		xmlStr = xmlStr & 		"<LanguageCode>US</LanguageCode>"					
		xmlStr = xmlStr & 	"</GetPackagesRequest>"
		GetPackagesXML=xmlStr
	End Function
	
	Function GetPackages(APIUser, APIPass, APIKey)
		Dim to_hash, xmlStr, result, currentTime
		On Error Resume Next

		'// 1.) Get Timestamp
		result = GetCurrentTime()
		If result = "ERROR" Then
			GetPackages="0"
			SB_ErrMsg="We could not contact SubscriptionBridge time server. This means SubscriptionBridge may be down temporarily."
			Exit Function 
		Else
			currentTime = pcf_GetNode(result, "CurrentTime", "//GetTimeResponse")
		End IF 

		'// 2.) Create Request		
		to_hash = APIPass & "|" & cstr(currentTime) '// cstr(ReverseParseDate(4, Now()))
		pcv_token = hex_hmac_sha1(APIKey, to_hash)	
		pcv_username = APIUser			
		xmlStr = GetPackagesXML
	
		result = methodCall("GetPackagesRequest", xmlStr, gv_EndPoint)

		'// 3.) Process Response
		IF result="ERROR" or result="TIMEOUT" THEN
			GetPackages="0"
			SB_ErrMsg="Cannot connect to SubscriptionBridge API."
		ELSE
		
			strStatus = pcf_GetNode(result, "Ack", "//GetPackagesResponse")
			SB_ErrMsg=""
			if strStatus="Success" then
				GetPackages=result
			else				
				tmpCode = pcf_GetNode(result, "ErrorCode", "//Error")
				GetPackages="0"	
				Select Case tmpCode
					Case "10002": SB_ErrMsg="Authentication/Authorization Failed"
					Case Else: SB_ErrMsg="Error Code: " & tmpCode
				End Select
			end if			
			Set Node = Nothing
			Set Nodes = Nothing
			Set myXmlDoc = Nothing
			
		END IF
	End Function
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Get Packages
	'//////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Get Current Time
	'//////////////////////////////////////////////////////////////////////////////
	Function GetCurrentTime()
		Dim xmlStr, results
		On Error Resume Next
		xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
		xmlStr = xmlStr & 	"<GetTimeRequest>"	
		xmlStr = xmlStr & 		"<Username></Username>"					
		xmlStr = xmlStr & 	"</GetTimeRequest>"		
		results = methodCall("GetTimeRequest",xmlStr, gv_EndPoint)		
		GetCurrentTime = ProcessResults(results)
	End Function
	
	Function ProcessResults(results)
		if err.number<>0 then
			err.number=0
			err.description=""
			ProcessResults="ERROR"
		else
			ProcessResults=results
		end if	
	End Function
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Get Current Time
	'//////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Get Current Time
	'//////////////////////////////////////////////////////////////////////////////
	Function GetTerms(LinkID)
		Dim xmlStr, result
		On Error Resume Next
		
		'// 1.) Prepare Request
		xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
		xmlStr = xmlStr & 	"<GetTermsRequest>"	
		xmlStr = xmlStr & 		"<LinkID>"& LinkID &"</LinkID>"					
		xmlStr = xmlStr & 	"</GetTermsRequest>"
		
		'// 2.) Send Requesst		
		result = methodCall("GetTermsRequest",xmlStr, gv_EndPoint)	
		'response.Write(results)
		'response.End()	
		
		'// 3.) Process Response
		IF result="ERROR" or result="TIMEOUT" THEN
			GetTerms="0"
			SB_ErrMsg="Cannot connect to SubscriptionBridge API."
		ELSE
		
			strStatus = pcf_GetNode(result, "Ack", "//GetTermsResponse")

			SB_ErrMsg=""
			if strStatus="Success" then
				GetTerms=result
			else				
				tmpCode = pcf_GetNode(result, "ErrorCode", "//Error")
				GetTerms="0"	
				Select Case tmpCode
					Case "10002": SB_ErrMsg="Authentication/Authorization Failed"
					Case Else: SB_ErrMsg="Error Code: " & tmpCode
				End Select
			end if			
			Set Node = Nothing
			Set Nodes = Nothing
			Set myXmlDoc = Nothing
			
		END IF
		
	End Function
	
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Get Current Time
	'//////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Ultility Methods
	'//////////////////////////////////////////////////////////////////////////////
	
	Function TermsWidget(LinkID)
		result = GetTerms(LinkID)
		pcv_strTermsBilling = pcf_GetNode(result, "TermsBilling", "//GetTermsResponse")
		pcv_strTermsTrialBilling = pcf_GetNode(result, "TermsTrialBilling", "//GetTermsResponse")
		pcv_strTermsCustom = pcf_GetNode(result, "TermsCustom", "//GetTermsResponse")
		TermsWidget = "<div id=""pcSBTerms"">"
		if pcIsSubTrial = True OR pcv_intIsTrial then 
			TermsWidget=TermsWidget&"<div class=""TermsMain"">" & pcv_strTermsTrialBilling & "</div>"
			TermsWidget=TermsWidget&"<div class=""TermsSub"">" & pcv_strTermsBilling & "</div>"		
		else
			TermsWidget=TermsWidget&"<div class=""TermsMain"">" & pcv_strTermsBilling & "</div>"
			TermsWidget=TermsWidget&"<div class=""TermsSub"">" & pcv_strTermsTrialBilling & "</div>"
		end if  
		TermsWidget=TermsWidget&"<div class=""TermsCustom"">" & pcv_strTermsCustom & "</div>"
		TermsWidget=TermsWidget&"</div>"
	End Function
	
	
	Function CheckExistTag(tagName)		
		Dim tmpNode
		Set tmpNode=ResponseObj.selectSingleNode(tagName)
		If tmpNode is Nothing Then
			CheckExistTag=False
		Else
			CheckExistTag=True
		End if
	End Function


	Function pcf_GetNode(responseXML, nodeName, nodeParent)
		Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)				 
		myXmlDoc.loadXml(responseXML)
		'response.Write(nodeParent)
		Set Nodes = myXmlDoc.selectnodes(nodeParent)	
		For Each Node In Nodes	
			pcf_GetNode = pcf_CheckNode(Node,nodeName,"")				
		Next
		Set Node = Nothing
		Set Nodes = Nothing
		Set myXmlDoc = Nothing
	End Function


	Function pcf_CheckNode(Node,tagName,default)		
		Dim tmpNode
		Set tmpNode=Node.selectSingleNode(tagName)
		If tmpNode is Nothing Then
			pcf_CheckNode=default
		Else
			pcf_CheckNode=Node.selectSingleNode(tagName).text
		End if
	End Function
	
	
	Function ParseDate(RFC)
		if RFC<>"" AND isNULL(RFC)=False then
			ParseDate = mid(RFC, 1, 10) & " " & mid(RFC, 12, 8 )
			ParseDate = formatdatetime(ParseDate)
		end if
	End Function
	
	
	Function ReverseParseDate(offset, pTime)
		pTime=LocalTime(offset,pTime)
		ReverseParseDate = ToDateStamp(pTime) & "T" & ToTimeStamp(pTime)
	End Function	
	
	
	Function LocalTime(offset, pTime)
		LocalTime=DateAdd("h", offset, pTime)
	End Function	
	
	
	Function UTCTime(offset, pTime)
		UTCTime=DateAdd("h", -(offset), pTime)
	End Function
	
	
	Function ToDateStamp(ByVal dt)
		ToDateStamp = Year(dt) & "-" & Right("00" & Month(dt), 2)  & "-" & Right("00" & Day(dt), 2)
	End Function
	
	
	Function ToTimeStamp(ByVal dt)
		ToTimeStamp = Right("00" & Hour(dt), 2) & ":" & Right("00" & Minute(dt), 2) & ":" & "00" '// Right("00" & Second(dt), 2)
	End Function
	
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Ultility Methods
	'//////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Get Sub ID
	'//////////////////////////////////////////////////////////////////////////////
	Public Function getSubscriptionID(idProduct)
		query = "Select * from SB_Packages where idProduct="& idProduct &";"
		set rsSubscriptionID=Server.CreateObject("ADODB.RecordSet")
		set rsSubscriptionID=connTemp.execute(query)
		If NOT rsSubscriptionID.EOF Then
			getSubscriptionID=rsSubscriptionID("SB_PackageID")
		Else
			getSubscriptionID=0
		End If			
		set rsSubscriptionID = Nothing
	End Function
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Get Sub ID
	'//////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Subscription Request
	'//////////////////////////////////////////////////////////////////////////////
	Public Function SubscriptionRequestXML()
		Dim xmlStr				
				xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
				xmlStr = xmlStr & 	"<SubscriptionRequest>"	
				xmlStr = xmlStr & 		"<Username>" & pcv_username & "</Username>"
				xmlStr = xmlStr & 		"<Token>" & pcv_token & "</Token>"
				xmlStr = xmlStr & 		"<Customer>"
				xmlStr = xmlStr & 			"<Email>" & CustomerEmail & "</Email>"
				xmlStr = xmlStr & 			"<FirstName>" & BillingFirstName & "</FirstName>"
				xmlStr = xmlStr & 			"<LastName>" & BillingLastName & "</LastName>"
				xmlStr = xmlStr & 			"<BillingAddress>"
				xmlStr = xmlStr & 				"<FirstName>" & BillingFirstName & "</FirstName>"
				xmlStr = xmlStr & 				"<LastName>" & BillingLastName & "</LastName>"
				If BillingCompany<>"" Then
					xmlStr = xmlStr & 			"<Company>" & BillingCompany & "</Company>"
				End If
				xmlStr = xmlStr & 				"<Address>" & BillingAddress & "</Address>"
				If BillingAddress2<>"" Then
					xmlStr = xmlStr & 			"<Address2>" & BillingAddress2 & "</Address2>"
				End If			
				xmlStr = xmlStr & 				"<City>" & BillingCity & "</City>"
				If BillingStateCode<>"" Then
					xmlStr = xmlStr & 				"<Region>" & BillingStateCode & "</Region>"
				Else
					xmlStr = xmlStr & 				"<Region>" & BillingProvince & "</Region>"		
				End If
				xmlStr = xmlStr & 				"<PostalCode>" & BillingPostalCode & "</PostalCode>"
				xmlStr = xmlStr & 				"<Country>" & BillingCountryCode & "</Country>"
				If BillingPhone<>"" Then
					xmlStr = xmlStr & 			"<Phone>" & BillingPhone & "</Phone>"
				End If
				xmlStr = xmlStr & 			"</BillingAddress>"
				xmlStr = xmlStr & 			"<ShippingAddress>"
				xmlStr = xmlStr & 				"<FirstName>" & ShippingFirstName & "</FirstName>"
				xmlStr = xmlStr & 				"<LastName>" & ShippingLastName & "</LastName>"
				If ShippingCompany<>"" Then
					xmlStr = xmlStr & 			"<Company>" & ShippingCompany & "</Company>"
				End If
				xmlStr = xmlStr & 				"<Address>" & ShippingAddress & "</Address>"
				If ShippingAddress2<>"" Then
					xmlStr = xmlStr & 			"<Address2>" & ShippingAddress2 & "</Address2>"
				End If			
				xmlStr = xmlStr & 				"<City>" & ShippingCity & "</City>"
				If ShippingStateCode<>"" Then
					xmlStr = xmlStr & 				"<Region>" & ShippingStateCode & "</Region>"
				Else
					xmlStr = xmlStr & 				"<Region>" & ShippingProvince & "</Region>"		
				End If
				xmlStr = xmlStr & 				"<PostalCode>" & ShippingPostalCode & "</PostalCode>"
				xmlStr = xmlStr & 				"<Country>" & ShippingCountryCode & "</Country>"
				If ShippingPhone<>"" Then
					xmlStr = xmlStr & 			"<Phone>" & ShippingPhone & "</Phone>"
				End If
				xmlStr = xmlStr & 			"</ShippingAddress>"
				xmlStr = xmlStr & 			"<Password>" & CustomerPassword & "</Password>"
				xmlStr = xmlStr & 			"<Account>" & CustomerAccount & "</Account>"
				xmlStr = xmlStr & 		"</Customer>"
				xmlStr = xmlStr & 		"<CreditCard>"
				xmlStr = xmlStr & 			"<CardNumber>" & PayInfoCardNumber & "</CardNumber>"
				xmlStr = xmlStr & 			"<CardType>" & PayInfoCardType & "</CardType>"
				xmlStr = xmlStr & 			"<ExpMonth>" & PayInfoExpMonth & "</ExpMonth>"
				xmlStr = xmlStr & 			"<ExpYear>" & PayInfoExpYear & "</ExpYear>"
				xmlStr = xmlStr & 			"<SecureCode>" & PayInfoCVVNumber & "</SecureCode>"
				xmlStr = xmlStr & 		"</CreditCard>"
			
				If LinkID<>"" Then
		
					'// Linked Subscription
					xmlStr = xmlStr & 		"<Cart>"
					xmlStr = xmlStr & 			"<RegularAmt>" & CartRegularAmt & "</RegularAmt>"
					xmlStr = xmlStr & 			"<TrialAmt>" & CartTrialAmt & "</TrialAmt>"
					xmlStr = xmlStr & 			"<RegularTax>" & CartRegularTax & "</RegularTax>"
					xmlStr = xmlStr & 			"<TrialTax>" & CartTrialTax & "</TrialTax>"
					xmlStr = xmlStr & 			"<RegularShipping>" & CartRegularShipping & "</RegularShipping>"
					xmlStr = xmlStr & 			"<TrialShipping>" & CartTrialShipping & "</TrialShipping>"
					xmlStr = xmlStr & 			"<IsShippable>" & lcase(CartIsShippable) & "</IsShippable>"
					xmlStr = xmlStr & 			"<ShipName>" & CartShipName & "</ShipName>"
					xmlStr = xmlStr & 			"<TaxName>" & CartTaxName & "</TaxName>"
					xmlStr = xmlStr & 			"<AgreedToTerms>" & lcase(CartAgreedToTerms) & "</AgreedToTerms>"
					xmlStr = xmlStr & 			"<LanguageCode>" & CartLanguageCode & "</LanguageCode>"
					xmlStr = xmlStr & 		"</Cart>"
					xmlStr = xmlStr & 		"<Package>"
					xmlStr = xmlStr & 			"<LinkID>" & LinkID & "</LinkID>"
					xmlStr = xmlStr & 			"<Plan>"
					xmlStr = xmlStr & 				"<Profile>"
					xmlStr = xmlStr & 					"<IsTrial>" & lcase(IsTrial) & "</IsTrial>"
					xmlStr = xmlStr & 				"</Profile>"
					xmlStr = xmlStr & 			"</Plan>"
					xmlStr = xmlStr & 		"</Package>"
		
				
				Else
				
					'// Subscription
					xmlStr = xmlStr & 		"<Cart>"
					xmlStr = xmlStr & 			"<RegularAmt>" & CartRegularAmt & "</RegularAmt>"
					xmlStr = xmlStr & 			"<TrialAmt>" & CartTrialAmt & "</TrialAmt>"
					xmlStr = xmlStr & 			"<RegularTax>" & CartRegularTax & "</RegularTax>"
					xmlStr = xmlStr & 			"<TrialTax>" & CartTrialTax & "</TrialTax>"
					xmlStr = xmlStr & 			"<RegularShipping>" & CartRegularShipping & "</RegularShipping>"
					xmlStr = xmlStr & 			"<TrialShipping>" & CartTrialShipping & "</TrialShipping>"
					xmlStr = xmlStr & 			"<IsShippable>" & CartIsShippable & "</IsShippable>"
					xmlStr = xmlStr & 			"<ShipName>" & CartShipName & "</ShipName>"
					xmlStr = xmlStr & 			"<TaxName>" & CartTaxName & "</TaxName>"
					xmlStr = xmlStr & 			"<AgreedToTerms>" & CartAgreedToTerms & "</AgreedToTerms>"
					xmlStr = xmlStr & 			"<LanguageCode>" & CartLanguageCode & "</LanguageCode>"
					xmlStr = xmlStr & 		"</Cart>"
					xmlStr = xmlStr & 		"<Package>"
					xmlStr = xmlStr & 			"<LinkID>" & LinkID & "</LinkID>"
					xmlStr = xmlStr & 			"<PackageName>Product Name - Plan Name</PackageName>"
					xmlStr = xmlStr & 			"<PackageDescription>Product Description - Plan Description</PackageDescription>"
					xmlStr = xmlStr & 			"<PackagePrice>100</PackagePrice>"
					xmlStr = xmlStr & 			"<TrialName>Name Entered In Interface</TrialName>"
					xmlStr = xmlStr & 			"<TrialDescription>Description Entered In Interfacet</TrialDescription>"
					xmlStr = xmlStr & 			"<TrialPrice>0</TrialPrice>"
					xmlStr = xmlStr & 			"<Plan>"
					xmlStr = xmlStr & 				"<Description>My ProductCart Option Group Name</Description>"
					xmlStr = xmlStr & 				"<Name>My ProductCart Option Value Name</Name>"
					xmlStr = xmlStr & 				"<Profile>"
					xmlStr = xmlStr & 					"<BillingPeriod>" & BillingPeriod & "</BillingPeriod>"
					xmlStr = xmlStr & 					"<BillingFrequency>" & BillingFrequency & "</BillingFrequency>"
					xmlStr = xmlStr & 					"<TotalBillingCycles>" & TotalBillingCycles & "</TotalBillingCycles>"
					xmlStr = xmlStr & 					"<IsTrial>" & IsTrial & "</IsTrial>"
					xmlStr = xmlStr & 					"<TrialBillingPeriod>" & TrialBillingPeriod & "</TrialBillingPeriod>"
					xmlStr = xmlStr & 					"<TrialBillingFrequency>" & TrialBillingFrequency & "</TrialBillingFrequency>"
					xmlStr = xmlStr & 					"<TrialTotalBillingCycles>" & TrialTotalBillingCycles & "</TrialTotalBillingCycles>"
					xmlStr = xmlStr & 					"<StartDate>" & StartDate & "</StartDate>"
					xmlStr = xmlStr & 					"<IsTrialShipping>" & IsTrialShipping & "</IsTrialShipping>"
					xmlStr = xmlStr & 					"<CurrencyCode>" & CurrencyCode & "</CurrencyCode>"
					xmlStr = xmlStr & 				"</Profile>"
					xmlStr = xmlStr & 			"</Plan>"
					xmlStr = xmlStr & 			"<Product>"
					xmlStr = xmlStr & 				"<Description></Description>"
					xmlStr = xmlStr & 				"<Name></Name>"
					'xmlStr = xmlStr & 				"<Terms></Terms>"
					xmlStr = xmlStr & 			"</Product>"
					xmlStr = xmlStr & 		"</Package>"
					'xmlStr = xmlStr & 		"<Features>"
					'xmlStr = xmlStr & 			"<Feature>"
					'xmlStr = xmlStr & 				"<Description>Option Group</Description>"
					'xmlStr = xmlStr & 				"<Name>Option Value</Name>"
					'xmlStr = xmlStr & 				"<Price>10</Price>"
					'xmlStr = xmlStr & 				"<TrialPrice>0</TrialPrice>"
					'xmlStr = xmlStr & 			"</Feature>"
					'xmlStr = xmlStr & 		"</Features>"
				
				End If
				'xmlStr = xmlStr & 		"<NotifyURL></NotifyURL>"			
				xmlStr = xmlStr & 	"</SubscriptionRequest>"
		SubscriptionRequestXML=xmlStr
	End Function
	
	Public Function SubscriptionRequestXML_Linked()
		Dim xmlStr				
				xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
				xmlStr = xmlStr & 	"<SubscriptionRequest>"	
				xmlStr = xmlStr & 		"<Username>" & pcv_username & "</Username>"
				xmlStr = xmlStr & 		"<Token>" & pcv_token & "</Token>"
				xmlStr = xmlStr & 		"<Customer>"
				xmlStr = xmlStr & 			"<Email>" & removeInvalidXML(CustomerEmail) & "</Email>"
				xmlStr = xmlStr & 			"<FirstName>" & removeInvalidXML(BillingFirstName) & "</FirstName>"
				xmlStr = xmlStr & 			"<LastName>" & removeInvalidXML(BillingLastName) & "</LastName>"
				xmlStr = xmlStr & 			"<BillingAddress>"
				xmlStr = xmlStr & 				"<FirstName>" & removeInvalidXML(BillingFirstName) & "</FirstName>"
				xmlStr = xmlStr & 				"<LastName>" & removeInvalidXML(BillingLastName) & "</LastName>"
				If BillingCompany<>"" Then
					xmlStr = xmlStr & 			"<Company>" & removeInvalidXML(BillingCompany) & "</Company>"
				End If
				xmlStr = xmlStr & 				"<Address>" & removeInvalidXML(BillingAddress) & "</Address>"
				If BillingAddress2<>"" Then
					xmlStr = xmlStr & 			"<Address2>" & removeInvalidXML(BillingAddress2) & "</Address2>"
				End If			
				xmlStr = xmlStr & 				"<City>" & removeInvalidXML(BillingCity) & "</City>"
				If BillingStateCode<>"" Then
					xmlStr = xmlStr & 				"<Region>" & removeInvalidXML(BillingStateCode) & "</Region>"
				Else
					xmlStr = xmlStr & 				"<Region>" & removeInvalidXML(BillingProvince) & "</Region>"		
				End If
				xmlStr = xmlStr & 				"<PostalCode>" & removeInvalidXML(BillingPostalCode) & "</PostalCode>"
				xmlStr = xmlStr & 				"<Country>" & removeInvalidXML(BillingCountryCode) & "</Country>"
				If BillingPhone<>"" Then
					xmlStr = xmlStr & 			"<Phone>" & removeInvalidXML(BillingPhone) & "</Phone>"
				End If
				xmlStr = xmlStr & 			"</BillingAddress>"
				'If len(ShippingFirstName)>0 Then
					xmlStr = xmlStr & 			"<ShippingAddress>"
					xmlStr = xmlStr & 				"<FirstName>" & removeInvalidXML(ShippingFirstName) & "</FirstName>"
					xmlStr = xmlStr & 				"<LastName>" & removeInvalidXML(ShippingLastName) & "</LastName>"
					If ShippingCompany<>"" Then
						xmlStr = xmlStr & 			"<Company>" & removeInvalidXML(ShippingCompany) & "</Company>"
					End If
					xmlStr = xmlStr & 				"<Address>" & removeInvalidXML(ShippingAddress) & "</Address>"
					If ShippingAddress2<>"" Then
						xmlStr = xmlStr & 			"<Address2>" & removeInvalidXML(ShippingAddress2) & "</Address2>"
					End If			
					xmlStr = xmlStr & 				"<City>" & removeInvalidXML(ShippingCity) & "</City>"
					If ShippingStateCode<>"" Then
						xmlStr = xmlStr & 				"<Region>" & removeInvalidXML(ShippingStateCode) & "</Region>"
					Else
						xmlStr = xmlStr & 				"<Region>" & removeInvalidXML(ShippingProvince) & "</Region>"		
					End If
					xmlStr = xmlStr & 				"<PostalCode>" & ShippingPostalCode & "</PostalCode>"
					xmlStr = xmlStr & 				"<Country>" & ShippingCountryCode & "</Country>"
					If ShippingPhone<>"" Then
						xmlStr = xmlStr & 			"<Phone>" & ShippingPhone & "</Phone>"
					End If
					xmlStr = xmlStr & 			"</ShippingAddress>"
					xmlStr = xmlStr & 			"<Password>" & CustomerPassword & "</Password>"
					xmlStr = xmlStr & 			"<Account>" & removeInvalidXML(CustomerAccount) & "</Account>"
				'End If
				xmlStr = xmlStr & 		"</Customer>"
				If NOT PayInfoType = "PP" Then
					xmlStr = xmlStr & 		"<CreditCard>"
					xmlStr = xmlStr & 			"<CardNumber>" & PayInfoCardNumber & "</CardNumber>"
					xmlStr = xmlStr & 			"<CardType>" & PayInfoCardType & "</CardType>"
					xmlStr = xmlStr & 			"<ExpMonth>" & PayInfoExpMonth & "</ExpMonth>"
					xmlStr = xmlStr & 			"<ExpYear>" & PayInfoExpYear & "</ExpYear>"
					xmlStr = xmlStr & 			"<SecureCode>" & PayInfoCVVNumber & "</SecureCode>"
					xmlStr = xmlStr & 		"</CreditCard>"
				End If
				xmlStr = xmlStr & 		"<Cart>"
				xmlStr = xmlStr & 			"<RegularAmt>" & CartRegularAmt & "</RegularAmt>"
				if CartTrialAmt>0 OR IsTrial then				
					xmlStr = xmlStr & 			"<TrialAmt>" & CartTrialAmt & "</TrialAmt>"
				end if
				xmlStr = xmlStr & 			"<RegularTax>" & CartRegularTax & "</RegularTax>"
				if CartTrialTax>0 OR IsTrial then	
					xmlStr = xmlStr & 			"<TrialTax>" & CartTrialTax & "</TrialTax>"
				end if
				if CartRegularShipping>0 then
					xmlStr = xmlStr & 			"<RegularShipping>" & CartRegularShipping & "</RegularShipping>"
				end if
				if CartTrialShipping>0 OR IsTrial then
					xmlStr = xmlStr & 			"<TrialShipping>" & CartTrialShipping & "</TrialShipping>"
				end if
				xmlStr = xmlStr & 			"<IsShippable>" & lcase(CartIsShippable) & "</IsShippable>"
				if len(CartShipName)>0 then
					xmlStr = xmlStr & 			"<ShipName>" & removeInvalidXML(CartShipName) & "</ShipName>"
				end if
				if len(CartTaxName) then
					xmlStr = xmlStr & 			"<TaxName>" & removeInvalidXML(CartTaxName) & "</TaxName>"
				end if
				xmlStr = xmlStr & 			"<AgreedToTerms>" & lcase(CartAgreedToTerms) & "</AgreedToTerms>"
				xmlStr = xmlStr & 			"<LanguageCode>" & CartLanguageCode & "</LanguageCode>"
				If PayInfoType = "PP" Then
					xmlStr = xmlStr & 			"<Token>" & PayInfoToken & "</Token>"
					xmlStr = xmlStr & 			"<PayerID>" & PayInfoPayerID & "</PayerID>"
				End If
				xmlStr = xmlStr & 		"</Cart>"
				xmlStr = xmlStr & 		"<Package>"
				xmlStr = xmlStr & 			"<LinkID>" & LinkID & "</LinkID>"
				if CartTrialAmt>0 OR IsTrial then
					'xmlStr = xmlStr & 			"<Plan>"
					'xmlStr = xmlStr & 				"<Profile>"
					'xmlStr = xmlStr & 					"<IsTrial>" & lcase(IsTrial) & "</IsTrial>"
					'xmlStr = xmlStr & 				"</Profile>"
					'xmlStr = xmlStr & 			"</Plan>"
				end if
				xmlStr = xmlStr & 		"</Package>"
				'xmlStr = xmlStr & 		"<NotifyURL></NotifyURL>"			
				xmlStr = xmlStr & 	"</SubscriptionRequest>"
		SubscriptionRequestXML_Linked=xmlStr
	End Function
	
	Public Function SubscriptionRequest(APIUser, APIPass, APIKey)	
		Dim to_hash, xmlStr, result, currentTime
		Dim dtDifTest, dtUnitToADD, sb_SubTransAmount
		'On Error Resume Next


		'// Get Subscription Products
		query = "SELECT idProductOrdered, quantity, unitPrice, pcSubscription_ID, pcPO_SubDetails, pcPO_SubType, pcPO_SubFrequency, pcPO_SubPeriod, pcPO_SubCycles, pcPO_SubStartDate, pcPO_SubAmount, pcPO_SubTrialPeriod, pcPO_SubTrialCycles, pcPO_SubTrialAmount, pcPO_NoShipping, pcPO_SubUPDStartDate, pcPO_LinkID, pcPO_IsTrial FROM ProductsOrdered WHERE idOrder="& sb_IdOrder&" AND pcSubscription_id > 0;"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)		
		If NOT rs.EOF Then

			Do While NOT rs.eof
				
				pcv_SubAmount = rs("pcPO_SubAmount")
				pcv_TrialAmount = rs("pcPO_SubTrialAmount")
				
				sb_idProductOrdered = rs("idProductOrdered")
				sb_SubUnitPrice = rs("unitPrice")
				sb_SubQty = rs("quantity")
				sb_SubscriptionID = rs("pcSubscription_ID")				
				sb_subTrial = rs("pcPO_SubTrialPeriod")
				sb_SubTrialOccur = rs("pcPO_SubTrialCycles") 
				sb_SubUnit = rs("pcPO_SubPeriod")
				sb_SubTotalOccur = rs("pcPO_SubCycles")
				sb_SubLength = rs("pcPO_SubFrequency")
				sb_subStartDate = rs("pcPO_SubStartDate")
				sb_NoShippingFlag = rs("pcPO_NoShipping")
				'sb_subType = rs("pcPO_SubscriptionType")
				sb_LinkID = rs("pcPO_LinkID")
				sb_IsTrial = rs("pcPO_IsTrial")

				'// Package
				LinkID = sb_LinkID
				'PackageName
				'PackageDescription
				'PackagePrice
				'TrialName
				'TrialDescription
				'TrialPrice
				
				'// Plan
				'Description
				'Name
				
				'// Plan Profile
				BillingPeriod = sb_SubUnit
				BillingFrequency = sb_SubLength
				TotalBillingCycles = sb_SubTotalOccur
				IsTrial = sb_IsTrial
				TrialBillingPeriod = sb_SubUnit
				TrialBillingFrequency = sb_SubLength
				TrialTotalBillingCycles = sb_SubTrialOccur
				StartDate = sb_subStartDate
				'IsTrialShipping
				If scSBCurrencyCode<>"" Then
					CurrencyCode = scSBCurrencyCode
				Else
					CurrencyCode = "USD"
				End If
			
				'// Product		
				'Description
				'Name
				'Terms
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// SET PROPERTIES
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// Amount
				CartRegularAmt = pcv_SubAmount
				
				'// Trial Amount
				CartTrialAmt = pcv_TrialAmount
				
				'// Tax 
				'CartRegularTax = / >>> order level set on gwreturn.asp
				
				'// Trial Tax
				'CartTrialTax = / >>> order level set on gwreturn.asp
				
				'// Shipping
				'CartRegularShipping = / >>> order level set on gwreturn.asp
				
				'// Trial Shipping
				'CartTrialShipping = / >>> order level set on gwreturn.asp
				

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// START PROCESSING
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// 1.) Get Timestamp
				result = GetCurrentTime()
				If result = "ERROR" Then
					SubscriptionRequest="0"
					SB_ErrMsg="We could not contact SubscriptionBridge time server. This means SubscriptionBridge may be down temporarily."
					Exit Function 
				Else
					currentTime = pcf_GetNode(result, "CurrentTime", "//GetTimeResponse")
				End IF 


				'// 2.) Create Request
				to_hash = APIPass & "|" & cstr(currentTime)
				pcv_token = hex_hmac_sha1(APIKey, to_hash)
				pcv_username = APIUser
				
				'// 3.)  Send It!
				If len(LinkID)>0 Then
				 	xmlStr = SubscriptionRequestXML_Linked
				Else
					xmlStr = SubscriptionRequestXML
				End If
				
				'response.Clear()
				'response.ContentType="text/xml"
				'response.Write(xmlStr)
				'response.End()
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Subscription Request XML
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		
		
		
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Send the XML
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
				result = methodCall("SubscriptionRequest", xmlStr, gv_EndPoint)
				'response.Write(result)
				'response.End()
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Send the XML
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
				
			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: Process the Response XML
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				IF result="ERROR" or result="TIMEOUT" THEN
				
						tmpCode = "0"
						tmpReason = "Error saving subscription"
						tmpDetail = "There was an unexpected error.  Please contact us so we can help."
						tmpSeverity = "Severe"
				
						SB_ErrMsg = "<div align=""left"">" & _
							"<ul>" &_
							"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
							"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
							"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
							"<li>" & "Error Severity: " & tmpSeverity & "</li>"
						SB_ErrMsg = SB_ErrMsg & "</ul></div>"
					
				ELSE
				
					strStatus = pcf_GetNode(result, "Ack", "//SubscriptionResponse")
					
					'response.Write(result)
					'response.End()
					
					SB_ErrMsg=""
					if strStatus="Success" then
						
						'// Success! Save the GUID for management
						Dim pcv_strGUID, pcv_strTerms
						pcv_strGUID = pcf_GetNode(result, "Guid", "//SubscriptionResponse")
						pcv_strTerms = pcf_GetNode(result, "Terms", "//SubscriptionResponse")
						
						pIdOrder=session("GWOrderId")
						pIdOrder=(int(pIdOrder)-scpre)
						
						query="INSERT INTO SB_Orders (idOrder, SB_GUID, SB_Terms) VALUES ("&pIdOrder&", '"&pcv_strGUID&"', '"&pcv_strTerms&"');"
						set rsSBORD=server.CreateObject("ADODB.RecordSet")
						set rsSBORD=connTemp.execute(query)
						set rsSBORD = nothing

					elseif strStatus="300" then
					
						'// Warning!  Alert the store owner, let customer procede
					
					else
						
						'// Error!  Display error to customer			
						tmpCode = pcf_GetNode(result, "ErrorCode", "//Error")
						tmpReason = pcf_GetNode(result, "ErrorShort", "//Error")
						tmpDetail = pcf_GetNode(result, "ErrorDetail", "//Error")
						tmpSeverity = pcf_GetNode(result, "SeverityLevel", "//Error")

						SB_ErrMsg = "<div align=""left"">" & _
							"<ul>" &_
							"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
							"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
							"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
							"<li>" & "Error Severity: " & tmpSeverity & "</li>"
						SB_ErrMsg = SB_ErrMsg & "</ul></div>"
		
					end if
					
					Set Node = Nothing
					Set Nodes = Nothing
					Set myXmlDoc = Nothing
					
				END IF
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End: Process the Response XML
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

				rs.movenext
			Loop 

		End If
		set rs=nothing
			
	End Function 
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Subscription Request
	'//////////////////////////////////////////////////////////////////////////////





	'//////////////////////////////////////////////////////////////////////////////
	'// START:  OneTime Payment Request
	'//////////////////////////////////////////////////////////////////////////////
	Public Function OneTimePaymentRequestXML()
		Dim xmlStr				
				xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
				xmlStr = xmlStr & 	"<OneTimePaymentRequest>"	
				xmlStr = xmlStr & 		"<Username>" & pcv_username & "</Username>"
				xmlStr = xmlStr & 		"<Token>" & pcv_token & "</Token>"
				xmlStr = xmlStr & 		"<Amount>" & CartRegularAmt & "</Amount>"
				xmlStr = xmlStr & 		"<GUID>" & GUID & "</GUID>"
				xmlStr = xmlStr & 		"<CreditCard>"
				xmlStr = xmlStr & 			"<CardNumber>" & PayInfoCardNumber & "</CardNumber>"
				xmlStr = xmlStr & 			"<CardType>" & PayInfoCardType & "</CardType>"
				xmlStr = xmlStr & 			"<ExpMonth>" & PayInfoExpMonth & "</ExpMonth>"
				xmlStr = xmlStr & 			"<ExpYear>" & PayInfoExpYear & "</ExpYear>"
				xmlStr = xmlStr & 			"<SecureCode>" & PayInfoCVVNumber & "</SecureCode>"
				xmlStr = xmlStr & 		"</CreditCard>"
				xmlStr = xmlStr & 		"<Customer>"
				xmlStr = xmlStr & 			"<Email>" & CustomerEmail & "</Email>"
				xmlStr = xmlStr & 			"<FirstName>" & BillingFirstName & "</FirstName>"
				xmlStr = xmlStr & 			"<LastName>" & BillingLastName & "</LastName>"
				xmlStr = xmlStr & 			"<BillingAddress>"
				xmlStr = xmlStr & 				"<FirstName>" & BillingFirstName & "</FirstName>"
				xmlStr = xmlStr & 				"<LastName>" & BillingLastName & "</LastName>"
				If BillingCompany<>"" Then
					xmlStr = xmlStr & 			"<Company>" & BillingCompany & "</Company>"
				End If
				xmlStr = xmlStr & 				"<Address>" & BillingAddress & "</Address>"
				If BillingAddress2<>"" Then
					xmlStr = xmlStr & 			"<Address2>" & BillingAddress2 & "</Address2>"
				End If			
				xmlStr = xmlStr & 				"<City>" & BillingCity & "</City>"
				If BillingStateCode<>"" Then
					xmlStr = xmlStr & 				"<Region>" & BillingStateCode & "</Region>"
				Else
					xmlStr = xmlStr & 				"<Region>" & BillingProvince & "</Region>"		
				End If
				xmlStr = xmlStr & 				"<PostalCode>" & BillingPostalCode & "</PostalCode>"
				xmlStr = xmlStr & 				"<Country>" & BillingCountryCode & "</Country>"
				If BillingPhone<>"" Then
					xmlStr = xmlStr & 			"<Phone>" & BillingPhone & "</Phone>"
				End If
				xmlStr = xmlStr & 			"</BillingAddress>"
				xmlStr = xmlStr & 			"<ShippingAddress>"
				xmlStr = xmlStr & 				"<FirstName>" & ShippingFirstName & "</FirstName>"
				xmlStr = xmlStr & 				"<LastName>" & ShippingLastName & "</LastName>"
				If ShippingCompany<>"" Then
					xmlStr = xmlStr & 			"<Company>" & ShippingCompany & "</Company>"
				End If
				xmlStr = xmlStr & 				"<Address>" & ShippingAddress & "</Address>"
				If ShippingAddress2<>"" Then
					xmlStr = xmlStr & 			"<Address2>" & ShippingAddress2 & "</Address2>"
				End If			
				xmlStr = xmlStr & 				"<City>" & ShippingCity & "</City>"
				If ShippingStateCode<>"" Then
					xmlStr = xmlStr & 				"<Region>" & ShippingStateCode & "</Region>"
				Else
					xmlStr = xmlStr & 				"<Region>" & ShippingProvince & "</Region>"		
				End If
				xmlStr = xmlStr & 				"<PostalCode>" & ShippingPostalCode & "</PostalCode>"
				xmlStr = xmlStr & 				"<Country>" & ShippingCountryCode & "</Country>"
				If ShippingPhone<>"" Then
					xmlStr = xmlStr & 			"<Phone>" & ShippingPhone & "</Phone>"
				End If
				xmlStr = xmlStr & 			"</ShippingAddress>"
				xmlStr = xmlStr & 			"<Password>" & CustomerPassword & "</Password>"
				xmlStr = xmlStr & 			"<Account>" & CustomerAccount & "</Account>"
				xmlStr = xmlStr & 		"</Customer>"
				xmlStr = xmlStr & 	"</OneTimePaymentRequest>"
		OneTimePaymentRequestXML=xmlStr
	End Function
	
	Public Function OneTimePaymentRequest(APIUser, APIPass, APIKey)	
		Dim to_hash, xmlStr, result, currentTime
		Dim dtDifTest, dtUnitToADD, sb_SubTransAmount
		On Error Resume Next

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// START PROCESSING
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// 1.) Get Timestamp
		result = GetCurrentTime()
		If result = "ERROR" Then
			SubscriptionRequest="0"
			SB_ErrMsg="We could not contact SubscriptionBridge time server. This means SubscriptionBridge may be down temporarily."
			Exit Function 
		Else
			currentTime = pcf_GetNode(result, "CurrentTime", "//GetTimeResponse")
		End IF 


		'// 2.) Create Request
		to_hash = APIPass & "|" & cstr(currentTime)
		pcv_token = hex_hmac_sha1(APIKey, to_hash)
		pcv_username = APIUser
		
		'// 3.)  Send It!
		xmlStr = OneTimePaymentRequestXML
		
		'response.Clear()
		'response.ContentType="text/xml"
		'response.Write(xmlStr)
		'response.End()
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Subscription Request XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Send the XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		result = methodCall("OneTimePaymentRequest", xmlStr, gv_EndPointManagement)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Send the XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Process the Response XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		IF result="ERROR" or result="TIMEOUT" THEN
		
				tmpCode = "0"
				tmpReason = "Error"
				tmpDetail = "Error"
				tmpSeverity = "Severe"
		
				SB_ErrMsg = "<div align=""left"">" & _
					"<ul>" &_
					"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
					"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
					"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
					"<li>" & "Error Severity: " & tmpSeverity & "</li>"
				SB_ErrMsg = SB_ErrMsg & "</ul></div>"
			
		ELSE
		
			strStatus = pcf_GetNode(result, "Ack", "//OneTimePaymentResponse")

			SB_ErrMsg=""
			if strStatus="Success" then
				
				'// Success! Save the GUID for management
				Dim pcv_strTransactionID
				pcv_strTransactionID = pcf_GetNode(result, "TransactionID", "//OneTimePaymentResponse")

				'query="INSERT INTO SB_Orders (idOrder, SB_GUID, SB_Terms) VALUES ("&pIdOrder&", '"&pcv_strGUID&"', '"&pcv_strTerms&"');"
				'set rsSBORD=server.CreateObject("ADODB.RecordSet")
				'set rsSBORD=connTemp.execute(query)
				'set rsSBORD = nothing

			elseif strStatus="300" then
			
				'// Warning!  Alert the store owner, let customer procede
			
			else
				
				'// Error!  Display error to customer			
				tmpCode = pcf_GetNode(result, "ErrorCode", "//Error")
				tmpReason = pcf_GetNode(result, "ErrorShort", "//Error")
				tmpDetail = pcf_GetNode(result, "ErrorDetail", "//Error")
				tmpSeverity = pcf_GetNode(result, "SeverityLevel", "//Error")

				SB_ErrMsg = "<div align=""left"">" & _
					"<ul>" &_
					"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
					"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
					"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
					"<li>" & "Error Severity: " & tmpSeverity & "</li>"
				SB_ErrMsg = SB_ErrMsg & "</ul></div>"

			end if
			
			Set Node = Nothing
			Set Nodes = Nothing
			Set myXmlDoc = Nothing
			
		END IF
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Process the Response XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	End Function 
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  OneTime Payment Request
	'//////////////////////////////////////////////////////////////////////////////
	




	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Cancel Request
	'//////////////////////////////////////////////////////////////////////////////
	Public Function CancellationRequestXML()
		Dim xmlStr				
				xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
				xmlStr = xmlStr & 	"<CancellationRequest>"	
				xmlStr = xmlStr & 		"<Username>" & pcv_username & "</Username>"
				xmlStr = xmlStr & 		"<Token>" & pcv_token & "</Token>"
				xmlStr = xmlStr & 		"<GUID>" & GUID & "</GUID>"
				xmlStr = xmlStr & 		"<Reason>" & Reason & "</Reason>"
				xmlStr = xmlStr & 	"</CancellationRequest>"
		CancellationRequestXML=xmlStr
	End Function


	Public Function CancellationRequest(APIUser, APIPass, APIKey)	
		Dim to_hash, xmlStr, result, currentTime
		Dim dtDifTest, dtUnitToADD, sb_SubTransAmount
		On Error Resume Next

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// START PROCESSING
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// 1.) Get Timestamp
		result = GetCurrentTime()
		If result = "ERROR" Then
			SubscriptionRequest="0"
			SB_ErrMsg="We could not contact SubscriptionBridge time server. This means SubscriptionBridge may be down temporarily."
			Exit Function 
		Else
			currentTime = pcf_GetNode(result, "CurrentTime", "//GetTimeResponse")
		End IF 


		'// 2.) Create Request
		to_hash = APIPass & "|" & cstr(currentTime)
		pcv_token = hex_hmac_sha1(APIKey, to_hash)
		pcv_username = APIUser
		
		'// 3.)  Send It!
		xmlStr = CancellationRequestXML
		
		'response.Clear()
		'response.ContentType="text/xml"
		'response.Write(xmlStr)
		'response.End()
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Subscription Request XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Send the XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		result = methodCall("CancellationRequest", xmlStr, gv_EndPointManagement)
		'response.Write(result)
		'response.End()
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Send the XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Process the Response XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		IF result="ERROR" or result="TIMEOUT" THEN
		
				tmpCode = "0"
				tmpReason = "Error"
				tmpDetail = "Error"
				tmpSeverity = "Severe"
		
				SB_ErrMsg = "<div align=""left"">" & _
					"<ul>" &_
					"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
					"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
					"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
					"<li>" & "Error Severity: " & tmpSeverity & "</li>"
				SB_ErrMsg = SB_ErrMsg & "</ul></div>"
			
		ELSE

			strStatus = pcf_GetNode(result, "Ack", "//CancellationResponse")

			SB_ErrMsg=""
			if strStatus="Success" then
				
				'// Success! Save the GUID for management
				Dim pcv_strTransactionID
				pcv_strTransactionID = pcf_GetNode(result, "TransactionID", "//CancellationResponse")

				'query="INSERT INTO SB_Orders (idOrder, SB_GUID, SB_Terms) VALUES ("&pIdOrder&", '"&pcv_strGUID&"', '"&pcv_strTerms&"');"
				'set rsSBORD=server.CreateObject("ADODB.RecordSet")
				'set rsSBORD=connTemp.execute(query)
				'set rsSBORD = nothing

			elseif strStatus="300" then
			
				'// Warning!  Alert the store owner, let customer procede
			
			else
				
				'// Error!  Display error to customer			
				tmpCode = pcf_GetNode(result, "ErrorCode", "//Error")
				tmpReason = pcf_GetNode(result, "ErrorShort", "//Error")
				tmpDetail = pcf_GetNode(result, "ErrorDetail", "//Error")
				tmpSeverity = pcf_GetNode(result, "SeverityLevel", "//Error")

				SB_ErrMsg = "<div align=""left"">" & _
					"<ul>" &_
					"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
					"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
					"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
					"<li>" & "Error Severity: " & tmpSeverity & "</li>"
				SB_ErrMsg = SB_ErrMsg & "</ul></div>"

			end if
			
			Set Node = Nothing
			Set Nodes = Nothing
			Set myXmlDoc = Nothing
			
		END IF
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Process the Response XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	End Function 
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Cancel Request
	'//////////////////////////////////////////////////////////////////////////////
	




	'//////////////////////////////////////////////////////////////////////////////
	'// START:  Get Subscription Details Request
	'//////////////////////////////////////////////////////////////////////////////
	Public Function GetSubscriptionDetailsRequestXML()
		Dim xmlStr				
				xmlStr = 			"<?xml version=""1.0"" encoding=""utf-8""?>"
				xmlStr = xmlStr & 	"<GetSubscriptionDetailsRequest>"	
				xmlStr = xmlStr & 		"<Username>" & pcv_username & "</Username>"
				xmlStr = xmlStr & 		"<Token>" & pcv_token & "</Token>"
				xmlStr = xmlStr & 		"<GUID>" & GUID & "</GUID>"
				xmlStr = xmlStr & 		"<LanguageCode>" & CartLanguageCode & "</LanguageCode>"
				xmlStr = xmlStr & 	"</GetSubscriptionDetailsRequest>"
		GetSubscriptionDetailsRequestXML=xmlStr
	End Function


	Public Function GetSubscriptionDetailsRequest(APIUser, APIPass, APIKey)	
		Dim to_hash, xmlStr, result, currentTime
		Dim dtDifTest, dtUnitToADD, sb_SubTransAmount
		On Error Resume Next

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// START PROCESSING
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// 1.) Get Timestamp
		result = GetCurrentTime()
		If result = "ERROR" Then
			GetSubscriptionDetailsRequest="0"
			SB_ErrMsg="We could not contact SubscriptionBridge time server. This means SubscriptionBridge may be down temporarily."
			Exit Function 
		Else
			currentTime = pcf_GetNode(result, "CurrentTime", "//GetTimeResponse")
		End IF 


		'// 2.) Create Request
		to_hash = APIPass & "|" & cstr(currentTime)
		pcv_token = hex_hmac_sha1(APIKey, to_hash)
		pcv_username = APIUser
		
		'// 3.)  Send It!
		xmlStr = GetSubscriptionDetailsRequestXML
		
		'response.Clear()
		'response.ContentType="text/xml"
		'response.Write(xmlStr)
		'response.End()
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Subscription Request XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Send the XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		result = methodCall("GetSubscriptionDetailsRequest", xmlStr, gv_EndPointManagement)
		'response.Write(result)
		'response.End()
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Send the XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Process the Response XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		IF result="ERROR" or result="TIMEOUT" THEN
		
				tmpCode = "0"
				tmpReason = "Error"
				tmpDetail = "Error"
				tmpSeverity = "Severe"
		
				SB_ErrMsg = "<div align=""left"">" & _
					"<ul>" &_
					"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
					"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
					"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
					"<li>" & "Error Severity: " & tmpSeverity & "</li>"
				SB_ErrMsg = SB_ErrMsg & "</ul></div>"
			
		ELSE

			strStatus = pcf_GetNode(result, "Ack", "//GetSubscriptionDetailsResponse")

			SB_ErrMsg=""
			if strStatus="Success" then
				
				'// Success! Save the GUID for management
				'Dim pcv_strTransactionID
				'pcv_strTransactionID = pcf_GetNode(result, "TransactionID", "//GetSubscriptionDetailsResponse")

				'query="INSERT INTO SB_Orders (idOrder, SB_GUID, SB_Terms) VALUES ("&pIdOrder&", '"&pcv_strGUID&"', '"&pcv_strTerms&"');"
				'set rsSBORD=server.CreateObject("ADODB.RecordSet")
				'set rsSBORD=connTemp.execute(query)
				'set rsSBORD = nothing
				
				GetSubscriptionDetailsRequest = result

			elseif strStatus="300" then
			
				'// Warning!  Alert the store owner, let customer procede
			
			else
				
				'// Error!  Display error to customer			
				tmpCode = pcf_GetNode(result, "ErrorCode", "//Error")
				tmpReason = pcf_GetNode(result, "ErrorShort", "//Error")
				tmpDetail = pcf_GetNode(result, "ErrorDetail", "//Error")
				tmpSeverity = pcf_GetNode(result, "SeverityLevel", "//Error")

				SB_ErrMsg = "<div align=""left"">" & _
					"<ul>" &_
					"<li>" & "<u>Error Code: " & tmpCode & "</u></li>" &_
					"<li>" & "Error Reason: " & tmpReason & "</li>" &_								
					"<li>" & "Error Detail: " & tmpDetail & "</li>" &_	
					"<li>" & "Error Severity: " & tmpSeverity & "</li>"
				SB_ErrMsg = SB_ErrMsg & "</ul></div>"

			end if
			
			Set Node = Nothing
			Set Nodes = Nothing
			Set myXmlDoc = Nothing
			
		END IF
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Process the Response XML
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	End Function 
	'//////////////////////////////////////////////////////////////////////////////
	'// END:  Get Subscription Details Request
	'//////////////////////////////////////////////////////////////////////////////








End Class
%>