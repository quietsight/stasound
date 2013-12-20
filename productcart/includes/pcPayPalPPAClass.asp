
<%
Dim API_ENDPOINT, API_HEADER, API_VERSION, objPayPalHttp, nvpstr
Dim pcPay_PayPalAd_Partner, pcPay_PayPalAd_MerchantLogin, pcPay_PayPalAd_Vendor, pcPay_PayPalAd_User, pcPay_PayPalAd_Password, pcPay_PayPalAd_TransType, pcPay_PayPalAd_CSC, pcPay_PayPalAd_Sandbox, pcPay_PayPal_Currency, pcPay_PayPal_CardTypes, PaymentAction, pcPay_PayPal_Method

Dim DeclinedString
Dim pErrNumber, pErrDescription, pErrSource, pErrSeverityCode
Dim pcv_strShippingFullName, pcv_strShippingCompany, pcv_strShippingAddress, pcv_strShippingPostalCode, pcv_strShippingStateCode, pcv_strShippingProvince, pcv_strShippingPhone, pcv_strShippingCity, pcv_strShippingCountryCode, pcv_strShippingAddress2

'///////////////////////////////////////////////////////////////////////////////////
'// START: Express Checkout for PayFlow Editions
'///////////////////////////////////////////////////////////////////////////////////
Class pcPayPalClass


	'// Initialize Class
	private sub Class_Initialize() 
		On Error Resume Next
		API_HEADER= "text/namevalue"
		API_VERSION= "2.0"
		Set objPayPalHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		If Err.Number<>0 Then
			Err.Number=0
		End If			
	end sub 

	'// Terminate Class
	private sub Class_Terminate()		
		Set objPayPalHttp = nothing
	end sub 

	'----------------------------------------------------------------------------------
	' Purpose: Make the API call to PayPal, using API signature.
	' Inputs:  Method name to be called & NVP string to be sent with the post method
	' Returns: NVP Collection object of Call Response.
	'----------------------------------------------------------------------------------		
	Public Function hash_call(methodName, nvpStr)	
		On Error Resume Next

		AddNVP "USER", pcPay_PayPalAd_User
		AddNVP "PWD", pcPay_PayPalAd_Password
		AddNVP "VENDOR", pcPay_PayPalAd_MerchantLogin
		AddNVP "PARTNER", pcPay_PayPalAd_Partner 
		
		If methodName="SetExpressCheckout" Then
			AddNVP "TRXTYPE", PaymentAction 'PaymentAction '// S = Sale transaction, A = Authorisation, C = Credit, D = Delayed Capture, V = Void 
			AddNVP "ACTION", "S" '// S = Set, G = Get, D = Do
			AddNVP "AMT", OrderTotal	
			AddNVP "CANCELURL", cancelURL
			'AddNVP "CUSTOM", "TRVV14459"
			'AddNVP "EMAIL", "buyer_name@abc.com"
			AddNVP "RETURNURL["&lenReturnURL&"]", returnURL
			AddNVP "TENDER", "P" '// C = credit card, P = PayPal
		End If
		
		If methodName="GetExpressCheckoutDetails" Then
			AddNVP "TRXTYPE", PaymentAction 'PaymentAction '// S = Sale transaction, A = Authorisation, C = Credit, D = Delayed Capture, V = Void 
			AddNVP "ACTION", "G" '// S = Set, G = Get, D = Do
			AddNVP "TENDER", "P" '// C = credit card, P = PayPal
			AddNVP "TOKEN", token
		End If
		
		If methodName="DoExpressCheckoutPayment" Then
			AddNVP "TRXTYPE", PaymentAction '// S = Sale transaction, A = Authorisation, C = Credit, D = Delayed Capture, V = Void 
			AddNVP "TENDER", "P" '// C = credit card, P = PayPal
			AddNVP "ACTION", "D" '// S = Set, G = Get, D = Do
			AddNVP "TOKEN", Token
			AddNVP "PAYERID", PayerID
			AddNVP "AMT", pcf_CurrencyField(money(paymentAmount))		
		End If
		
		If methodName="DOCapture" Then
			AddNVP "TRXTYPE", "D" '// S = Sale transaction, A = Authorisation, C = Credit, D = Delayed Capture, V = Void 
			AddNVP "TENDER", "P" '// C = credit card, P = PayPal
			AddNVP "ORIGID", pidAuthOrder
		End If
		Set Session("nvpReqArray") = deformatNVP(nvpStr)
				
		API_ENDPOINT = GetPayPalURL(pcPay_PayPal_Method)
		objPayPalHttp.open "POST", API_ENDPOINT, False

		objPayPalHttp.setOption(2) = (objPayPalHttp.getOption(2) - SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID)
		objPayPalHttp.setRequestHeader "Content-Type", API_HEADER
		
		'// PayPal Protocol Headers 
		objPayPalHttp.setRequestHeader "Content-Length", "233"
		objPayPalHttp.setRequestHeader "Host", API_ENDPOINT
		objPayPalHttp.setRequestHeader "Connection", "close"
		objPayPalHttp.setRequestHeader "X-VPS-Timeout", "30"		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-OS-Name", "Windows"  '// Name of your Operating System (OS)		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-OS-Version", "XPSP2"  '// OS Version		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-Client-Type", "ASP/MSXML"  '// Language you are using		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-Client-Version", "3.5"  '// For your info		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-Client-Architecture", "x86"  '// For your info		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-Client-Certification-Id", "44baf5893fc2123d8b191d2d011b7fdf" '// This header requirement will be removed		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-Integration-Product", "ProductCart"  '// For your info, would populate with application name		
		objPayPalHttp.setRequestHeader "X-VPS-VIT-Integration-Version", "0.01" '// Application version		
		objPayPalHttp.setRequestHeader "X-VPS-Request-ID", createGuid()	

		'//PAYPAL LOGGING START
		If scPPLogging = "1" Then
			OutputFile.WriteBlankLines(1)
			OutputFile.WriteLine nvpstr
			OutputFile.WriteBlankLines(1)
		End If
		'//PAYPAL LOGGING END
		
		objPayPalHttp.Send nvpStr
		
		If Err.Number <> 0 Then 			
			DeclinedString = DeclinedString & ErrorFormatter(Err.Number, Err.Description, "hash_call")
			Session("nvpReqArray") =  Null
		End If
		
		Set Session("nvpReqArray") = deformatNVP(nvpStr)
		Set  nvpResponseCollection = deformatNVP(objPayPalHttp.responseText)
		Set  hash_call = nvpResponseCollection
		
		If Err.Number <> 0 Then 			
			DeclinedString = DeclinedString & ErrorFormatter(Err.Number, Err.Description, "hash_call")
			Session("nvpReqArray") =  Null
		End If
			
	End Function
	
	

	'----------------------------------------------------------------------------------
	' Purpose: Creates a unique id
	' Inputs:  none
	' Returns: GUID
	'----------------------------------------------------------------------------------
	Function createGuid()
		On Error Resume Next
		Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
		tg = TypeLib.Guid
		createGuid = left(tg, len(tg)-2)
		createGuid = replace(createGuid,"-","")
		createGuid = replace(createGuid,"{","")
		createGuid = replace(createGuid,"}","")
		Set TypeLib = Nothing
	End Function
	
	

	'----------------------------------------------------------------------------------
	' Purpose: Append a new name value pair to the NVP string.
	' Inputs:  Name and Value
	' Returns: Properly Formatted String
	'----------------------------------------------------------------------------------
	Public Sub AddNVP(pName, pValue)
		On Error Resume Next		
		nvpstr = nvpstr & "&" & pName & "=" & pValue	
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "AddNVP")
		End If	
	End Sub
	
	

	'----------------------------------------------------------------------------------
	' Purpose: It generates a random number.
	' Inputs:  The limit, or the highest possible random number.
	' Returns: A random number between 1 and the limit.
	'----------------------------------------------------------------------------------
	Public Function randomNumber(limit)
		randomize
		randomNumber=int(rnd*limit)+2
	End Function


	'----------------------------------------------------------------------------------
	' Purpose: It gives out decoded url path to dispaly.
	' Inputs:  Url string to be decoded.
	' Returns: Decoded Url string.
	'----------------------------------------------------------------------------------
	Function URLDecode(str) 
		On Error Resume Next
		
		str = Replace(str, "+", " ")		
		For i = 1 To Len(str) 
		sT = Mid(str, i, 1) 
			If sT = "%" Then 		
 					sR = sR & Chr(CLng("&H" & Mid(str, i+1, 2))) 
					i = i+2 	
			Else 
				sR = sR & sT 
			End If 
		Next 				   
		URLDecode = sR 
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "URLDecode")
		End If
		
	End Function




	'----------------------------------------------------------------------------------
	' Purpose: It's Workaround Method for Response.Redirect
	'          It will redirect the page to the specified url without urlencoding
	' Inputs: Url to redirect the page
	'----------------------------------------------------------------------------------
	Function ReDirectURL(url)	
		On Error Resume Next
			
		response.clear
		response.status="302 Object moved"
		response.AddHeader "location",url
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "ReDirectURL")
		End If
		
	End Function
	'----------------------------------------------------------------------------------
	' Purpose: It will Format error Messages into a HTML string.
	' Inputs:  NVP string.
	' Returns: NVP Collection object deformated from NVP string.
	'----------------------------------------------------------------------------------
	Function ErrorFormatter(errNumber, errDesc , errlocation)
		
		ErrorFormatter = "<div align=""left"">" & _
							"<ul>" &_
							"<li>" & "<u>Error " & errNumber & "</u></li>" &_
							"<li>" & "Error Description: " & errDesc & "</li>"
							if pcPay_PayPal_Method = "sandbox" then
		ErrorFormatter = ErrorFormatter & "<li>" & "Error Location: " & errlocation & "</li>"
							end if
		ErrorFormatter = ErrorFormatter & "</ul></div>"
		
		If Err.Number <> 0 Then
			Err.Clear
		End If
	End Function 



	'----------------------------------------------------------------------------------
	' Purpose: Append Our HTML error strings into one report.
	' Inputs:  pcv_PayPalErrMessage, DeclinedString.
	' Returns: pcv_PayPalErrMessage + DeclinedString as one formatted string.
	'----------------------------------------------------------------------------------
	Public Sub GenerateErrorReport()
		On Error Resume Next
		
		pErrNumber = resArray("RESULT")
		pErrDescription = resArray("RESPMSG")
		
		If pErrDescription <> "" Then			
			pcv_PayPalErrMessage = pcv_PayPalErrMessage & objPayPalClass.ErrorFormatter(pErrNumber, pErrDescription, "PayPal Service")
		End If
		
		if DeclinedString<>"" then
			pcv_PayPalErrMessage=pcv_PayPalErrMessage & "<hr/><div>API Errors</div><hr/>"		
			pcv_PayPalErrMessage=pcv_PayPalErrMessage & "<div>" & DeclinedString & "</div>"
			pcv_PayPalErrMessage=pcv_PayPalErrMessage & "<hr/>"
		end if		
	End Sub
	
	
	
	'----------------------------------------------------------------------------------
	' Purpose: It gives url path for the cancel & return  page.
	' Returns: Url path of current page without file name.
	'----------------------------------------------------------------------------------
	Public Function GetURL() 
		On Error Resume Next		
		
		if scSSL = "1" then
			Virtual_Path = scSslURL &"/"& scPcFolder & "/pc/"
		else
			Virtual_Path = scStoreURL &"/"& scPcFolder & "/pc/"
		end if
		
		GetURL = Virtual_Path
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "GetURL")
		End If
		
	End Function


	'----------------------------------------------------------------------------------
	' Purpose: It gives url to the PayPal server.
	' Inputs:  PayPal method "sandbox" or "live"
	' Returns: Sandbox or Live Server URL
	'----------------------------------------------------------------------------------
	Public Function GetPayPalURL(pcPay_PayPal_Method)		
		if pcPay_PayPal_Method = "sandbox" then
			GetPayPalURL = "https://pilot-payflowpro.paypal.com" 
		else
			GetPayPalURL = "https://payflowpro.paypal.com"
		end if
	End Function
	
	Public Function GetECURL(pcPay_PayPal_Method)		
		if pcPay_PayPal_Method = "sandbox" then
			GetECURL = "https://pilot-payflowpro.paypal.com/"
		else
			GetECURL = "https://payflowpro.paypal.com"
		end if
	End Function


	'----------------------------------------------------------------------------------
	' Purpose: It will convert nvp string to Collection object.
	' Inputs:  NVP string.
	' Returns: NVP Collection object deformated from NVP string.
	'----------------------------------------------------------------------------------
	Public Function deformatNVP(nvpstr)
		On Error Resume Next
		
		Dim AndSplitedArray, EqualtoSplitedArray, Index1, Index2, NextIndex
		Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
		AndSplitedArray = Split(nvpstr, "&", -1, 1)
		NextIndex=0
		For Index1 = 0 To UBound(AndSplitedArray)
			EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
			For Index2 = 0 To UBound(EqualtoSplitedArray)
				NextIndex=Index2+1
				NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
				Index2=Index2+1
			Next
		Next
		Set deformatNVP = NvpCollection
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "deformatNVP")
		End If
		
	End Function
	
	
	'----------------------------------------------------------------------------------
	' Purpose: Provides a clean way to set all the PayPal variables to local.
	' Inputs:  None. Requires an open database connection
	' Returns: pcPay_PayPal_TransType, PaymentAction, pcPay_PayPal_Username, pcPay_PayPal_Password, pcPay_PayPal_Sandbox, pcPay_PayPal_Method, pcPay_PayPal_Signature
	'----------------------------------------------------------------------------------	
	Public Sub pcs_SetAllVariables()
		'On Error Resume Next
		
		'// Query PayPal Table
		query="SELECT pcPay_PayPalAd_Partner, pcPay_PayPalAd_MerchantLogin, pcPay_PayPalAd_Vendor, pcPay_PayPalAd_User, pcPay_PayPalAd_Password, pcPay_PayPalAd_TransType, pcPay_PayPalAd_CSC, pcPay_PayPalAd_Sandbox FROM pcPay_PayPalAdvanced WHERE pcPay_PayPalAd_ID=1;"
		set rsPayPalVar=server.CreateObject("ADODB.RecordSet")
		set rsPayPalVar=conntemp.execute(query)
	
		'// Set Local Var
		pcPay_PayPalAd_Partner=trim(rsPayPalVar("pcPay_PayPalAd_Partner"))
		pcPay_PayPalAd_MerchantLogin=trim(rsPayPalVar("pcPay_PayPalAd_MerchantLogin"))
		pcPay_PayPalAd_MerchantLogin=enDeCrypt(pcPay_PayPalAd_MerchantLogin, scCrypPass)
		pcPay_PayPalAd_User=trim(rsPayPalVar("pcPay_PayPalAd_User"))
		pcPay_PayPalAd_User=enDeCrypt(pcPay_PayPalAd_User, scCrypPass)
		pcPay_PayPalAd_Password = trim(rsPayPalVar("pcPay_PayPalAd_Password"))
		pcPay_PayPalAd_Password=enDeCrypt(pcPay_PayPalAd_Password, scCrypPass)
		pcPay_PayPalAd_TransType = trim(rsPayPalVar("pcPay_PayPalAd_TransType"))
		pcPay_PayPalAd_CSC = trim(rsPayPalVar("pcPay_PayPalAd_CSC"))
		pcPay_PayPalAd_Sandbox = trim(rsPayPalVar("pcPay_PayPalAd_Sandbox"))
		
		' Check pcPay_PayPal_Currency for NULL
		pcPay_PayPal_Currency="USD"
		
		' Check pcPay_PayPal_CVV for NULL
		if pcPay_PayPalAd_CSC = "YES" then
			pcPay_PayPalAd_CSC = 1
		end if
		if pcPay_PayPalAd_CSC = "N0" then
			pcPay_PayPalAd_CSC = 0
		end if
		if isNULL(pcPay_PayPalAd_CSC)=True or pcPay_PayPalAd_CSC="" then
			pcPay_PayPalAd_CSC=1
		end if
		
		' Check pcPay_PayPal_CardTypes for NULL
		pcPay_PayPal_CardTypes="V, M, D"
		
		' Authorize or Capture
		if pcPay_PayPalAd_TransType="S" then
			PaymentAction="S"	
		else
			PaymentAction="A"
		end if
		
		' Sandbox or Live
		if pcPay_PayPalAd_Sandbox="YES" then
			pcPay_PayPal_Method = "sandbox"
		else
			pcPay_PayPal_Method = "live"
		end if
		
		'// Close our Db connections
		set rsPayPalVar=nothing
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "pcs_SetAllVariables")
		End If
		
	End Sub
	
	
	'----------------------------------------------------------------------------------
	' Purpose: Provides a clean way to obtain the latest Address.
	'----------------------------------------------------------------------------------	
	Public Sub pcs_SetShipAddress(OrderID)
		On Error Resume Next
		
		'// Query PayPal Table
		query="SELECT ShippingFullName, shippingCompany, shippingAddress, shippingZip, shippingStateCode, shippingState, pcOrd_shippingPhone, shippingCity, shippingCountryCode, shippingAddress2 FROM orders WHERE idorder="&OrderID&";"
		set rsPayPalVar=server.CreateObject("ADODB.RecordSet")
		set rsPayPalVar=conntemp.execute(query)

		'// Set Local Var
		pcv_strShippingFullName=rsPayPalVar("ShippingFullName")	
		pcv_strShippingCompany=rsPayPalVar("shippingCompany")
		pcv_strShippingAddress=rsPayPalVar("shippingAddress")
		pcv_strShippingPostalCode=rsPayPalVar("shippingZip")
		pcv_strShippingStateCode=rsPayPalVar("shippingStateCode")
		pcv_strShippingProvince=rsPayPalVar("shippingState")
		pcv_strShippingPhone=rsPayPalVar("pcOrd_shippingPhone")
		pcv_strShippingCity=rsPayPalVar("shippingCity")
		pcv_strShippingCountryCode=rsPayPalVar("shippingCountryCode")
		pcv_strShippingAddress2=rsPayPalVar("shippingAddress2")						
		
		'// Close our Db connections
		set rsPayPalVar=nothing
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "pcs_SetShipAddress")
		End If
		
	End Sub


End Class
'///////////////////////////////////////////////////////////////////////////////////
'// END: WEBSITE PAYMENTS PRO - United Kingdom - Payflo Edition - 2.0
'///////////////////////////////////////////////////////////////////////////////////


'// Format For Field
Public Function pcf_CurrencyField(moneyAMT)	
	if scDecSign = "," then
		moneyAMT=replace(moneyAMT,".","")
		moneyAMT=replace(moneyAMT,",",".")		
	else
		moneyAMT=replace(moneyAMT,",","")
	end if
	pcf_CurrencyField=moneyAMT
End Function
%>