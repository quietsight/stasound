<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#INCLUDE file="pcPay_GoogleCheckout_base64.asp"--> 
<!--#include file="../includes/GoogleCheckoutConstants.asp"--> 
<!--#include file="../includes/rc4_GoogleCheckout.asp"--> 
<%
Dim strMerchantId
Dim strMerchantKey
Dim strMsxmlDomDocument
Dim strXmlHttp
Dim strXmlns
Dim strXmlVersionEncoding
Dim attrCurrency
Dim logFilename
Dim baseUrl
Dim checkoutUrl
Dim checkoutDiagnoseUrl
Dim requestUrl
Dim requestDiagnoseUrl
Dim errorReportType
Dim BreakPointMode
Dim LogMessagesMode
Dim DeclaredValueMode
	
setGlobalVariables

'// Set Globals
Function setGlobalVariables
	pcv_tmpLogFile=server.MapPath("pcPay_GoogleCheckout_Global.asp")
	pcv_tmpLogFile=left(pcv_tmpLogFile,instr(pcv_tmpLogFile,"\pc\"))
	pcv_tmpLogFile=pcv_tmpLogFile&"includes/GoogleCheckoutLog.out"
    logFilename = pcv_tmpLogFile
	errorReportType = 1
	BreakPointMode = GOOGLELOGGING
	LogMessagesMode = "0"
	
	'******************************************************************
	'// UPS Insured Value
	'******************************************************************
	'// To change this value from the default insured value of 100.00 
	'// you will need to change the variable below to the value of 0.
	'//
	'// For Example: 
	'// DeclaredValueMode = "1"	
	'******************************************************************
	DeclaredValueMode = "1"
	'******************************************************************
	
	'******************************************************************
	'// U.S.P.S. OPTIONAL VARIABLES
	'******************************************************************
	'// USPS Value of Content for International Rates Only
	'// If specified, it is used to compute Insurance fee 
	'// (if insurance is available for service and destination) and
	'// indemnity coverage. 
	'// To turn this variable on, change the value to "1" 
	'//
	'// For Example: 
	'// pcv_UseValueOfContents=1
	'******************************************************************
	pcv_UseValueOfContents=1
	'******************************************************************
	
	attrCurrency = GOOGLECURRENCY
    strXmlns = "http://checkout.google.com/schema/2"
    strMerchantId = getMerchantId
    strMerchantKey = getMerchantKey    
	if GOOGLETESTMODE="YES" then
    	baseUrl = "https://sandbox.google.com/checkout/api/checkout/v2/"
	else
		baseUrl = "https://checkout.google.com/api/checkout/v2/"
	end if
    checkoutUrl = baseUrl & "merchantCheckout/Merchant/" & strMerchantId
	requestUrl = baseUrl & "request/Merchant/" & strMerchantId
    checkoutDiagnoseUrl = baseUrl & "/diagnose"
	if scXML="" then
		tmpscXML=".3.0"
	else
		tmpscXML=scXML
	end if
    strMsxmlDomDocument = "Msxml2.DOMDocument"&tmpscXML
	strXmlHttp = "Msxml2.serverXmlHttp"&tmpscXML
    strXmlVersionEncoding = "version=""1.0"" encoding=""UTF-8"""	
End Function

'// Get Merchant ID
Function getMerchantId()
	if GOOGLETESTMODE="YES" then
    	getMerchantId = GOOGLESANDBOXID
	else
    	getMerchantId = GOOGLEMERCHANTID
	end if
End Function

'// Get Merchant Key
Function getMerchantKey()
    if GOOGLETESTMODE="YES" then
		getMerchantKey = GDeCrypt(GOOGLESANDBOXKEY, scCrypPass)
	else
    	getMerchantKey = GDeCrypt(GOOGLEMERCHANTKEY, scCrypPass)
	end if
End Function


'// Send Request and Verify Data
Function sendRequest(request, strPostUrl)

    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "sendRequest()"

    ' Check for missing parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "request", request
    checkForError errorType, strFunctionName, "strPostUrl", strPostUrl
    checkForError errorType, strFunctionName, "strMerchantId", strMerchantId
    checkForError errorType, strFunctionName, "strMerchantKey", strMerchantKey

    ' Define objects used to send the HTTP request
    Dim xmlHttp
    Dim strAuthentication 
    Dim bRequest

    ' Log the outgoing message
    logMessage logFilename, request

    ' Create the XMLHttpRequest object
    Set xmlHttp = Server.CreateObject(strXmlHttp)

    ' The HTTP request method is POST
    xmlHttp.open "POST", strPostUrl, False

    ' Do NOT ignore Server SSL Cert Errors
    Const SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2
    Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
    xmlHttp.setOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, _
        (xmlHttp.getOption(SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS) - _
        SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)

    bRequest = InStr(strPostUrl, "request")

	' Build HTTP Basic Authentication scheme
	strAuthentication = createHttpBasicAuthentication(strMerchantId, strMerchantKey)

	' Set HTTP headers
	xmlHttp.SetRequestHeader "Authorization", strAuthentication
	xmlHttp.SetRequestHeader "Content-Type", "application/xml"
	xmlHttp.SetRequestHeader "Accept", "application/xml"

    ' Transmit the request
    xmlHttp.send request

    ' Log the HTTP response
    logMessage logFilename, xmlHttp.responseText

    ' Return the response from the Google server
    sendRequest = xmlHttp.responseText

    ' Release the object used to send the request
    Set xmlHttp = Nothing

End Function



'// The createHttpBasicAuthentication creates a string in the format
Function createHttpBasicAuthentication(strMerchantId, strMerchantKey)

    Dim strCredential
    Dim b64credential
    Dim strAuthentication

    ' Create "userid:password" 
    strCredential = strMerchantId & ":" & strMerchantKey
	b64credential = Base64_Encode(strCredential)
    strAuthentication = "Basic " & b64credential

    ' Return the HTTP Basic Authentication string
    createHttpBasicAuthentication = strAuthentication

End Function



'// The displayDiagnoseResponse function is a debugging function that
Function displayDiagnoseResponse(request, strPostUrl, xml, action)

    ' Define objects used to diagnose the API response
    Dim diagnoseResponse
    Dim bValidated
    Dim domResponse
    Dim strRootTag
    Dim nodeList
    Dim strResult

    ' Execute the API request and capture the Google Checkout server's response
    diagnoseResponse = sendRequest(request, strPostUrl)

    ' If the function finds that the request contained valid XML, the
    ' $validated variable will be set to true
    bValidated = false

    Set domResponse = Server.CreateObject(strMsxmlDomDocument)
    domResponse.loadXml diagnoseResponse

    strRootTag = domResponse.documentElement.tagName

    ' This if-else block determines whether the API response indicates
    ' that the response contained invalid XML or if there was some other
    ' problem associated with the request, such as an invalid signature.
    If strRootTag = "diagnosis" Then
        Set nodeList = _
            domResponse.documentElement.getElementsByTagName("string")
        If nodeList.length > 0 Then
            strResult = nodeList(0).text
        Else
            bValidated = True
        End If
    Elseif strRootTag = "error" Then
        Set nodeList = _
            domResponse.documentElement.getElementsByTagName("error-message")
        strResult = nodeList(0).text
    ElseIf strRootTag = "request-received" Then
        bValidated = true
    End If

    ' If the request is invalid, print the reason that the request is
    ' invalid if the errorReportType variable indicates that errors
    ' should be displayed in the user's browser. Also display a link 
    ' to a tool where the user can edit the XML request unless the
    ' validation request was submitted from that tool.
    If bValidated = False And (errorReportType = 2 Or errorReportType = 3) Then
        Response.write "<tr><td style=""color:red""><p>" & _
            "<span style=""text-align:center""><h2>" & _
            "This XML is NOT Validated!</h2></span></p>"
        Response.write "<p style=""text-align:left""><b>" & _
            Server.HTMLEncode(strResult) & "</b></p>"
        If action = "debug" Then
            Response.write "<p><form method=POST action=DebuggingTool.asp>"
            Response.write "<input type=""hidden"" name=""xml"" value=""" & _
                Server.HTMLEncode(xml) & """/>"
            Response.write "<input type=""hidden"" name=""toolType"" " & _
                "value=""Validate XML""/>"
            Response.write "<input type=""submit"" name=""Debug"" " & _
                "value=""Debug XML""/>"
            Response.write "</form></p></td></tr>"
        End If
    End If

    ' Return a Boolean value indicating whether the request
    ' contained valid XML.
    displayDiagnoseResponse = bValidated
End Function



'// The CheckForError function determines whether a parameter has a null
Function checkForError(errorType, strFunctionName, strParamName, strParamValue)
    If strParamValue = "" Then
        errorHandler errorType, strFunctionName, strParamName, strParamValue
    End If
End Function


'// The errorHandler function returns the error message that should be logged
Function errorHandler(errorType, errorFunctionName, errorParamName, errorParamValue) 

    Select Case errorType 

        ' MISSING_PARAM error
        ' A function call omits a required parameter.
        Case "MISSING_PARAM"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: """ & errorParamName _
                & """ must be provided."

        ' MISSING_PARAM_NONE error
        ' A function call must have a value for at least one parameter.
        Case "MISSING_PARAM_NONE"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: " _
                & "At least one parameter should be provided."

        ' INVALID_INPUT_ARRAY error
        ' AddAreas() function called with invalid value for
        ' $state_areas or $zip_areas parameter
        Case "INVALID_INPUT_ARRAY"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Invalid Input: """ & errorParamName _
                & """ should be an array."

        ' MISSING_CURRENCY error
        ' The attrCurrency value is empty.
        Case "MISSING_CURRENCY"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: ""attrCurrency"" " _
                & "should be set when the ""elemAmount"" is set."

        ' MISSING_TRACKING error
        ' The ChangeShippingInfo() function in
        ' OrderProcessingAPIFunctions.asp is being called without
        ' specifying a tracking number even though a shipping
        ' carrier is specified.
        Case "MISSING_TRACKING"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: ""elemTrackingNumber"" " _
                & "should be set when the ""elemCarrier"" is set."

        Case Else

    End Select 

    ' Print the error message to the screen
    If (errorReportType = 2) Or (errorReportType = 3) Then 

        Dim errstrHtml
        errstrHtml = errstr & "<br><br>"

        Response.write errstrHtml

    End If

    ' Write out the error message to the IIS Log File
    If (errorReportType = 1) Or (errorReportType = 3) Then 

        Response.appendToLog errstr

    End If

    Exit Function

End Function



'//  The logMessage function logs a message to a local file. 
Function logMessage(logFilename, message)
	Dim oFs
	Dim oTextFile
	If LogMessagesMode = "1" Then
		' Print out the notification message to a local file
		Set oFs = Server.createobject("Scripting.FileSystemObject")
		Const ioMode = 8
		Set oTextFile = oFs.openTextFile(logFilename, ioMode, True)
		oTextFile.writeLine now
		oTextFile.writeLine message
		oTextFile.close
		' Free object
		Set oTextFile = Nothing
		Set oFS = Nothing
	End If
End Function


'//  Set Break Points for Debugging the Web Service 
Function BreakPoint(logFilename, logDescription, message, logErr)
	Dim oFs
	Dim oTextFile
	If (BreakPointMode = "3" AND logErr<>"") OR BreakPointMode = "1" Then		
			if message = "" then
				message = "none"
			end if
			if logErr = "" then
				logErr = "no errors"
			end if
			' Print out the notification message to a local file
			Set oFs = Server.createobject("Scripting.FileSystemObject")
			Const ioMode = 8
			Set oTextFile = oFs.openTextFile(logFilename, ioMode, True)
			oTextFile.writeLine now
			oTextFile.writeLine logDescription
			oTextFile.writeLine "Variable Output: " & message
			if err.number>0 then
				oTextFile.writeLine "Error Report: " & logErr		
			end if
			oTextFile.writeLine " "
			oTextFile.close
			' Clear errors
			err.clear
			err.number=0
			' Free object
			Set oTextFile = Nothing
			Set oFS = Nothing		
	End If
End Function



'//  Set Break Points for Debugging the Web Service 
Function TrackBug(logFilename, message)
	Dim oFs
	Dim oTextFile
	if message = "" then
		message = "no error"
	end if
	' Print out the notification message to a local file
	Set oFs = Server.createobject("Scripting.FileSystemObject")
	Const ioMode = 8
	Set oTextFile = oFs.openTextFile(logFilename, ioMode, True)
	oTextFile.writeLine "Tracking Report: " & message
	oTextFile.close
	' Clear errors
	err.clear
	err.number=0
	' Free object
	Set oTextFile = Nothing
	Set oFS = Nothing
End Function


' Save Cart
Public Sub pcs_SaveCartArrayToDB

	for f=1 to pcCartIndex
		query="INSERT INTO pcCartArray (pcCartArray_Key, pcCartArray_Date, pcCartArray_0, pcCartArray_1, pcCartArray_2, pcCartArray_3, pcCartArray_4, pcCartArray_5, pcCartArray_6, pcCartArray_7, pcCartArray_8, pcCartArray_9, pcCartArray_10, pcCartArray_11, pcCartArray_12, pcCartArray_13, pcCartArray_14, pcCartArray_15, pcCartArray_16, pcCartArray_17, pcCartArray_18, pcCartArray_19, pcCartArray_20, pcCartArray_21, pcCartArray_22, pcCartArray_23, pcCartArray_24, pcCartArray_25, pcCartArray_26,pcCartArray_27, pcCartArray_28, pcCartArray_29,pcCartArray_30, pcCartArray_31, pcCartArray_32,pcCartArray_33, pcCartArray_34, pcCartArray_35, pcCartArray_36, pcCartArray_37, pcCartArray_38, pcCartArray_39, pcCartArray_40, pcCartArray_41, pcCartArray_42, pcCartArray_43, pcCartArray_44, pcCartArray_45) "
		if pcCartArray(f,0)="" OR isNULL(pcCartArray(f,0))=True then
			pcCartArray(f,0)=0
		end if
		if pcCartArray(f,2)="" OR isNULL(pcCartArray(f,2))=True then
			pcCartArray(f,2)=0
		end if
		if pcCartArray(f,5)="" OR isNULL(pcCartArray(f,5))=True then
			pcCartArray(f,5)=0
		end if
		if pcCartArray(f,6)="" OR isNULL(pcCartArray(f,6))=True then
			pcCartArray(f,6)=0
		end if
		if pcCartArray(f,8)="" OR isNULL(pcCartArray(f,8))=True then
			pcCartArray(f,8)=0
		end if
		if pcCartArray(f,12)="" OR isNULL(pcCartArray(f,12))=True then
			pcCartArray(f,12)=0
		end if
		if pcCartArray(f,13)="" OR isNULL(pcCartArray(f,13))=True then
			pcCartArray(f,13)=0
		end if
		if pcCartArray(f,14)="" OR isNULL(pcCartArray(f,14))=True then
			pcCartArray(f,14)=0
		end if
		if pcCartArray(f,15)="" OR isNULL(pcCartArray(f,15))=True then
			pcCartArray(f,15)=0
		end if
		if pcCartArray(f,17)="" OR isNULL(pcCartArray(f,17))=True then
			pcCartArray(f,17)=0
		end if
		if pcCartArray(f,19)="" OR isNULL(pcCartArray(f,19))=True then
			pcCartArray(f,19)=0
		end if
		if pcCartArray(f,20)="" OR isNULL(pcCartArray(f,20))=True then
			pcCartArray(f,20)=0
		end if
		if pcCartArray(f,22)="" OR isNULL(pcCartArray(f,22))=True then
			pcCartArray(f,22)=0
		end if
		query=query& "VALUES ("& pcv_intGoogleRandomKey &", "& queryDate &","
		query=query& ""& fixstring(pcCartArray(f,0)) &", "
		query=query& "'"& fixstring(pcCartArray(f,1)) &"', "
		query=query& ""& pcCartArray(f,2) &", "
		query=query& ""& pcCartArray(f,3) &", "
		query=query& "'"& fixstring(pcCartArray(f,4)) &"', "
		query=query& ""& pcCartArray(f,5) &", "
		query=query& ""& pcCartArray(f,6) &", "
		query=query& "'"& fixstring(pcCartArray(f,7)) &"', "
		query=query& ""& pcCartArray(f,8) &", "
		query=query& "'"& fixstring(pcCartArray(f,9)) &"', "
		query=query& ""& pcCartArray(f,10) &", "
		query=query& "'"& fixstring(pcCartArray(f,11)) &"', "
		query=query& ""& pcCartArray(f,12) &", "
		query=query& ""& pcCartArray(f,13) &", "
		query=query& ""& pcCartArray(f,14) &", "
		query=query& ""& pcCartArray(f,15) &", "
		query=query& "'"& fixstring(pcCartArray(f,16)) &"', "
		query=query& ""& pcCartArray(f,17) &", "
		query=query& ""& pcCartArray(f,18) &", "
		query=query& ""& pcCartArray(f,19) &", "
		query=query& ""& pcCartArray(f,20) &", "
		query=query& "'"& fixstring(pcCartArray(f,21)) &"', "
		query=query& ""& pcCartArray(f,22) &", "
		query=query& "'"& pcCartArray(f,23) &"', "
		query=query& "'"& fixstring(pcCartArray(f,24)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,25)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,26)) &"', "
		query=query& ""& pcCartArray(f,27) &", "
		query=query& ""& pcCartArray(f,28) &", "
		query=query& "'"& pcCartArray(f,29) &"', "
		query=query& "'"& pcCartArray(f,30) &"', "
		query=query& "'"& pcCartArray(f,31) &"', "
		query=query& "'"& pcCartArray(f,32) &"', "
		query=query& "'"& pcCartArray(f,33) &"', "
		query=query& "'"& pcCartArray(f,34) &"', "
		query=query& "'"& pcCartArray(f,35) &"', "
		query=query& "'"& pcCartArray(f,36) &"', "
		query=query& "'"& pcCartArray(f,37) &"', "
		query=query& "'"& pcCartArray(f,38) &"', "
		query=query& "'"& pcCartArray(f,39) &"', "
		query=query& "'"& pcCartArray(f,40) &"', "
		query=query& "'"& pcCartArray(f,41) &"', "
		query=query& "'"& pcCartArray(f,42) &"', "
		query=query& "'"& pcCartArray(f,43) &"', "
		query=query& "'"& pcCartArray(f,44) &"', "
		query=query& "'"& pcCartArray(f,45) &"');"		
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if			
		set rs=nothing
	next
	
End Sub

Public Sub pcs_RestoreCartArray		
	query="SELECT *  FROM pcCartArray  WHERE pcCartArray_Key="& pcStrCustomerRefKey &""
	set rs2=server.CreateObject("ADODB.RecordSet")
	set rs2=conntemp.execute(query)		
	f=1
	Do while NOT rs2.eof
		for x=0 to 45
			pcCartArray(f,x)=rs2("pcCartArray_"&x)
		next
	f=f+1
	rs2.movenext
	loop	
	set rs2=nothing
End Sub

Private Function pcf_ValidateKey(pcStrCustomerRefKey)
	
	'// Optimize Performance/ Purge 30 Day Records
	if scDB="SQL" then
		strDtDelim="'"
	else
		strDtDelim="#"
	end if
	query="DELETE FROM pcCartArray WHERE pcCartArray_Date<"&strDtDelim&dateadd("d",-30,Date())&strDtDelim&" ;"	
	set rsOptimize=server.CreateObject("ADODB.RecordSet")
	set rsOptimize=conntemp.execute(query)	
	set rsOptimize=nothing
	
	'// Validate the Key
	Dim pcv_intOrderKey, pcv_strcurrentKey, pcv_intcurrentKeyUnique, pcv_strUniqueKey	
	pcv_strcurrentKey=pcStrCustomerRefKey
	pcv_intcurrentKeyUnique=0	
	do while pcv_intcurrentKeyUnique<1
		
		'// Check the pcCartArray Table
		query="SELECT pcCartArray.pcCartArray_ID FROM pcCartArray WHERE pcCartArray_Key="& pcv_strcurrentKey &""
		set rsValidateKey=server.CreateObject("ADODB.RecordSet")
		set rsValidateKey=conntemp.execute(query)		
		If NOT rsValidateKey.eof Then
			pcv_intcurrentKeyUnique=0
			pcv_strcurrentKey=randomNumber(99999999)		
		Else
			pcv_intcurrentKeyUnique=1
			pcv_strUniqueKey=pcv_strcurrentKey			
		End If
		set rsValidateKey=nothing
		
		'// Check the Orders Table
		query="SELECT orders.randomNumber FROM orders WHERE randomNumber="& pcv_strcurrentKey &""
		set rsValidateKey=server.CreateObject("ADODB.RecordSet")
		set rsValidateKey=conntemp.execute(query)		
		If NOT rsValidateKey.eof Then
			pcv_intcurrentKeyUnique=0
			pcv_strcurrentKey=randomNumber(99999999)		
		Else
			pcv_intcurrentKeyUnique=1
			pcv_strUniqueKey=pcv_strcurrentKey			
		End If
		set rsValidateKey=nothing
			
	loop	
	
	pcf_ValidateKey=pcv_strUniqueKey	
	
End Function

'// UK Support
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