<!-- #include file="Centinel_Hash.js" -->
<%
'======================================================================================
'=  Cardinal Commerce (http://www.cardinalcommerce.com)
'= 
'=  The CentinelClient class is defined to assist integration efforts with the Centinel
'=  XML message integration. The class implements helper methods to construct, send, and
'=  receive XML messages with respect to the Centinel XML Message APIs.
'======================================================================================
%>
<%

Function sendMsg(ccRequest, transactionUrl, resolveTimeout, connectTimeout, sendTimeout, receiveTimeout)

    Dim fields
    Dim fieldName
    Dim xmlRequest
    Dim xmlResponse
    
    'Build Request
    xmlRequest = "<CardinalMPI>"
    fields = ccRequest.Keys
    For i = 0 To UBound(fields)
        fieldName = fields(i)
        xmlRequest = xmlRequest & generateXMLTag(fieldName, ccRequest.item(fieldName))    
    Next
    xmlRequest = xmlRequest & "<Source>ASPTC</Source><SourceVersion>1.1.0</SourceVersion></CardinalMPI>"    
    'Send Request
	'response.write xmlRequest
	'response.end
	On Error Resume Next	

	Dim requestMessage, objXMLSender

	' Timeout Values represented as Milliseconds, receiveTimeout is the most important
	' The receiveTimeout will allow the checkout process to continue in the event a timely response
	' is not received from CardinalCommerce
	
	requestMessage = "cmpi_msg=" & Server.URLEncode(xmlRequest)

    Set objXMLSender = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")

	objXMLSender.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout

    objXMLSender.open "POST", transactionUrl, False
        
	objXMLSender.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXMLSender.send requestMessage
	
	If Err.Number <> 0 Then
		
		' Check for specific Error Codes, if you required enhanced error handling

		If Err.Number = "5050" Then
		    xmlResponse = "<CardinalMPI><ErrorNo>X100</ErrorNo><ErrorDesc>Communication Timeout</ErrorDesc></CardinalMPI>"
		ElseIf Err.Number = "5051" Then
		    xmlResponse = "<CardinalMPI><ErrorNo>X110</ErrorNo><ErrorDesc>The server name or address could not be resolved</ErrorDesc></CardinalMPI>"
        Else 
            xmlResponse = "<CardinalMPI><ErrorNo>5052</ErrorNo><ErrorDesc>Error Sending/Receiving XML Message (General)</ErrorDesc></CardinalMPI>"
		End If
    Else
        xmlResponse = objXMLSender.responseText
	End If

	objXMLSender = nothing

    'Parse Response
    Dim oDOM
    Dim oNodeList
    Dim Item
	Set oDOM = Server.CreateObject("Msxml2.DOMDocument")
	oDOM.async = False 
	oDOM.LoadXML xmlResponse
    Dim dicResp
    Set dicResp = Server.CreateObject("Scripting.Dictionary")
    Set oNodeList = oDom.documentElement.childNodes         
    Dim i
    Dim val
    For i=0 To oNodeList.length - 1
        Set Item = oNodeList.item(i)
        If Item.NodeType = oDom.documentElement.NodeType Then
            dicResp.add CStr(Item.nodeName), Item.Text
        End If
    Next
    Set sendMsg = dicResp
End Function



Function generateXMLTag(strTagName, strValue)
	Dim tmpValue, strReturn
	tmpValue = escapeValue(strValue)
	strReturn = "<" & strTagName & ">" & tmpValue & "</" & strTagName & ">"
	generateXMLTag = strReturn
End Function

Function getUnparsedMessage(ccObject)

    Dim fields, i
    Dim fieldName
    Dim xmlMessage

    'Build Request
    xmlMessage = "<CardinalMPI>"
    fields = ccObject.Keys
    For i = 0 To UBound(fields)
        fieldName = fields(i)
        xmlMessage = xmlMessage & generateXMLTag(fieldName, ccObject.item(fieldName))    
    Next
    xmlMessage = xmlMessage & "</CardinalMPI>"    

    getUnparsedMessage = xmlMessage

End Function



Function generatePayload(ccRequest, transactionPwd)

	dim fields, fieldName, i
	dim payload, hashString, first

	payload = ""

	
	fields = ccRequest.Keys
	For i = 0 To UBound(fields)
		fieldName = fields(i)
		
		if fieldName = "TransactionPwd" then
		'do nothing
		else

			payload = payload & fieldName & "=" & Server.URLEncode(ccRequest.item(fieldName)) & "&"
		
		end if

	Next
	
	hashString = hex_sha1(payload & transactionPwd)
	payload = payload & "Hash=" & hashString

	generatePayload = payload

End Function

Function escapeValue (originalValue)

	dim inLen, returnString, outBuf

	inLen = Len(originalValue)
	outBuf = ""
	returnString = originalValue

	If inLen > 0 Then
	
		Dim current, f, currentInt, i, idx, arrLength, inArrLen
		arrLength = CInt(inLen-1)

		Dim chars()
		ReDim chars(arrLength)
		f = 0

		For i = 1 To inLen
			chars(f) = Mid(originalValue,i,1)
			f=f+1
		Next
		
		For idx = 0 To UBound(chars)
			
			current = chars(idx)
			currentInt = Asc(current)

			If current = "&" Then
				outBuf = outBuf &"&amp;"
			ElseIf current = "<" Then
				outBuf= outBuf & "&lt;"
			ElseIf currentInt > 126 Then
				outBuf = outBuf & "&#" & currentInt & ";"
			Else
				outBuf = outBuf & current
			End If
		Next

			returnString = outBuf
	
	End if
		
		escapeValue = returnString
			

End Function

%>


