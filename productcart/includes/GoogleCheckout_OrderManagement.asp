<%
Select Case pcv_strGoogleMethod


	Case "archive"
		'// This section creates an <archive-order> request
		attrGoogleOrderNumber = pcv_strGoogleIDOrder
		xmlRequest = createArchiveOrder(attrGoogleOrderNumber)
		
		' Validate Request XML
		'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"
		
		transmitResponse = SendRequest(xmlRequest, requestUrl)
		
		' Process the response
		ProcessXmlData(transmitResponse)
		
		' Archived Status
		os=12
		
		'// Update Fullment Status
		query="UPDATE orders SET orderstatus="& os &" WHERE pcOrd_GoogleIDOrder='"& attrGoogleOrderNumber &"' "
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		set rs=nothing
	
	
	
	
	
	Case "charge"
		'// This section creates a <charge-order> request
		attrGoogleOrderNumber = pcv_strGoogleIDOrder
		
		xmlRequest = createChargeOrder(attrGoogleOrderNumber, elemAmount)
		
		' Validate Request XML
		'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"
		
		transmitResponse = SendRequest(xmlRequest, requestUrl)
		
		' Process the response
		ProcessXmlData(transmitResponse)
		
		
		
		'// This section creates a <process-order> request
		attrGoogleOrderNumber = pcv_strGoogleIDOrder
		xmlRequest = createProcessOrder(attrGoogleOrderNumber)
		
		' Validate Request XML
		'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"
		
		transmitResponse = SendRequest(xmlRequest, requestUrl)
		
		' Process the response
		ProcessXmlData(transmitResponse)
	
	
	
	
	Case "mark" 
	

		'// Only Run for Google Checkout Orders
		if isNULL(pcv_strGoogleIDOrder)=True then
			pcv_strGoogleIDOrder = ""
		end if
		If pcv_strGoogleIDOrder<>"" Then
		
			'// Look up the pcPackageInfo Details for the Carrier and Tracking Number
			query="SELECT pcPackageInfo_ShipMethod, pcPackageInfo_TrackingNumber, pcPackageInfo_ShippedDate, pcPackageInfo_Comments FROM pcPackageInfo WHERE idOrder=" & pIdOrder & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			if NOT rs.eof then
				pcv_tempTrackingNumber=rs("pcPackageInfo_TrackingNumber")
			else
				pcv_tempTrackingNumber="na"
			end if
			
			if isNULL(pcv_tempTrackingNumber)=True OR pcv_tempTrackingNumber="" then
				pcv_tempTrackingNumber="na"
			end if

			if pcv_tempTrackingNumber<>"na" then
				' This section creates an <add-tracking-data> request
				attrGoogleOrderNumber = pcv_strGoogleIDOrder
				if pcv_strShipper<>"" AND isNULL(pcv_strShipper)=False then
					elemCarrier = pcv_strShipper
				else
					elemCarrier = "Other"
				end if
				elemTrackingNumber = pcv_tempTrackingNumber
				xmlRequest = createAddTrackingData(attrGoogleOrderNumber, elemCarrier, elemTrackingNumber)
				
				' Validate Request XML
				'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"

				transmitResponse = SendRequest(xmlRequest, requestUrl)
				
				' Process the response
				ProcessXmlData(transmitResponse)
			else
				elemCarrier = ""
				pcv_tempTrackingNumber = ""
			end if
			
			' This section creates a <deliver-order> request
			attrGoogleOrderNumber = pcv_strGoogleIDOrder
			elemTrackingNumber = pcv_tempTrackingNumber
			elemSendEmail = "true"
			xmlRequest = createDeliverOrder(attrGoogleOrderNumber, elemCarrier, elemTrackingNumber, elemSendEmail)

			' Validate Request XML
			'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"

			transmitResponse = SendRequest(xmlRequest, requestUrl)
			
			' Process the response
			ProcessXmlData(transmitResponse)
			
		End If	


	Case "cancel"
		
		'// This section creates a <cancel-order> request
		attrGoogleOrderNumber = pcv_strGoogleIDOrder
		
		elemReason = pcv_strReason
		elemComment = pcv_strComment
		xmlRequest = createCancelOrder(attrGoogleOrderNumber, elemReason, elemComment)
		
		' Validate Request XML
		'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"
		
		transmitResponse = SendRequest(xmlRequest, requestUrl)
		
		' Process the response
		ProcessXmlData(transmitResponse)
	
	
	
	
	Case "message"
	
		'// This section creates a <send-buyer-message> request
		attrGoogleOrderNumber = pcv_strGoogleIDOrder
		
		elemMessage = pcv_strBuyerMessage
		elemSendEmail = "true"
		xmlRequest = createSendBuyerMessage(attrGoogleOrderNumber, elemMessage, elemSendEmail)
		
		' Validate Request XML
		'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"
		
		transmitResponse = SendRequest(xmlRequest, requestUrl)
		
		' Process the response
		ProcessXmlData(transmitResponse)
	
	
	
	
	Case "refund"
	
		
		'// This section creates a <refund-order> request
		attrGoogleOrderNumber = pcv_strGoogleIDOrder
		
		elemReason = pcv_strRefundReason '// "Buyer requested refund."
		'elemAmount = pcv_strAmount
		elemComment = pcv_strRefundComment '//"Buyer is not happy with the product."
		xmlRequest = createRefundOrder(attrGoogleOrderNumber, elemReason, elemAmount, elemComment)
		
		' Validate Request XML
		'DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, xmlRequest, "debug"
		
		transmitResponse = SendRequest(xmlRequest, requestUrl)
		
		' Process the response
		ProcessXmlData(transmitResponse)
		
		' Processing Status
		ps=6
		
		'// Update Fullment Status
		query="UPDATE orders SET pcOrd_PaymentStatus="& ps &" WHERE pcOrd_GoogleIDOrder='"& attrGoogleOrderNumber &"' "
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		set rs=nothing
	
	
	
	
	Case "everything else"
	
		' The next three lines of code specify the information for
		' a <charge-order> command and then call the CreateChargeOrder
		' function to construct that command.
		'
		' Following the <charge-order> command, there are several snippets 
		' of commented code that could be used to create other types of 
		' Order Processing API commands for the same order.
		
		
		' This section creates an <unarchive-order> request
		attrGoogleOrderNumber = "841171949013218"
		xmlRequest = createUnarchiveOrder(attrGoogleOrderNumber)


End Select
%>

