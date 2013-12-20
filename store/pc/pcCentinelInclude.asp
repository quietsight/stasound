<%
If request.QueryString("centinel")<>"Y" Then
	session("reqCardNumber")=request.Form("cardNumber")
	session("reqExpMonth")=request.Form("expMonth")
	session("reqExpYear")=request.Form("expYear")
	session("reqCVV")=request.Form("CVV")
End If
	
' validates expiration
if DateDiff("d", Month(Now)&"/"&Year(now), session("reqExpMonth")&"/20"&session("reqExpYear"))<=-1 then
	response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_6")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
     
end if

'Check if Centinel is Active
pcPay_CentByPass=0
If request.QueryString("centinel")="Y" Then
	pcPay_CentByPass=1
End If

dim pcPay_CardType
pcPay_CardType="YES"

If determineCardType(session("reqCardNumber"))<>"MASTERCARD" AND determineCardType(session("reqCardNumber"))<>"VISA" then
	pcPay_CentByPass=1
	pcPay_CardType="NO"
end if

If pcPay_Cent_Active=1 AND pcPay_CentByPass=0 Then
	' Send the XML message and Retrieve the Response
	Dim xmlRequest, xmlResponse
	xmlRequest = "<CardinalMPI>"
	xmlRequest = xmlRequest & generateXMLTag("Version", pcPay_Cent_MessageVersion)
	xmlRequest = xmlRequest & generateXMLTag("MsgType", "cmpi_lookup")
	xmlRequest = xmlRequest & generateXMLTag("ProcessorId", pcPay_Cent_ProcessorId)
	xmlRequest = xmlRequest & generateXMLTag("MerchantId", pcPay_Cent_MerchantID)
	xmlRequest = xmlRequest & generateXMLTag("TransactionPwd", pcPay_Cent_TransactionPwd)
	xmlRequest = xmlRequest & generateXMLTag("TransactionType", "C")
	xmlRequest = xmlRequest & generateXMLTag("OrderNumber", Cstr(session("GWOrderID")))
	GWOrderTotal=replace(money(pcBillingTotal),".","")
	GWOrderTotal=replace(GWOrderTotal,",","")
	xmlRequest = xmlRequest & generateXMLTag("Amount", Cstr(GWOrderTotal))
	xmlRequest = xmlRequest & generateXMLTag("CurrencyCode", Cstr("840"))
	xmlRequest = xmlRequest & generateXMLTag("CardNumber", Cstr(session("reqCardNumber")))
	dtTempExpYear=Cstr(session("reqExpYear"))
	if len(dtTempExpYear)=2 then
		dtTempExpYear="20"&dtTempExpYear
	end if
	xmlRequest = xmlRequest & generateXMLTag("CardExpMonth", Cstr(session("reqExpMonth")))
	xmlRequest = xmlRequest & generateXMLTag("CardExpYear", dtTempExpYear)
	xmlRequest = xmlRequest & generateXMLTag("OrderDescription", Cstr(replace(scCompanyName,",","-") & " Order: " & session("GWOrderID")))
	xmlRequest = xmlRequest & generateXMLTag("UserAgent", Request.ServerVariables("HTTP_USER_AGENT"))
	xmlRequest = xmlRequest & generateXMLTag("BrowserHeader", Request.ServerVariables("HTTP_ACCEPT"))
	xmlRequest = xmlRequest & generateXMLTag("Installment", Cstr(request("Installment")))
	xmlRequest = xmlRequest & generateXMLTag("IPAddress", pcCustIpAddress)
	xmlRequest = xmlRequest & generateXMLTag("EMail", Cstr(pcCustomerEmail))
	xmlRequest = xmlRequest & generateXMLTag("FirstName", Cstr(pcBillingFirstName))
	xmlRequest = xmlRequest & generateXMLTag("LastName", Cstr(pcBillingLastName))
	xmlRequest = xmlRequest & generateXMLTag("Address1", Cstr(pcBillingAddress))
	xmlRequest = xmlRequest & generateXMLTag("Address2", Cstr(pcBillingAddress2))
	xmlRequest = xmlRequest & generateXMLTag("City", Cstr(pcBillingCity))
	xmlRequest = xmlRequest & generateXMLTag("State", Cstr(pcBillingState))
	xmlRequest = xmlRequest & generateXMLTag("CountryCode", Cstr(pcBillingCountryCode))
	xmlRequest = xmlRequest & generateXMLTag("PostalCode", Cstr(pcBillingPostalCode))
	xmlRequest = xmlRequest & generateXMLTag("Recurring", "N")
	xmlRequest = xmlRequest & "</CardinalMPI>"

	xmlResponse = sendMessage(xmlRequest, "xmlErrors.asp")
	'======================================================================================
	' If communication Times Out the xmlResponse will be empty
	'======================================================================================
	
	If xmlResponse <> "" Then
		on error resume next
		
		' Retrieve the Elements from the XML Response
		Dim oDOM
		Set oDOM = CreateObject("Msxml2.DOMDocument")
		oDOM.async = False 
		oDOM.LoadXML xmlResponse
		
		Dim oErrorNo, oErrorDesc, oEnrolled, oTransactionId, oAcsURL, oPayload
		
		Set oErrorNo = oDOM.documentElement.selectSingleNode("ErrorNo") 
		Set oErrorDesc = oDOM.documentElement.selectSingleNode("ErrorDesc")
		Set oTransactionId = oDOM.documentElement.selectSingleNode("TransactionId")
		Set oEnrolled = oDOM.documentElement.selectSingleNode("Enrolled")

		'======================================================================================
		' Assert that there was no error code returned and the Cardholder is enrolled in the PI
		' prior to starting the Authentication process.
		'======================================================================================
	
		Dim strCardEnrolled, strErrorNo, strErrorDesc, strTransactionId
	
		strCardEnrolled = oEnrolled.text
		strErrorNo = oErrorNo.text
		strErrorDesc = oErrorDesc.text
		strTransactionId = oTransactionId.text
	
		Session("Centinel_ErrorNo") = strErrorNo
		Session("Centinel_ErrorDesc") = strErrorDesc
		Session("Centinel_TransactionId") = strTransactionId
		Session("Centinel_Enrolled") = oEnrolled.text
		
		intDebug=0
		
		if intDebug=1 then
			response.write Session("Centinel_Enrolled")&"<BR>"
			response.write Session("Centinel_ErrorNo")&"<BR>"
			response.write Session("Centinel_ErrorDesc")&"<BR>"
			response.End
		end if
		'======================================================================================
		' Handle ALL Payer Authentication Logic
		'======================================================================================

		If strErrorNo = "0" AND strCardEnrolled = "Y" Then 
	
			'======================================================================================
			' Proceed to Payer Authentication
			'======================================================================================
	
			Set oAcsURL = oDOM.documentElement.selectSingleNode("ACSUrl")
			Set oPayload = oDOM.documentElement.selectSingleNode("Payload")
			
			Session("Centinel_ACSURL") = oAcsURL.text
			Session("Centinel_PAYLOAD") = oPayload.text
			%>
					<p>For your security, please fill out the form below to compete your order. Do not click the refresh or back button or this transaction may be interrupted or cancelled.</p>
					</td>
				</tr>
			</table>
			</div>
			<IFRAME src="pcPay_Cent_Ecauth.asp" width="100%" height="450" frameborder="0" ALLOWTRANSPARENCY="true"> </IFRAME> 
			</body>
			</html>
			<% 
			conlayout.Close 
			Set conlayout=nothing 
			Set RSlayout = nothing 
			Set rsIconObj = nothing 
			%>


			<% response.end
		ElseIf strErrorNo = "0" AND strCardEnrolled = "U" Then
	
			'======================================================================================
			' Proceed to Authorization, Payer Authentication Not Available
			' Set the proper ECI value based on the Card Type
			'======================================================================================
			
			If (determineCardType(session("reqCardNumber")) = "VISA") Then
				Session("Centinel_ECI") = "07"
				ElseIf (determineCardType(session("reqCardNumber")) = "MASTERCARD") Then
				Session("Centinel_ECI") = "01"
				ElseIf (determineCardType(session("reqCardNumber")) = "JCB") Then
				Session("Centinel_ECI") = "07"
			End If

		ElseIf strErrorNo = "0" AND strCardEnrolled = "N" Then
			'======================================================================================
			' Proceed to Authorization, Payer Authentication Not Available
			' Set the proper ECI value based on the Card Type
			'======================================================================================
			If (determineCardType(session("reqCardNumber")) = "VISA") Then
				Session("Centinel_ECI") = "06"
				ElseIf (determineCardType(session("reqCardNumber")) = "MASTERCARD") Then
				Session("Centinel_ECI") = "01"
				ElseIf (determineCardType(session("reqCardNumber")) = "JCB") Then
				Session("Centinel_ECI") = "06"
			End If
		Else 
				'==================================================================================
				' An error was encountered
				' Log Error Message, this is an unexpected state
				' Proceed to authorization to complete the transaction.
				'==================================================================================
				If (determineCardType(session("reqCardNumber")) = "VISA") Then
					Session("Centinel_ECI") = "07"
					ElseIf (determineCardType(session("reqCardNumber")) = "MASTERCARD") Then
					Session("Centinel_ECI") = "01"
					ElseIf (determineCardType(session("reqCardNumber")) = "JCB") Then
					Session("Centinel_ECI") = "07"
				Else
					if strErrorDesc<>"" then
						session("Message2")=strErrorDesc
					else
						session("Message2")="An error has occurred during the checkout process, please resubmit your payment information."
					end if
					response.redirect "msgb.asp?message="&server.URLEncode("<b>Errors&nbsp;</b>:&nbsp;"&session("Message2")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("reqCustomerID")&"&idOrder="&session("reqOrderID")&"&ordertotal="&session("GWOrderTotal")&"""><img src="""&rslayout("back")&"""></a>")
				End if
		End If
	Else
		Session("Message") = "Communication timed out with Cardinal Centinel"
		err.clear
		err.number=0
	End If 
	

	'// Insert all values for Centinel
	call opendb()
	
	query="INSERT INTO pcPay_Centinel_Orders (pcPay_CentOrd_OrderID,pcPay_CentOrd_Enrolled,pcPay_CentOrd_ErrorNo, pcPay_CentOrd_ErrorDesc,pcPay_CentOrd_PAResStatus,pcPay_CentOrd_SignatureVerification,pcPay_CentOrd_EciFlag, pcPay_CentOrd_Xid, pcPay_CentOrd_Cavv, pcPay_CentOrd_rErrorNo,pcPay_CentOrd_rErrorDesc) VALUES ("&session("GWOrderID")&",'"&Session("Centinel_Enrolled")&"','"&Session("Centinel_ErrorNo")&"','"&Session("Centinel_ErrorDesc")&"','"&Session("Centinel_PAResStatus")&"','"&Session("Centinel_SignatureVerification")&"','"&Session("Centinel_ECI")&"','"&Session("Centinel_XID")&"','"&Session("Centinel_CAVV")&"','"&Session("Centinel_ErrorNo")&"','"&Session("Centinel_ErrorDesc") &"');"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	set rs=nothing
	
	call closedb()

End If 
%>