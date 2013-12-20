<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<%
'//Set redirect page to the current file name
session("redirectPage")="gwSkipJack.asp"

'//Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

'//Declare URL path to gwSubmit.asp
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If

'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if

'//Retrieve customer data from the database using the current session id
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

dim connTemp, rs
call openDb()
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE
'//Functions parse response string
'======================================================================================
Dim FinalArr
FUNCTION ParseSJAPI_AuthorizeRequest( strResponse )

	 'Variable to hold the Final Array of values

	IF NOT ISNULL( strResponse ) THEN

			Dim arrLines 'Variable to hold each line of the Response

		'Split the Response string by carriage return
		'First line is the Header Record and will be stored in arrLines(0)
		'Second line is the Returned data and will be stored in arrLines(1)
			arrLines = Split( strResponse, vbCRLF, -1, vbTextCompare )

		'Check to ensure we have an array with the Upper Bounds greater than 1
			IF IsArray( arrLines ) AND UBound( arrLines ) >= 1 THEN

				Dim arrResponse, arrHeaders  'Variable to hold the response in as an array

			'Split the Response Data, arrLines(1), out into an array using ","
				arrHeaders = Split( arrLines(0), CHR(34) & "," & CHR(34), -1, vbTextCompare )
				arrResponse = Split( arrLines(1), CHR(34) & "," & CHR(34), -1, vbTextCompare )

			'Check to ensure we have an array of Response Data
				IF IsArray( arrResponse ) THEN
				IF UBound( arrResponse ) >= 13 THEN

					'Create variables to hold each piece of the response data
					Dim ApprovalCode, SerialNumber, TransactionAmount, DeclineMessage
					Dim AVSResponseCode, AVSResponseMessage, OrderNumber, AuthorizationResponseCode
					Dim IsApproved, CVV2ResponseCode, CVV2ResponseMessage, ReturnCode, TransactionID, CAVVResponseCode

					'Assign the response data to it's corresponding variable.
					'For the first and last element in the array, remove the double quote from it.
					ApprovalCode = TRIM( REPLACE( arrResponse(0), CHR(34), "", 1, -1, vbTextCompare ) )
					SerialNumber = TRIM( arrResponse(1) )
					TransactionAmount = TRIM( arrResponse(2) )
					DeclineMessage = TRIM( arrResponse(3) )
					AVSResponseCode = TRIM( arrResponse(4) )
					AVSResponseMessage = TRIM( arrResponse(5) )
					OrderNumber = TRIM( arrResponse(6) )
					AuthorizationResponseCode = TRIM( arrResponse(7) )
					IsApproved = TRIM( arrResponse(8) )
					CVV2ResponseCode = TRIM( arrResponse(9) )
					CVV2ResponseMessage = TRIM( arrResponse(10) )
					ReturnCode = TRIM( arrResponse(11) )
					TransactionID = TRIM( arrResponse(12) )
					CAVVResponseCode = TRIM( REPLACE( arrResponse(13), CHR(34), "", 1, -1, vbTextCompare ) )

					'Format the Transaction amount by adding in a decimal place,
					'Skipjack returns the amount in cents, so 500 = 5.00
					Dim LenAmt
					LenAmt = LEN( TransactionAmount )

					IF LenAmt >= 3 THEN
						TransactionAmount = LEFT( TransactionAmount, LenAmt - 2 ) & "." & RIGHT( TransactionAmount, 2 )
					END IF

					'Redim the FinalArray and assign the Response Data to it
					Redim FinalArr(14)
					FinalArr(0) = ApprovalCode
					FinalArr(1) = SerialNumber
					FinalArr(2) = TransactionAmount
					FinalArr(3) = DeclineMessage
					FinalArr(4) = AVSResponseCode
					FinalArr(5) = AVSResponseMessage
					FinalArr(6) = OrderNumber
					FinalArr(7) = AuthorizationResponseCode
					FinalArr(8) = IsApproved
					FinalArr(9) = CVV2ResponseCode
					FinalArr(10) = CVV2ResponseMessage
					FinalArr(11) = ReturnCode
					FinalArr(12) = TransactionID
					FinalArr(13) = CAVVResponseCode

					'FOR TESTING - write out the Response Data
					'Response.Write( "Approval Code = " & ApprovalCode & "<br />" & vbCRLF )
					'Response.Write( "Serial Number = " & SerialNumber & "<br />" & vbCRLF )
					'Response.Write( "Transaction Amount = " & TransactionAmount & "<br />" & vbCRLF )
					'Response.Write( "Decline Message = " & DeclineMessage & "<br />" & vbCRLF )
					'Response.Write( "AVS Response Code = " & AVSResponseCode & "<br />" & vbCRLF )
					'Response.Write( "AVS Response Message = " & AVSResponseMessage & "<br />" & vbCRLF )
					'Response.Write( "Order Number = " & OrderNumber & "<br />" & vbCRLF )
					'Response.Write( "Authorization Response Code = " & AuthorizationResponseCode & "<br />" & vbCRLF )
					'Response.Write( "IsApproved = " & IsApproved & "<br />" & vbCRLF )
					'Response.Write( "CVV2 Response Code = " & CVV2ResponseCode & "<br />" & vbCRLF )
					'Response.Write( "CVV2 Response Message = " & CVV2ResponseMessage & "<br />" & vbCRLF )
					'Response.Write( "Return Code = " & ReturnCode & "<br />" & vbCRLF )
					'Response.Write( "Transaction ID = " & TransactionID & "<br />" & vbCRLF )
					'Response.Write( "CAVV Response Code = " & CAVVResponseCode & "<br />" & vbCRLF )

					END IF   'End if Ubound(arrResponse) >= 13

				END IF   'End If IsArray( arrResponse )

			END IF    'End If IsArray( arrLines ) AND UBound( arrLines ) >= 1

	END IF   'End If NOT ISNULL( strResponse )

	'Return the FinalArray of Response Data.
	'ParseSJAPI_AuthorizeRequest = FinalArr

END FUNCTION

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT pcPay_SkipJack_SerialNumber, pcPay_SkipJack_TestMode, pcPay_SkipJack_Cvc2 FROM pcPay_SkipJack WHERE pcPay_SkipJack_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_SkipJack_SerialNumber=rs("pcPay_SkipJack_SerialNumber")
pcPay_SkipJack_TestMode=rs("pcPay_SkipJack_TestMode")
pcPay_CVV=rs("pcPay_SkipJack_Cvc2")
pcPay_SkipJack_SSLProvider="1"
set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	err.clear

	Dim objXMLHTTP,DataToSend,  xml, XmlSend, strStatus
	' make sure the CVV is Required
	if pcPay_CVV = "1" and (not isNumeric(request.form("CVV")) or  len(request.form("CVV")) < 3 ) Then
		Msg = "Please Supply a Security Code."
		response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwSkipJack.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		Response.end
	End if
	' in test amount must be ending in .00
	if pcPay_SkipJack_TestMode=1 then
		pcBillingTotal = "1.00"
	End if

	'// Get Card Details
	pcv_CardNumber=Request.Form( "CardNumber" )
	pcv_expMonth=Request.Form( "expMonth" )
	pcv_expYear=Request.Form( "expYear" )
	pcv_CVVRequest=request.form("CVV")
	'Send the transaction info as part of the a Post
	if pcShippingFirstName<>"" then
	   pcShipToName=pcShippingFirstName&" "&pcShippingLastName
	else
	   pcShipToName=pcBillingFirstName&" "&pcBillingLastName
	end if

	DataToSend = "SerialNumber=" & Server.URLEncode(pcPay_SkipJack_SerialNumber) &_
	 "&SJName=" & Server.URLEncode(pcBillingFirstName & " " & pcBillingLastName) &_
	 "&Email=" & Server.URLEncode(pcCustomerEmail) &_
	 "&StreetAddress=" & Server.URLEncode(pcBillingAddress) &_
	 "&City=" & Server.URLEncode(pcBillingCity) & _
	 "&State=" & Server.URLEncode(pcBillingState) & _
	 "&ZipCode=" & Server.URLEncode(pcBillingPostalCode) & _
	 "&Country=" & Server.URLEncode(pcBillingCountryCode) &_
	 "&AccountNumber=" &  server.URLEncode(pcv_CardNumber) &_
	 "&Month=" &  server.URLEncode(pcv_expMonth) &_
	 "&Year=" &  server.URLEncode(pcv_expYear) &_
	 "&TransactionAmount=" & money(pcBillingTotal) &_
	 "&OrderNumber=" &	server.URLEncode(session("GWOrderID")) & _
	 "&OrderString=" &	server.URLEncode("1~None~0.00~0~N~||") & _
	 "&ShipToName=" & Server.URLEncode(pcShipToName) & _
	 "&ShipToStreetAddress=" & Server.URLEncode(pcBillingAddress) &_
	 "&ShipToCity=" & Server.URLEncode(pcBillingCity) & _
	 "&ShipToState=" & Server.URLEncode(pcBillingState) & _
	 "&ShipToZipCode=" & Server.URLEncode(pcBillingPostalCode) & _
	 "&ShipToCountry=" & Server.URLEncode(pcBillingCountryCode) &_
	 "&ShipToPhone=" & Server.URLEncode(pcShippingPhone)

	if pcPay_CVV = "1" Then
		DataToSend = DataToSend &    "&CVV2=" & Server.URLEncode(pcv_CVVRequest)
	End if

	' determine where what url to send to
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	if pcPay_SkipJack_TestMode=1 then
		xmlSend = "https://developer.skipjackic.com/scripts/evolvcc.dll?AuthorizeAPI"
	else
		xmlSend = "https://www.skipjackic.com/scripts/evolvcc.dll?AuthorizeAPI"
	end if
	'Send the request to the SkipJack processor.
	xml.open "POST", xmlSend , false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send(DataToSend)
	if err.number<>0 then
		pcResultErrorMsg = err.description
	end if
	strStatus = xml.Status

	if strStatus = 200 then
		'store the response
		strRetVal = xml.responseText


			'/////////////////////////////////////////////////////
			'// Create Log of response and save in includes
			'/////////////////////////////////////////////////////
			dim authLogging
			authLogging=0 'Change to 1 to log

			if authLogging=1 then

				if PPD="1" then
					pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/SJLOG.txt")
				else
					pcStrFileName=Server.Mappath ("../includes/SJLOG.txt")
				end if

				dim strFileName
				dim fs
				dim OutputFile

				'Specify directory and file to store silent post information
				strFileName = pcStrFileName
				Set fs = CreateObject("Scripting.FileSystemObject")
				Set OutputFile = fs.OpenTextFile (strFileName, 8, True)

				OutputFile.WriteLine now()
				OutputFile.WriteLine "===================================="
				OutputFile.WriteLine "URL: " & xmlSend
				OutputFile.WriteBlankLines(1)
				OutputFile.WriteLine "Request from ProductCart: " & DataToSend
				OutputFile.WriteBlankLines(1)


				OutputFile.WriteLine "Response from Skip Jack: " & strRetVal
				OutputFile.WriteBlankLines(2)

				OutputFile.Close
			end if
			'/////////////////////////////////////////////////////
			'// End - Create Log of response and save in includes
			'/////////////////////////////////////////////////////

		if strRetVal <> "" Then
			' parse and load array
			Call ParseSJAPI_AuthorizeRequest( strRetVal )
			' Process Returns
			pcResultAuthCode =  FinalArr(0)
			pcResultErrorMsg = FinalArr(3)
			pcApproved = FinalArr(8)
			pcResultResponseCode = FinalArr(11)
			pcResultTransRefNumber = FinalArr(12)
			If pcResultErrorMsg&""<>"" Then
			Else
				select case pcResultResponseCode
					case "-1"
						pcResultErrorMsg="There was an error in your request."
				case "1"
					pcResultErrorMsg="Status complete."
				case "-34"
					pcResultErrorMsg="Error authorization failed."
				case "-35"
					pcResultErrorMsg="Error invalid credit card number."
				case "-37"
					pcResultErrorMsg="Error failed dial."
				case "-39"
					pcResultErrorMsg="Error length serial number."
				case "-51"
					pcResultErrorMsg="Error length zip code."
				case "-52"
					pcResultErrorMsg="Error length shipto zip code."
				case "-53"
					pcResultErrorMsg="Error length expiration date."
				case "-54"
					pcResultErrorMsg="Error length account number date."
				case "-55"
					pcResultErrorMsg="Error length street address."
				case "-56"
					pcResultErrorMsg="Error length shipto street address."
				case "-57"
					pcResultErrorMsg="Error length transaction amount."
				case "-58"
					pcResultErrorMsg="Error length name."
				case "-59"
					pcResultErrorMsg="Error length location."
				case "-60"
					pcResultErrorMsg="Error length state."
				case "-61"
					pcResultErrorMsg="Error length shipto state."
				case "-62"
					pcResultErrorMsg="Error length order string."
				case "64"
					pcResultErrorMsg="Error invalid phone number."
				case "-79"
					pcResultErrorMsg="Error length customer name."
				case "-80"
					pcResultErrorMsg="Error length shipto customer name."
				case "-81"
					pcResultErrorMsg="Error length customer location."
				case "-82"
					pcResultErrorMsg="Error length customer state."
				case "-83"
					pcResultErrorMsg="Error length shipto phone."
				case "-65"
					pcResultErrorMsg="Error empty name."
				case "-66"
					pcResultErrorMsg="Error empty email."
				case "-67"
					pcResultErrorMsg="Error empty street address."
				case "-68"
					pcResultErrorMsg="Error empty city."
				case "-69"
					pcResultErrorMsg="Error empty state."
				case "-70"
					pcResultErrorMsg="Error empty zip code."
				case "-71"
					pcResultErrorMsg="Error empty order number."
				case "-72"
					pcResultErrorMsg="Error empty account (credit card) number."
				case "-73"
					pcResultErrorMsg="Error empty credit card  expiration month."
				case "-74"
					pcResultErrorMsg="Error empty credit card expiration year."
				case "-75"
					pcResultErrorMsg="Error empty serial number."
				case "-76"
					pcResultErrorMsg="Error empty transaction amount."
				case "-77"
					pcResultErrorMsg="Error empty order string."
				case "-78"
					pcResultErrorMsg="Error empty phone number."
				case "-84"
					pcResultErrorMsg="Transaction with a duplicate order number was submitted (and is valid only when the merchant account is set up to reject duplicate orders)."
				case "-91"
					pcResultErrorMsg="CVV2 error."
				case "93"
					pcResultErrorMsg="Blind credits not allowed."
				case "-94"
					pcResultErrorMsg="Blind credits failed."
				case "95"
					pcResultErrorMsg="Voice Authorization not allowed."
				case "-96"
					pcResultErrorMsg="Voice Authorization failed."
					case "-97"
						pcResultErrorMsg="Fraud Rejection: Violates Velocity Setting."
					case "-98"
						pcResultErrorMsg="Invalid Discount Amount"
					case "-99"
						pcResultErrorMsg="POS PIN Debit Pin Block: Debit-specific"
					case "-100"
						pcResultErrorMsg="POS PIN Debit Invalid Key Serial Number: Debit-specific"
					case "-101"
						pcResultErrorMsg="Invalid Authentication Data: Data for Verified by Visa/MC Secure Code is invalid."
					case "-102"
						pcResultErrorMsg="Authentication Data Not Allowed"
					case "-103"
						pcResultErrorMsg="POS Check Invalid Birth Date: POS check dateofbirth variable contains a birth date in an incorrect format. Use MM/DD/YYYY format for this variable."
					case "-104"
						pcResultErrorMsg="POS Check Invalid Identification Type: POS check identificationtype variable contains a identification type value which is invalid. Use the single digit value where Social Security Number=1, Drivers License=2 for this variable."
					case "-105"
						pcResultErrorMsg="Invalid trackdata: Track Data is in invalid format."
					case "-106"
						pcResultErrorMsg="POS Check Invalid Account Type"
					case "-107"
						pcResultErrorMsg="POS PIN Debit Invalid Sequence Number"
					case "-108"
						pcResultErrorMsg="Invalid Transaction ID: For TSYS PIN-based debit transactions the Unqtransactionid is not valid. Resubmit transaction using the correct Unqtransactionid."
					case "-109"
						pcResultErrorMsg="Invalid From Account Type"
					case "-110"
						pcResultErrorMsg="Pos Error Invalid To Account Type"
					case "-112"
						pcResultErrorMsg="Pos Error Invalid Auth Option: Options selected for account type are incorrect (must be in the range 1 to 8) or incorrect for the protocol type."
					case "-113"
						pcResultErrorMsg="Pos Error Transaction Failed"
					case "-114"
						pcResultErrorMsg="Pos Error Invalid Incoming ECI"
					case "-115"
						pcResultErrorMsg="POS Check Invalid Check Type"
					case "-116"
						pcResultErrorMsg="POS Check Invalid Lane Number: POS Check lane or cash register number is invalid. Use a valid lane or cash register number that has been configured in the Skipjack Merchant Account."
					case "-117"
						pcResultErrorMsg="POS Check Invalid Cashier Number"
					case "-118"
						pcResultErrorMsg="Invalid POST URL: The URL posted to is incorrect. Confirm URL that is being posted and resubmit the transaction."
			End Select
			End If
		Else
			'//ERROR
			pcResultErrorMsg = "Transaction error or declined.  Error Message: " & pcResultErrorMsg
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& pcResultErrorMsg &"<br><br><a href="""&tempURL&"?psslurl=gwSkipJack.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			Response.end
		End If
		If pcApproved = "1"  then
			session("GWAuthCode")=pcResultAuthCode
			session("GWTransId")=pcResultTransRefNumber
			response.redirect "gwReturn.asp?s=true&gw=SkipJack"
		Else
			if pcResultErrorMsg="" then
			  pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"
			end if
			Msg=pcResultErrorMsg
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& pcResultErrorMsg &"<br><br><a href="""&tempURL&"?psslurl=gwSkipJack.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			Response.end
		End if
	Else
		if pcResultErrorMsg="" then
			pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"
		end if
		Msg=pcResultErrorMsg
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& pcResultErrorMsg &"<br><br><a href="""&tempURL&"?psslurl=gwSkipJack.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		Response.end
	End if

	'*************************************************************************************
	' END
	'*************************************************************************************
end if
%>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<img src="images/checkout_bar_step5.gif" alt="">
			</td>
		</tr>
		<tr>
			<td>
				<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
				<input type="hidden" name="PaymentSubmitted" value="Go">
					<table class="pcShowContent">

					<% if Msg<>"" then %>
						<tr valign="top">
							<td colspan="2">
								<div class="pcErrorMessage"><%=Msg%></div>							</td>
						</tr>
					<% end if %>
					<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></p></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_1")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingFirstName&" "&pcBillingLastName%></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingAddress%></p></td>
					</tr>
					<% if pcBillingAddress2<>"" then %>
						<tr>
							<td colspan="2"><p><%=pcBillingAddress2%></p></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2"><p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% if pcPay_SkipJack_TestMode="1" then %>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						</tr>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
					<% end if %>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td>
							<input type="text" name="CardNumber" value="">						</td>
					</tr>
					<tr>
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%>
							<select name="expMonth">
								<option value="01">1</option>
								<option value="02">2</option>
								<option value="03">3</option>
								<option value="04">4</option>
								<option value="05">5</option>
								<option value="06">6</option>
								<option value="07">7</option>
								<option value="08">8</option>
								<option value="09">9</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="12">12</option>
							</select>
							<% dtCurYear=Year(date()) %>
							&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%>
							<select name="expYear">
								<option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,2)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,2)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,2)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,2)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,2)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,2)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,2)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,2)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,2)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,2)%>"><%=dtCurYear+10%></option>
							</select>						</td>
					</tr>
					<% If pcPay_CVV="1" Then %>
						<tr>
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td>
								<input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4">
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% End If %>
					<tr>
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>

					<tr>
						<td colspan="2" align="center">
							<!--#include file="inc_gatewayButtons.asp"-->
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->