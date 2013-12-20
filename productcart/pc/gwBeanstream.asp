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
'======================================================================================
'// Set redirect page
'======================================================================================
' The redirect page tells the form where to post the payment information. Most of the 
' time you will redirect the form back to this page.
'======================================================================================
session("redirectPage")="gwBeanStream.asp"  'ALTER

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE	
'======================================================================================
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

': Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
': End Declare and Retrieve Customer's IP Address	

': Declare URL path to gwSubmit.asp	
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
': End Declare URL path to gwSubmit.asp

': Get Order ID and Set to session
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if
': End Get Order ID
	
': Get customer and order data from the database for this order	
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<%
': End Get customer and order data


': Reset customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
': End Reset customer session

': Open Connection to the DB
dim connTemp, rs 'DELETE FOR HARD CODED VARS
call openDb() 'DELETE FOR HARD CODED VARS
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'//Functions BeanStream parse response string
'======================================================================================
Function DecodeQueryValue(qValue)
		'Purpose: To URL decode a string
		'Pre: qValue is set to a url encoded value of a query string parameter. ex: "one+two"
		'Post: none
		'Returns: Returns the url decoded value of qValue. ex: "one two"
		Dim i
		Dim qChar
		dim newString
		if IsNull(qValue) = false then
			For i = 1 To Len(qValue)
			qChar = Mid(qValue, i, 1)
				If qChar = "%" Then
				on error resume next
				newString = newString & Chr("&H" & Mid(qValue, i + 1, 2))
				on error goto 0
				i = i + 2
				ElseIf qChar = "+" Then
				newString = newString & " "
				Else
				newString = newString & qChar
				End If
			Next
			
			DecodeQueryValue = Replace(newString, "&lt;", "<")
		else
			DecodeQueryValue = ""
		end if

End Function

Function GetQueryValue(queryString, paramName)
		'Purpose: To return the value of a parameter in an HTTP query string.
		'Pre: queryString is set to the full query string of url encoded name value pairs. ex:
		'"value1=one&value2=two&value3=3"
		' paramName is set to the name of one of the parameters in the queryString. ex: "value2"
		'Post: None
		'Returns: The function returns the query string value assigned to the paramName parameter. ex: "two"
		Dim pos1
		dim pos2
		Dim qString
		qString = "&" & queryString & "&"
		pos1 = InStr(1, qString, paramName & "=")
		If pos1 > 0 Then
		    pos1 = pos1 + Len(paramName) + 1
		    pos2 = InStr(pos1, qString, "&")
			If pos2 > 0 Then
			GetQueryValue = DecodeQueryValue(Mid(qString, pos1, pos2 - pos1))
			End If
		End If
End Function
'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your 
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
query="SELECT pcPay_BS_MerchantID,pcPay_BS_MerchantPassword,pcPay_BS_TransType,pcPay_BS_VBV,pcPay_BS_Interac,pcPay_BS_cardTypes, pcPay_BS_CVC,pcPay_BS_AccountID,pcPay_BS_TestMode FROM pcPay_BeanStream Where pcPay_BS_ID=1;"
			
			
'ALTER :: DELETE FOR HARD CODED VARS
'======================================================================================
'// End custom query
'======================================================================================

': Create recordset and execute query
set rs=server.CreateObject("ADODB.RecordSet") 'DELETE FOR HARD CODED VARS
set rs=connTemp.execute(query) 'DELETE FOR HARD CODED VARS

': Capture any errors
if err.number<>0 then 'DELETE FOR HARD CODED VARS
	call LogErrorToDatabase() 'DELETE FOR HARD CODED VARS
	set rs=nothing 'DELETE FOR HARD CODED VARS
	call closedb() 'DELETE FOR HARD CODED VARS
	response.redirect "techErr.asp?err="&pcStrCustRefID 'DELETE FOR HARD CODED VARS
end if 'DELETE FOR HARD CODED VARS

'======================================================================================
'// Set gateway specific variables - These can be your "hard coded variables" or 
'// Variables retrieved from the database.
'======================================================================================
	pcPay_BS_MerchantID=rs("pcPay_BS_MerchantID")
	pcPay_BS_MerchantID=enDeCrypt(pcPay_BS_MerchantID, scCrypPass)
	pcPay_BS_MerchantPassword=rs("pcPay_BS_MerchantPassword")
	pcPay_BS_MerchantPassword=enDeCrypt(pcPay_BS_MerchantPassword, scCrypPass)
	pcPay_BS_TransType = TRIM(rs("pcPay_BS_TransType"))
	pcPay_BS_VBV = rs("pcPay_BS_VBV")
	pcPay_BS_Interac = rs("pcPay_BS_Interac")
	pcPay_BS_cardTypes=rs("pcPay_BS_cardTypes")
	pcPay_BS_CVC=rs("pcPay_BS_CVC")
	pcPay_BS_AccountID = rs("pcPay_BS_AccountID")
	pcPay_BS_AccountID=enDeCrypt(pcPay_BS_AccountID, scCrypPass)
	pcPay_BS_TestMode=rs("pcPay_BS_TestMode")
'======================================================================================
'// End gateway specific variables
'======================================================================================

': Clear recordset and close db connection
set rs=nothing 'DELETE FOR HARD CODED VARS
call closedb() 'DELETE FOR HARD CODED VARS

'======================================================================================
'// If you are posting back to this page from the gateway form, all actions will happen 
'// here. 
'====================================================================================== 
 Dim objXMLHTTP, xml,xmlSend, RequestData_POST
 Dim pcv_SuccessURL
 '//        post from payment form ------ Redirect Back FRom VERify By Visa  -------- Funded Response FRom Bank(Interac Online)
if request("PaymentSubmitted")="Go" or Trim(Request.QueryString("MD")) <> "" or  Trim(Request.QueryString("funded")) <> "" then
  

	'*************************************************************************************
	'// This is where you would post and retrieve info to and from the gateway
	'// START below this line
	'*************************************************************************************
	   ' Post from Payemnt Page
		if pcBillingCountryCode="US" OR pcBillingCountryCode="CA" then
		else
			pcBillingStateCode="--"
		end if
		if pcShippingCountryCode="US" OR pcShippingCountryCode="CA" then
		else
			pcShippingStateCode="--"
		end if
		
		if pcShippingAddress&""="" OR pcShippingStateCode&""="" Then
			pcShippingAddress = pcBillingAddress
			pcShippingPostalCode = pcBillingPostalCode
			pcShippingCity = pcBillingCity
			pcShippingStateCode = pcBillingStateCode
 		end if
		
		if request("PaymentSubmitted")="Go" Then
					'Send the transaction info as part of the post
					RequestData_POST = 	"requestType=BACKEND" & _
					"&merchant_id=" & Server.URLEncode(pcPay_BS_MerchantID) & _
					"&password=" &  Server.URLEncode(pcPay_BS_MerchantPassword) & _
					"&username=" &  Server.URLEncode(pcPay_BS_AccountID) & _
					"&trnOrderNumber="&  Server.URLEncode(session("GWOrderID")) & _
					"&trnAmount=" &  Server.URLEncode(pcBillingTotal)	& _
					"&trnType="&  Server.URLEncode(pcPay_BS_TransType )& _
					"&trnCardOwner=" &  Server.URLEncode(pcBillingFirstName & " " & pcBillingLastName )& _	
					"&trnCardNumber=" &  Server.URLEncode(Request.Form( "CardNumber" )) & _
					"&trnExpMonth=" &  Server.URLEncode(Request.Form( "expMonth" ) ) & _
					"&trnExpYear=" &  Server.URLEncode(Request.Form( "expYear" )) & _
					"&ordName=" &  Server.URLEncode(pcBillingFirstName & " " & pcBillingLastName) & _	
					"&ordEmailAddress=" &  Server.URLEncode(pcCustomerEmail) & _
					"&ordPhoneNumber=" &  Server.URLEncode(pcBillingPhone) & _			
					"&ordAddress1="&  Server.URLEncode(pcBillingAddress) & _
					"&ordPostalCode=" &  Server.URLEncode(pcBillingPostalCode) & _
					"&ordCity=" &  Server.URLEncode(pcBillingCity) & _			
					"&ordProvince="&  (left(pcBillingStateCode,2)) & _
					"&ordCountry=" &  Server.URLEncode(pcBillingCountryCode) & _
					"&shipName=" &  Server.URLEncode(pcShippingFirstName & " " & pcShippingLastName )& _	
					"&shipEmailAddress=" &  Server.URLEncode(pcShippingEmail) & _
					"&shipPhoneNumber=" &  Server.URLEncode(pcShippingPhone) & _			
					"&shipAddress1="&  Server.URLEncode(pcShippingAddress) & _
					"&shipPostalCode=" &  Server.URLEncode(pcShippingPostalCode) & _
					"&shipCity=" &  Server.URLEncode(pcShippingCity) & _			
					"&shipProvince="&  (left(pcShippingStateCode,2)) & _
					"&shipCountry=" &  Server.URLEncode(pcShippingCountryCode)
					if pcPay_BS_CVC = 1 Then 
						RequestData_POST = RequestData_POST &"&trnCardCvd=" & Server.URLEncode(request.form("CVV"))
					End if 
					if pcPay_BS_Interac = 1 and Request.form("CardType") = "INTER" then
						 RequestData_POST = RequestData_POST  & "&paymentMethod=IO" 
						 TERM_URL =  "https://www.beanstream.com/scripts/process_transaction_auth.asp" 
						 RequestData_POST = RequestData_POST  & "&TermUrl=" & server.urlEncode(TERM_URL)	
				
					else
						  RequestData_POST = RequestData_POST  & "&paymentMethod=CC"
					End if 
					if pcPay_BS_VBV = 1 and Request.form("CardType") = "VISA" then
						  RequestData_POST = RequestData_POST  & "&vbvEnabled=" &  Server.URLEncode(pcPay_BS_VBV)
						  TERM_URL =  Replace(tempURL,"gwSubmit", "GwBeanstream")
						  RequestData_POST = RequestData_POST  & "&TermUrl=" & server.urlEncode(TERM_URL)	
						'  Response.write RequestData_POST
						 ' response.end 			
					else
						  RequestData_POST = RequestData_POST  & "&vbvEnabled=0" 
					End if 			
		    		xmlSend = "https://www.beanstream.com/scripts/process_transaction.asp"
		    ''' Verify By Visa Redirect From Gateway to Product Cart 
			''' Set up HTTPS VAlues To Complete Transaction with BeanStream
			Elseif Trim(Request.QueryString("MD")) <> "" Then
			      RequestData_POST = "PaRes=" &  Server.URLEncode(request("PaRes")) & "&MD=" &  Server.URLEncode(request("MD"))
                  xmlSend = "https://www.beanstream.com/scripts/process_transaction_auth.asp"	
		    ''' Intertac Online Redirect From Gateway to Product Cart 
			''' Set up HTTPS VAlues To Complete Transaction with BeanStream
			Elseif Trim(Request.QueryString("funded")) =  "1" Then
				      
                   xmlSend = "https://www.beanstream.com/scripts/process_transaction_auth.asp"
			      
			       RequestData_POST = "bank_choice=" & Server.URLEncode(Request.form("bank_choice")) & _
				   "&merchant_name=" & Server.URLEncode(Request.form("merchant_name")) & _
                   "&confirmValue=" & Server.URLEncode(Request.form("confirmValue")) & _
				   "&headerText=" & Server.URLEncode(Request.form("headerText")) & _
                   "&IDEBIT_MERCHDATA=" & Server.URLEncode(Request.form("IDEBIT_MERCHDATA")) & _
				   "&IDEBIT_INVOICE=" & Server.URLEncode(Request.form("IDEBIT_INVOICE")) & _
                   "&IDEBIT_AMOUNT=" & Server.URLEncode(Request.form("IDEBIT_AMOUNT")) & _
				   "&IDEBIT_FUNDEDURL="& Server.URLEncode(Request.form("IDEBIT_FUNDEDURL")) & _
				   "&IDEBIT_NOTFUNDEDURL="& Server.URLEncode(Request.form("IDEBIT_NOTFUNDEDURL")) & _
                   "&IDEBIT_ISSLANG=" &Server.URLEncode(Request.form("IDEBIT_ISSLANG")) & _
				   "&IDEBIT_TRACK2=" & Server.URLEncode(Request.form("IDEBIT_TRACK2")) & _
                   "&IDEBIT_ISSCONF=" & Server.URLEncode(Request.form("IDEBIT_ISSCONF")) & _
				   "&IDEBIT_ISSNAME=" & Server.URLEncode(Request.form("IDEBIT_ISSNAME")) & _
				   "&IDEBIT_VERSION=" &Server.URLEncode(Request.form("IDEBIT_VERSION"))
		    
		    Else ' Intertac Online Not Funnded Eror
			      Msg = "Transaction cancelled or declined."
				  response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwBeanStream.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			      Response.end						
			End if 
			
		 
		     'Send the transaction info as part of the post			
			 'determine where what url to send to			
			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)			
			'Send the request to the BeanStream processor.
			xml.open "POST", xmlSend , false
			xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
			xml.setOption(2) = 4096
            xml.setOption(3) = ""
			xml.send(RequestData_Post)			
			if err.number<>0 then
				pcResultErrorMsg = err.description
				response.end
			end if
			 strStatus = xml.Status
			if strStatus = 200 then 	
				'store the response				
			    strRetVal = xml.responseText
			    'Response.write  xmlSend &"<BR>sent--------------------------Returned<BR>"& strRetVal &"<BR>"
				'Response.end
				if strRetVal <> "" Then
					' Get the results responsetype=R Redirect or responsetype=T Complete then Go to verfify
					'  Verify By Visa OR Intertac Online Redirect to GateWay
					if  GetQueryValue(strRetVal,"responseType") ="R" then
					     Response.write GetQueryValue(strRetVal,"pageContents") 
					     Response.end
				     Else
					       ' Process Returns 
							pcResultApproveCode = GetQueryValue(strRetVal,"trnApproved")
							pcResultTransRefNumber = GetQueryValue(strRetVal,"trnId") 
							pcResultErrorMsg = GetQueryValue(strRetVal,"messageText")  
							pcResultAuthCode =  GetQueryValue(strRetVal,"authCode")
							' Save For Later Get Financial Institions name Confirmation Number
						    'If GetQueryValue(strRetVal,"paymentMethod") = "IO" Then
							'	pcResultFinacInstName = GetQueryValue(strRetVal,"ioInstName")
							'	pcResultConfNum = GetQueryValue(strRetVal,"ioConfCode")	
							'End if 
				 	end if   
					
				Else
		        	'//ERROR
					pcResultErrorMsg = "Transaction error or declined.  Error Message: " & pcResultErrorMsg
					response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwBeanStream.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			        Response.end	
				 End If
				 If pcResultApproveCode = "1" then
				 	session("GWAuthCode")=pcResultAuthCode
					session("GWTransId")=pcResultTransRefNumber
					session("GWTransType")=pcPay_BS_TransType
					' Save For Later
					'session("GWIOBankName") = pcResultFinacInstName
					'session("GWIOBankConfNum") = pcResultConfNum
					response.redirect "gwReturn.asp?s=true&gw=Ogone"
				  Else
					if pcResultErrorMsg="" then
					   pcResultErrorMsg="Transaction error or declined."   &    pcResultErrorMsg        
					end if
					Msg=pcResultErrorMsg
					response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "&Msg &"<br><br><a href="""&tempURL&"?psslurl=gwBeanStream.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			        Response.end	
				End if
			Else
		 
			if pcResultErrorMsg="" then
				pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"
			end if
			Msg=pcResultErrorMsg
		  		response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwBeanStream.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
				Response.end	
			End if 
		
			If scSSL="" OR scSSL="0" Then
				pcv_SuccessURL=replace((scStoreURL&"/"&scPcFolder&"/pc/pcPay_TD_Receipt.asp"),"//","/")
				pcv_SuccessURL=replace(pcv_SuccessURL,"https:/","https://")
				pcv_SuccessURL=replace(pcv_SuccessURL,"http:/","http://") 
			Else
				pcv_SuccessURL=replace((scSslURL&"/"&scPcFolder&"/pc/pcPay_TD_Receipt.asp"),"//","/")
				pcv_SuccessURL=replace(pcv_SuccessURL,"https:/","https://")
				pcv_SuccessURL=replace(pcv_SuccessURL,"http:/","http://")
			End If

	

	'*************************************************************************************
	' END
	'*************************************************************************************
	
end if 

'======================================================================================
'// End post back 
 
'======================================================================================


'======================================================================================
'// Show customer the payment form 
'======================================================================================
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
								<div class="pcErrorMessage"><%=Msg%></div>
							</td>
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
					<% 
					'======================================================================================
					'// If your gateway supports a Testing environment, create a variable for it and then
					'// if the cart is in testmode, alert the customer that this is not a live transaction.
					'// NOTE :: If no testing environment exists, delete the table row below
					'======================================================================================
					if pcPay_BS_Testmode=1 then %>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
					<% end if
					'======================================================================================
					'// End Testing environment variable
					'// NOTE :: If no testing environment exists, delete the table row above
					'======================================================================================
					%>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
                    <tr> 
						<td align="left"><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
						<td align="left">
                            <select name="CardType" onChange="hideCardInfo(this)">
								<% dim ArryCardTypes, strCardType, j
                                ArryCardTypes=split(pcPay_BS_CardTypes,", ")
                                for j=0 to ubound(ArryCardTypes) 
									strCardType=ArryCardTypes(j) 
									select case strCardType
										case "VISA"
											response.write "<option value='VISA'>VISA</option>"
										case "MAST"
											response.write "<option value='MAST'>Master Card</option>"
										case "AMER"
											response.write "<option value='AMER'>American Express</option>"
										case "DINE"
											response.write "<option value='DINE'>Diners Club</option>"
										case "SEARS"
											response.write "<option value='SEARS'>Sears Card</option>"
										end select
								next 
                                if pcPay_BS_Interac = 1 and pcPay_BS_TransType = "P"  Then
								    response.write "<option value='INTER'>Interac&copy; Online</option>"
							    End if	%>
                            </select>
						</td>
					</tr>
					<tr id="ccInfo1" > 
						<td align="left" ><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td align="left" > 
							<input type="text" name="CardNumber" value="">
						</td>
					</tr>
					<tr id="ccInfo2" > 
						<td align="left" ><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td align="left" ><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
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
							</select>
						</td>
					</tr>
					<% 
					'======================================================================================
					'// If your gateway supports Credit Card Security Code (such as CVV and CVV2), create
					'// a variable for it and then show the row below.
					'// NOTE :: If no Security Code support exists, delete the table row below
					'======================================================================================
					If pcPay_BS_CVC="1" Then %>
						<tr id="ccInfo3" > 
							<td align="left" ><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td align="left" > 
								<input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4">
							</td>
						</tr>
						<tr id="ccInfo4" > 
							<td>&nbsp;</td>
							<td align="left" ><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% end if
					'======================================================================================
					'// End Security Code support
					'// NOTE :: If no Security Code support exists, delete the table row above
					'======================================================================================
				 	%>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					
					<tr > 
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

<script language="javascript">

	function hideCardInfo(CC){
	var browserName=navigator.appName; 
	var showRow

		 if (browserName=="Microsoft Internet Explorer")
		  showRow = "block";
		 else
		  showRow = "table-row";
 
	
		 
				 if (CC.options[CC.selectedIndex].value == "INTER" ){
				  document.getElementById("ccInfo1").style.display ="none";
				  document.getElementById("ccInfo2").style.display ="none";
				<% If pcPay_BS_CVC="1" Then %>
				  document.getElementById("ccInfo3").style.display ="none";
				  document.getElementById("ccInfo4").style.display ="none";
				 <% end if %>
				  }else{
					document.getElementById("ccInfo1").style.display =showRow;
					document.getElementById("ccInfo2").style.display =showRow;
				<% If pcPay_BS_CVC="1" Then %>
				  document.getElementById("ccInfo3").style.display =showRow;
				  document.getElementById("ccInfo4").style.display =showRow;
				 <% end if %>
			
				  }
			
	 }

</script>
<% 
'======================================================================================
'// End Show customer the payment form 
'======================================================================================
%>
<!--#include file="footer.asp"-->