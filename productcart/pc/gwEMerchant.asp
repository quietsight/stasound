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
session("redirectPage")="gwEMerchant.asp"  'ALTER

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE	
'======================================================================================
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
dim connTemp, rs
call openDb()
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'======================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database 
'======================================================================================
query="SELECT pcPay_eMerch_MerchantID, pcPay_eMerch_PaymentKey, pcPay_eMerch_CVD, pcPay_eMerch_CardType, pcPay_eMerch_TransType, pcPay_eMerch_TestMode FROM pcPay_eMerchant WHERE pcPay_eMerch_ID=1;" 

'======================================================================================
'// End custom query
'======================================================================================

': Create recordset and execute query
set rs=server.CreateObject("ADODB.RecordSet") 'DELETE FOR HARD CODED VARS
set rs=connTemp.execute(query) 'DELETE FOR HARD CODED VARS

': Capture any errors
if err.number<>0 then 
	call LogErrorToDatabase() 
	set rs=nothing 
	call closedb() 
	response.redirect "techErr.asp?err="&pcStrCustRefID 
end if 

'======================================================================================
'// Set gateway specific variables - hard code is not using database to store gateway
'// information
'======================================================================================
pcv_MerchantID=rs("pcPay_eMerch_MerchantID")
pcv_PaymentKey=rs("pcPay_eMerch_PaymentKey")
pcv_CVV=rs("pcPay_eMerch_CVD")
pcv_CardType=rs("pcPay_eMerch_CardType")
pcv_TransType=rs("pcPay_eMerch_TransType")
pcv_TestMode=rs("pcPay_eMerch_TestMode")

'======================================================================================
'// End gateway specific variables
'======================================================================================

': Clear recordset and close db connection
set rs=nothing
call closedb()

'======================================================================================
'// If you are posting back to this page from the gateway form, all actions will happen 
'// here. 
'======================================================================================
if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	'// This is where you would post and retrieve info to and from the gateway
	'// START below this line
	'*************************************************************************************
	
	Dim objHTTP, strRequest, DLLUrl
	
	if pcv_TestMode="1" then
		DLLUrl="https://www.e-merchant.co.uk/api/test/"
	else
		DLLUrl="https://www.e-merchant.co.uk/api/"
	end if
	dim pcPassedTotal
	pcPassedTotal=replace(pcBillingTotal,",","")
	pcPassedTotal=replace(pcPassedTotal,".","")
	
	strRequest ="merchantID="& pcv_MerchantID&_
	"&paymentKey="& pcv_PaymentKey&_
	"&amount="& pcPassedTotal&_
	"&cardNumber="& request.Form("CardNumber")&_ 
	"&cardExp="& request.Form("expMonth")&"/"&request.Form("expYear")
	if pcv_CVV=1 then
		strRequest=strRequest&"&cvdIndicator=1"&_	
		"&cvdValue="& request.Form("CVV")
	else
		strRequest=strRequest&"&cvdIndicator=0"	
	end if
	'//IF THE CARD TYPE SUBMITTED IS SWITCH/MAESTRO OR SOLO, THIS NUMBER IS REQUIRED
	if (request.Form("cardType")="SW" OR request.Form("cardType")="SO") AND request("IssueNumber")<>"" then
		strRequest=strRequest&"&issueNumber="& request.Form("IssueNumber")
	end if
	
	strRequest=strRequest&"&cardType="&request.Form("cardType")&_
	"&operation="& pcv_TransType&_	
	"&clientVersion=1.1"&_	
	
	"&custName1="&pcBillingFirstName&" "&pcBillingLastName&_
	"&streetAddr="&pcBillingAddress
	if pcBillingAddress2<>"" then
		strRequest=strRequest&"&streetAddr2="&pcBillingAddress2
	end if
	strRequest=strRequest&"&city="&pcBillingCity&_
	"&country="&pcBillingCountryCode&_
	"&zip="&pcBillingPostalCode&_
	"&phone="&pcBillingPhone&_
	"&email="&pcCustomerEmail&_
	"&merchantTxn="&session("GWOrderID")

	set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP"&scXML)
	objHttp.open "POST", DLLUrl, false
	objHttp.setRequestHeader "Host", "www.e-merchant.co.uk:1402"
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objHttp.setRequestHeader "Content-Length", Len(strRequest)
	objHttp.Send strRequest

	If objHTTP.Status = 200 Then  ' HTTP_STATUS_OK=200 
	
		Dim objDictResponse, intDelimiterPos, ResponseArray
		Dim strResponse, strNameValuePair, strName, strValue
		strResponse = replace(objHTTP.ResponseText,chr(10)," ")
		strResponse = rtrim(strResponse)

		' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
		ResponseArray = Split(strResponse, "&")
		Set objDictResponse = server.createobject("Scripting.Dictionary")
		For each ResponseItem in ResponseArray
			NameValue = Split(ResponseItem, "=")
			objDictResponse.Add NameValue(0), NameValue(1)
		Next
	
		' Parse the response into local vars
		strstatus=objDictResponse.Item("status")
		if strstatus="E" then
			errCode=objDictResponse.Item("errCode")
			errString=objDictResponse.Item("errString")
			errstring=replace(errstring,"%2E",".<br>")
			errstring=replace(errstring,"+"," ")
			subError=objDictResponse.Item("subError")
			actionCode=objDictResponse.Item("actionCode")
			clientVersion=objDictResponse.Item("clientVersion")
			authCode=objDictResponse.Item("authCode")
			subErrorString=objDictResponse.Item("subErrorString")
			subErrorString=replace(subErrorString,"%2E",".<br>")
			subErrorString=replace(subErrorString,"+"," ")
			avsInfo=objDictResponse.Item("avsInfo")
			cvdInfo=objDictResponse.Item("cvdInfo")
			pmtError=objDictResponse.Item("pmtError")
		end if
		
		'save and update eMerchant table
		If strstatus = "SP" OR strstatus="A" Then
			pcPay_eMerch_authCode=objDictResponse.Item("authCode")
			pcPay_eMerch_authTime=objDictResponse.Item("authTime")
			pcPay_eMerch_avsInfo=objDictResponse.Item("avsInfo")
			pcPay_eMerch_curAmount=objDictResponse.Item("curAmount")
			pcPay_eMerch_amount=objDictResponse.Item("amount")
			pcPay_eMerch_TxnNumber=objDictResponse.Item("txnNumber")
			pcPay_eMerch_clientVersion=objDictResponse.Item("clientVersion")
			pcPay_eMerch_cvdInfo=objDictResponse.Item("cvdInfo")
			
			call opendb()
			
			pcv_SecurityPass = pcs_GetSecureKey
			pcv_SecurityKeyID = pcs_GetKeyID
			
			dim pCardNumber, pCardNumber2
			pCardNumber=request.Form("cardNumber")
			pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)
			'save info in NetbillOrders if AUTH_CAPTURE
			pTempOrderID=(int(session("GWOrderId"))-scpre)
			
			if pcv_TransType="A" then 'Authorisation
				query="INSERT INTO pcPay_eMerch_Orders (idOrder, idCustomer, pcPay_eMerch_Ord_Amount, pcPay_eMerch_Ord_CardType, pcPay_eMerch_Ord_CardNumber, pcPay_eMerch_Ord_CardExp, pcPay_eMerch_Ord_TxnNumber, pcPay_eMerch_Ord_fname,pcPay_eMerch_Ord_lname, pcPay_eMerch_Ord_streetAddr, pcPay_eMerch_Ord_Country, pcPay_eMerch_Ord_Zip, pcPay_eMerch_Ord_Captured, pcSecurityKeyID) VALUES ("&pTempOrderID&", "&session("idCustomer")&", "&pcPay_eMerch_amount&",'"&request.Form("cardType")&"', '"&pCardNumber2&"', '"&request.Form("expMonth")&"/"&request.Form("expYear")&"', '"&pcPay_eMerch_TxnNumber&"', '"&pcBillingFirstName&"', '"&pcBillingLastName&"', '"&pcBillingAddress&"', '"&pcBillingCountry&"', '"&pcBillingPostalCode&"', 0, "&pcv_SecurityKeyID&");"

				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				set rs=nothing
			end if

			call closedb()

			session("GWAuthCode")=pcPay_eMerch_authCode
			session("GWTransId")=pcPay_eMerch_avsInfo
			session("GWTransType")=pcv_TransType
			
			Response.redirect "gwReturn.asp?s=true&gw=eMerchant"
				
		Elseif strstatus = "E" Then
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;"&errCode&":&nbsp;"&(errString)&"<br><br><a href="""&tempURL&"?psslurl=gwEMerchant.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			response.end
		End If  
			
	Else   
		strStatus= "Transaction Failed...<BR>"
		strStatus= strStatus&"<BR> Https Response <BR>" 
		strStatus= strStatus&"Status Code= " & objHttp.status       & "<BR>"
		strStatus= strStatus&"Error Status = " & objHttp.statusText   & "<BR>"
		strStatus= strStatus&"Header     = " & objHttp.getAllResponseHeaders &"<BR>"
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error:&nbsp;&nbsp;"&errCode&"&nbsp;due to&nbsp;"&lcase(errString)&"<br><br><a href="""&tempURL&"?psslurl=gwEMerchant.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		response.end
	End If   
	Set objHttp   = Nothing

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
					if pcv_TestMode=1 then %>
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
						<td><p>Card Type:</p></td>
						<td>
							<select name="cardType">
							<% 	'MC, MD, SO, SW, VI, VD, VE
								cardTypeArray=split(pcv_CardType,", ")
								i=ubound(cardTypeArray)
								cardCnt=0
								do until cardCnt=i+1
									cardVar=cardTypeArray(cardCnt)
									select case cardVar
										case "VI"
											response.write "<option value=""VI"" selected>Visa</option>"
											cardCnt=cardCnt+1
										case "MC" 
											response.write "<option value=""MC"">MasterCard</option>"
											cardCnt=cardCnt+1
										case "MD"
											response.write "<option value=""MD"">Maestro</option>"
											cardCnt=cardCnt+1
										case "SO"
											response.write "<option value=""SO"">SOLO</option>"
											cardCnt=cardCnt+1
										case "SW"
											response.write "<option value=""SW"">Switch/Maestro</option>"
											cardCnt=cardCnt+1
										case "VD"
											response.write "<option value=""VD"">Visa Delta</option>"
											cardCnt=cardCnt+1
										case "VE"
											response.write "<option value=""VE"">Visa Electron</option>"
											cardCnt=cardCnt+1
											
									end select
								loop
							%>
							</select>
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td> 
							<input type="text" name="CardNumber" value="">
						</td>
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
							</select>
						</td>
					</tr>
					<% 
					'======================================================================================
					'// If your gateway supports Credit Card Security Code (such as CVV and CVV2), create
					'// a variable for it and then show the row below.
					'// NOTE :: If no Security Code support exists, delete the table row below
					'======================================================================================
					If pcv_CVV="1" Then %>
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
					
					<tr> 
						<td colspan="2" align="center">
						<% 
						tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp"),"//","/")
						tempURL=replace(tempURL,"https:/","https://")
						tempURL=replace(tempURL,"http:/","http://") 
						%> 
						<a href="<%=tempURL%>"><img src="<%=rslayout("back")%>"></a>&nbsp;&nbsp; 
						<input type="image" name="Continue" src="<%=rslayout("pcLO_placeOrder")%>" id="submit">
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
</div>
<% 
'======================================================================================
'// End Show customer the payment form 
'======================================================================================
%>
<!--#include file="footer.asp"-->