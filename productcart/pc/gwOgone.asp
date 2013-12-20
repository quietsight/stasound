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
session("redirectPage")="gwOgone.asp"  'ALTER

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
'======================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your 
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
query="SELECT pcPay_OG_MerchantID,pcPay_OG_MerchantPassword,pcPay_OG_TransType,pcPay_OG_Lang,pcPay_OG_CurCode,pcPay_OG_cardTypes, pcPay_OG_CVC,pcPay_OG_AccountID,pcPay_OG_TestMode FROM pcPay_Ogone Where pcPay_OG_ID=1;"
			
			
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
	pcPay_OG_MerchantID=rs("pcPay_OG_MerchantID")
	pcPay_OG_MerchantID=enDeCrypt(pcPay_OG_MerchantID, scCrypPass)
	pcPay_OG_MerchantPassword=rs("pcPay_OG_MerchantPassword")
	pcPay_OG_MerchantPassword=enDeCrypt(pcPay_OG_MerchantPassword, scCrypPass)
	pcPay_OG_TransType = rs("pcPay_OG_TransType")
	pcPay_OG_Lang = rs("pcPay_OG_Lang")
	pcPay_OG_CurCode = rs("pcPay_OG_CurCode")
	pcPay_OG_cardTypes=rs("pcPay_OG_cardTypes")
	pcPay_OG_CVC=rs("pcPay_OG_CVC")
	pcPay_OG_AccountID = rs("pcPay_OG_AccountID")
	pcPay_OG_AccountID=enDeCrypt(pcPay_OG_AccountID, scCrypPass)
	pcPay_OG_TestMode=rs("pcPay_OG_TestMode")
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
if request("PaymentSubmitted")="Go" then
  

	'*************************************************************************************
	'// This is where you would post and retrieve info to and from the gateway
	'// START below this line
	'*************************************************************************************
	
		    Dim objXMLHTTP, xml
		    'Send the transaction info as part of the post
	        RequestData_POST = 	"PSPID=" & pcPay_OG_MerchantID & _
			"&PSWD=" & pcPay_OG_MerchantPassword & _
			"&orderID="& session("GWOrderID") & _
			"&amount=" & (pcBillingTotal*100)	& _
			"&CURRENCY="& pcPay_OG_CurCode & _
			"&CN=" & pcBillingFirstName & " " & pcBillingLastName & _	
			"&BRAND=" & Request.form("CardType") & _
			"&CARDNO=" & Request.Form( "CardNumber" ) & _
		    "&ED=" & Request.Form( "expMonth" )& "/" & Request.Form( "expYear" ) & _
			"&EMAIL=" & pcCustomerEmail & _
			"&REMOTE_ADDR=" & pcCustIpAddress & _
			"&LANGUAGE=" & pcPay_OG_Lang & _
			"&OWNERADDRESS="& pcBillingAddress & _
			"&OWNERZIP=" & pcBillingPostalCode & _
			"&USERID=" & pcPay_OG_AccountID 
			
			if pcPay_OG_CVC = 1 Then 
				RequestData_POST = RequestData_POST &"&CVC=" &request.form("CVV")
			End if 
		    'Send the transaction info as part of the post			
			' determine where what url to send to 
			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			if pcPay_OG_TestMode=1 then
				xmlSend = "https://www.secure.neos-solution.com/ncol/test/orderdirect.asp"
			else
				xmlSend = "https://www.secure.neos-solution.com/ncol/prod/orderdirect.asp"
			end if
			'Send the request to the Ogone processor.
			xml.open "POST", xmlSend , false
			xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
			xml.send(RequestData_Post)			
			if err.number<>0 then
				pcResultErrorMsg = err.description
				response.end
			end if
			 strStatus = xml.Status
			if strStatus = 200 then 	
				'store the response				
			    strRetVal = xml.responseText
				 'Response.write  xmlSend &"<BR>"& strRetVal &"<BR>"
				 'response.end 
				Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
				xmlDoc.async = False
				If xmlDoc.loadXML(strRetVal) Then
					' Get the results
					pcResultErrorCode = xmldoc.documentElement.getAttribute("NCERROR")
					pcResultTransRefNumber = xmldoc.documentElement.getAttribute("PAYID")
					pcResultResponseMess = xmldoc.documentElement.getAttribute("NCERRORPLUS")
					pcResultApproval = xmldoc.documentElement.getAttribute("ACCEPTANCE")					
				Else
					'//ERROR
					pcResultErrorCode = "Transaction error or declined.  Error Message: " & pcResultErrorMsg
					response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwOgone.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
					response.end
				End If

				If pcResultErrorCode = "0" then
				 	session("GWAuthCode")=pcResultApproval
					session("GWTransId")=pcResultTransRefNumber
					session("GWTransType")=pcPay_OG_TransType
					response.redirect "gwReturn.asp?s=true&gw=Ogone"
				Else				
					if pcResultErrorMsg="" then
					   pcResultErrorMsg="Transaction error or declined. Error:"   &    pcResultResponseMess        
					end if
					Msg=pcResultErrorMsg
					response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwOgone.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			       Response.end	
				End if
			Else
		 
			if pcResultErrorMsg="" then
				pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"
			end if
			Msg=pcResultErrorMsg
		  		response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& Msg &"<br><br><a href="""&tempURL&"?psslurl=gwOgone.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			Response.end	
			End if 
	

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
					if pcPay_OG_TestMode=1 then %>
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
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
						<td>
                            <select name="CardType">
								<% dim ArryCardTypes, strCardType, j
                                ArryCardTypes=split(pcPay_OG_CardTypes,", ")
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
										case "JCB"
											response.write "<option value='JCB'>JCB</option>"
										case "AURORA"
											response.write "<option value='AURORA'>Aurora</option>"
										case "AURORE"
											response.write "<option value='AURORE'>Aurore</option>"
										case "MaestroUK"
											response.write "<option value='MaestroUK'>MaestroUK</option>"
										case "EUROCARD"
											response.write "<option value='EUROCARD'>Euro Card</option>"
									end select
								next %>
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
					If pcPay_OG_CVC="1" Then %>
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
							<!--#include file="inc_gatewayButtons.asp"-->
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
