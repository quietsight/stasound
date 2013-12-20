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
session("redirectPage")="gwACH.asp"

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

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT pcPay_ACH_MerchantID, pcPay_ACH_PWD, pcPay_ACH_TransType, pcPay_ACH_TestMode, pcPay_ACH_CVV, pcPay_ACH_CardTypes FROM pcPay_ACHDirect WHERE pcPay_ACH_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_ACH_MerchantID=rs("pcPay_ACH_MerchantID")
pcPay_ACH_MerchantId=enDeCrypt(pcPay_ACH_MerchantId, scCrypPass)
pcPay_ACH_PWD=rs("pcPay_ACH_PWD")
pcPay_ACH_PWD=enDeCrypt(pcPay_ACH_PWD, scCrypPass)
pcPay_ACH_TransType=rs("pcPay_ACH_TransType")
pcPay_ACH_TestMode=rs("pcPay_ACH_TestMode")
pcPay_CVV=rs("pcPay_ACH_CVV")
pcPay_ACH_CardTypes=rs("pcPay_ACH_CardTypes")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************

	Set gwObject=Server.CreateObject("SendPmt.clsSendPmt")
	
	if err.number<>0 then
		ACH_success=0
		strErrorInfo="Unable to send payment information, the required COM Object is not installed on this server."
	else
		'Formatting the data
		DIM strData
		
		strData="pg_merchant_id=" & pcPay_ACH_MerchantID & chr(10)
		strData=strData & "pg_password=" & pcPay_ACH_PWD & chr(10)
		strData=strData & "pg_total_amount=" & money(pcBillingTotal) & chr(10)

		'The transaction type is hard coded and can be changed if needed
		strData=strData & "pg_transaction_type=10" & chr(10)
		
		strData=strData & "ecom_billto_postal_name_first=" & pcBillingFirstName & chr(10)
		strData=strData & "ecom_billto_postal_name_last=" & pcBillingLastName & chr(10)
		strData=strData & "ecom_billto_postal_street_line1=" & pcBillingAddress & chr(10)
		strData=strData & "ecom_billto_postal_stateprov=" & pcBillingState & chr(10)
		strData=strData & "ecom_billto_postal_postalcode=" & pcBillingPostalCode & chr(10)
		strData=strData & "ecom_billto_telecom_phone_number=" & pcBillingPhone & chr(10)
		strData=strData & "ecom_billto_online_email=" & pcCustomerEmail & chr(10)
		strData=strData & "ecom_payment_card_name=" & pcBillingFirstName & " "& pcBillingLastName & chr(10)
		strData=strData & "ecom_payment_card_type=" & Request.Form("CardType") & chr(10)
		strData=strData & "ecom_payment_card_number=" & Request.Form( "CardNumber" ) & chr(10)
		strData=strData & "ecom_payment_card_expdate_month=" & Request.Form( "expMonth" ) & chr(10)
		strData=strData & "ecom_payment_card_expdate_year=" & Request.Form( "expYear" ) & chr(10)
		strData=strData & "ecom_payment_card_verification=" & Request.Form( "CVV" ) & chr(10)
		strData=strData & "pg_avs_method=220000" & chr(10)
		strData=strData & "endofdata" & chr(10)
	
		response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
		'-========================

		'Send the data over to our test server and get the response back
		IF pcPay_ACH_TestMode=1 THEN
			strReturn=gwObject.SendPayment(strData, "test")
		ELSE
			strReturn=gwObject.SendPayment(strData, "live")
		END IF
		
		'Parse out the return
		Call WriteReturnData(strReturn)

		'free memory
		SET gwObject = NOTHING

		'----------------------------------------------------
		'Utility functions for parsing out the return data
		'----------------------------------------------------
		'This function takes the string name such as "pg_merchant_id=" and parses out 
		'the value of "pg_merchant_id" from the return data
		Function ParseString(strReturn,strName)
			i=InStr(strReturn,Trim(strName))
			IF i > 0 THEN
				j=InStr(i,strReturn,chr(10))
				ParseString=Mid(strReturn,i+Len(strName),j-i-Len(strName)+1)
			ELSE
				ParseString=""
			END IF
		End Function

		'This function writes out the parsed return data in a table format
		Sub WriteReturnData(strMsg)
			DIM arrNames(9), arrValues(9)
			arrNames(0)="Merchant Id"
			arrNames(1)="Transaction Type"
			arrNames(2)="Total Amount"
			arrNames(3)="First Name"
			arrNames(4)="Last Name"
			arrNames(5)="Response Type"
			arrNames(6)="Response Code"
			arrNames(7)="Response Description"
			arrNames(8)="Authorization Code"
			arrNames(9)="Trace Number"
	
			arrValues(0)=ParseString(strMsg,"pg_merchant_id=")
			arrValues(1)=ParseString(strMsg,"pg_transaction_type=")
			arrValues(2)=ParseString(strMsg,"pg_total_amount=")
			arrValues(3)=ParseString(strMsg,"ecom_billto_postal_name_first=")
			arrValues(4)=ParseString(strMsg,"ecom_billto_postal_name_last=")
			arrValues(5)=ParseString(strMsg,"pg_response_type=")
			pcv_response_code=trim(ParseString(strMsg,"pg_response_code="))
			arrValues(6)=pcv_response_code
			pcv_response_description=trim(ParseString(strMsg,"pg_response_description="))
			arrValues(7)=pcv_response_description
			pcv_authorization_code=replace(trim(ParseString(strMsg,"pg_authorization_code="))," ","")
			arrValues(8)=pcv_authorization_code
			pcv_trace_number=trim(ParseString(strMsg,"pg_trace_number="))
			arrValues(9)=pcv_trace_number
	
			Response.Write "<center><P><B><font size='5'>Response For Your Credit Card Transaction</font></B></P>"
			Response.Write "<Table border=0>"
			FOR i=0 TO UBOUND(arrValues)
				IF arrValues(i) > "" THEN
					Response.Write "<tr>"
					Response.Write "<td align=right>Array #: "&i&" :" & arrNames(i) & ":&nbsp;</td>"
					Response.Write "<td>" & arrValues(i) & "</td>"
					Response.Write "</tr>"
				END IF
			NEXT
			Response.Write "</Table></center>"

			'===================
			
			ACH_success=0

			If cstr(left(pcv_response_code,3))="A01" Then
				 ACH_success=1
			End if

			Dim ACH_rd_successurl, ACH_rd_resultfailurl
		
			If ACH_success=1 Then
				session("GWAuthCode")=pcv_authorization_code
				session("GWTransId")=pcv_trace_number
				session("GWTransType")=pcPay_ACH_TransType
				ACH_rd_successurl="gwReturn.asp?s=true&gw=ACH"
			end if

			If (ACH_success <> 1) then
				strErrorInfo=""
				If instr(pcv_response_code,"U") then
					strErrorInfo="Declined: " & pcv_response_description
				End If
				If instr(pcv_response_code,"F") then
					strErrorInfo="Formatting Error: " & pcv_response_description
				End If
				If instr(pcv_response_code,"E") then
					strErrorInfo="Fatal Exception Error: " & pcv_response_description
				End If

				If (strErrorInfo="") Then
					strErrorInfo="There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."
				End if
			End if
			ACH_rd_resultfailurl="msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& strErrorInfo &"<br><br><a href="""&tempURL&"?psslurl=gwACH.asp&idCustomer="&session("idCustomer")&"&idOrder="&	session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")

			If ACH_success <> 1 Then
				Response.Redirect ACH_rd_resultfailurl
			ElseIf ACH_success=1 Then
				call closeDb()
				Response.Redirect ACH_rd_successurl
			End If
		End Sub
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
						<td><p><%=pcBillingFirstName&" "&pcBillingLastName%></p></td>
					</tr>
					<tr>
						<td><p><%=pcBillingAddress%></p></td>
					</tr>
					<% if pcBillingAddress2<>"" then %>
					<tr>
						<td><p><%=pcBillingAddress2%></p></td>
					</tr>
					<% end if %>
					<tr>
						<td><p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% if pcPay_ACH_TestMode=1 then %>
					<tr>
						<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
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
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
						<td>
                            <select name="CardType">
								<% dim ArryCardTypes, strCardType, j
                                ArryCardTypes=split(pcPay_ACH_CardTypes,",")
                                for j=0 to ubound(ArryCardTypes) 
									strCardType=ArryCardTypes(j) 
									select case strCardType
										case "VISA"
											response.write "<option value='VISA'>VISA</option>"
										case "MAST"
											response.write "<option value='MAST'>Master Card</option>"
										case "AMER"
											response.write "<option value='AMER'>American Express</option>"
										case "DISC"
											response.write "<option value='DISC'>Discover Card</option>"
										case "DINE"
											response.write "<option value='DINE'>Diners Club</option>"
										case "JCB"
											response.write "<option value='JCB'>JCB</option>"
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