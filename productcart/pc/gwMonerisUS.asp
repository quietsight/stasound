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
session("redirectPage")="gwMonerisUS.asp"

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
query="SELECT pcPay_Moneris_StoreId, pcPay_Moneris_Key, pcPay_Moneris_TransType, pcPay_Moneris_Lang, pcPay_Moneris_Testmode, pcPay_Moneris_CVVEnabled, pcPay_Moneris_Meth FROM pcPay_Moneris Where pcPay_Moneris_Id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_Moneris_StoreId=rs("pcPay_Moneris_StoreId")
pcPay_Moneris_StoreId=enDeCrypt(pcPay_Moneris_StoreId, scCrypPass)
pcPay_Moneris_Key=rs("pcPay_Moneris_Key")
pcPay_Moneris_Key=enDeCrypt(pcPay_Moneris_Key, scCrypPass)
pcPay_Moneris_TransType=rs("pcPay_Moneris_TransType")
pcPay_Moneris_Lang=rs("pcPay_Moneris_Lang")
pcPay_Moneris_Testmode=rs("pcPay_Moneris_Testmode")
pcv_CVV=rs("pcPay_Moneris_CVVEnabled")
pcPay_Moneris_Meth = rs("pcPay_Moneris_Meth")
set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

'*************************************************************************************
' This is where you would post info to the gateway
' START
'*************************************************************************************
	'// Validate Input fields
	strMessage=""
	'CC Number
	if request.Form("CardNumber")="" then
		'CC Number is required
		strMessage=strMessage&"Credit Card Number is required.<br>"
	end if
	
	If pcv_CVV="1" Then
		if request.Form("avs_street_number")="" then
			strMessage=strMessage&"Street Number is required.<br>"
		end if
		if request.Form("avs_street_name")="" then
			strMessage=strMessage&"Street Name is required.<br>"
		end if
		'//  2 required variables for CVD
		if request.Form("CVV")="" then
			strMessage=strMessage&dictLanguage.Item(Session("language")&"_GateWay_11")&" is required.<br>"
		end if
	end if
	
	if strMessage<>"" then
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Errors&nbsp;</b>:<br>"&strMessage&"<br><br><a href="""&tempURL&"?psslurl=gwMonerisUS.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
	end if

	'// Validates expiration
	if DateDiff("d", Month(Now)&"/"&Year(now), request("expMonth")&"/20"&request("expYear"))<=-1 then
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_6")&"<br><br><a href="""&tempURL&"?psslurl=gwMonerisUS.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
	end if

	if NOT CheckCC(request.Form("CardNumber")) then
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_5")&"<br><br><a href="""&tempURL&"?psslurl=gwMonerisUS.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
	end if
		
	Dim objXMLHTTP, xml
	
	if pcPay_Moneris_TestMode="1" then
		pcBillingTotal="1.00"
	end if

		'Send the request 
		stext="dp_id="&pcPay_Moneris_StoreId
		stext=stext & "&dp_key="&pcPay_Moneris_Key
		stext=stext & "&amount=" & replace(money(pcBillingTotal),",","")
		stext=stext & "&crypt_type=7"
		'stext=stext & "&lang=" & pcPay_Moneris_Lang
		'// eFraud Information
	If pcv_CVV="1" Then
		'// add the 3 required variables for AVS
		stext=stext & "&avs_street_number="& request.Form("avs_street_number")
		stext=stext & "&avs_street_name="& request.Form("avs_street_name")
		stext=stext & "&avs_zipcode="&pcBillingPostalCode
		'//  2 required variables for CVD
		stext=stext & "&cvd_value="&request.Form("CVV")
		'stext=stext & "&cvd_indicator=1"
	end if
		stext=stext & "&cc_num="& request.Form("CardNumber")
		stext=stext & "&exp_month="& request.Form("expMonth")
		stext=stext & "&exp_year="& request.Form("expYear")
		pcStrCustRefID = Session.SessionID & "-" & Hour(Now) & Minute(Now) & Second(Now)
		
		stext=stext & "&order_no="&pcStrCustRefID
		stext=stext & "&cust_id=" & session("GWOrderID")
		stext=stext & "&email=" & pcCustomerEmail
		stext=stext & "&bill_first_name=" & pcBillingFirstName
		stext=stext & "&bill_last_name=" & pcBillingLastName
		stext=stext & "&bill_company_name=" & replace(pcBillingCompany,",","||")
		stext=stext & "&bill_address_one=" & replace(pcBillingAddress,",","||")
		stext=stext & "&bill_city=" & pcBillingCity
		stext=stext & "&bill_state_or_province=" & pcBillingState
		stext=stext & "&bill_postal_code=" & pcBillingPostalCode
		stext=stext & "&bill_country=" & pcBillingCountryCode
		stext=stext & "&bill_phone=" & pcBillingPhone
		stext=stext & "&ship_first_name=" & pcShippingFirstName
		stext=stext & "&ship_last_name=" & pcShippingLastName
		stext=stext & "&ship_company_name=" & pcShippingCompany
		stext=stext & "&ship_address_one=" & replace(pcShippingAddress,",","||")
		stext=stext & "&ship_city=" & pcShippingCity
		stext=stext & "&ship_state_or_province=" & pcShippingState
		stext=stext & "&ship_postal_code=" & pcShippingPostalCode
		stext=stext & "&ship_country=" & pcShippingCountryCode
 
	if pcPay_Moneris_TestMode="1" then

	    strHostURL="https://esplusqa.moneris.com/DPHPP/index.php"
	
	else

	    strHostURL="https://esplus.moneris.com/DPHPP/index.php"
	
	end if

	dim resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	
	resolveTimeout	= 5000
	connectTimeout	= 5000
	sendTimeout		= 5000
	receiveTimeout	= 10000
	
	'Send the transaction info as part of the querystring
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	if pcPay_Moneris_Meth ="1"  then 
		xml.open "POST", strHostURL &"", false
		xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xml.send(stext)
	 Else
		xml.open "GET", strHostURL &"?" &stext & "", false
		xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xml.send "" 
	End if  
  ' Response.write strHostURL & stext & "<BR><BR>"
  ' response.end	
	strRetVal = xml.responseText
	Session("MonerisTransKey")=strRetVal
	response.write strRetVal
	response.end	

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
								<div class="pcErrorMessage"><%=Msg%></div></td>
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
					<% if pcPay_Moneris_TestMode="1" then %>
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
						<td width="23%"><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td width="77%"> 
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
					<% If pcv_CVV="1" Then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
						  <td colspan="2"><p>Please verify your credit card billing address information below. This is the address that your credit card statement is mailed to. If you need to alter the postal code, please click on the <a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a> above to change your billing address. </p>				  
						</tr>
						<tr>
							<td><p>Street number:</p>
							<td><p>
								<input name="avs_street_number" type="text" value="" size="10">
							&nbsp;&nbsp;Street Name: &nbsp;&nbsp;
							<input type="text" name="avs_street_name" value="">
							</p></td>
							</tr>
						<tr>
							<td><p>Postal Code :</p>
							<td><p><%=pcBillingPostalCode%></p></td>
						</tr>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
								<p><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></p></td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% End If %>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><p><%=money(pcBillingTotal)%></p></td>
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
'****************************************************************************
' Name: Credit Card Validation (Luhn Formula)
' Description:Uses the Luhn formula to quickly validate a credit card. Basically
'     all the digits except for the last one are summed together and the output is a s
'     ingle digit (0 to 9). This digit is compared with the last digit ensure a proper
'     credit card number is entered (Does not actually confirm that is is a real numbe
'     r, just that it is likely to be one. Example: Entering "4000-0000-0000-0002" wil
'     l pass the check, but "4000-0000-0000-0003" will not pass.)
' By: BrewGuru99
'
'This code is copyrighted and has    ' limited warranties.Please see
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6581&lngWId=4
'for details.
'****************************************************************************
    
function CheckCC(CCNo)
	Dim i, w, x, y
	
	y = 0
	CCNo = Replace(Replace(Replace(CStr(CCNo), "-", ""), " ", ""), ".", "") 'Ensure proper format of the input
	'Process digits from right to left, drop
	'     last digit if total length is even
	w = 2 * (Len(CCNo) Mod 2)
	For i = Len(CCNo) - 1 To 1 Step -1
		x = Mid(CCNo, i, 1)
		if IsNumeric(x) Then
			Select Case (i Mod 2) + w
				Case 0, 3 'Even Digit - Odd where total length is odd (eg. Visa vs. Amx)
					y = y + CInt(x)
				Case 1, 2 'Odd Digit - Even where total length is odd (eg. Visa vs. Amx)
					x = CInt(x) * 2
					if x > 9 Then
						'Break the digits (eg. 19 becomes 1 + 9)
						'     
						y = y + (x \ 10) + (x - 10)
					Else
						y = y + x
					End if
			End Select
   		End if
    Next
    'Return the 10's complement of the total
    '     
    y = 10 - (y Mod 10)
    if y > 9 Then y = 0
    CheckCC = (CStr(y) = Right(CCNo, 1))
End function
%>
<!--#include file="footer.asp"-->