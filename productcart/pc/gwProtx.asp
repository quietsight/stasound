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
<!--#include file="protx_functions.asp"-->

<!--#include file="header.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwProtx.asp"

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
query="SELECT Protxid,ProtxPassword,ProtxTestmode,ProtxCurcode,TxType,avs FROM protx Where idProtx=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_VendorName=rs("Protxid")
pcv_rVendorName=pcv_VendorName
pcv_password=rs("ProtxPassword")
	'decrypt
	pcv_password=enDeCrypt(pcv_password, scCrypPass)
ProtxTestmode=rs("ProtxTestmode")
pcv_CurCode=rs("ProtxCurcode")
pcv_TxType=rs("TxType")
pcv_AVVS = rs("avs")

If ProtxTestmode=1 Then
	'Simulator Account
	vspsite="https://test.sagepay.com/Simulator/VSPFormGateway.asp"
Else ' Live Mode
	If ProtxTestmode=2 then
		vspsite="https://test.sagepay.com/gateway/service/vspform-register.vsp"
	Else
		vspsite="https://live.sagepay.com/gateway/service/vspform-register.vsp"
	End if
End If	

If request.QueryString("crypt")<>"" then
	' ** Decrypt the plaintext string for inclusion in the hidden field **
	pcv_ReceiptCrypt=request.QueryString("crypt")
	pcv_ReceiptString =SimpleXor(base64decode(pcv_ReceiptCrypt),pcv_password)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
	pcv_ResponseArray = Split(pcv_ReceiptString, "&") 
	Set objDictResponse = server.createobject("Scripting.Dictionary")
	For each ResponseItem in pcv_ResponseArray
		pcv_NameValue = Split(ResponseItem, "=")
		objDictResponse.Add pcv_NameValue(0), pcv_NameValue(1)
	Next
       
	' Parse the response into local vars
	pcv_StrStatus = objDictResponse.Item("Status")
	pcv_StrVendorTxCode  = objDictResponse.Item("VendorTxCode")
	If ProtxTestmode =1 Then
		pcv_StrVendorTxCode=replace(pcv_StrVendorTxCode,""&pcv_rVendorName&"","")
	End if
	pcv_StrTxAuthNo = objDictResponse.Item("TxAuthNo")
	pcv_StrAVSCV2 = objDictResponse.Item("AVSCV2")
	pcv_StrVPSTxID = objDictResponse.Item("VPSTxID")
       
	If pcv_StrStatus="OK" then
		session("GWAuthCode")=pcv_StrTxAuthNo
		session("GWTransId")=pcv_StrVPSTxID
		session("GWTransType")=pcv_TxType
		if session("GWOrderId")="" then
			session("GWOrderId")=session("ProtxOrdno")
		end if
		session("GWSessionID")=Session.SessionID 
		Response.redirect "gwReturn.asp?s=true&gw=SagePay"
	Else
		if pcv_strStatus = "ABORT" then 
			   StatusOutput = "You elected to cancel your online payment<BR>Any credit/debit card details you entered have not been sent to the bank. You will not be charged for this transaction. Press the BACK button to try again."
		elseif pcv_strStatus = "NOTAUTHED" then 
			   StatusOutput = "The VSP was unable to authorise your payment<BR>The acquiring bank would not authorise your selected method of payment. You will not be charged for this transaction. Press the BACK button to try again."
		else
			   StatusOutput = "An error has occurred at SagePay<br>Because an error occurred in the payment process, you will not be charged for this transaction, even if an authorisation was given by the bank. Press the BACK button to try again."
		end if
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;"&StatusOutput&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&session("amount")&"""><img src="""&rslayout("back")&""" border=0></a>")
		response.end
	end if
end if

if pcShippingFullName<>"" then
	pcShippingNameArry=split(pcShippingFullName, " ")
	if ubound(pcShippingNameArry)>0 then
		pcShippingFirstName=pcShippingNameArry(0)
		if ubound(pcShippingNameArry)>1 then
			 tmpShipFirstName = pcShippingFirstName&" "
			 pcShippingLastName = replace(pcShippingFullName,tmpShipFirstName,"")
		else
			pcShippingLastName=pcShippingNameArry(1)
		end if
	else
		pcShippingFirstName=pcShippingFullName
		pcShippingLastName=pcShippingFullName
	end if
else
	pcShippingFirstName=pcBillingFirstName
	pcShippingLastName=pcBillingLastName
end if

if len(pcShippingLastName)> 0 then
	pcShippingLastName =  left(pcBillingLastName,20)
end if

if len(pcShippingFirstName)> 0 then
	pcShippingFirstName = left(pcBillingFirstName,20)
end if

if len(pcShippingAddress)> 0 then
	pcShippingAddress =  left(pcBillingAddress,100)
	pcShippingAddress2 = left(pcBillingAddress2,100)
	pcShippingCity = left(pcBillingCity,40)
	pcShippingPostalCode = left(pcBillingPostalCode,10)
	pcShippingCountryCode = left(pcBillingCountryCode,2)
	pcShippingPhone = left(pcBillingPhone,20)
end if

if len(pcShippingStateCode)> 0 then
	If ucase(pcBillingCountryCode) = "US" then
		pcShippingStateCode = left(pcBillingStateCode,2)
	End If	
end if
                    
if scSSL="1" then
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwprotx.asp"),"//","/")
else
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwprotx.asp"),"//","/")
end if
tempURL=replace(tempURL,"http:/","http://")
tempURL=replace(tempURL,"https:/","https://")

ThisVendorTxCode = session("GWOrderId") & "_" & Hour(Now) & Minute(Now) & Second(Now)
If ProtxTestmode =1 Then
	ThisVendorTxCode = ThisVendorTxCode & timer() & rnd()
End if

pcvPostString = "VendorTxCode=" & ThisVendorTxCode & "&"
pcvPostString = pcvPostString + "Amount=" & money(pcBillingTotal) & "&"
pcvPostString = pcvPostString + "Currency=" & pcv_CurCode & "&"
DescString = "" 
for f = 1 to ppcCartIndex
 DescString = DescString  & replace(pcCartArray(f,1),"'","") & " || " 
next
if len(DescString)>0 then
	DescString = left(left(DescString,(len(DescString) - 3)), 100)
else
	DescString=scCompanyName& " Order"
end if
pcvPostString = pcvPostString + "Description="& DescString &"&"
pcvPostString = pcvPostString + "SuccessURL=" & tempURL & "&"
pcvPostString = pcvPostString + "FailureURL=" & tempURL & "&"
if pcCustomerEmail<>"" then
	pcvPostString = pcvPostString + "CustomerEMail=" & pcCustomerEmail & "&"
end if

pcvPostString = pcvPostString + "BillingSurname=" & pcBillingLastName & "&"
pcvPostString = pcvPostString + "BillingFirstnames=" & pcBillingFirstName & "&"
pcvPostString = pcvPostString + "BillingAddress1=" & pcBillingAddress & "&"
pcvPostString = pcvPostString + "BillingAddress2=" & pcBillingAddress2 & "&"
pcvPostString = pcvPostString + "BillingCity=" & pcBillingCity & "&"
pcvPostString = pcvPostString + "BillingCountry=" & pcBillingCountryCode & "&"
pcvPostString = pcvPostString + "DeliverySurname=" & pcShippingFirstName & "&"
pcvPostString = pcvPostString + "DeliveryFirstnames=" & pcShippingLastName & "&"
pcvPostString = pcvPostString + "DeliveryAddress1=" & pcShippingAddress & "&"
pcvPostString = pcvPostString + "DeliveryCity=" & pcShippingCity & "&"
pcvPostString = pcvPostString + "DeliveryPostCode=" & pcShippingPostalCode & "&"
pcvPostString = pcvPostString + "DeliveryCountry=" & pcShippingCountryCode & "&"
pcvPostString = pcvPostString + "BillingPhone=" & pcBillingPhone & "&"
pcvPostString = pcvPostString + "CustomerName=" & pcBillingFirstName&" "&pcBillingLastName & "&"
pcvPostString = pcvPostString + "BillingPostCode=" & pcBillingPostalCode &"&" 
If Ucase(pcBillingCountryCode)="US" then
	pcvPostString = pcvPostString + "BillingState=" & pcBillingStateCode &"&"
End If
If Ucase(pcBillingCountryCode)="US" then
	pcvPostString = pcvPostString + "DeliveryState=" & pcShippingStateCode &"&"
End If
pcvPostString = pcvPostString + "ApplyAVSCV2=" & pcv_AVVS
' ** Encrypt the plaintext string for inclusion in the hidden field **
pcv_Crypt = base64Encode(SimpleXor(pcvPostString,pcv_password))

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

session("redirectPage2")=vspsite

set rs=nothing
call closedb()

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
				<form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="VPSProtocol" value="2.23">
					<input type="hidden" name="TxType" value="<%=ucase(pcv_TxType)%>">
					<input type="hidden" name="Vendor" value="<%=pcv_VendorName%>">
					<input type="hidden" name="Crypt" value="<%=pcv_Crypt %>">
					<input type="hidden" name="BillingFirstnames" value="<%=pcBillingFirstName%>">
					<input type="hidden" name="BillingSurname" value="<%=pcBillingLastName%>">
					<input type="hidden" name="BillingAddress1" value="<%=pcBillingAddress%>">
					<input type="hidden" name="BillingCity" value="<%=pcBillingCity%>">
                    <% If Ucase(pcBillingCountryCode)="US" then %>
						<input type="hidden" name="BillingState" value="<%=pcBillingStateCode%>">
                    <% End If %>
					<input type="hidden" name="BillingPostCode" value="<%=pcBillingPostalCode%>">
					<input type="hidden" name="BillingCountry" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="BillingPhone" value="<%=pcBillingPhone%>">
					
                    <input type="hidden" name="DeliveryFirstnames" value="<%=left(pcShippingFirstName,20)%>">
					<input type="hidden" name="DeliverySurname" value="<%=left(pcShippingLastName,20)%>">
					<input type="hidden" name="DeliveryAddress1" value="<%=pcShippingAddress%>">
					<input type="hidden" name="DeliveryCity" value="<%=pcShippingCity%>">
                    <% If Ucase(pcBillingCountryCode)="US" then %>
						<input type="hidden" name="DeliveryState" value="<%=pcShippingStateCode%>">
                    <% End If %>
					<input type="hidden" name="DeliveryPostCode" value="<%=pcShippingPostalCode%>">
					<input type="hidden" name="DeliveryCountry" value="<%=pcShippingCountryCode%>">
					<input type="hidden" name="DeliveryPhone" value="<%=pcShippingPhone%>">
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
					<% if ProtxTestmode="1" OR ProtxTestmode="2" then %>
					<tr>
						<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% end if %>

					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					<tr>
						<td class="pcSpacer" colspan="2"></td>
					</tr>
					<tr>
						<td colspan="2">
						<p>NOTE: When you click on the 'Place Order' button, you will temporarily leave our Web site and will be taken to a secure payment page on the SagePay Web site. You will be redirected back to our store once the transaction has been processed. We have partnered with SagePay, a leader in secure Internet payment processing, to ensure that your transactions are processed securely and reliably.</p>
						</td>
					</tr>
					<tr>
						<td class="pcSpacer" colspan="2"></td>
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