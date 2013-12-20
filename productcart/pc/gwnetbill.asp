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
session("redirectPage")="gwNetBill.asp"

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
query="SELECT NBAccountID, NBCVVEnabled, NBAVS, NBTranType, NBSiteTag FROM netbill Where idNetbill=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
NBAccountID=rs("NBAccountID")
'decrypt
NBAccountID=enDeCrypt(NBAccountID, scCrypPass)
pcv_CVV=rs("NBCVVEnabled")
If pcv_CVV ="1" Then
	NBdisableCVV="0"
Else 
	NBdisableCVV="1"   
End If
NBAVS=rs("NBAVS")
If NBAVS ="1" Then
	NBdisableAVS="0"
Else 
	NBdisableAVS="1"   
End If
NBTranType=rs("NBTranType")
NBSiteTag=rs("NBSiteTag")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	
	Dim objHTTP, strRequest, DLLUrl
	
	DLLUrl="https://secure.netbilling.com:1402/gw/sas/direct3.0"
	
	strRequest ="pay_type="&"C"&_
	"&account_id="&NBAccountID&_
	"&site_tag="&NBSiteTag&_
	"&tran_type="&NBTranType&_
	"&amount="&pcBillingTotal&_   
	"&card_number="&request.Form("CardNumber")&_ 
	"&card_cvv2="& request.Form("CVV")&_	
	"&card_expire="&request.Form("expMonth")&request.Form("expYear")&_
	"&disable_avs="&NBdisableAVS&_
	"&disable_cvv2="&NBdisableCVV&_
	"&bill_name1="&pcBillingFirstName&_
	"&bill_name2="&pcBillingLastName&_
	"&bill_street="&pcBillingAddress&_
	"&bill_city="&pcBillingCity&_
	"&bill_state="&pcBillingState&_
	"&bill_zip="&pcBillingPostalCode&_
	"&bill_country="&pcBillingCountryCode&_
	"&cust_phone="&pcBillingPhone&_
	"&cust_email="&pcCustomerEmail&_
	"&ship_name1"&pcShippingFirstName&" "&pcShippingLastName&_
	"&ship_street="&pcShippingaddress&_
	"&ship_city="&pcShippingCity&_
	"&ship_state="&pcShippingSstate&_
	"&ship_zip"&pcShippingPostalCode&_
	"&ship_country="&pcShippingCountryCode&_
	"&cust_ip="&pcCustIpAddress
					
	set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP"&scXML)
	objHttp.open "POST", DLLUrl, false
	objHttp.setRequestHeader "Host", "secure.netbilling.com:1402"
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
		strstatus=objDictResponse.Item("status_code")
		strstatusmsg=objDictResponse.Item("auth_msg")
		strAuthorizationNumber=objDictResponse.Item("auth_code")
		strtrans_id=objDictResponse.Item("trans_id")
			
		if strstatus = "0" Then
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error:&nbsp;&nbsp;"&strstatus&"&nbsp;due to&nbsp;"&lcase(strstatusmsg)&"<br><br><a href="""&tempURL&"?psslurl=gwnetbill.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			response.end
		End If  
		
		'save and update Netbillorders table
		If strstatus = "T" or strstatus = "1" Then
			
			call opendb()
			
			pcv_SecurityPass = pcs_GetSecureKey
			pcv_SecurityKeyID = pcs_GetKeyID
		
			dim pCardNumber, pCardNumber2
			pCardNumber=request.Form("cardNumber")
			pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)
			'save info in NetbillOrders if AUTH_CAPTURE
			pTempOrderID=(int(session("GWOrderId"))-scpre)
			if NBTranType="A" then
				query="INSERT INTO netbillorders (idOrder, amount,paymentmethod,transtype,authcode,ccnum,ccexp,idCustomer,fname,lname,address,zip,trans_id,captured,pcSecurityKeyID) VALUES ("&pTempOrderID&", "&pcBillingTotal&",'CC','"&NBTranType&"','"&strAuthorizationNumber&"','"&pCardNumber2&"','"&request.Form("expMonth")&request.Form("expYear")&"',"&session("idCustomer")&",'"&replace(pcBillingFirstName,"'","''")&"','"&replace(pcBillingFirstName,"'","''")&"','"&replace(pcBillingAddress,"'","''")&"','"&pcBillingPostalCode&"','"&strtrans_id&"',0,"&pcv_SecurityKeyID&");"
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
			
			session("GWAuthCode")=strAuthorizationNumber
			session("GWTransId")=strtrans_id
			session("GWTransType")=NBTranType
			Response.redirect "gwReturn.asp?s=true&gw=Netbill"
				
		End if
	Else   
		strStatus= "Transaction Failed...<BR>"
		strStatus= strStatus&"<BR> Https Response <BR>" 
		strStatus= strStatus&"Status Code= " & objHttp.status       & "<BR>"
		strStatus= strStatus&"Error Status = " & objHttp.statusText   & "<BR>"
		strStatus= strStatus&"Header     = " & objHttp.getAllResponseHeaders &"<BR>"
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error:&nbsp;&nbsp;"&strStatus&"<br><a href="""&tempURL&"?psslurl=gwnetbill.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		response.end
	End If   
	Set objHttp   = Nothing

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
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
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
					<% if pcv_CVV="1" then %>
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