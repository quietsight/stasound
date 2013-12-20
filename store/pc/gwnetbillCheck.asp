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
session("redirectPage")="gwNetBillCheck.asp"

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
query="SELECT NBAccountID,NBCVVEnabled,NBAVS,NBSiteTag FROM netbill Where idNetbill=1;"
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
NBAVS=rs("NBAVS")
NBTranType="S"
NBSiteTag=rs("NBSiteTag")
If NBAVS ="1" Then
	NBAVS="y"
Else 
	NBAVS="n"   
End If

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim objHTTP, strRequest, DLLUrl
	
	DLLUrl="https://secure.netbilling.com:1402/gw/sas/direct3.0"
				
	strRequest ="pay_type="&"K"&_
	"&account_id="&NBAccountID&_
	"&site_tag="&NBSiteTag&_
	"&tran_type="&NBTranType&_
	"&amount="&pcBillingTotal&_   
	"&account_number="&request.Form("bank_aba_code")&"%3A"&request.Form("bank_acct_num")&_ 
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
	"&ship_street="&pcShippingAddress&_
	"&ship_city="&pcShippingCity&_
	"&ship_state="&pcShippingState&_
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
		strAuthorizationNumber=objDictResponse.Item("auth_date")
		strTransactionID=objDictResponse.Item("trans_id")
			
		If strstatus = "1" and strTransactionID <> "" Then 
			session("GWAuthCode")=strAuthorizationNumber
			session("GWTransId")=strTransactionID
			Response.redirect "gwReturn.asp?s=true&gw=NetbillCheck"
		Else

			if strstatus = "0" then
				strstatus = "Failed transaction"
			end if
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error:&nbsp;&nbsp;"&strstatus&"&nbsp;due to&nbsp;"&lcase(strstatusmsg)&"<br><br><a href="""&tempURL&"?psslurl=gwNetBillCheck.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
			response.end
		End If  
	Else   
		strStatus= "Transaction Failed...<BR>"
		strStatus= strStatus&"<BR> Https Response <BR>" 
		strStatus= strStatus&"Status Code= " & objHttp.status       & "<BR>"
		strStatus= strStatus&"Error Status = " & objHttp.statusText   & "<BR>"
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error:&nbsp;&nbsp;"&strStatus&"<br><a href="""&tempURL&"?psslurl=gwNetBillCheck.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		
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
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
					  <td colspan="2" align="center"><img src="images/sampleck.gif" width="390" height="230"></td>
				  </tr>
						<tr> 
							<td><p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_12")%></p></td>
							<td>  
								<input name="bank_acct_name" type="text" size="35" maxlength="50">
							</td>
						</tr>
						<tr> 
							<td><p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_13")%></p></td>
							<td>  
									<input name="bank_aba_code" type="text" size="35">
							</td>
						</tr>
						<tr> 
							<td><p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_14")%></p></td>
							<td>  
								<input name="bank_acct_num" type="text" size="35">
							</td>
						</tr>
						<tr> 
							<td><p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_27")%></p></td>
							<td>  
								<input name="check_num" type="text" size="15">
							</td>
						</tr>
						<tr> 
							<td><p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_15")%></p></td>
							<td>
							<select name="bank_acct_type">
								<option value="CHECKING">Checking Account</option>
								<option value="SAVINGS">Savings Account</option>
							</select>
							</td>
						</tr>
						<tr> 
							<td><p><%response.write dictLanguage.Item(Session("language")&"_gwAuthorizeCheck_16")%></p></td>
							<td>  
								<input name="bank_name" type="text" size="20" maxlength="20">
							</td>
						</tr>
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