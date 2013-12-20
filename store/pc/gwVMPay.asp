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
Dim iRoot,xmldoc
Function CheckExistTag(tagName)
Dim tmpNode
	Set tmpNode=iRoot.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		CheckExistTag=False
	Else
		CheckExistTag=True
	End if
End Function

' thisd is to clean out nay bad characters for VM XML parcer
function getUserVM_XMLOutPut(input,stringLength)
 dim tempStr

 known_bad= array("*","--")
 if stringLength>0 then
  tempStr	= left(trim(input),stringLength)
 else
  tempStr	= trim(input)
 end if
 for i=lbound(known_bad) to ubound(known_bad)
	if (instr(1,tempStr,known_bad(i),vbTextCompare)<>0) then
		tempStr	= replace(tempStr,known_bad(i),"")
	end if
 next
 tempStr	= replace(tempStr,"'","''")
 tempStr	= replace(tempStr,"<","")
 tempStr	= replace(tempStr,">","")
 tempStr	= replace(tempStr,"%0d","")
 tempStr	= replace(tempStr,"%0D","")
 tempStr	= replace(tempStr,"%0a","")
 tempStr	= replace(tempStr,"%0A","")
 tempStr	= replace(tempStr,"\r\n","")
 tempStr	= replace(tempStr,"\r","")
 tempStr	= replace(tempStr,"\n","")
 tempStr	= replace(tempStr,"\R\N","")
 tempStr	= replace(tempStr,"\R","")
 tempStr	= replace(tempStr,"\N","")
 tempStr	= replace(tempStr,"&","")
 tempStr	= replace(tempStr,"#","")
 tempStr	= replace(tempStr,"%","")
 tempStr	= replace(tempStr,"EXEC(","",1,-1,1)


	if tempStr<>"" then
		if IsNumeric(tempStr) then
			if InStr(Cstr(10/3),",")>0 then
				if Instr(tempStr,".")>0 then
					tempStr=FormatNumber(tempStr,,,,0)
					tempStr=replace(tempStr,".",",")
				end if
			end if
		end if
	end if

 getUserVM_XMLOutPut	= tempStr
end function

'//Set redirect page to the current file name
session("redirectPage")="gwVMPay.asp"

'//VirtualMerchant Gateway URL
Dim pcVMPayURL

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
<!--#include file="pcGatewayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables
query= "SELECT TOP 1 pcPay_VM_MerchantID,pcPay_VM_UserID,pcPay_VM_Pin,pcPay_VM_TransType,pcPay_VM_TestMode,pcPay_VM_CVV2 FROM pcPay_VirtualMerchant;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_VM_MerchantID=rs("pcPay_VM_MerchantID")
pcPay_VM_UserID=rs("pcPay_VM_UserID")
pcPay_VM_Pin=rs("pcPay_VM_Pin")
pcPay_VM_Pin=enDeCrypt(pcPay_VM_Pin, scCrypPass)
pcPay_VM_TransType=rs("pcPay_VM_TransType")
pcPay_VM_TestMode=rs("pcPay_VM_TestMode")
pcPay_VM_CVV2=rs("pcPay_VM_CVV2")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	strCardNumber=request("CardNumber")
	strCardNumber=replace(strCardNumber,"-","")
	strCardNumber=replace(strCardNumber," ","")
	strCardNumber=replace(strCardNumber,".","")

	pcVMPayURL="https://www.myvirtualmerchant.com/VirtualMerchant/processxml.do?xmldata="

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim SrvVMPayXmlHttp, pcVMPayXMLPostData
	pcVMPayXMLPostData=""
	pcVMPayXMLPostData = pcVMPayXMLPostData & "<txn>"

	if pcPay_VM_TestMode = "1" Then
	 pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_test_mode>true</ssl_test_mode>"
	else
	 pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_test_mode>false</ssl_test_mode>"
	end if
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_test_mode>" & pcPay_VM_TestMode &"</ssl_test_mode>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_transaction_type>"&pcPay_VM_TransType&"</ssl_transaction_type>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_merchant_id>"&pcPay_VM_MerchantID&"</ssl_merchant_id>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_pin>"&pcPay_VM_Pin&"</ssl_pin>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_user_id>"&pcPay_VM_UserID&"</ssl_user_id>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_amount>"&pcBillingTotal&"</ssl_amount>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_salestax>"&pcBillingTaxAmount&"</ssl_salestax>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_card_number>"&getUserVM_XMLOutPut(strCardNumber,20)&"</ssl_card_number>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_exp_date>"&getUserVM_XMLOutPut(request("expMonth")&request("expYear"),0)&"</ssl_exp_date>"
	if pcPay_VM_CVV2="1" then
		pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_cvv2cvc2_indicator>"&getUserVM_XMLOutPut(pcPay_VM_CVV2,0)&"</ssl_cvv2cvc2_indicator>"
		pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_cvv2cvc2>"&getUserVM_XMLOutPut(request("CVV"),0)&"</ssl_cvv2cvc2>"
	end if
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_description>Payment for the Order ID:" & session("GWOrderID") & "</ssl_description>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_invoice_number>" & session("GWOrderID") & "</ssl_invoice_number>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_customer_code>" & session("idCustomer") & "</ssl_customer_code>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_company>"&getUserVM_XMLOutPut(pcBillingCompany,50)&"</ssl_company>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_first_name>"&getUserVM_XMLOutPut(trim(pcBillingFirstName),20)&"</ssl_first_name>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_last_name>"&getUserVM_XMLOutPut(trim(pcBillingLastName),30)&"</ssl_last_name>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_avs_address>"&getUserVM_XMLOutPut(pcBillingAddress,20)&"</ssl_avs_address>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_address2>"&getUserVM_XMLOutPut(pcBillingAddress2,30)&"</ssl_address2>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_city>"&getUserVM_XMLOutPut(pcBillingCity,30)&"</ssl_city>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_state>"&getUserVM_XMLOutPut(pcBillingState,30)&"</ssl_state>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_avs_zip>"&getUserVM_XMLOutPut(pcBillingPostalCode,9)&"</ssl_avs_zip>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_country>"&getUserVM_XMLOutPut(pcBillingCountryCode,50)&"</ssl_country>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_phone>"&getUserVM_XMLOutPut(pcBillingPhone,20)&"</ssl_phone>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_email>"&getUserVM_XMLOutPut(pcCustomerEmail,100)&"</ssl_email>"
	If pcShippingAddress&""="" Then
		pcShippingCompany = pcBillingCompany
		pcShippingFirstName = pcBillingFirstName
		pcShippingLastName =pcBillingLastName
		pcShippingAddress = pcBillingAddress
		pcShippingAddress2 =  pcBillingAddress2
		pcShippingCity =pcBillingCity
		pcShippingState = pcBillingState
		pcShippingPostalCode = pcBillingPostalCode
		pcShippingCountryCode = pcBillingCountryCode
		pcShippingPhone = pcBillingPhone
	End If


	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_company>"&getUserVM_XMLOutPut(pcShippingCompany,50)&"</ssl_ship_to_company>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_first_name>"&getUserVM_XMLOutPut(trim(pcShippingFirstName),15)&"</ssl_ship_to_first_name>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_last_name>"&getUserVM_XMLOutPut(trim(pcShippingLastName),14)&"</ssl_ship_to_last_name>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_address1>"&getUserVM_XMLOutPut(pcShippingAddress,30)&"</ssl_ship_to_address1>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_address2>"&getUserVM_XMLOutPut(pcShippingAddress2,30)&"</ssl_ship_to_address2>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_city>"&getUserVM_XMLOutPut(pcShippingCity,30)&"</ssl_ship_to_city>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_state>"&getUserVM_XMLOutPut(pcShippingState,30)&"</ssl_ship_to_state>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_zip>"&getUserVM_XMLOutPut(pcShippingPostalCode,10)&"</ssl_ship_to_zip>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_country>"&getUserVM_XMLOutPut(pcShippingCountryCode,50)&"</ssl_ship_to_country>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ssl_ship_to_phone>"&getUserVM_XMLOutPut(pcShippingPhone,20)&"</ssl_ship_to_phone>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"</txn>"

	Set SrvVMPayXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	err.clear
	SrvVMPayXmlHttp.open "POST", pcVMPayURL & pcVMPayXMLPostData , false
	SrvVMPayXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	SrvVMPayXmlHttp.send()
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	VMPayResult = SrvVMPayXmlHttp.responseText
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.async = False
	pcResultVMPay_Result=""
	pcResultVMPay_Msg=""
	pcResultVMPay_ErrCode=""
	pcResultVMPay_ErrName=""
	pcResultVMPay_ErrMsg=""
	pcResultVMPay_TransID=""
	pcResultVMPay_AuthCode=""

	If xmlDoc.loadXML(SrvVMPayXmlHttp.responseText) Then
		Set iRoot=xmlDoc.documentElement
		' Get the results
		if CheckExistTag("ssl_result_message") then
			pcResultVMPay_Msg=iRoot.selectSingleNode("ssl_result_message").Text
			if pcResultVMPay_Msg="APPROVED" OR pcResultVMPay_Msg="APPROVAL" then
				pcResultVMPay_TransID=iRoot.selectSingleNode("ssl_txn_id").Text
				pcResultVMPay_AuthCode=iRoot.selectSingleNode("ssl_approval_code").Text
				pcResultVMPay_AVSCode=iRoot.selectSingleNode("ssl_avs_response").Text
				pcResultVMPay_CVV2Code=iRoot.selectSingleNode("ssl_cvv2_response").Text
			else
				pcResultVMPay_ErrMsg=pcResultVMPay_Msg
			end if
		else
			pcResultVMPay_ErrCode=iRoot.selectSingleNode("errorCode").Text
			pcResultVMPay_ErrName=iRoot.selectSingleNode("errorName").Text
			pcResultVMPay_ErrMsg=iRoot.selectSingleNode("errorMessage").Text
		end if
	Else
		'//ERROR
		response.write "Failed to process response"
		response.end
	End If

	if pcResultVMPay_ErrMsg="" then
		'create sessions
		session("GWAuthCode")=pcResultVMPay_AuthCode
		session("GWTransId")=pcResultVMPay_TransID
		session("GWTransType")=pcPay_VM_TransType
		session("AVSCode")=pcResultVMPay_AVSCode
		session("CVV2Code")=pcResultVMPay_CVV2Code
		closedb()

		Response.redirect "gwReturn.asp?s=true&gw=VM"
		response.end
	else
		closedb()
		response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;</b>: "& pcResultVMPay_ErrMsg &"<br><br><a href="""&tempURL&"?psslurl=gwVMPay.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
		response.end
	end if

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
					<%if pcPay_VM_Testmode="1" then %>
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
					<% If pcPay_VM_CVV2="1" Then %>
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