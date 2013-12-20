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
session("redirectPage")="gwHSBC.asp"

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
query="SELECT pcPay_HSBC_UserId, pcPay_HSBC_Password, pcPay_HSBC_ClientId, pcPay_HSBC_TransType, pcPay_HSBC_CVV, pcPay_HSBC_Currency, pcPay_HSBC_TestMode FROM pcPay_HSBC Where pcPay_HSBC_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_HSBC_UserID=rs("pcPay_HSBC_UserId")
pcPay_HSBC_Password=rs("pcPay_HSBC_Password")
pcPay_HSBC_Password=enDeCrypt(pcPay_HSBC_Password, scCrypPass)
pcPay_HSBC_ClientID=rs("pcPay_HSBC_ClientId")
pcPay_HSBC_TransType=rs("pcPay_HSBC_TransType")
pcv_CVV=rs("pcPay_HSBC_CVV")
if len(pcv_CVV)<1 then
	pcv_CVV=0
end if
pcPay_HSBC_Currency=rs("pcPay_HSBC_Currency")
pcPay_HSBC_TestMode=rs("pcPay_HSBC_TestMode")
if len(pcPay_HSBC_TestMode)<1 then
	pcPay_HSBC_TestMode=0
end if

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

'*************************************************************************************
' This is where you would post info to the gateway
' START
'*************************************************************************************
	Dim strPostURL
	strPostURL="https://www.secure-epayments.apixml.hsbc.com"
	
	'// Format Total				
	AMT=money(pcBillingTotal)
	if Instr(AMT,",")>0 then
		A=split(AMT,",")
	else
		if Instr(AMT,".")>0 then
			A=split(AMT,".")
		else
			AMT=AMT & ".00"
			A=split(AMT,".")
		end if
	end if
				
	if A(1)="" then
		A(1)="00"
	end if
				
	for i=len(A(1))+1 to 2
		A(1)=A(1) & "0"
	next
				
	AMT=A(0) & A(1)
	
	Dim strRequest
				
	strRequest=""
	
	strRequest="<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	strRequest=strRequest & "<EngineDocList>" & vbCrLf
	strRequest=strRequest & "	<DocVersion>1.0</DocVersion>" & vbCrLf
	strRequest=strRequest & "	<EngineDoc>" & vbCrLf
	strRequest=strRequest & "		<ContentType>OrderFormDoc</ContentType>" & vbCrLf
	strRequest=strRequest & "		<User>" & vbCrLf
	strRequest=strRequest & "			<Name>" & pcPay_HSBC_UserID & "</Name>" & vbCrLf
	strRequest=strRequest & "			<Password>" & pcPay_HSBC_Password & "</Password>" & vbCrLf
	strRequest=strRequest & "			<ClientId DataType=""S32"">" & pcPay_HSBC_ClientID & "</ClientId>" & vbCrLf
	strRequest=strRequest & "		</User>" & vbCrLf
	strRequest=strRequest & "		<Instructions>" & vbCrLf
	strRequest=strRequest & "			<Pipeline>PaymentNoFraud</Pipeline>" & vbCrLf
	strRequest=strRequest & "		</Instructions>" & vbCrLf
	strRequest=strRequest & "		<OrderFormDoc>" & vbCrLf
	if pcPay_HSBC_TestMode="1" then
		strRequest=strRequest & "			<Mode>Y</Mode>" & vbCrLf
	else
		strRequest=strRequest & "			<Mode>P</Mode>" & vbCrLf
	end if
	strRequest=strRequest & "			<Comments/>" & vbCrLf
	strRequest=strRequest & "			<Consumer>" & vbCrLf
	if trim(Request.Form( "EMAIL" ))<>"" then
		strRequest=strRequest & "				<Email>" & pcCustomerEmail & "</Email>" & vbCrLf
	else
		strRequest=strRequest & "				<Email/>" & vbCrLf
	end if
	strRequest=strRequest & "				<PaymentMech>" & vbCrLf
	strRequest=strRequest & "					<CreditCard>" & vbCrLf
	strRequest=strRequest & "						<Number>" & replace(Request.Form( "CardNumber" )," ","") & "</Number>" & vbCrLf
	strRequest=strRequest & "						<Expires DataType=""ExpirationDate"" Locale=""840"">" & Request.Form( "expMonth" ) & "/" & Request.Form( "expYear" ) & "</Expires>" & vbCrLf
	if pcPay_HSBC_CVV="1" then
		strRequest=strRequest & "						<Cvv2Val>" & Request.Form( "CVV" ) & "</Cvv2Val>" & vbCrLf
		strRequest=strRequest & "						<Cvv2Indicator>1</Cvv2Indicator>" & vbCrLf
	end if
	strRequest=strRequest & "					</CreditCard>" & vbCrLf
	strRequest=strRequest & "				</PaymentMech>" & vbCrLf
	strRequest=strRequest & "				<BillTo>" & vbCrLf
	strRequest=strRequest & "					<Location>" & vbCrLf
	if trim(pcBillingPhone)<>"" then
		strRequest=strRequest & "						<TelVoice>" & pcBillingPhone & "</TelVoice>" & vbCrLf
	else
		strRequest=strRequest & "						<TelVoice/>" & vbCrLf
	end if
	strRequest=strRequest & "						<TelFax/>" & vbCrLf
	strRequest=strRequest & "						<Address>" & vbCrLf
	strRequest=strRequest & "							<Name>" & pcBillingFirstName & " " & pcBillingLastName & "</Name>" & vbCrLf
	strRequest=strRequest & "							<Street1>" & pcBillingAddress & "</Street1>" & vbCrLf
	if pcBillingAddress2<>"" then
		strRequest=strRequest & "							<Street2>" & pcBillingAddress2 & "</Street2>" & vbCrLf
	else
		strRequest=strRequest & "						<Street2/>" & vbCrLf
	end if
	strRequest=strRequest & "							<City>" & pcBillingCity & "</City>" & vbCrLf
	strRequest=strRequest & "							<StateProv>" & pcBillingState & "</StateProv>" & vbCrLf
	strRequest=strRequest & "							<PostalCode>" & pcBillingPostalCode & "</PostalCode>" & vbCrLf
	strRequest=strRequest & "							<Country/>" & vbCrLf
	strRequest=strRequest & "							<Company/>" & vbCrLf
	strRequest=strRequest & "						</Address>" & vbCrLf
	strRequest=strRequest & "					</Location>" & vbCrLf
	strRequest=strRequest & "				</BillTo>" & vbCrLf
	strRequest=strRequest & "			</Consumer>" & vbCrLf
	strRequest=strRequest & "			<Transaction>" & vbCrLf
	strRequest=strRequest & "				<Type>" & pcPay_HSBC_TransType & "</Type>" & vbCrLf
	strRequest=strRequest & "				<CurrentTotals>" & vbCrLf
	strRequest=strRequest & "					<Totals>" & vbCrLf
	strRequest=strRequest & "						<Total DataType=""Money"" Currency=""" & pcPay_HSBC_Currency & """>" & AMT & "</Total>" & vbCrLf
	strRequest=strRequest & "					</Totals>" & vbCrLf
	strRequest=strRequest & "				</CurrentTotals>" & vbCrLf
	strRequest=strRequest & "			</Transaction>" & vbCrLf
	strRequest=strRequest & "		</OrderFormDoc>" & vbCrLf
	strRequest=strRequest & "	</EngineDoc>" & vbCrLf
	strRequest=strRequest & "</EngineDocList>" & vbCrLf

	'If Request.Form("HSBC")="Go" Then
	response.buffer=true
	response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
		
	hsbc_success=0
	
	Set HSBCXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		ErrString = ErrString&"2: "&err.description&"<BR>"

	err.clear
	HSBCXmlHttp.open "POST", strPostURL, false
	HSBCXmlHttp.send(strRequest)
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	HSBC_result = HSBCXmlHttp.responseText
	'response.write strRequest&"<HR>"
	'response.write HSBC_result
	'response.end
	
	'=========== NEW CODE ===================		
	Dim xmldoc
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument"&scXML)
	ErrString = ErrString&"3: "&err.description&"<BR>"
	xmlDoc.async = False
	pcResultHSBC_Result=""
	pcResultHSBC_Msg=""
	pcResultHSBC_ErrCode=""
	pcResultHSBC_ErrName=""
	pcResultHSBC_ErrMsg=""
	pcResultHSBC_TransID=""
	pcResultHSBC_AuthCode=""


	Dim hsbc_rd_successurl, hsbc_rd_resultfailurl


	If xmlDoc.loadXML(HSBC_result) Then
			'pcResultHSBC_Msg = xmlDoc.documentElement.selectSingleNode("CcReturnMsg").Text
			pcResultHSBC_Msg = xmlDoc.documentElement.selectSingleNode("//EngineDocList/EngineDoc/Overview/CcReturnMsg").Text

			'response.write "<HR>"&pcResultHSBC_Msg
			'response.end
			ErrString = ErrString&"5: "&err.description&"<BR>"
			if pcResultHSBC_Msg="Approved." then
				ErrString = ErrString&"5: "&err.description&"<BR>"

				hsbc_success=1
				pcResultHSBC_TransID=xmlDoc.documentElement.selectSingleNode("//EngineDocList/EngineDoc/Overview/TransactionId").Text
				ErrString = ErrString&"6: "&err.description&"<BR>"

				pcResultHSBC_AuthCode=xmlDoc.documentElement.selectSingleNode("//EngineDocList/EngineDoc/Overview/AuthCode").Text
				session("GWAuthCode")=pcResultHSBC_AuthCode
				session("GWTransId")=pcResultHSBC_TransID
				session("GWTransType")=pcPay_HSBC_TransType
				hsbc_rd_successurl="gwReturn.asp?s=true&gw=HSBC"
			else
				pcResultHSBC_ErrCode=xmlDoc.documentElement.selectSingleNode("//EngineDocList/EngineDoc/Overview/CcErrCode").Text
				pcResultHSBC_ErrMsg=xmlDoc.documentElement.selectSingleNode("//EngineDocList/EngineDoc/Overview/CcReturnMsg").Text
				strErrorInfo=pcResultHSBC_ErrCode&": "&pcResultHSBC_ErrMsg'&" - "&ErrString&" - "&err.description
			end if

	Else
		'//ERROR
		strErrorInfo = "Failed to process response - possible gateway communication failure"
	End If

	'=========== NEW CODE ===================		
				
	If (hsbc_success <> 1) and (strErrorInfo="") Then
		strErrorInfo="There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."
	End if
				
	hsbc_rd_resultfailurl="msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error</b>: "&strErrorInfo&"<br><br><a href="""&tempURL&"?psslurl=gwHSBC.asp&idCustomer="&session("idcustomer")&"&idOrder="&Session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
			
	If hsbc_success <> 1 Then
		Response.Redirect hsbc_rd_resultfailurl
	ElseIf hsbc_success=1 Then
		call closeDb()
		Response.Redirect hsbc_rd_successurl
	End If
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
					<% if pcPay_HSBC_TestMode="1" then %>
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
					<% If pcv_CVV="1" Then %>
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