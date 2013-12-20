<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<% 'Gateway specific files %>
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td><img src="images/checkout_bar_step5.gif" alt=""></td>
	</tr>
	<tr>
		<td class="pcSpacer"></td>
	</tr>
	<tr>
		<td>
			
		<% if session("GWOrderDone")="YES" then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://") 
			session("GWOrderDone")=""
			response.redirect tempURL
		end if
		
		dim connTemp, rs
		%>
		<!-- #Include File="pcPay_Cent_XMLFunctions.asp"-->
		<!-- #Include File="pcPay_Cent_Utility.asp"-->

		<% session("redirectPage")="gwPaymentech.asp" %>
		
		<% Dim pcCustIpAddress
		pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
		
		dim tempURL
		If scSSL="" OR scSSL="0" Then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://") 
		Else
			tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://")
		End If
		
		' Get Order ID
		if session("GWOrderId")="" then
			session("GWOrderId")=request("idOrder")
		end if
		
		pcGatewayDataIdOrder=session("GWOrderID")
		%>
		<!--#include file="pcGateWayData.asp"-->
		<% session("idCustomer")=pcIdCustomer
		
		If Request.Form("PaymentGWCentinel")="Go" OR request.QueryString("centinel")<>"" Then %>
			
			<!--#include file="pcCentinelInclude.asp"-->
		
			<% call opendb()
			
			query="SELECT pcPay_PT_MerchantId, pcPay_PT_BIN, pcPay_PT_APIType, pcPay_PT_Testing, pcPay_PT_TransType, pcPay_PT_CVC, pcPay_PT_CurrencyCode FROM pcPay_Paymentech WHERE pcPay_PT_Id=1"
			set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
	
			pcPay_PT_MerchantId=rs("pcPay_PT_MerchantId")
			pcPay_PT_BIN=rs("pcPay_PT_BIN")
			pcPay_PT_APIType=rs("pcPay_PT_APIType")
			if len(pcPay_PT_APIType)<1 then
				pcPay_PT_APIType="XML"
			end if
			pcPay_PT_Testing=rs("pcPay_PT_Testing")
			pcPay_PT_TransType=rs("pcPay_PT_TransType")
			if pcPay_PT_TransType="" then
				pcPay_PT_TransType="A"
			end if
			pcPay_PT_CVC=rs("pcPay_PT_CVC")
			if len(pcPay_PT_CVC)<1 then
				pcPay_PT_CVC="0"
			end if
			pcPay_PT_CurrencyCode=rs("pcPay_PT_CurrencyCode")
			if len(pcPay_PT_CurrencyCode)<1 then
				pcPay_PT_CurrencyCode="840"
			end if
			pt_CurrencyExponent="2"
			if pcPay_PT_CurrencyCode="392" then
				pt_CurrencyExponent="0"
			end if
			
			set rs=nothing
			
			call closedb()

			pt_Amount=pcBillingTotal
			pt_Amount=replace(pt_Amount,",","")
			pt_Amount=replace(pt_Amount,".","")
		
			if pcBillingCountryCode<>"US" AND pcBillingCountryCode<>"UK" AND pcBillingCountryCode<>"GB" AND pcBillingCountryCode<>"CA" then
				pcBillingCountryCode=" "
			end if
			
			pt_Comments = "" 'comment about transactio
			pt_LangInd="00" 'English
			paymentech_URL="https://orbital1.paymentech.net/authorize"

			if pcPay_PT_Testing="1" then
				paymentech_URL="https://orbitalvar1.paymentech.net/authorize"
				session("reqCardNumber") = "4012888888881"
				session("reqExpDate") = "1206"
				pt_Comments = "Test Authorize Request"
				pcPay_PT_CVC="0"
			end if
	
			if pcPay_PT_APIType="XML" then
				Dim srvpaymentechXmlHttp, paymentech_postdata
				paymentech_postdata=""
				paymentech_postdata=paymentech_postdata&"<?xml version=""1.0"" encoding=""UTF-8""?>"
				paymentech_postdata=paymentech_postdata&"<Request>"
				paymentech_postdata=paymentech_postdata&"<AC>"
				paymentech_postdata=paymentech_postdata&"<CommonData>"
				paymentech_postdata=paymentech_postdata&"<CommonMandatory AuthOverrideInd=""N"" LangInd="""&pt_LangInd&""" CardHolderAttendanceInd=""01"" HcsTcsInd=""T"" TxCatg=""7"" MessageType="""&pcPay_PT_TransType&""" Version=""2"" TzCode="""&pt_TzCode&""">"
				paymentech_postdata=paymentech_postdata&"<AccountNum AccountTypeInd=""91"">"&session("reqCardNumber")&"</AccountNum>"
				paymentech_postdata=paymentech_postdata&"<POSDetails POSEntryMode=""01""/>"
				paymentech_postdata=paymentech_postdata&"<MerchantID>"&pcPay_PT_MerchantId&"</MerchantID>"
				paymentech_postdata=paymentech_postdata&"<TerminalID TermEntCapInd=""05"" CATInfoInd=""06"" TermLocInd=""01"" CardPresentInd=""N"" POSConditionCode=""59"" AttendedTermDataInd=""01"">001</TerminalID>"
				paymentech_postdata=paymentech_postdata&"<BIN>"&pcPay_PT_BIN&"</BIN>"
				paymentech_postdata=paymentech_postdata&"<OrderID>"&session("GWOrderId")&"</OrderID>"
				paymentech_postdata=paymentech_postdata&"<AmountDetails>"
				paymentech_postdata=paymentech_postdata&"<Amount>"&pt_Amount&"</Amount>"
				paymentech_postdata=paymentech_postdata&"</AmountDetails>"
				paymentech_postdata=paymentech_postdata&"<TxTypeCommon TxTypeID=""G""/>"
				paymentech_postdata=paymentech_postdata&"<Currency CurrencyCode="""&pcPay_PT_CurrencyCode&""" CurrencyExponent="""&pt_CurrencyExponent&"""/>"
				paymentech_postdata=paymentech_postdata&"<CardPresence>"
				paymentech_postdata=paymentech_postdata&"<CardNP>"
				paymentech_postdata=paymentech_postdata&"<Exp>"&session("reqExpMonth") & session("reqExpYear")&"</Exp>"
				paymentech_postdata=paymentech_postdata&"</CardNP>"
				paymentech_postdata=paymentech_postdata&"</CardPresence>"
				paymentech_postdata=paymentech_postdata&"<TxDateTime/>"
				paymentech_postdata=paymentech_postdata&"</CommonMandatory>"
				paymentech_postdata=paymentech_postdata&"<CommonOptional>"
				paymentech_postdata=paymentech_postdata&"<Comments>"&pt_Comments&"</Comments>"
				paymentech_postdata=paymentech_postdata&"<PCCore>"
				paymentech_postdata=paymentech_postdata&"<PCOrderNum>"&session("GWOrderId")&"</PCOrderNum>"
				paymentech_postdata=paymentech_postdata&"<PCDestZip>"&pcShippingPostalCode&"</PCDestZip>"
				paymentech_postdata=paymentech_postdata&"<PCDestName>"&pcShippingFirstName&" "&session("reqShipLastName")&"</PCDestName>"
				paymentech_postdata=paymentech_postdata&"<PCDestAddress1>"&pcShippingAddress&"</PCDestAddress1>"
				paymentech_postdata=paymentech_postdata&"<PCDestAddress2>"&pcShippingAddress2&"</PCDestAddress2>"
				paymentech_postdata=paymentech_postdata&"<PCDestCity>"&pcShippingCity&"</PCDestCity>"
				paymentech_postdata=paymentech_postdata&"<PCDestState>"&pcShippingState&"</PCDestState>"
				paymentech_postdata=paymentech_postdata&"</PCCore>"
				if pcPay_PT_CVC="1" then
					paymentech_postdata=paymentech_postdata&"<CardSecVal CardSecInd=""1"">"&session("reqCVV")&"</CardSecVal>"
				end if
				if session("Centinel_ECI")<>"" then
					paymentech_postdata=paymentech_postdata&"<ECommerceData ECSecurityInd="""&Session("Centinel_ECI")&""">"
				else
					paymentech_postdata=paymentech_postdata&"<ECommerceData ECSecurityInd=""07"">"
				end if
				paymentech_postdata=paymentech_postdata&"<ECOrderNum>"&session("GWOrderId")&"</ECOrderNum>"
				paymentech_postdata=paymentech_postdata&"</ECommerceData>"
				paymentech_postdata=paymentech_postdata&"</CommonOptional>"
				paymentech_postdata=paymentech_postdata&"</CommonData>"
				paymentech_postdata=paymentech_postdata&"<Auth>"
				paymentech_postdata=paymentech_postdata&"<AuthMandatory FormatInd=""H""/>"
				paymentech_postdata=paymentech_postdata&"<AuthOptional>"
				paymentech_postdata=paymentech_postdata&"<AVSextended>"
				paymentech_postdata=paymentech_postdata&"<AVSphoneNum>"&pcBillingPhone&"</AVSphoneNum>"
				paymentech_postdata=paymentech_postdata&"<AVSname>"&pcBillingFirstName&" "&pcBillingLastName&"</AVSname>"
				paymentech_postdata=paymentech_postdata&"<AVSaddress1>"&pcBillingAddress&"</AVSaddress1>"
				paymentech_postdata=paymentech_postdata&"<AVSaddress2>"&pcBillingAddress2&"</AVSaddress2>"
				paymentech_postdata=paymentech_postdata&"<AVScity>"&pcBillingCity&"</AVScity>"
				paymentech_postdata=paymentech_postdata&"<AVSstate>"&pcBillingState&"</AVSstate>"
				paymentech_postdata=paymentech_postdata&"<AVSzip>"&pcBillingPostalCode&"</AVSzip>"
				paymentech_postdata=paymentech_postdata&"<AVScountryCode>"&pcBillingCountryCode&"</AVScountryCode>"
				paymentech_postdata=paymentech_postdata&"</AVSextended>"
				
				If pcPay_Cent_Active=1 AND pcPay_CentByPass=1 AND pcPay_CardType="YES" Then
					paymentech_postdata=paymentech_postdata&"<VerifiedByVisa>"
					paymentech_postdata=paymentech_postdata&"<XID>"&Session("Centinel_XID")&"</XID>"
					paymentech_postdata=paymentech_postdata&"<CAVV>"&Session("Centinel_CAVV")&"</CAVV>"
					paymentech_postdata=paymentech_postdata&"</VerifiedByVisa>"
				End If
			
				paymentech_postdata=paymentech_postdata&"</AuthOptional>"
				paymentech_postdata=paymentech_postdata&"</Auth>"
				paymentech_postdata=paymentech_postdata&"<Cap>"
				paymentech_postdata=paymentech_postdata&"<CapMandatory>"
				paymentech_postdata=paymentech_postdata&"<EntryDataSrc>02</EntryDataSrc>"
				paymentech_postdata=paymentech_postdata&"</CapMandatory>"
				paymentech_postdata=paymentech_postdata&"<CapOptional/>"
				paymentech_postdata=paymentech_postdata&"</Cap>"
				paymentech_postdata=paymentech_postdata&"</AC>"
				paymentech_postdata=paymentech_postdata&"</Request>"

				'//View Post Data
				'response.write paymentech_postdata
				'response.End()
				
				
				Set srvpaymentechXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
				err.clear
				srvpaymentechXmlHttp.open "POST", paymentech_URL, false
				srvpaymentechXmlHttp.setRequestHeader "MIME-Version", "1.0"
				srvpaymentechXmlHttp.setRequestHeader "Content-type", "application/PTI34"
				srvpaymentechXmlHttp.setRequestHeader "Content-length", Len(paymentech_postdata)
				'//srvpaymentechXmlHttp.setRequestHeader "Content-length", "876"
				srvpaymentechXmlHttp.setRequestHeader "Content-transfer-encoding", "text"
				srvpaymentechXmlHttp.setRequestHeader "Request-number", "1"
				srvpaymentechXmlHttp.setRequestHeader "Document-type", "Request"
				srvpaymentechXmlHttp.setRequestHeader "Interface-Version", "Test 1.4"
				srvpaymentechXmlHttp.send(paymentech_postdata)
				
				if err.number<>0 then
					response.write "ERROR: "&err.description
					response.end
				end if
				paymentech_result = srvpaymentechXmlHttp.responseText
				if err.number<>0 then
					response.write "ERROR: "&err.description
					response.end
				end if
	
				'clear Centinel sessions
				Session("Centinel_Enrolled")=""
				Session("Centinel_ErrorNo")=""
				Session("Centinel_ErrorDesc")=""
				Session("Centinel_PAResStatus")=""
				Session("Centinel_SignatureVerification")=""
				Session("Centinel_ECI")=""
				Session("Centinel_XID")=""
				Session("Centinel_CAVV")=""
				Session("Centinel_ErrorNo")=""
				Session("Centinel_ErrorDesc")=""
				Session("Centinel_TransactionId")=""
				Session("Centinel_ACSURL")=""
				Session("Centinel_PAYLOAD")=""
	
				set srvpaymentechXmlHttp=nothing
				intStartProcStatus = InStr(1,paymentech_result, "<ProcStatus>", 0) +12
				intEndProcStatus = InstrRev(paymentech_result,"</ProcStatus>")
				ProcStatus=Mid(paymentech_result, intStartProcStatus, (intEndProcStatus-intStartProcStatus))
				Response.write "ProcStatus: "&ProcStatus&"<BR>"
				if ProcStatus="0" then
					intStartRespCode = InStr(1,paymentech_result, "<RespCode>", 0) +10
					intEndRespCode = InstrRev(paymentech_result,"</RespCode>")
					RespCode=Mid(paymentech_result, intStartRespCode, (intEndRespCode-intStartRespCode))
					Response.write "RespCode: "&RespCode&"<BR>"
					if RespCode="00" then	
						intStartTxRefNum = InStr(1,paymentech_result, "<TxRefNum>", 0) +10
						intEndTxRefNum = InstrRev(paymentech_result,"</TxRefNum>")
						TxRefNum=Mid(paymentech_result, intStartTxRefNum, (intEndTxRefNum-intStartTxRefNum))
						
						intStartTxRefIdx = InStr(1,paymentech_result, "<TxRefIdx>", 0) +10
						intEndTxRefIdx = InstrRev(paymentech_result,"</TxRefIdx>")
						TxRefIdx=Mid(paymentech_result, intStartTxRefIdx, (intEndTxRefIdx-intStartTxRefIdx))
				
						intStartAuthCode = InStr(1,paymentech_result, "<AuthCode>", 0) +10
						intEndAuthCode = InstrRev(paymentech_result,"</AuthCode>")
						AuthCode=Mid(paymentech_result, intStartAuthCode, (intEndAuthCode-intStartAuthCode))
						
						session("GWAuthCode")=strAuthorizationNumber
						session("GWTransId")=strTransactionID
						'session("TxRefIdx")=MethodName
						session("TransType")=pcPay_PT_TransType
						Response.redirect "gwReturn.asp?s=true&gw=Paymentech"
					else
						pt_processErr=1
					end if
				else
					pt_processErr=1
				end if
				
				if pt_processErr=1 then
					intStartStatusMsg = InStr(1,paymentech_result, "<StatusMsg", 0) +29
					intEndStatusMsg = InstrRev(paymentech_result,"</StatusMsg>")
					StatusMsg=Mid(paymentech_result, intStartStatusMsg, (intEndStatusMsg-intStartStatusMsg))
					response.redirect "msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error&nbsp;#"& ProcStatus &"</b>&nbsp;:&nbsp;"&StatusMsg&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&""" border=0></a>")
				end if
		
			end if
	
		end if
			
		call opendb()

		query="SELECT pcPay_PT_CVC, pcPay_PT_Testing FROM pcPay_Paymentech WHERE pcPay_PT_Id=1;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		pcPay_PT_CVC=rs("pcPay_PT_CVC")
		if isNULL(pcPay_PT_CVC) OR pcPay_PT_CVC="" then
			pcPay_PT_CVC="0"
		end if
		pcPay_PT_Testing=rs("pcPay_PT_Testing")
		if isNULL(pcPay_PT_Testing) OR pcPay_PT_Testing="" then
			pcPay_PT_Testing="0"
		end if
		set rs=nothing
		call closedb()
	
		%>
		<form action="gwPaymentech.asp" method="post" name="form1" class="pcForms">
			<input type="hidden" name="PaymentGWCentinel" value="Go">
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
					<% if pcPay_PT_Testing="1" then %>
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
					<% If pcPay_PT_CVC="1" Then %>
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
