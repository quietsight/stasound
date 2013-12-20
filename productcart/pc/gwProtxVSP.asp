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
session("redirectPage")="gwProtxVSP.asp"

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
query="SELECT Protxid, ProtxTestmode, ProtxCurcode, CVV, avs, TxType, ProtxCardTypes, ProtxApply3DSecure FROM protx Where idProtx=1;"

set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcv_StrVendorName=rs("Protxid")
pcv_StrTestVendorName=pcv_StrVendorName

pcv_CVV=rs("CVV")
pcv_StrProtxTestmode=rs("ProtxTestmode")
pcv_ProtxCurcode=rs("ProtxCurcode")
avs=rs("avs")
pcv_StrTxType=rs("TxType")
pcv_StrProtxCardTypes=rs("ProtxCardTypes")
pcv_IntProtxApply3DSecure=rs("ProtxApply3DSecure")
if pcv_IntProtxApply3DSecure<>"0" AND pcv_IntProtxApply3DSecure<>"3" then
	pcv_IntProtxApply3DSecure="3"
end if


' ******************************************************************
' SagePay system to connect to
' ******************************************************************
pcv_StrProtocolVersion="2.23"

set rs=nothing
call closedb()

pcv_StrProtxAmex=0
pcv_StrProtxMaestro=0
strFormCardTypes=""

cardTypeArray=split(pcv_StrProtxCardTypes,", ")

for i=lbound(cardTypeArray) to ubound(cardTypeArray)
	cardVar=cardTypeArray(i)
	select case ucase(cardVar)
		case "VISA"
			strFormCardTypes=strFormCardTypes&"<option value=""VISA"" selected>VISA</option>"
		case "MC"
			strFormCardTypes=strFormCardTypes&"<option value=""MC"" selected>MasterCard</option>"
		case "UKE" 
			strFormCardTypes=strFormCardTypes&"<option value=""UKE"">Visa Debit/Visa Electron</option>"
		case "AMEX"
			strFormCardTypes=strFormCardTypes&"<option value=""AMEX"">American Express</option>"
			pcv_StrProtxAmex=1
		case "DELTA"
			strFormCardTypes=strFormCardTypes&"<option value=""DELTA"">Delta</option>"
		case "DC"
			strFormCardTypes=strFormCardTypes&"<option value=""DC"">Diners Club</option>"
		case "MAESTRO", "SWITCH"
			strFormCardTypes=strFormCardTypes&"<option value=""MAESTRO"">Maestro</option>"
			pcv_StrProtxAmex=1
			pcv_StrProtxMaestro=1
		case "SOLO"
			strFormCardTypes=strFormCardTypes&"<option value=""SOLO"">Solo</option>"
			pcv_StrProtxAmex=1
			pcv_StrProtxMaestro=1
		case "JCB"
			strFormCardTypes=strFormCardTypes&"<option value=""JCB"">JCB</option>"
	end select
next

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	if pcv_StrProtxTestmode=1 then
		purchaseURL	= "https://test.sagepay.com/Simulator/VSPDirectGateway.asp"
	elseif pcv_StrProtxTestmode=2 then
		purchaseURL = "https://test.sagepay.com/gateway/service/vspdirect-register.vsp"
	elseif pcv_StrProtxTestmode=0 then
		purchaseURL = "https://live.sagepay.com/gateway/service/vspdirect-register.vsp"
	else
		'// problem
	end if
	
	ThisVendorTxCode=session("GWOrderId")
		
	Randomize
	
	If pcv_StrProtxTestmode =1 Then
		ThisVendorTxCode=pcv_StrTestVendorName & timer() & rnd()
	End if
	
	'// Validate all input
	pcv_CardNumber=request("CardNumber")
	pcv_CardNumber=replace(pcv_CardNumber,"-","")
	pcv_CardNumber=replace(pcv_CardNumber,".","")
	pcv_CardNumber=replace(pcv_CardNumber," ","")
	if NOT isNumeric(pcv_CardNumber) then
		'Card Number is not a valid format
	end if
		
	'set all the required outgoing properties
	postData = _
	"VPSProtocol=" & pcv_StrProtocolVersion & _
	"&TxType=" & pcv_StrTxType & _
	"&Vendor=" & pcv_StrVendorName

	postData = postData & "&VendorTxCode=" & ThisVendorTxCode
	postData = postData & "&Amount=" & replace(money(pcBillingTotal),",","")
	postData = postData & "&Currency=" & pcv_ProtxCurcode
	postData = postData & "&Description=" & Server.URLEncode( session("GWOrderId") )
	postData = postData & "&CardHolder=" & pcBillingFirstName & " " & pcBillingLastName
	postData = postData & "&CardNumber=" & pcv_CardNumber
	if pcv_StrProtxAmex=1 then
		postData = postData & "&StartDate=" & request.form( "StartDate" )
	end if
	postData = postData & "&ExpiryDate=" & request.form( "expMonth" ) & request.form( "expYear" )
	if pcv_StrProtxMaestro=1 then
	 postData = postData & "&IssueNumber=" & request.form( "IssueNumber" )
	end if
	postData = postData & "&CardType=" & request.form( "ProtxCardTypes" )
	If pcv_CVV ="1" Then
		postData = postData & "&CV2=" & request.form( "CVV" )
	End If

	'//New for 2.23
	postData = postData & "&BillingSurname=" & left(pcBillingLastName,20)
	postData = postData & "&BillingFirstnames=" & left(pcBillingFirstName,20)
	postData = postData & "&BillingAddress1=" & left(pcBillingAddress,100)
	postData = postData & "&BillingAddress2=" & left(pcBillingAddress2,100)
	postData = postData & "&BillingCity=" & left(pcBillingCity,40)
	postData = postData & "&BillingPostCode=" & left(pcBillingPostalCode,10)
	postData = postData & "&BillingCountry=" & left(pcBillingCountryCode,2)
	If ucase(pcBillingCountryCode) = "US" then
		postData = postData & "&BillingState=" & left(pcBillingStateCode,2)
	End If
	postData = postData & "&BillingPhone=" & left(pcBillingPhone,20)
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
		postData = postData & "&DeliverySurname=" & left(pcShippingLastName,20)
        else
		postData = postData & "&DeliverySurname=" & left(pcBillingLastName,20)
	end if
	
	if len(pcShippingFirstName)> 0 then
		postData = postData & "&DeliveryFirstnames=" & left(pcShippingFirstName,20)
        else
		postData = postData & "&DeliveryFirstnames=" & left(pcBillingFirstName,20)
	end if
	
	if len(pcShippingAddress)> 0 then
		postData = postData & "&DeliveryAddress1=" & left(pcShippingAddress,100)
        else
		postData = postData & "&DeliveryAddress1=" & left(pcBillingAddress,100)
	end if
	
	if len(pcShippingAddress2)> 0 then
		postData = postData & "&DeliveryAddress2=" & left(pcShippingAddress2,100)
        else
		postData = postData & "&DeliveryAddress2=" & left(pcBillingAddress2,100)
	end if
	
	if len(pcShippingCity)> 0 then
		postData = postData & "&DeliveryCity=" & left(pcShippingCity,40)
        else
		postData = postData & "&DeliveryCity=" & left(pcBillingCity,40)
	end if
	
	if len(pcShippingPostalCode)> 0 then
		postData = postData & "&DeliveryPostCode=" & left(pcShippingPostalCode,10)
        else
		postData = postData & "&DeliveryPostCode=" & left(pcBillingPostalCode,10)
	end if
	
	if len(pcShippingCountryCode)> 0 then
		postData = postData & "&DeliveryCountry=" & left(pcShippingCountryCode,2)
        else
		postData = postData & "&DeliveryCountry=" & left(pcBillingCountryCode,2)
	end if
	
	if len(pcShippingStateCode)> 0 then
		If ucase(pcBillingCountryCode) = "US" then
			postData = postData & "&DeliveryState=" & left(pcShippingStateCode,2)
        	else
			postData = postData & "&DeliveryState=" & left(pcBillingStateCode,2)
		End If	
	end if
	
	if len(pcShippingPhone)> 0 then
		postData = postData & "&DeliveryPhone=" & left(pcShippingPhone,20)
        else
		postData = postData & "&DeliveryPhone=" & left(pcBillingPhone,20)
	end if	
	
	postData = postData & "&CustomerEMail=" &pcCustomerEmail
	
	If pcv_CVV ="1" Then
		ApplyAVSCV2=1
		postData = postData & "&ApplyAVSCV2=" & ApplyAVSCV2
	End If 
	postData = postData & "&ClientIPAddress=" & pcCustIpAddress
	
	'** Send the account type to be used for this transaction.  Web sites should us E for e-commerce **
	'** If you are developing back-office applications for Mail Order/Telephone order, use M **
	'** If your back office application is a subscription system with recurring transactions, use C **
	'** Your SagePay account MUST be set up for the account type you choose.  If in doubt, use E **
	postData = postData & "&AccountType=E"

	'// Use this variable to turn your BASKET feature ON/OFF - Default is "OFF"
	ThisShoppingBasket="OFF"

	if ThisShoppingBasket="ON" then
		'select all products from the ProductsOrdered table to insert them into the 2Checkout db.
		call opendb()
		query="SELECT products.idproduct, products.description, quantity, unitPrice FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& session("GWOrderId")
		set rsBasketObj=server.CreateObject("ADODB.Recordset")
		set rsBasketObj=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsBasketObj=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		IntProdCnt=0
		do until rsBasketObj.eof
			tempIntIdProduct=rsBasketObj("idproduct")
			tempStrDescription=rsBasketObj("description")
			tempIntQuantity=rsBasketObj("quantity")
			tempDblUnitPrice=rsBasketObj("unitPrice")
			IntProdCnt=IntProdCnt+1

			strPCBasket = strPCBasket & ":" & tempStrDescription & ":" & tempIntQuantity & ":" & tempDblUnitPrice & ":::" &tempDblUnitPrice

			rsBasketObj.moveNext
		loop 
		set rsBasketObj=nothing
		call closedb() 
		
		postData = postData & "&Basket="&IntProdCnt & strPCBasket
	end if
	
	'** Allow fine control over 3D-Secure checks and rules by changing this value. 0 is Default **
	if pcv_IntProtxApply3DSecure="0" then
		pcv_IntProtxApply3DSecure="2"
	end if
	'** It can be changed dynamically, per transaction, if you wish.  See the VSP Server Protocol document **
	postData = postData & "&Apply3DSecure="&pcv_IntProtxApply3DSecure&""
	'0 = If 3D-Secure checks are possible and rules allow, perform the checks and apply the authorisation rules (default).
	'1 = Force 3D-Secure checks for this transaction only (if your account is 3D-enabled) and apply rules for authorisation. 
	'2 = Do not perform 3D-Secure checks for this transaction only and always authorise. 
	'3 = Force 3D-Secure checks for this transaction (if your account is 3D-enabled) but ALWAYS obtain an auth code, irrespective of rule base.

	'send to SagePay
	set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
	
	' *** open connection to SagePay
	httpRequest.Open "POST", purchaseURL, False
	
	httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	httpRequest.send postData
	
	responseData = httpRequest.responseText
	
	'*** Following line shows the whole reply for debugging purposes ***
	'response.write "ERROR: "&Err.number & " - "&Err.description &"<br/>"
	'response.write "<b>Response was:</b> " & responseData & "<br/>"
	'response.end
	
	'** An non zero Err.number indicates an error of some kind **
	'** Check for the most common error... unable to reach the purchase URL **  
	strPageError="" 
	if err.number<>0 then
		if Err.number = -2147012889 then
			strPageError="Your server was unable to register this transaction with SagePay." &_
						"  Check that you do not have a firewall restricting the POST and " &_
						"that your server can correctly resolve the address " & strPurchaseURL
		else
			strPageError="An Error has occurred whilst trying to register this transaction.<BR>" &_
						"The Error Number is: " & Err.number & "<BR>" &_
						"The Description given is: " & Err.Description
		end If 
		if strPageError<>"" then
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:&nbsp;"&strProtxErrorMsg&"<br><br><a href="""&tempURL&"?psslurl=gwProtxVSP.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
		end if
	end if
	' ******************************************************************
	' Determine next action
	'** No transport level errors, so the message got the SagePay **
	'** Analyse the response from VSP Direct to check that everything is okay **
	'** Registration results come back in the Status and StatusDetail fields **
	strStatus=findField("Status",responseData)
	strStatusDetail=findField("StatusDetail",responseData)

	if strStatus="3DAUTH" then
		'** This is a 3D-Secure transaction, so we need to redirect the customer to their bank **
		'** for authentication.  First get the pertinent information from the response **
		strMD=findField("MD",responseData)
		strACSURL=findField("ACSURL",responseData)
		strPAReq=findField("PAReq",responseData)
		strPageState="3DRedirect"
	else
		'** If this isn't 3D-Auth, then this is an authorisation result (either successful or otherwise) **
		'** Get the results form the POST if they are there **
		strVPSTxId=findField("VPSTxId",responseData)
		strSecurityKey=findField("SecurityKey",responseData)
		strTxAuthNo=findField("TxAuthNo",responseData)
		strAVSCV2=findField("AVSCV2",responseData)
		strAddressResult=findField("AddressResult",responseData)
		strPostCodeResult=findField("PostCodeResult",responseData)
		strCV2Result=findField("CV2Result",responseData)
		str3DSecureStatus=findField("3DSecureStatus",responseData)
		strCAVV=findField("CAVV",responseData)
	
	
		if strStatus="OK" then
			session("GWAuthCode")=strTxAuthNo
			session("GWTransId")=strVPSTxId
			session("GWTransType")=pcv_StrTxType
			Response.redirect "gwReturn.asp?s=true&gw=SagePay"
		else
			if strStatus="AUTHENTICATED" then
				session("GWAuthCode")=strTxAuthNo
				session("GWTransId")=strVPSTxId
				session("GWTransType")=pcv_StrTxType
				Response.redirect "gwReturn.asp?s=true&gw=SagePay"
			end if
			' ** Something has gone wrong, record the error and redirect etc.
			 strProtxErrorType=strStatus
			 strProtxErrorMsg=strStatusDetail
			'REJECTED, NOTAUTHED, ERROR redirect back to payment form
			response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:&nbsp;"&strProtxErrorType&" - "&strProtxErrorMsg&"<br><br><a href="""&tempURL&"?psslurl=gwProtxVSP.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
			' ** Write VPSTxID, SecurityKey, Status and StatusDetail to the screen, log file or database
			response.write "<b>Failed</b><br/>"
			response.end
		end if
	
		' ******************************************************************
		' remove the reference to the object
		set httpRequest = nothing
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
        <% '** A 3D-Auth response has been returned, so show the bank page inline if possible, or redirect to it otherwise
		if strPageState="3DRedirect" then %>
		<tr>
			<td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
				  <td class="subheader" align="center">3D-Secure Authentication with your Bank</td>
              </tr>
              <tr>
                <td valign="top">
					<table border="0" width="100%">
						<tr>
							<td width="80%">To increase the security of Internet transactions Visa and Mastercard have introduced 3D-Secure (like an online version of Chip and PIN). <br>
							  <br>
						    You have chosen to use a card that is part of the 3D-Secure scheme, so you will need to authenticate yourself with your bank in the section below.</td>
							<td width="20%" align="center"><img src="images/vbv_logo_small.gif" alt="Verified by Visa"><BR><BR><img src="images/mcsc_logo.gif" alt="MasterCard SecureCode"></td>
						</tr>
					</table>
				</td>
              </tr>
			  
			  <tr>
                <td valign="top"><%
					'** Attempt to set up an inline frame here.  If we can't, set up a standard full page redirection **
					Session("MD")=strMD
					Session("PAReq")=strPAReq
					Session("ACSURL")=strACSURL
					Session("VendorTxCode")=strVendorTxCode
				%>
					<IFRAME SRC="gwProtx_3DRedirect.asp" NAME="3DIFrame" WIDTH="100%" HEIGHT="500" FRAMEBORDER="0">
					<% 'Non-IFRAME browser support
					response.write "<SCRIPT LANGUAGE=""Javascript""> function OnLoadEvent() { document.form.submit(); }</" & "SCRIPT>" 		
					response.write "<html><head><title>3D Secure Verification</title></head>"

					response.write "<body OnLoad=""OnLoadEvent();"">"
					response.write "<FORM name=""form"" action=""" & strACSURL &""" method=""POST"">"
					response.write "<input type=""hidden"" name=""PaReq"" value=""" & strPAReq &"""/>"
					response.write "<input type=""hidden"" name=""TermUrl"" value=""" & strYourSiteFQDN & strVirtualDir & "/3DCallback.asp?VendorTxCode=" & strVendorTxCode & """/>"
					response.write "<input type=""hidden"" name=""MD"" value=""" & strMD &"""/>"

					response.write "<NOSCRIPT>" 
					response.write "<center><p>Please click button below to Authenticate your card</p><input type=""submit"" value=""Go""/></p></center>"
					response.write "</NOSCRIPT>"
					response.write "</form></body></html>"%>
					</IFRAME>
				</td>
			  </tr>

			</table>
           </td>
           </tr>
        <% else %>
		<tr>
			<td>
				<form method="POST" autocomplete="off" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
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
					<% if pcv_StrProtxTestmode<>0 then %>
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
								<select name="ProtxCardTypes">
									<%=strFormCardTypes%>
								</select>
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td> 
							<input type="text" name="CardNumber" autocomplete="off"  size="18" maxlength="18" value="">
						</td>
					</tr>
					<% if pcv_StrProtxAmex=1 then %>
					<tr>
						<td><p>Start Date:</p></td>
					<td>
						<input name="StartDate" autocomplete="off" type="text" size="6" maxlength="4">
						<span class="pcSmallText">Required for some Maestro, Solo and Amex; <strong>mmyy</strong> format.</span>
						</td>
					</tr>
					<% end if %>
					<% if pcv_StrProtxMaestro=1 then %>
						<tr>
							<td><p>Issue Number:</p>
							</td>
						<td>
							<input name="IssueNumber" autocomplete="off" type="text" size="4" maxlength="2">
							<span class="pcSmallText">Required some Maestro and Solo cards only.</span>
						</td>
						</tr>
					<% end if %>
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
					<% If pcv_CVV ="1" Then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
								<input name="CVV" autocomplete="off" type="text" id="CVV" value="" size="4" maxlength="4">
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
   <% end if %>
</table>
</div>
<% 
'***********************************************
' Useful methods
'***********************************************

function findField( fieldName, postResponse )
  items = split( postResponse, chr( 13 ) )
  for idx = LBound( items ) to UBound( items )
    item = replace( items( idx ), chr( 10 ), "" )
    if InStr( item, fieldName & "=" ) = 1 then
      ' found
      findField = right( item, len( item ) - len( fieldName ) - 1 )
      Exit For
    end if
  next 
end function
%>
<!--#include file="footer.asp"-->