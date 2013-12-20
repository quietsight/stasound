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
session("redirectPage")="gwMonerisInterac.asp"
	
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
query="SELECT pcPay_Moneris_StoreId, pcPay_Moneris_Key, pcPay_Moneris_TransType, pcPay_Moneris_Lang, pcPay_Moneris_Testmode, pcPay_Moneris_CVVEnabled, pcPay_Moneris_Meth, pcPay_Moneris_Interac, pcPay_Moneris_Interac_MerchantID FROM pcPay_Moneris Where pcPay_Moneris_Id=1;"
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
pcPay_Moneris_Interac = rs("pcPay_Moneris_Interac")
pcPay_Moneris_Interac_MerchantID = rs("pcPay_Moneris_Interac_MerchantID")
pcPay_Moneris_Interac_MerchantID = enDeCrypt(pcPay_Moneris_Interac_MerchantID, scCrypPass)
set rs=nothing
call closedb()
'response.write pcPay_Moneris_Interac



if request("IDEBIT_VERSION") <> ""  or request("IDEBIT_INVOICE") <> "" Then 
	if pcPay_Moneris_TestMode="1" then
		pcBillingTotal="1.00"
	end if
	
		IDEBIT_INVOICE = request("IDEBIT_INVOICE") 
		IDEBIT_ISSLANG = request("IDEBIT_ISSLANG")
		IDEBIT_ISSCONF = request("IDEBIT_ISSCONF")
		IDEBIT_ISSNAME = request("IDEBIT_ISSNAME")
		IDEBIT_TRACK2 = request("IDEBIT_TRACK2") 
		IDEBIT_VERSION = request("IDEBIT_VERSION")
        IDEBIT_VERSION = request("IDEBIT_VERSION") 
   For Each Item In Request.Form
	fieldName = Item
	fieldValue = Request.Form(Item) 
	response.write fieldName &"-" &fieldValue &"<BR>"
Next
'Response.end

     if IDEBIT_ISSCONF = "" or IDEBIT_ISSNAME = "" or IDEBIT_INVOICE ="" OR IDEBIT_ISSLANG = "" or IDEBIT_TRACK2 = "" or IDEBIT_VERSION = ""  Then
	  		 strDeclinedRedirect = "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>: The INTERAC&reg; Online transaction was declined <br><br><a href="""&tempURL&"?psslurl=gwMonerisInterac.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
	   		 response.redirect strDeclinedRedirect
	  		 response.end 
	  End if 

      session("IDEBIT_ISSCONF") = IDEBIT_ISSCONF
	  session("IDEBIT_ISSNAME") = IDEBIT_ISSNAME

	'Send the request to the Authorize.NET processor.
	stext="ps_store_id="&pcPay_Moneris_StoreId
	stext=stext & "&hpp_key="&pcPay_Moneris_Key
	stext=stext & "&charge_total=" & replace(money(pcBillingTotal),",","")
	stext=stext & "&lang=" & pcPay_Moneris_Lang
	stext=stext & "&cc_num="& IDEBIT_TRACK2
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
	
	if pcPay_Moneris_TestMode="1" or pcPay_Moneris_TestMode="2" then
		strHostURL="https://esqa.moneris.com/HPPDP/index.php"
	else
		strHostURL="https://www3.moneris.com/HPPDP/index.php"
	end if
    'response.write strHostURL &stext
	'response.end	
	
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
		xml.open "GET", strHostURL & "?"&stext & "", false
		xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xml.send "" 
	 End if 
	
	strRetVal = xml.responseText
	Session("MonerisTransKey")=strretVal
	
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
			<%
			
   Select case pcPay_Moneris_TestMode   
   case "0"
    	strHostURL = "https://gateway.interaconline.com/merchant_processor.do"
   case "1"
		pcBillingTotal="1.00"
		strHostURL = "https://merchant.interacidebit.ca/testtools/merchant_test_tool_request.do"
   case "2"
	 	strHostURL = "https://merchant.interacidebit.ca/testtools/merchant_certification_request.do"
   End Select
			
			
			%>
			<form action='<%=strHostURL%>' method='post'>
			<input type='hidden' name='IDEBIT_INVOICE' value='<%=pcStrCustRefID%>'>
			<input type='hidden' name='DEBIT_MERCHDATA' value='<%=session("GWOrderID")%>'>
			<input type='hidden' name='IDEBIT_AMOUNT' value='<%=pcBillingTotal * 100%>'>
			<input type='hidden' name='IDEBIT_MERCHNUM' value='<%=pcPay_Moneris_Interac_MerchantID%>'>
			<input type='hidden' name='IDEBIT_CURRENCY' value='CAD'>
			<input type='hidden' name='IDEBIT_FUNDEDURL' value='<%=replace(tempURL, "gwSubmit.asp", "gwMonerisInterac.asp" )%>'>
			<input type='hidden' name='IDEBIT_NOTFUNDEDURL' value='<%=replace(tempURL, "gwSubmit.asp", "gwMonerisIntNoFund.asp" )%>'>
			<!--input type='hidden' name='IDEBIT_MERCHLANG' value='<%=left(pcPay_Moneris_Lang,2)%>'-->			
			<input type='hidden' name='IDEBIT_VERSION' value='1'>
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
					<% if pcPay_Moneris_TestMode="2" then %>
					<tr>
						<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_23")%></div></td>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% end if %>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_25")%></p></td>
					</tr>
					<tr height="30">
					<td colspan="2"><p><a href="#"onclick="javascript:window.open('http://www.interaconline.com/learn','learn','height=400,width=550, toolbar=0, scrollbars=1, location=0, statusbar=0, menubar=0, resizeable=1');">Learn
						More</a>
						&nbsp;
						<a href="#"
						onclick="javascript:window.open('http://www.interacenligne.com/renseignements','learn','height=400,width=550,toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizeable=1');">Rens
						eignements</a></p></td>
					
					</tr>
					<tr height="30">
					<td colspan="2"><p>You have 30 minutes to complete this transaction otherwise it will time out.</p></td>
					
					</tr>
					
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_24")%></p></td>
						<td align="left"><select name="IDEBIT_MERCHLANG" size="1">
						<option value="en">English</option>
						<option value="fr">French</option>						
						</select></td>
					</tr>
			       
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><p><%=money(pcBillingTotal)%>&nbsp;(CAD)</p></td>
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