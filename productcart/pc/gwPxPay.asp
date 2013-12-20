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
'======================================================================================
'// Set redirect page
'======================================================================================
' The redirect page tells the form where to post the payment information. Most of the
' time you will redirect the form back to this page.
'======================================================================================
session("redirectPage")="gwPxPay.asp"  'ALTER

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE
'======================================================================================
': Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
': End Declare and Retrieve Customer's IP Address

': Declare URL path to gwSubmit.asp
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
': End Declare URL path to gwSubmit.asp

': Get Order ID and Set to session
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if
': End Get Order ID

': Get customer and order data from the database for this order
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<%
': End Get customer and order data


': Reset customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
': End Reset customer session

': Open Connection to the DB
dim connTemp, rs
call openDb()
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE
'======================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your
'// new database table.
'//
'// If you are going to hard-code these variables, delete lines 77 - 93 (All lines are
'// commented as 'DELETE FOR HARD CODED VARS') and then set variables starting
'// at line 100 below.
'======================================================================================
query="SELECT pcPay_PxPay.pcPay_PxPay_PxPayUserId, pcPay_PxPay.pcPay_PxPay_PxPayTestUserId, pcPay_PxPay.pcPay_PxPay_PxPayKey, pcPay_PxPay.pcPay_PxPay_TxnType, pcPay_PxPay.pcPay_PxPay_TestMode, pcPay_PxPay.pcPay_PxPay_CurrencyInput FROM pcPay_PxPay WHERE (((pcPay_PxPay.pcPay_PxPay_ID)=1));"

'======================================================================================
'// End custom query
'======================================================================================

': Create recordset and execute query
set rs=server.CreateObject("ADODB.RecordSet") 'DELETE FOR HARD CODED VARS
set rs=connTemp.execute(query) 'DELETE FOR HARD CODED VARS

': Capture any errors
if err.number<>0 then 'DELETE FOR HARD CODED VARS
	call LogErrorToDatabase() 'DELETE FOR HARD CODED VARS
	set rs=nothing 'DELETE FOR HARD CODED VARS
	call closedb() 'DELETE FOR HARD CODED VARS
	response.redirect "techErr.asp?err="&pcStrCustRefID 'DELETE FOR HARD CODED VARS
end if 'DELETE FOR HARD CODED VARS

'======================================================================================
'// Set gateway specific variables - hard code is not using database to store gateway
'// information
'======================================================================================
pcv_PxPayUserId=rs("pcPay_PxPay_PxPayUserId")
pcv_PxPayTestUserId=rs("pcPay_PxPay_PxPayTestUserId")
pcv_PxPayKey=rs("pcPay_PxPay_PxPayKey")
pcv_TxnType=rs("pcPay_PxPay_TxnType")
pcv_CurrencyInput=rs("pcPay_PxPay_CurrencyInput")
pcv_TestMode=rs("pcPay_PxPay_TestMode")
if pcv_TestMode=1 then
	pcv_PxPayUserId=pcv_PxPayTestUserId
end if
'======================================================================================
'// End gateway specific variables
'======================================================================================

': Clear recordset and close db connection
set rs=nothing
call closedb()

'======================================================================================
'// If you are posting back to this page from the gateway form, all actions will happen
'// here.
'======================================================================================
if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	'// This is where you would post and retrieve info to and from the gateway
	'// START below this line
	'*************************************************************************************
	pcKeys_Key_ID= Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & session("idcustomer")

	Dim objXMLHTTP, xml

	pxPay_postdata=""
	pxPay_postdata = pxPay_postdata & "<GenerateRequest>"
	pxPay_postdata = pxPay_postdata & "<PxPayUserId>"&pcv_PxPayUserId&"</PxPayUserId>"
	pxPay_postdata = pxPay_postdata & "<PxPayKey>"&pcv_PxPayKey&"</PxPayKey>"
	pxPay_postdata = pxPay_postdata & "<TxnType>"&pcv_TxnType&"</TxnType>"
	pxPay_postdata = pxPay_postdata & "<TxnId>"&pcKeys_Key_ID&"</TxnId>"
	pxPay_postdata = pxPay_postdata & "<CurrencyInput>"&pcv_CurrencyInput&"</CurrencyInput>"
	pxPay_postdata = pxPay_postdata & "<AmountInput>"&replace(money(pcBillingTotal),",","")&"</AmountInput>"
	pxPay_postdata = pxPay_postdata & "<MerchantReference>"&session("GWOrderID")&"</MerchantReference>"
	pxPay_postdata = pxPay_postdata & "<ReceiptEmail>"&pcCustomerEmail&"</ReceiptEmail>"

	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwPXPay_receipt.asp"),"//","/")
	tempURL=replace(tempURL,"http:/","http://")
	tempURL=replace(tempURL,"https:/","https://")

	pxPay_postdata = pxPay_postdata & "<UrlSuccess>"&tempURL&"</UrlSuccess>"
	pxPay_postdata = pxPay_postdata & "<UrlFail>"&tempURL&"</UrlFail>"
	pxPay_postdata = pxPay_postdata & "</GenerateRequest>"

	pxPay_URL="https://sec.paymentexpress.com/pxpay/pxaccess.aspx"

	Set objXMLhttp =  Server.CreateObject("Msxml2.serverXmlHttp"&scXML) ' server.Createobject("MSXML2.XMLHTTP")
	objXMLhttp.Open "POST", pxPay_URL ,False
	objXMLhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXMLhttp.send pxPay_postdata

	Dim oXML, URI
	Set oXML = Server.CreateObject("MSXML2.DomDocument")
	oXML.loadXML(objXMLhttp.responseText)
	URI = oXML.selectSingleNode("//URI").text
	Response.Redirect URI

	Set objXMLhttp = nothing


	'*************************************************************************************
	' END
	'*************************************************************************************

end if
'======================================================================================
'// End post back
'======================================================================================


'======================================================================================
'// Show customer the payment form
'======================================================================================
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
					<%
					'======================================================================================
					'// If your gateway supports a Testing environment, create a variable for it and then
					'// if the cart is in testmode, alert the customer that this is not a live transaction.
					'// NOTE :: If no testing environment exists, delete the table row below
					'======================================================================================
					if pcv_TestMode=1 then %>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
					<%
					'======================================================================================
					'// End Testing environment variable
					'// NOTE :: If no testing environment exists, delete the table row above
					'======================================================================================
					%>
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
						<p>NOTE: When you click on the 'Place Order' button, you will temporarily leave our Web site and will be taken to a secure payment page on the PxPay Web site. You will be redirected back to our store once the transaction has been processed. We have partnered with DPS, a leader in secure Internet payment processing, to ensure that your transactions are processed securely and reliably.</p>
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
<%
'======================================================================================
'// End Show customer the payment form
'======================================================================================
%>
<!--#include file="footer.asp"-->