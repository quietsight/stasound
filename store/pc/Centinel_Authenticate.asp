<%@ language="vbscript" %>
<% 'option explicit %>
<%response.expires=-1%>
<%response.buffer=true%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!-- #include file="Centinel_Config.asp"-->

<!-- #include file="Centinel_Config.asp"-->
<!-- #include file="Centinel_ThinClient.asp"-->
<% dim conntemp, rs, query
'==========================================================================================
'= CardinalCommerce (http://www.cardinalcommerce.com)
'= Sample page that represents the TermUrl and used process the Authenticate message
'==========================================================================================
%>
<%

	dim pares, merchant_data, redirectPage
	dim centinelRequest
	dim centinelResponse
	'==========================================================================================
	' Retrieve the PaRes and MD values from the Card Issuer's Form POST to this Term URL page.
	' If you like, the MD data passed to the Card Issuer could contain the OrderId
	' that would enable you to reestablish the transaction session. This would be the 
	' alternative to using the Client Session Cookies
	'==========================================================================================

	pares = request.Form("PaRes")
	merchant_data = request.Form("MD")
		
	'==========================================================================================
	' If the PaRes is Not Empty then process the cmpi_authenticate message
	'==========================================================================================
	 if (pares <> "") then

		set centinelRequest = Server.CreateObject("Scripting.Dictionary")
		centinelRequest.Add "Version", Cstr(MessageVersion)
		centinelRequest.Add "MsgType", "cmpi_authenticate"
		centinelRequest.Add "ProcessorId", Cstr(ProcessorId)
		centinelRequest.Add "MerchantId", Cstr(MerchantId)
		centinelRequest.Add "TransactionPwd", Cstr(TransactionPwd)
		centinelRequest.Add "TransactionType", "C"
		'centinelRequest.Add "OrderId", Session("Centinel_OrderId")
		centinelRequest.Add "TransactionId", Session("Centinel_TransactionId")
		centinelRequest.Add "PAResPayload", Cstr(pares)

		'==========================================================================================
		' Send the XML Msg to the MAPS Server
		' SendHTTP will send the cmpi_authenticate message to the MAPS Server (requires fully qualified Url)
		' The Response is the CentinelResponse Object
		'==========================================================================================

		set centinelResponse = sendMsg(centinelRequest, Cstr(TransactionURL), CLng(ResolveTimeout), CLng(ConnectTimeout), CLng(SendTimeout), CLng(ReceiveTimeout))	

		'==========================================================================================
		' ************************************************************************************
		'								** Important Note **
		' ************************************************************************************
		'
		' Here you should persist the transaction results to your commerce system. 
		'
		' Be sure not to simply 'pass' the transaction results around from page to page on the
		' URL, since the values pass thru the consumer's browser and can be manipulated.
		' 
		'==========================================================================================

		' Using the centinelResponse object, we need to retrieve the results as follows
		Session("Centinel_PAResStatus") = centinelResponse.item("PAResStatus")
		Session("Centinel_ErrorNo") = centinelResponse.item("ErrorNo")
		Session("Centinel_ErrorDesc") = centinelResponse.item("ErrorDesc")
		Session("Centinel_SignatureVerification") = centinelResponse.item("SignatureVerification")
		Session("Centinel_ECI") = centinelResponse.item("EciFlag")
		Session("Centinel_XID") = centinelResponse.item("Xid")
		Session("Centinel_CAVV") = centinelResponse.item("Cavv")


		set centinelRequest = nothing
		set centinelResponse = nothing
	else
		Session("Centinel_ErrorDesc") = "No Pares Returned"
	end if

	'==========================================================================================
    ' Determine if the result was Successful or Error
    '
    ' If the Authentication results (PAResStatus) is a Y or A, and the SignatureVerification is Y, 
	' then the Payer Authentication was successful. The Authorization should be processed,
    ' and the User taken to a Order Confirmation location.
    '
    ' Note that it is also important that you account for cases when your flow logic can account
    ' for error cases, and the flow can be broken after 'N' number of attempts
	'==========================================================================================

	redirectPage = "gwAuthorizeAIM.asp?Centinel=Y"

	if ( Session("Centinel_ErrorNo") = "0" or Session("Centinel_ErrorNo") = "1140" ) _
	and Session("Centinel_SignatureVerification") = "Y" _
	and ( Session("Centinel_PAResStatus") = "Y" or Session("Centinel_PAResStatus") = "A") then

        '==========================================================================================
        '     If no errors were returned, the signature verification passed, and
        '     the transaction status was either "Y" (authenticated) or "A"
        '     (attempted), Payer Authentication was successful.
		'==========================================================================================

		Session("Centinel_Message") = "Your transaction completed successfully."
		
	elseif ( Session("Centinel_ErrorNo") = "0" or Session("Centinel_ErrorNo") = "1140" ) _
	and Session("Centinel_SignatureVerification") = "Y" _
	and ( Session("Centinel_PAResStatus") = "N") then

		'==========================================================================================
        '       Customer was presented with the authentication screen however
        '       either clicked the "exit" option or was unable to provide the
        '       correct password
		'==========================================================================================
		If scSSL="" OR scSSL="0" Then
			tempCAURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempCAURL=replace(tempCAURL,"https:/","https://")
			tempCAURL=replace(tempCAURL,"http:/","http://") 
		Else
			tempCAURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempCAURL=replace(tempCAURL,"https:/","https://")
			tempCAURL=replace(tempCAURL,"http:/","http://")
		End If

		Session("Centinel_Message") = "Your transaction was unable to authenticate. Please provide another form of payment. (PAResStatus = N)"
		redirectPage = tempCAURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&session("x_amount")

	else

		'==========================================================================================
		' Continue to authorization, either an error occurred or an unexpected status
        ' was returned ("U" for example). 
		'==========================================================================================

		Session("Centinel_Message") = "Your transaction completed however is pending review. Your order will be shipped once payment is verified."

	end if
	

%>
<html>
<head>
<script language="javascript">
<!--
	function onLoadHandler(){
		document.frmResultsPage.submit();
	}
//-->
</script>
<title>Processing Your Transaction</title>
</head>
<body onLoad="onLoadHandler();">

<form name="frmResultsPage" method="post" action="<%=redirectPage%>" target="_parent">
<noscript>
	<br><br>
	<center>
	<font color="red">
	<h1>Processing Your Transaction</h1>
	<h2>JavaScript is currently disabled or is not supported by your browser.<br></h2>
	<h3>Please click Submit to continue the processing of your transaction.</h3>
	</font>
	<input type="submit" value="submit">
	</center>
</noscript>
</form>
</body>
</html>
