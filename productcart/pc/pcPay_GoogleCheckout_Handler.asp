<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#INCLUDE file="pcPay_GoogleCheckout_Notification.asp"--> 
<!--#INCLUDE file="pcPay_GoogleCheckout_Calculations.asp"--> 
<%
'***********************************************************************************
' START: PROCESS XML
'***********************************************************************************
Function processXmlData(xmlData)

    Dim domResponseObj

    Set domResponseObj = Server.CreateObject(strMsxmlDomDocument)
    domResponseObj.loadXml xmlData
	
	Dim messageRecognizer
    messageRecognizer = domResponseObj.documentElement.tagName

   	Select Case messageRecognizer

        ' <request-received> received
        Case "request-received"
            processRequestReceivedResponse domResponseObj
         
        ' <error> received
        Case "error"
            processErrorResponse domResponseObj

        ' <diagnosis> received
        Case "diagnosis"
            processDiagnosisResponse domResponseObj

        ' <checkout-redirect> received
        Case "checkout-redirect"
			processCheckoutRedirect domResponseObj

        ' the Merchant Calculations API, you may ignore this case.
        ' <merchant-calculation-callback> received
        Case "merchant-calculation-callback"
			processMerchantCalculationCallback domResponseObj

        ' <new-order-notification> received
        Case "new-order-notification"
			processNewOrderNotification domResponseObj
     
        ' <order-state-change-notification> received
        Case "order-state-change-notification"
			processOrderStateChangeNotification domResponseObj
     
        ' <charge-amount-notification> received
        Case "charge-amount-notification"
			processChargeAmountNotification domResponseObj
         
        ' <chargeback-amount-notification> received
        Case "chargeback-amount-notification"
			processChargebackAmountNotification domResponseObj
         
        ' <refund-amount-notification> received
        Case "refund-amount-notification"
			processRefundAmountNotification domResponseObj
         
        ' <risk-information-notification> received
        Case "risk-information-notification"
			processRiskInformationNotification domResponseObj
         
        ' None of the above: message is not recognized.
        ' You should not remove this case.
        Case Else

    End Select 

End Function
'***********************************************************************************
' END: PROCESS XML
'***********************************************************************************




'********* Functions for processing synchronous response messages ********




'***********************************************************************************
' START: PROCESS REQUEST RECIEVED
'***********************************************************************************
Function processRequestReceivedResponse(domResponseObj)
    '// Response.write Server.HTMLEncode(domResponseObj.xml)
End Function
'***********************************************************************************
' END: PROCESS REQUEST RECIEVED
'***********************************************************************************




'***********************************************************************************
' START: PROCESS ERROR RESPONSE
'***********************************************************************************
Function processErrorResponse(domResponseObj)
	On Error Resume Next
	errstr="There was an error processing your request: "& domResponseObj.selectSingleNode("//error/error-message").text
End Function
'***********************************************************************************
' END: PROCESS ERROR RESPONSE
'***********************************************************************************




'***********************************************************************************
' START: PROCESS DIAGNOSTIC RESPONSE
'***********************************************************************************
Function processDiagnosisResponse(domResponseObj)   	
    errstr="<i>Diagnosis response message received:</i> "& Server.HTMLEncode(domResponseObj.xml)
End Function
'***********************************************************************************
' END: PROCESS DIAGNOSTIC RESPONSE
'***********************************************************************************




'***********************************************************************************
' START: PROCESS CHECKOUT REDIRECT
'***********************************************************************************
Function processCheckoutRedirect(domResponseObj)

    '// Define objects used to process <checkout-redirect> response
    Dim domResponseObjRoot
    Dim redirectUrlList
    Dim strRedirectUrl

    '// Identify the URL to which the customer should be redirected
    Set domResponseObjRoot = domResponseObj.documentElement
    Set redirectUrlList =  domResponseObjRoot.getElementsByTagname("redirect-url")
    strRedirectUrl = redirectUrlList(0).text

    '// Redirect the customer to the URL
    Response.redirect strRedirectUrl

    '// Release objects used to process <checkout-redirect> response
    Set domResponseObjRoot = Nothing
    Set redirectUrlList = Nothing

End Function
'***********************************************************************************
' END: PROCESS CHECKOUT REDIRECT
'***********************************************************************************
%>

