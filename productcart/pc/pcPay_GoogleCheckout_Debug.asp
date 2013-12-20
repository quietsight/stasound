<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="pcPay_GoogleCheckout_Global.asp"-->  
<!--#include file="pcPay_GoogleCheckout_Handler.asp"--> 
<%
Dim transmitResponse
Dim diagnoseResponse
Dim bValidated
Dim xml
Dim b64signature
Dim b64cart
Dim checkoutPostData


If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    ' Get <checkout-shopping-cart> XML
    xml = Trim(Request.Form("xml"))

    If Request.Form("toolType") = "Submit Cart to Google Checkout" Then

        transmitResponse = SendRequest(xml, requestUrl)
        ProcessXmlData(transmitResponse)
        Response.End

    ElseIf Request.Form("toolType") = "Display HTML Form for Checkout" Then

        ' Use the cart XML and your Merchant Key to calculate the HMAC-SHA1 
        ' value and Base64-encode the Cart XML and the signature before posting
'****************** MODIFIED TO UTILISE PURE ASP BASE64 AND HMAC_SHA1

        b64cart = Base64_Encode(xml)
        b64signature = b64_hmac_sha1(strMerchantKey, xml)
        
'****************** MODIFIED TO UTILISE PURE ASP BASE64 AND HMAC_SHA1

        checkoutPostData = "cart=" & Server.urlencode(b64cart) & _
            "&signature=" & Server.urlencode(b64signature)


        ' Log <checkout-shopping-cart> XML
        LogMessage logFilename, checkoutPostData

%>
<html>
<head>
    <style type="text/css">@import url(gbuy.css);</style>
</head>
<body>
    <p style="text-align:center">
    <table class="table-1" cellspacing="5" cellpadding="5">
        <tr><td style="padding-bottom:20px"><h2>
        Place a New Order
        </h2></td></tr>

        <tr><td>

            <!-- Print the steps for posting a shopping cart XML -->
            <p><b>Follow these steps to post an XML shopping cart:</b></p>
            <p><ol>
                <li>Create a &lt;checkout-shopping-cart&gt; XML structure
                    containing information about the buyer's order.</li>
                <li>Create an HMAC_SHA1 signature for the shopping cart.
                <li>Base64-encode the shopping cart XML.</li>
                <li>Base64-encode the HMAC_SHA1 signature.</li>
                <li>Put the cart and signature into a form that displays
                    a Google Checkout button.</li>
            </ol></p>
            <p>&nbsp;</p>

            <!-- Print the shopping cart XML -->
            <p><b>1. &lt;checkout-shopping-cart&gt; XML:</b></p>
            <p><%=Server.HTMLEncode(xml)%></p>
            <p>&nbsp;</p>

            <!-- Print the base64-encoded shopping cart XML -->
            <p><b>3. Base64-encoded &lt;checkout-shopping-cart&gt; XML:</b></p>
            <p><%=Server.HTMLEncode(b64cart)%></p>
            <p>&nbsp;</p>

            <!-- Print the base64-encoded HMAC-SHA1 signature -->
            <p><b>2 & 4. Base64-encoded HMAC-SHA1 Signature:</b></p>
            <p><%=Server.HTMLEncode(b64signature)%></p>
            <p>&nbsp;</p>

        </td></tr>

        <!-- Print Error message if the cart XML is invalid -->
        <%
            displayDiagnoseResponse checkoutPostData, checkoutDiagnoseUrl, _
                xml, "debug"
        %>

        <!-- Print the Google Checkout button in a form 
             containing the shopping cart data -->
        <tr><td>
            <p><b>Click on the button to post this cart.</b></p>

            <%
                ' Google Checkout button implementation

                Dim buttonW
                Dim buttonH
                Dim buttonStyle
                Dim buttonVariant
                Dim buttonLoc
                Dim buttonSrc
                
                buttonW = "180"
                buttonH = "46"
                buttonStyle = "white"
                buttonVariant = "text"
                buttonLoc = "en_US"
                buttonSrc = _
                    "https://sandbox.google.com/buttons/checkout.gif" & _
                    "?merchant_id=" & strMerchantId & _
                    "&w=" & buttonW & _
                    "&h=" & buttonH & _
                    "&style=" & buttonStyle & _
                    "&variant=" & buttonVariant & _
                    "&loc=" & buttonLoc
            %>

            <p><form method="POST" action="<%=checkoutUrl%>">
                <input type="hidden" name="cart" value="<%=b64cart%>">
                <input type="hidden" name="signature" value="<%=b64signature%>">
                <input type="image" name="Checkout" alt="Checkout" 
                src="<%=buttonSrc%>" height="<%=buttonH%>" width="<%=buttonW%>">
                </form></p>
        </td></tr>
    </table>
    </p>
</body>
</html>
<%
        Response.End
    End If
Else
    xml = ""
End If
%>
<html>
<head>
    <style type="text/css">@import url(gbuy.css);</style>
</head>
<body>
    <p style="text-align:center">
    <table class="table-1" cellspacing="5" cellpadding="5">
        <tr><td width="100%" style="text-align:center">
            <h2>Google Checkout API XML Debugging Tool</h2>
        </td></tr>
        <tr><td style="text-align:left">
            <form method="POST" 
            action="<%=Request.ServerVariables("REQUEST_URI")%>">
            <p><b>XML:</b></p>
            <p><textarea name="xml" cols="80" rows="20"><%=xml%></textarea></p>
            <p><table style="text-align:left" cellspacing="5" cellpadding="5">
                <tr><td><input name="toolType" type="submit" 
                value="Validate XML"></td>
                <td><input name="toolType" type="submit" 
                value="Display HTML Form for Checkout"></td>
                </tr><tr>
                <td><input name="toolType" type="submit" 
                value="Send Order Processing Command"></td>
                <td><input name="toolType" type="submit" 
                value="Submit Cart to Google Checkout"></td>
                </tr>
            </table></p>
            </form>
        </td></tr>
    </table>
    </p>
    <p style="text-align:center">
    <table class="table-1" cellspacing="5" cellpadding="5">
<%    
If Request.ServerVariables("REQUEST_METHOD") = "POST" And _
    Request.Form("toolType") = "Validate XML" Then

    ' Print Error message if the XML is invalid
    bValidated = displayDiagnoseResponse(xml, requestDiagnoseUrl, _
        xml, "diagnose")

    If bValidated = true Then
%>
        <tr><td colspan="2">
            <span style="text-align:center;color:green">
            <h2>This XML is Validated!</h2></span>
        </td></tr>
<% 
    End If

    Response.write "</table>"

ElseIf Request.ServerVariables("REQUEST_METHOD") = "POST" And _
    Request.Form("toolType") = "Send Order Processing Command" Then
%>
    <table class="table-1" cellspacing="5" cellpadding="5">
        <tr><td style="padding-bottom:20px;text-align:center"><h2>
            Order Processing Command
        </h2></td></tr>
        <tr><td style="padding-bottom:20px">
            <p><b>Order Processing Command XML:</b></p>
            <p><%=Server.HTMLEncode(xml)%></p>
        </td></tr>
<%
    ' Validate Request XML
    displayDiagnoseResponse xml, requestDiagnoseUrl, xml, "diagnose"

    Response.write "<tr><td style=""padding-bottom:20px"">" & _
        "<p><b>Synchronous Response Received:</b></p>"

    ' Send the request and receive a response
    transmitResponse = SendRequest(xml, requestUrl)

    ' Process the response
    Response.write "<p>" & ProcessXmlData(transmitResponse) & "</p></td></tr>"
    Response.write "</table>"

End If
%>
</p>
</body>
</html>