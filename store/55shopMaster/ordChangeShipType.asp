<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/ShipFromSettings.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<%
Dim pageTitle, Section, pOrdShipType
pageTitle="Change Shipping Type"
pageIcon="pcv4_icon_orders.gif"
Section="orders" 
%>
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<% Dim connTemp, qry_ID, query, rs, shiptype %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
Session.LCID = 1033
%>
<html>
<head>
<title>ProductCart v4 - Control Panel</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width: 400px; background-image: none;">
<h1 style="margin-left: 10px;"><%=pageTitle%></h1>
<%

' Get the integer order ID (if valid numeric)
qry_ID = 0
If Len(Trim(request.querystring("id")))>0 Then
   If IsNumeric(request.querystring("id")) Then
      qry_ID = CLng(request.querystring("id"))
   End if
End If

If qry_ID=0 Then
   response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query. Invalid input.")
End If


call openDb()

If UCase(request.servervariables("request_method"))="POST" Then

   shiptype = -1
   If Len(Trim(request.Form("shiptype")))>0 Then
      If IsNumeric(request.Form("shiptype")) Then
         shiptype = CLng(request.Form("shiptype"))
      End if
   End If
   
   If shipType<>-1 then
      query = "UPDATE orders SET ordShipType=" & shipType & " WHERE idOrder=" & qry_ID & ";"

      conntemp.execute query

      response.write "<p style=""margin: 12px;"">The shipping address type has been successfully updated. Please note: shipping charges have not been recalculated. Shipping charges can be recalculated on the <a href=""#"" onclick=""window.opener.location.href='AdminEditOrder.asp?ido=" & qry_ID & "';window.close();return false;"">Edit Order</a> page.<br /><br /><a href=""#"" onclick=""window.opener.location.href='Orddetails.asp?id=" & qry_ID & "&ActiveTab=6&s=1&msg=" & server.urlencode("The shipping address type has been successfully updated. Please note: shipping charges have not been recalculated. Shipping charges can be recalculated on the <a href=""AdminEditOrder.asp?ido=" & qry_ID & """>Edit Order</a> page.") & "';window.close();return false;"">Click here</a> to close this window.</p>"

   Else

      response.write "<p style=""margin: 12px;"">Error: Unknown shipping type selected.<br /><br /><a href=""#"" onclick=""window.close();return false;"">Click here</a> to close this window.</p>"

   End if

Else

   query="SELECT ordShipType FROM orders WHERE idOrder=" & qry_ID & ";"

   Set rs=Server.CreateObject("ADODB.Recordset")
   Set rs=connTemp.execute(query)

   If rs.eof Then

      Call closeDB()
      response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query. Order not found.")
   
   Else

      If IsNull(rs("ordShipType")) Then
         pOrdShipType=0
      else
         pOrdShipType=CLng(rs("ordShipType"))
      End If
      
      response.write "<p style=""margin: 12px;"">The shipping type for order #" & (scpre+int(qry_ID)) & " is currently set to "
      if pOrdShipType=0 then
         response.write "<strong>Residential</strong>"
      else
         response.write "<strong>Commercial</strong>" 
      end If
      response.write ". If you'd like to change the shipping type, please choose below and press 'Save Shipping Type' to save your change.<br /><br />"

      response.write "<form method=""post"" action=""ordChangeShipType.asp?id=" & qry_ID & """ class=""pcForms"">"
      response.write "<center>"
      response.write "<input type=""radio"" name=""shiptype"" value=""0"""
      If pOrdShipType=0 Then response.write " CHECKED"
      response.write "> Residential "
      response.write "<input type=""radio"" name=""shiptype"" value=""1"""
      If pOrdShipType=1 Then response.write " CHECKED"
      response.write "> Commercial"
      response.write "<br /><br /><input type=""submit"" value="" Save Shipping Type "" class=""submit2"">"

      response.write "</center>"
      response.write "</form>"

      response.write "</p>"

   End if




End If

Call closeDB()
%>
</div>
</body>
</html>
