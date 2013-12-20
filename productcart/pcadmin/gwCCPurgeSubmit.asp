<%@ Language=VBScript %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/productcartinc.asp"--> 
<!--#INCLUDE FILE="../includes/opendb.asp"-->
<% 
dim conntemp, rs, query
call opendb()

' extract real idorder (without prefix)
pTrueOrderId=request("POID")

'verify that this order doesn't alreay exists and that the idCustomer is only that of the customer logged in.
query="SELECT creditCards.idOrder, creditCards.cardnumber, creditCards.pcSecurityKeyID, orders.orderDate, orders.orderDate, orders.orderStatus FROM orders INNER JOIN creditCards ON orders.idOrder = creditCards.idOrder WHERE creditCards.idOrder="&pTrueOrderId
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

Var_OrderID = rs("idOrder")
set rs=nothing
call closedb()
%>
<HTML>
<HEAD>
</HEAD>
<body onLoad="document.FormCCP.submit();">
<form name="FormCCP" method="POST" action="CreditCardPurge.asp">
    <input type="hidden" name="idOrder1" value="1">
    <input type="hidden" name="ccOrderID1" value="<%=Var_OrderID%>"><br>
    <input type="hidden" name="pOrderID1" value="<%=Var_OrderID%>"><br>
    <input type="hidden" name="iCnt" value="1"><br>
    <input type="hidden" name="GW" value=""><br>
    <input type="hidden" name="PurgeNumbers" value="Yes"><br>
</form>
