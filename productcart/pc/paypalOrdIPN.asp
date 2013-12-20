<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: PayPal IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function pcf_AuthorizeQuery(pcv_OrderID)
	call opendb()
	query="SELECT orders.gwTransID FROM orders WHERE orders.idOrder=" & pcv_OrderID &";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if not rs.eof then
		gwTransID=rs("gwTransID")	
	else
		gwTransID=0	
	end if	
	if gwTransID=session("GWSessionID") then
		pcf_AuthorizeQuery=True
	end if
	set rs=nothing
	call closedb()
End Function

Private Function pcf_SetGateway()
	call opendb()
	query="SELECT payTypes.idPayment FROM payTypes WHERE payTypes.gwCode=3;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if not rs.eof then
		session("pcSFIdPayment")=rs("idPayment")
	end if
	set rs=nothing
	call closedb()
End Function

If request("gw")="PayPal" Then
	pcv_AuthorizeQuery=False
	session("GWOrderId")=getUserInput(Request("GWOrderId"),0)	
	session("GWAuthCode")=getUserInput(Request("GWAuthCode"),0)
	session("GWTransId")=getUserInput(Request("GWTransId"),0)
	session("GWSessionID")=getUserInput(Request("GWSessionID"),0)
	pcv_AuthorizeQuery=pcf_AuthorizeQuery(session("GWOrderId"))
	If pcv_AuthorizeQuery=True Then
		Session("idCustomer")=getUserInput(Request("GWCustomerID"),0)
		pcf_SetGateway()
	End If
	session("GWOrderId")=(int(session("GWOrderId"))+scpre)
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: PayPal IPN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>