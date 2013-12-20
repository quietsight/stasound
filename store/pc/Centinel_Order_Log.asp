<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<html>
<head>
<title>CardinalCommerce's Centinel System: Transaction Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 0;">
<div id="pcMain">
<%
if request.Form("Submit")<>"" then 
	intOrderId=getUserInput(request.Form("orderId"),0)
	if NOT validNum(intOrderId) then
		response.redirect "msg.asp?message=38"
	end if
	dim query, conntemp, rs
	call opendb()

	query="SELECT pcPay_Centinel_Orders.pcPay_CentOrd_Enrolled, pcPay_Centinel_Orders.pcPay_CentOrd_ErrorNo, pcPay_Centinel_Orders.pcPay_CentOrd_ErrorDesc, pcPay_Centinel_Orders.pcPay_CentOrd_PAResStatus, pcPay_Centinel_Orders.pcPay_CentOrd_SignatureVerification, pcPay_Centinel_Orders.pcPay_CentOrd_EciFlag, pcPay_Centinel_Orders.pcPay_CentOrd_Xid, pcPay_Centinel_Orders.pcPay_CentOrd_Cavv, pcPay_Centinel_Orders.pcPay_CentOrd_rErrorNo, pcPay_Centinel_Orders.pcPay_CentOrd_rErrorDesc FROM pcPay_Centinel_Orders WHERE (((pcPay_Centinel_Orders.pcPay_CentOrd_OrderID)="&intOrderId&"));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
		
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
 
	pcPay_CentOrd_Enrolled=rs("pcPay_CentOrd_Enrolled")
	pcPay_CentOrd_ErrorNo=rs("pcPay_CentOrd_ErrorNo")
	pcPay_CentOrd_ErrorDesc=rs("pcPay_CentOrd_ErrorDesc")
	pcPay_CentOrd_PAResStatus=rs("pcPay_CentOrd_PAResStatus")
	pcPay_CentOrd_SignatureVerification=rs("pcPay_CentOrd_SignatureVerification")
	pcPay_CentOrd_EciFlag=rs("pcPay_CentOrd_EciFlag")
	pcPay_CentOrd_Xid=rs("pcPay_CentOrd_Xid")
	pcPay_CentOrd_Cavv=rs("pcPay_CentOrd_Cavv")
	pcPay_CentOrd_rErrorNo=rs("pcPay_CentOrd_rErrorNo")
	pcPay_CentOrd_rErrorDesc=rs("pcPay_CentOrd_rErrorDesc")
	%>
	<table class="pcMainTable">
  	<tr> 
    	<td colspan="2"><h1>CardinalCommerce's Centinel System: Transaction Information</h1></td>
  	</tr>
		<tr> 
			<td colspan="2" class="pcSpacer"></td>
		</tr>
		<tr> 
			<td colspan="2"><p><strong>Cmpi Lookup Response</strong></p></td>
		</tr>
		<tr> 
			<td width="14%"><p>Enrolled:</p></td>
			<td width="86%"><%=pcPay_CentOrd_Enrolled%></td>
		</tr>
		<tr> 
			<td><p>ErrorNo:</p></td>
			<td><%=pcPay_CentOrd_ErrorNo%></td>
		</tr>
		<tr> 
			<td><p>ErrorDesc:</p></td>
			<td><%=pcPay_CentOrd_ErrorDesc%></td>
		</tr>
		<tr> 
			<td colspan="2" class="pcSpacer"></td>
		</tr>
		<tr> 
			<td colspan="2"><p><strong>Cmpi Authenticate Response</strong></p></td>
		</tr>
		<tr> 
			<td><p>PAResStatus:</p></td>
			<td><%=pcPay_CentOrd_PAResStatus%></td>
		</tr>
		<tr> 
			<td><p>SignatureVerification:</p></td>
			<td><%=pcPay_CentOrd_SignatureVerification%></td>
		</tr>
		<tr> 
			<td><p>EciFlag:</p></td>
			<td><%=pcPay_CentOrd_EciFlag%></td>
		</tr>
		<tr> 
			<td><p>Xid:</p></td>
			<td><%=pcPay_CentOrd_Xid%></td>
		</tr>
		<tr> 
			<td><p>Cavv:</p></td>
			<td><%=pcPay_CentOrd_Cavv%></td>
		</tr>
		<tr> 
			<td><p>ErrorNo:</p></td>
			<td><%=pcPay_CentOrd_rErrorNo%></td>
		</tr>
		<tr> 
			<td><p>ErrorDesc:</p></td>
			<td><%=pcPay_CentOrd_rErrorDesc%></td>
		</tr>
		<tr> 
			<td colspan="2" class="pcSpacer"></td>
		</tr>
		<tr>
			<td colspan="2"><p><a href="Centinel_Order_Log.asp">Look up a new order log</a></p></td>
		</tr>
	</table>
<% else %>
	<form action="Centinel_Order_Log.asp" method="post">
		<table class="pcMainTable">
			<tr> 
				<td colspan="2"><h1>CardinalCommerce's Centinel System: Transaction Information</h1></td>
			</tr>
			<tr> 
				<td colspan="2" class="pcSpacer"></td>
			</tr>
			<tr> 
				<td>Order ID: <input type="text" name="OrderID">
					</td>
			</tr>
			<tr> 
				<td class="pcSpacer"></td>
			</tr>
			<tr> 
				<td> 
					<input type="submit" name="Submit" value="Submit" class="submit2">
				</td>
			</tr>
			<tr> 
				<td class="pcSpacer"></td>
			</tr>
		</table>
	</form>
<% end if %>
</div>
</body>
</html>