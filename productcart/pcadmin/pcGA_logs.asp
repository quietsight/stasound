<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Adjust Google Analytics Statistics" %>
<% Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<% Session.LCID = 1033 %>
<div style="padding: 15px;">
The following Google Analytics adjustments have been logged - <a href="pcGA_refund.asp">Back</a>
</div>
<%

'// GOOGLE ANALYTICS
'// E-commerce transaction tracking: order refunds and cancellations
'// http://www.google.com/support/analytics/bin/answer.py?answer=27203&topic=7282

'// Show transaction logs
					
		Dim sDSNFile
		sDSNFile = "gaLog.dsn"
		
		' Let's now dynamically retrieve the current directory
		Dim sScriptDir
		sScriptDir = Request.ServerVariables("SCRIPT_NAME")
		sScriptDir = StrReverse(sScriptDir)
		sScriptDir = Mid(sScriptDir, InStr(1, sScriptDir, "/"))
		sScriptDir = StrReverse(sScriptDir)
		
		' Time to construct our dynamic DSN
		Dim sPath, sDSN
		sPath = Server.MapPath(sScriptDir) & "\GAlogs\"
		sDSN = "FileDSN=" & sPath & sDSNFile & _
					 ";DefaultDir=" & sPath & _
					 ";DBQ=" & sPath & ";"
		
		Dim newConn
		Set newConn = Server.CreateObject("ADODB.Connection")
		newConn.Open sDSN

		query = "SELECT ORDERNUMBER,DATE,TRANSACTIONINFO,ITEMINFO FROM gaLog.txt"
		set rs = newConn.execute(query)
		
		'Print out the contents of our recordset
			rs.MoveNext
			Response.Write "<div style='padding: 15px'>"
			Do While Not rs.EOF
				pcArrayOrderInfo=split(rs("TRANSACTIONINFO"),"|")
				Response.Write "<div><strong>Order Number</strong>: " & rs("ORDERNUMBER") & "</div>"
				Response.Write "<div><strong>Adjustment posted to Google Analytics on</strong>: " & rs("DATE") & "</div>"
				Response.Write "<div><strong>Adjustment Amount</strong>: " & pcArrayOrderInfo(3) & "</div>"
				replaceString="UTM:I|"&(int(pOID)+scpre)&"|"
				itemInfo=replace(rs("ITEMINFO"),replaceString,"")
				itemInfo=replace(itemInfo,"|","&nbsp;&nbsp;&nbsp;")
				Response.Write "<strong>Adjustment Items</strong> (Part Number, Name, Category, Unit Price, Units):<br>" & itemInfo & "<hr>"
				rs.MoveNext
			Loop
			response.write "</div>"
		
		'Close our recordset and connection
		rs.close
		set rs = nothing
		newConn.close
		set newConn = nothing

%>
<div style="padding: 15px;">
<a href="pcGA_refund.asp">Back</a>
</div>
<!--#include file="AdminFooter.asp"-->