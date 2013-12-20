<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Affiliate Table" %>
<% section="" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<%
response.Buffer=true
Response.Expires=0
	
dim mySQL, conntemp, rstemp

call openDb()
' Choose the records to display
idaffiliate=request.form("idaffiliate")
affiliateName=request.form("affiliateName")
affiliateEmail=request.form("affiliateEmail")
commission=request.form("commission")
	
strSQL="SELECT idaffiliate, affiliateName,affiliateEmail,commission FROM affiliates WHERE idaffiliate>1"
	
set rstemp=Server.CreateObject("ADODB.Recordset")     
rstemp.Open strSQL, conntemp, adOpenForwardOnly, adLockReadOnly, adCmdText

IF rstemp.eof then
	set rstemp=nothing
	call closedb()
%>
<!--#include file="AdminHeader.asp"-->
		<div class="pcCPmessage">
			Your search did not return any results. <a href="exportData.asp#affiliates">Back</a>.
		</div>
<!--#include file="AdminFooter.asp"-->
<%
response.End()
ELSE
		HTMLResult=""
		set StringBuilderObj = new StringBuilder
		If idaffiliate="1" then
			StringBuilderObj.append "<th>" & "Affiliate ID"& "</th>"
		End If
		If affiliateName="1" then
			StringBuilderObj.append "<th>" & "Name"& "</th>"
		End If
		If affiliateEmail="1" then
			StringBuilderObj.append "<th>" & "Email"& "</th>"
		End If
		If commission="1" then
			StringBuilderObj.append "<th>" & "Commission"& "</th>"
		End If
		HTMLResult="<table><tr>" & StringBuilderObj.toString() & "</tr>"
		set StringBuilderObj = nothing
		
		Do Until rstemp.EOF
		
			pidaffiliate=rstemp("idaffiliate")
			paffiliateName=rstemp("affiliateName")
			paffiliateEmail=rstemp("affiliateEmail")
			pcommission=rstemp("commission")
			
			set StringBuilderObj = new StringBuilder
			If idaffiliate="1" then 
				StringBuilderObj.append "<td width=""5%"">" & pidaffiliate& "</td>"
			End If
			If affiliateName="1" then
				StringBuilderObj.append "<td width=""40%"">" & paffiliateName& "</td>"
			End If
			If affiliateEmail="1" then
				StringBuilderObj.append "<td width=""40%"">" & paffiliateEmail& "</td>"
			End If
			If commission="1" then
				StringBuilderObj.append "<td>" & pcommission& "</td>"
			End If
			HTMLResult=HTMLResult & "<tr>" & StringBuilderObj.toString() & "</tr>"
			set StringBuilderObj = nothing
			rstemp.MoveNext			
		Loop
set rstemp=nothing
HTMLResult=HTMLResult & "</table>"
END IF
closedb()

Response.ContentType = "application/vnd.ms-excel"
%>
<%=HTMLResult%>