<%
response.Buffer=true
Response.Expires = -1
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Configure Payment Option"
pageIcon="pcv4_icon_pg.png"
section="paymntOpt" 
pcStrPageName="pcConfigurePayment.asp"
%>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="RTGatewayConstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
dim query, conntemp, rstemp
call openDb()

dim pcSpryCP

pcSpryCP="PP"
'====================================================
'// Check if Centinel has previously been activated.
'====================================================
dim intCentActive
dim pcPay_Cent_Active

intCentActive=0

err.clear
err.number=0
call openDb()  

query="SELECT pcPay_Cent_Active FROM pcPay_Centinel WHERE pcPay_Cent_ID=1;"
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=connTemp.execute(query)
pcPay_Cent_Active=rs("pcPay_Cent_Active")
if pcPay_Cent_Active=1 then
	intCentActive=1
end if

set rs=nothing
'====================================================
'====================================================

Dim i, strErr, isComErr, strComErr
	
sMode=Request.Form("submitMode")
eMode=Request.Form("mode")
iMode=Request.QueryString("mode")
'Delete
If iMode="Del" Then
	iActivate=request("activate")
	pcv_processOrder=request.Form("pcv_processOrder")
	if pcv_processOrder="" then
		pcv_processOrder="0"
	end if
	pcv_setPayStatus=request.Form("pcv_setPayStatus")
	if pcv_setPayStatus="" then
		pcv_setPayStatus="3"
	end if
	idPayment=Request.QueryString("id")
	gwCode= Request.QueryString("gwChoice")
	call openDb()
	If request.QueryString("TYPE")="CC" then
		CCcode=request.queryString("CCCode")
		query="UPDATE CCTypes SET active=0 WHERE CCcode='" & CCcode & "'"
		set rs=Server.CreateObject("ADODB.Recordset")  
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?v=1&error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		query="SELECT * FROM CCTypes WHERE active<>0"
		set rs=Server.CreateObject("ADODB.Recordset")  
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?v=2&error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		if rs.eof then
			query= "UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",active=0, paymentNickName='' WHERE gwCode=6"
			set rs=Server.CreateObject("ADODB.Recordset")  
			set rs=conntemp.execute(query)
		end if
	Else
		If gwCode="6" then
			query= "UPDATE CCTypes SET active=0 WHERE active<>0"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				strErrorDescription=err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?v=3&error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			query= "UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",active=0, paymentNickName='' WHERE gwCode=6;"
			set rs=Server.CreateObject("ADODB.Recordset")  
			set rs=conntemp.execute(query)
		End If
		
		If gwCode="7" then
			query= "DELETE FROM payTypes WHERE idPayment="& idPayment
		End If
		
		If gwCode<>"6" AND gwCode<>"7" then
			query= "DELETE FROM payTypes WHERE gwCode="& gwCode
			
			if gwCode="1" then 'Delete authorize check
				querytemp="DELETE FROM payTypes WHERE gwCode=16"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="24" then
				querytemp="DELETE FROM payTypes WHERE gwCode=25" 'Delete TCLink check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="27" then
				querytemp="DELETE FROM payTypes WHERE gwCode=28" 'Delete Netbilling check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="32" then 'Delete CyberSource Echeck
				querytemp="DELETE FROM payTypes WHERE gwCode=62"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="35" then
				querytemp="DELETE FROM payTypes WHERE gwCode=36" 'Delete USAePay check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="37" then
				querytemp="DELETE FROM payTypes WHERE gwCode=38" 'Delete FastCharge check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="60" then 'Delete DowCom check
				querytemp="DELETE FROM payTypes WHERE gwCode=61"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="11" then 'Delete authorize check
				querytemp="DELETE FROM payTypes WHERE gwCode=66"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
		End if
	End If
	
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	call Closedb()
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?v=4&error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	if iActivate&""<>"" then
		response.redirect "pcConfigurePayment.asp?gwchoice="&iActivate
	else
		if request("page")&""<>"" then
			response.redirect request("page")
		else
			response.redirect "paymentOptions.asp"
		end if
	end if
End If

if eMode="Edit" Then
	pcv_processOrder=request.Form("pcv_processOrder")
	if pcv_processOrder="" then
		pcv_processOrder="0"
	end if
	pcv_setPayStatus=request.Form("pcv_setPayStatus")
	if pcv_setPayStatus="" then
		pcv_setPayStatus="3"
	end if

	PaymentDesc=request.Form("PaymentDesc")
	idPayment=request.Form("idPayment")
	priceToAddType=request.Form("priceToAddType")
	gwCode=request.form("gwCode")
	If priceToAddType="price" Then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
		if priceToAdd="" then
			priceToAdd="0"
		end if
	Else
		percentageToAdd=request.Form("percentageToAdd")
		priceToAdd="0"
		if priceToAddType="null" then
			percentageToAdd="0"
		else
			if percentageToAdd="" then
				percentageToAdd="0"
			end if
		end if
	End If
	
	sslUrl=request.Form("sslUrl")
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")

	'//Gateway to edit
	pcv_EditGW=request("addGW")
	call gwCallEdit()
	call opendb()
	query="SELECT idPayment FROM payTypes WHERE gwCode="&pcv_EditGW
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	pcTempIdPayment = rs("idPayment")
	call closedb()

	response.redirect "pcPGActionCompleted.asp?id="&pcTempIdPayment&"&gwchoice="&pcv_EditGW&"&msg="&Server.URLEncode("Gateway has been successfully edited.")
end if

If sMode <> "" Then
	If sMode="Add Gateway" Then
		dim varCheck
		varCheck=0
		'Start SDBA
		pcv_processOrder=request.Form("pcv_processOrder")
		if pcv_processOrder="" then
			pcv_processOrder="0"
		end if
		pcv_setPayStatus=request.Form("pcv_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="3"
		end if
		'End SDBA
		'//Gateway to activate
		pcv_AddGW=request("addGW")
		call gwCallAdd()
	end if
	
	If sMode="Add Gateway" and varCheck=1 then
		call opendb()
		query="SELECT idPayment FROM payTypes WHERE gwCode="&pcv_AddGW
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		pcTempIdPayment = rs("idPayment")
		call closedb()
		response.redirect "pcPGActionCompleted.asp?id="&pcTempIdPayment&"&gwchoice="&pcv_AddGW&"&msg="&Server.URLEncode("Gateway has been successfully enabled.")
	else
		response.redirect "pcPaymentSelection.asp?msg="&Server.URLEncode("You did not specify a payment option to add. Make sure that you check the box next to the payment option that you wish to add.")
	end if
end if

if request("mode")="Edit" then
	gwCode=request("gwChoice")
	idPayment=Request.QueryString("id")
	query= "SELECT paymentDesc, priceToAdd, cvv, percentageToAdd, sslUrl, terms, CReq, Cprompt, Cbtob, paymentNickName, pcPayTypes_processOrder, pcPayTypes_setPayStatus FROM payTypes WHERE gwCode= "& gwCode &" AND idPayment= "& idPayment
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?v=5&error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if NOT rs.eof then
		paymentDesc=rs("paymentDesc")
		priceToAdd=rs("priceToAdd")
		cvv=0
		percentageToAdd=rs("percentageToAdd")
		sslUrl=rs("sslUrl")
		terms=rs("terms")
		CReq=rs("CReq")
		Cprompt=rs("Cprompt")
		Cbtob=rs("Cbtob")
		paymentNickName=rs("paymentNickName")
		'Start SDBA
		pcv_processOrder=rs("pcPayTypes_processOrder")
		pcv_setPayStatus=rs("pcPayTypes_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="3"
		end if
		'End SDBA
	End If
	
	set rs=nothing
	if percentageToAdd<>"0" then
		priceToAddType="percentage"
	end if
	if priceToAdd<>"0" then
		priceToAddType="price"
	end if
else
	paymentNickName="Credit Card"
	paymentNickName2="Check"
	percentageToAdd="0"
	priceToAdd="0"
end if


call closedb()
%>

<form name="formname" method="post" action="<%=pcStrPageName%>" class="pcForms">  
	<table class="pcCPcontent">
		<tr> 
			<td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
		</tr>
		<tr> 
			<td colspan="2">
            	<!-- Gateway Include Start -->
				<!--#include file="RTGatewayIncludes.asp"-->
            	<!-- Gateway Include End -->
            </td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
<% Function IsObjInstalled(intClassNum)
	On Error Resume Next
	CONST CLASSBOUND = 2
	Dim objTest, j
	Dim strError
	strError = ""
	for j = 0 to CLASSBOUND
		If Not IsEmpty(strClass(intClassNum, j)) Then 		
			Set objTest = Server.CreateObject(strClass(intClassNum, j))	
			If Err.Number = 0 Then
				Set objTest = Nothing
				If Not isEmpty(strClass(intClassNum, 3)) Then
				strError = ""
				End if
			Else
				If IsObject(objTest) Then Set objTest = Nothing
				If strError = "" Then
					strError = strClass(intClassNum, j)
				Else
					strError = strError & ",<br>" & strClass(intClassNum, j)
				End If
			End If
		Else
			If Not isEmpty(strClass(intClassNum, 3)) Then
				errArray = split(strError, ",<br>", -1)
				if ubound(errArray) = 1 then
					strError = errArray(0)
					strError = strError & "<BR> or " & errArray(1)
				end if 
			End if  				
		End If	
	Next
	IsObjInstalled = strError
End Function
%>
