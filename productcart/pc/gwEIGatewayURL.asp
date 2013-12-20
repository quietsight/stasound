<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<% 'Gateway specific files %>
<%
'On Error Resume Next
Dim conntemp

Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If

'// Get Order ID
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% session("idCustomer")=pcIdCustomer
pcv_IncreaseCustID=(scCustPre + int(pcIdCustomer)) %>

<% 
'// ADD TO VAULT
Dim pcv_strIsSaved
pcv_strIsSaved=getUserInput(request.querystring("IsSaved"),5)
If len(pcv_strIsSaved)>0 Then
	Session("CustomerVaultID")=""
	Session("VaultID")=""
	Session("SF_IsSaved")=pcv_strIsSaved
End If

'// Save the Card Type
Dim pcv_strCardType
pcv_strCardType=getUserInput(request.querystring("CardType"),10)
If len(pcv_strCardType)>0 Then
	Session("CardType")=pcv_strCardType
End If

'// Save Card Number
Dim pcv_strCardNum
pcv_strCardNum=getUserInput(request.querystring("CardNum"),20)
If len(pcv_strCardNum)>0 Then
	Session("CardNum")=pcv_strCardNum
End If

'// Save Exp Date
Dim pcv_strExpDate
pcv_strExpDate=getUserInput(request.querystring("ExpDate"),20)
If len(pcv_strExpDate)>0 Then
	Session("CardExp")=pcv_strExpDate
End If


'// LOAD SSETTINGS
call opendb()
query="SELECT pcPay_EIG_Type, pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key, pcPay_EIG_Curcode, pcPay_EIG_CVV, pcPay_EIG_TestMode, pcPay_EIG_SaveCards, pcPay_EIG_UseVault FROM pcPay_EIG WHERE pcPay_EIG_ID=1"
set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)		
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if	
x_Username=rs("pcPay_EIG_Username")
x_Username=enDeCrypt(x_Username, scCrypPass)
x_Password=rs("pcPay_EIG_Password")
x_Password=enDeCrypt(x_Password, scCrypPass)
x_Key=rs("pcPay_EIG_Key")
x_Key=enDeCrypt(x_Key, scCrypPass)
x_CVV=rs("pcPay_EIG_CVV")
x_Type=rs("pcPay_EIG_Type")
x_TypeArray=Split(x_Type,"||")
x_TransType=x_TypeArray(0)
x_Curcode=rs("pcPay_EIG_Curcode")
x_TestMode=rs("pcPay_EIG_TestMode")
x_SaveCards=rs("pcPay_EIG_SaveCards")
x_UseVault=rs("pcPay_EIG_UseVault")
set rs=nothing
call closedb()

'/////////////////////////////////////////////////////////////////////////////////////////////
'// START: GET FORM URL
'/////////////////////////////////////////////////////////////////////////////////////////////
Dim VaultID, pcv_strCustomerVaultID
VaultID=getUserInput(request("VaultID"),0)
If len(VaultID)>0 Then

	Session("SF_IsSaved")=""
	
	call openDB()
	query="SELECT pcPay_EIG_Vault_Token FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_ID="& VaultID &""
	set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)		
	if NOT rs.eof then
		pcv_strCustomerVaultID = rs("pcPay_EIG_Vault_Token")
		If len(pcv_strCustomerVaultID)=0 Then
			pcv_strCustomerVaultID="NA"
		Else
			pcv_strCustomerVaultID=enDeCrypt(pcv_strCustomerVaultID, scCrypPass)
		End If
		Session("VaultID")=VaultID
		Session("CustomerVaultID")=pcv_strCustomerVaultID
	else
		response.Clear
		response.Write(dictLanguage.Item(Session("language")&"_EIG_8"))
		response.End
	end if
	call closeDB()				
	
End If

Dim returnURL
If scSSL="" OR scSSL="0" Then
	returnURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwEIGateway.asp"),"//","/")
	returnURL=replace(returnURL,"https:/","https://")
	returnURL=replace(returnURL,"http:/","http://") 
Else
	returnURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwEIGateway.asp"),"//","/")
	returnURL=replace(returnURL,"https:/","https://")
	returnURL=replace(returnURL,"http:/","http://")
End If

Dim strTest
strTest = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
If x_TransType="AUTH_ONLY" Then
	strTest = strTest & "<auth>"
Else
	strTest = strTest & "<sale>"		
End If
strTest = strTest & "<api-key>" & x_Key & "</api-key>"
strTest = strTest & "<redirect-url>" & returnURL & "</redirect-url>"
strTest = strTest & "<amount>" & pcBillingTotal & "</amount>"
strTest = strTest & "<currency>" & x_Curcode & "</currency>"
strTest = strTest & "<order-id>" & session("GWOrderId") & "</order-id>"
'strTest = strTest & "<customer-id>" & session("idCustomer") & "</customer-id>"

If len(Session("CustomerVaultID"))>0 Then
	strTest = strTest & "<customer-vault-id>" & Session("CustomerVaultID") & "</customer-vault-id>"
End If

'// If ("AUTH_ONLY" & Vault Storage Enabled) OR (Customer Save & Not Yet Saved) AND (No Existing Vault ID)
If ((x_TransType="AUTH_ONLY" AND x_UseVault="1") OR (x_SaveCards="1" AND Session("SF_IsSaved")="true")) AND (len(Session("CustomerVaultID"))=0) Then
	strTest = strTest & "<add-customer>"
	strTest = strTest & "<customer-vault-id></customer-vault-id>"
	strTest = strTest & "</add-customer>"
End If		

If x_TransType="AUTH_ONLY" Then
	strTest = strTest & "</auth>"
Else
	strTest = strTest & "</sale>"		
End If

'response.Clear()
'response.ContentType="text/xml"
'response.Write(strTest)
'response.End()

set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
xml.open "POST", "https://secure.nmi.com/api/v2/three-step", false
xml.setRequestHeader "Content-Type", "text/xml"
xml.send strTest
strStatus = xml.Status
strRetVal = xml.responseText
Set xml = Nothing

'response.Clear()
'response.ContentType="text/xml"
'response.Write(strRetVal)
'response.End()

strResult = pcf_GetNode(strRetVal, "result", "*")
strResultText = pcf_GetNode(strRetVal, "result-text", "*")
strTransactionID = pcf_GetNode(strRetVal, "transaction-id", "*") 
strResultCode = pcf_GetNode(strRetVal, "result-code", "*")
strFormURL = pcf_GetNode(strRetVal, "form-url", "*")
pcv_strCustomerVaultID = pcf_GetNode(strRetVal, "customer_vault_id", "*")
If len(pcv_strCustomerVaultID)>0 Then
	Session("CustomerVaultID")=pcv_strCustomerVaultID
End If
'response.Write("strResult:  " & strResult & "<br />")
'response.Write("strResultText:  " & strResultText & "<br />")
'response.Write("strTransactionID:  " & strTransactionID & "<br />")
'response.Write("strResultCode:  " & strResultCode & "<br />")
'response.Write("strFormURL:  " & strFormURL & "<br />")
'response.Write("pcv_strCustomerVaultID:  " & pcv_strCustomerVaultID & "<br />")
'response.End()

If len(strFormURL)=0 Then
	If x_TestMode="1" Then
		strError=strResultText
	Else
		strError=strResultText
	End If				
End If

'/////////////////////////////////////////////////////////////////////////////////////////////
'// END: GET FORM URL
'/////////////////////////////////////////////////////////////////////////////////////////////


response.Clear
If len(strError)>0 Then
	response.write strError
Else	
	response.write "OK||" & strFormURL
End If
response.End()


Function pcf_GetNode(responseXML, nodeName, nodeParent)
	Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)				 
	myXmlDoc.loadXml(responseXML)
	Set Nodes = myXmlDoc.selectnodes(nodeParent)	
	For Each Node In Nodes	
		pcf_GetNode = pcf_CheckNode(Node,nodeName,"")				
	Next
	Set Node = Nothing
	Set Nodes = Nothing
	Set myXmlDoc = Nothing
End Function

Function pcf_CheckNode(Node,tagName,default)		
	Dim tmpNode
	Set tmpNode=Node.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		pcf_CheckNode=default
	Else
		pcf_CheckNode=Node.selectSingleNode(tagName).text
	End if
End Function

Function pcf_FixXML(str)	
	str=replace(str, "&","and")	
	pcf_FixXML=str
End Function
%>