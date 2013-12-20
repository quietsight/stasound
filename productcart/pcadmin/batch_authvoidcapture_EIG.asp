<%PmAdmin=9%>
<% 
err.clear

'// VOID
strTest = ""
strTest = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
strTest = strTest & "<void>"
strTest = strTest & "<api-key>" & x_Key & "</api-key>"
strTest = strTest & "<transaction-id>" & Request.Form("transid"&r) & "</transaction-id>"
strTest = strTest & "</void>"

set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
xml.open "POST", "https://secure.nmi.com/api/v2/three-step", false
xml.setRequestHeader "Content-Type", "text/xml"
xml.send strTest
strStatus = xml.Status
strRetVal = xml.responseText
Set xml = Nothing

strResult = pcf_GetNode(strRetVal, "result", "*")

'// VOID SUCCESS
if strResult = 1 then

	'// CAPTURE NEW TRANSACTION, if no errors	
	strTest = ""
	strTest = strTest & "username=" & x_Username
	strTest = strTest & "&password=" & x_Password
	strTest = strTest & "&type=sale"
	If len(x_VaultToken)>0 Then
		strTest = strTest & "&customer_vault_id=" & x_VaultToken
	Else
		strTest = strTest & "&ccnumber=" & Request.Form("ccnum"&r)
		strTest = strTest & "&ccexp=" & Request.Form("ccexp"&r)
		'strTest = strTest & "&cvv=" & ""
	End If
	strTest = strTest & "&amount=" & curamount

	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.open "POST", "https://secure.nmi.com/api/transact.php", false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send strTest
	strStatus = xml.Status
	strRetVal = xml.responseText
	Set xml = Nothing
	
	Set resArray = deformatNVP(strRetVal)

	strResult = UCase(resArray("result"))
	strResultText = UCase(resArray("result-text"))
	strTransactionID = UCase(resArray("transaction-id"))
	strResultCode = UCase(resArray("result-code"))
	strAuthorizationCode = UCase(resArray("authorization-code"))
	pcv_strCustomerVaultID = UCase(resArray("customer-vault-id"))  

	'response.Write(strRetVal & ".<br />")
	'response.Write(strResult & ".<br />")
	'response.Write(strResultText & ".<br />")
	'response.Write(strTransactionID & ".<br />")
	'response.Write(strResultCode & ".<br />")
	'response.Write(authorization-code & ".<br />")
	'response.Write(pcv_strCustomerVaultID & ".<br />")
	'response.End

end if
err.clear %>
