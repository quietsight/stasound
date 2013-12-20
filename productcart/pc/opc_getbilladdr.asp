<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="opc_contentType.asp" -->
<% On Error Resume Next
dim rs,connTemp,query

Call SetContentType()

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

call openDb()


pcErrMsg=""
if session("idCustomer")>"0" then

	query="SELECT * FROM Customers WHERE idcustomer=" & session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closeDb()
		response.clear
		Call SetContentType()
		response.write "ERROR"
		response.End
	end if
	if NOT rs.eof then
		pcv_strCustName=rs("name") & " " & rs("lastName")
		pcv_strAddress = rs("address")
		pcv_strAddress2 = rs("address2")
		pcv_strPostalCode = rs("zip")
		pcv_strProvince = rs("stateCode")
		pcv_strState = rs("state")
		pcv_strCity= rs("city")
		pcv_strCountry = rs("countryCode")
		
		pcv_strFormatted = ""
		pcv_strFormatted = pcv_strFormatted & pcv_strCustName & " - " & pcv_strAddress
		
		If len(pcv_strAddress2)>0 Then
			pcv_strFormatted = pcv_strFormatted & " " & pcv_strAddress2
		End If
		
		If len(pcv_strCity)>0 Then
			pcv_strFormatted = pcv_strFormatted & " " & pcv_strCity
		End If

		If len(pcv_strProvince)>0 Then
			pcv_strFormatted = pcv_strFormatted & ", " & pcv_strProvince
		Else
			If len(pcv_strState)>0 then
				pcv_strFormatted = pcv_strFormatted & ", " & pcv_strState	
			End If
		End If

		If len(pcv_strCountry)>0 Then
			pcv_strFormatted = pcv_strFormatted & ", " & pcv_strCountry
		End If
		
		If len(pcv_strPostalCode)>0 Then
			pcv_strFormatted = pcv_strFormatted & " " & pcv_strPostalCode
		End If

		pcv_strEditLink = "<a id=""btnEditCO"" href=""javascript:;"" onclick=""javascript: acc1.openPanel('opcLogin'); GoToAnchor('opcLoginAnchor'); $('#BillingArea').show(); $('#ShippingArea').hide(); $('#TaxContentArea').hide(); $('#BillingLoader').hide();"">" & dictLanguage.Item(Session("language")&"_opc_53") & "</a>"
		
		OKmsg = ""
		OKmsg = OKmsg & "<div class=""editbox"" style=""margin-top: 6px;"">"
		OKmsg = OKmsg & dictLanguage.Item(Session("language")&"_opc_8a") & pcv_strFormatted & " " & pcv_strEditLink	
		OKmsg = OKmsg & "</div>"

	end if
	set rs = Nothing
	
end if

if pcErrMsg<>"" then
	response.write pcErrMsg
else
	response.write OKmsg
end if
call closeDb()
response.End()
%>
 