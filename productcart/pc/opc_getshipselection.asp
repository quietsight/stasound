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

	query="SELECT pcCustSession_ShippingArray FROM pcCustomerSessions WHERE idcustomer=" & session("idCustomer") & " AND (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
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
		pcv_strShipArray = rs("pcCustSession_ShippingArray")
		
		pcv_strFormatted = ""
		
		If len(pcv_strShipArray)>0 Then

			shipping=split(pcv_strShipArray,",")
			if ubound(shipping)>1 then
				if NOT isNumeric(trim(shipping(2))) then
					varShip="0"
					'response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
					Service="No shipping required"
				else
					Shipper=shipping(0)
					Service=shipping(1)
				end if
			else
				Service="No shipping required"
				varShip="0"
			end if 	
				
			pcv_strFormatted = Service
			
			pcv_strEditLink = "<a id=""btnEditCO"" href=""javascript:;"" onclick=""javascript: acc1.openPanel('opcShipping'); $('#TaxContentArea').hide(); $('#ShippingChargeArea').show(); GoToAnchor('opcLoginAnchor'); "">" & dictLanguage.Item(Session("language")&"_opc_53") & "</a>"
			
		OKmsg = ""
		OKmsg = OKmsg & "<div class=""editbox"" style=""margin-top: 6px;"">"
		OKmsg = OKmsg & dictLanguage.Item(Session("language")&"_opc_54") & pcv_strFormatted & " " & pcv_strEditLink	
		OKmsg = OKmsg & "</div>"
			
		Else


		End If

	end if
	set rs = Nothing
	
end if

if pcErrMsg<>"" then
	pcErrMsg=""
	response.write pcErrMsg
else
	response.write OKmsg
end if
call closeDb()
response.End()
%>


