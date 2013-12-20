<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<%
dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT ProtxTestmode FROM protx Where idProtx=1;"

set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'//Declare URL path to gwSubmit.asp	
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

'// Set gateway specific variables
pcv_StrProtxTestmode=rs("ProtxTestmode")

set rs=nothing
call closedb()

if pcv_StrProtxTestmode=0 then
	strVSPDirect3DCallbackPage="https://live.sagepay.com/gateway/service/direct3dcallback.vsp"
elseif pcv_StrProtxTestmode=2 then
  	strVSPDirect3DCallbackPage="https://test.sagepay.com/gateway/service/direct3dcallback.vsp"
else
	strVSPDirect3DCallbackPage="https://test.sagepay.com/Simulator/VSPDirectCallback.asp"
end if

'**************************************************************************************************
' Description
' ===========
'
' This page is the 3D-Secure completion page that redeives the MD and PaRes from the Issuing Bank 
' site, POSTs it to Protx, then reads the authorisation response and updates the database accordingly.
'**************************************************************************************************

' ** Otherwise, create the POST for Protx ensuring to URLEncode the PaRes before sending it **
strMD=request.form("MD")
strPaRes=request.form( "PaRes" )
strVendorTxCode=Session("VendorTxCode")

'** POST for Protx VSP Direct 3D completion page **
strPost = "MD=" & strMD & "&PARes=" & Server.URLEncode(strPaRes)

'** Use the Windows WinHTTP object to POST the data directly from this server to Protx **
set httpRequest = CreateObject("WinHttp.WinHttprequest.5.1")
		
on error resume next
httpRequest.Open "POST", CStr(strVSPDirect3DCallbackPage), false
httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
httpRequest.send strPost
responseData = httpRequest.responseText

'** An non zero Err.number indicates an error of some kind **
'** Check for the most common error... unable to reach the purchase URL ** 
strPageError="" 
if err.number<>0 then
	if Err.number = -2147012889 then
		strPageError="Your server was unable to register this transaction with SagePay." &_
					"  Check that you do not have a firewall restricting the POST and " &_
					"that your server can correctly resolve the address " & strPurchaseURL
	else
		strPageError="An Error has occurred whilst trying to register this transaction.<BR>" &_
					"The Error Number is: " & Err.number & "<BR>" &_
					"The Description given is: " & Err.Description
	end If 
	if strPageError<>"" then
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:&nbsp;"&strProtxErrorMsg&"<br><br><a href="""&tempURL&"?psslurl=gwProtxVSP.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
	end if
end if

' ******************************************************************
' Determine next action
'** No transport level errors, so the message got the Protx **
'** Analyse the response from VSP Direct to check that everything is okay **
'** Registration results come back in the Status and StatusDetail fields **
strStatus=findField("Status",responseData)
strStatusDetail=findField("StatusDetail",responseData)
		
'** If this isn't 3D-Auth, then this is an authorisation result (either successful or otherwise) **
'** Get the results form the POST if they are there **
strVPSTxId=findField("VPSTxId",responseData)
strSecurityKey=findField("SecurityKey",responseData)
strTxAuthNo=findField("TxAuthNo",responseData)
strAVSCV2=findField("AVSCV2",responseData)
strAddressResult=findField("AddressResult",responseData)
strPostCodeResult=findField("PostCodeResult",responseData)
strCV2Result=findField("CV2Result",responseData)
str3DSecureStatus=findField("3DSecureStatus",responseData)
strCAVV=findField("CAVV",responseData)
	
if strStatus="OK" then
	session("GWAuthCode")=strTxAuthNo
	session("GWTransId")=strVPSTxId
	session("GWTransType")=pcv_StrTxType
	Response.redirect "gwReturn.asp?s=true&gw=SagePay"
else
	if strStatus="AUTHENTICATED" then
		session("GWAuthCode")=strTxAuthNo
		session("GWTransId")=strVPSTxId
		session("GWTransType")=pcv_StrTxType
		Response.redirect "gwReturn.asp?s=true&gw=SagePay"
	end if
	' ** Something has gone wrong, record the error and redirect etc.
	 strProtxErrorType=strStatus
	 strProtxErrorMsg=strStatusDetail
	'REJECTED, NOTAUTHED, ERROR redirect back to payment form
	response.redirect "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:&nbsp;"&strProtxErrorType&" - "&strProtxErrorMsg&"<br><br><a href="""&tempURL&"?psslurl=gwProtxVSP.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
	' ** Write VPSTxID, SecurityKey, Status and StatusDetail to the screen, log file or database
	response.write "<b>Failed</b><br/>"
	response.end
end if

' ******************************************************************
' remove the reference to the object
set httpRequest = nothing
%>
<!--#include file="footer.asp"-->
<% 
'***********************************************
' Useful methods
'***********************************************

function findField( fieldName, postResponse )
  items = split( postResponse, chr( 13 ) )
  for idx = LBound( items ) to UBound( items )
    item = replace( items( idx ), chr( 10 ), "" )
    if InStr( item, fieldName & "=" ) = 1 then
      ' found
      findField = right( item, len( item ) - len( fieldName ) - 1 )
      Exit For
    end if
  next 
end function
%>
