<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<% 
on error resume next
Dim connTemp,query,rs
dim pcCartArray, ppcCartIndex

Sub RmvFrmCart(tmpvalue)
	for i=tmpvalue to UBound(pcCartArray,1)-1
	for j=0 to UBound(pcCartArray,2)
		pcCartArray(i,j)=pcCartArray(i+1,j)
		if j=27 AND pcCartArray(i,j) = i then
		    pcCartArray(i,j) = i-1
		end if
	next
	next
End Sub

call opendb()

'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=server.HTMLEncode(request.querystring("pcCartIndex"))

if ppcCartIndex="" then
  response.redirect "msg.asp?message=19"
end if

DelPrds=1

'// Check if this is a Cross Sell Child
if (cint(pcCartArray(ppcCartIndex,27))>0) then
    '//  Is this a Bundle?
    if (cint(pcCartArray(ppcCartIndex,12))=0) then  
        pcCartArray(cint(pcCartArray(ppcCartIndex,27)),28)=cint(0)  '// Set Parent discount to 0
        pcCartArray(cint(pcCartArray(ppcCartIndex,27)),8)=""        '// Set no bundle child
    end if
else
    '// Check if this is a Cross Sell Parent
    for i=(ppcCartIndex+1) to UBound(pcCartArray,1)-1
        '// Is this a Child?
        if (cint(pcCartArray(i,27))=cint(ppcCartIndex)) then
				'// Check "Not For Sale" Product
				query="SELECT formQuantity FROM Products WHERE idproduct=" & pcCartArray(i,0) & " AND formQuantity<>0;"
				set rs=connTemp.execute(query)
				if not rs.eof then
					'// If "YES" then remove it
					pcCartArray(i,10)=1
					RmvFrmCart(i)
					DelPrds=DelPrds+1
				else
					'// Set to normal product (show Remove Button)
					pcCartArray(i,12)=cint(0)
				end if
				set rs=nothing
            pcCartArray(i,27)=0  '// Set parent to 0
            pcCartArray(i,28)=0  '// Set discount to 0
            pcCartArray(i,8)=""  '// Set no bundle child
        end if
    next
end if

' mark item as removed
'pcCartArray(ppcCartIndex,10)=1

if session("sf_FQuotes")<>"" then 
'Delete from temponary Finalized Quotes array
	session("sf_FQuotes")=replace(session("sf_FQuotes"),"****" & pcCartArray(ppcCartIndex,0) & "****","")
end if

for i=ppcCartIndex to UBound(pcCartArray,1)-1
	for j=0 to UBound(pcCartArray,2)
		pcCartArray(i,j)=pcCartArray(i+1,j)
		if j=27 AND pcCartArray(i,j) = i then
		    pcCartArray(i,j) = i-1
		end if
	next
next

ppcCartIndex=session("pcCartIndex")-DelPrds%>
<!--#include file="inc-UpdPrdQtyDiscounts.asp"-->
<%
call closedb()
session("pcCartSession")=pcCartArray
session("pcCartIndex")=session("pcCartIndex")-DelPrds

set pcCartArray=nothing
set pcCartIndex=nothing
set ppcCartIndex=nothing

call clearLanguage()
response.redirect "viewCart.asp"
%>