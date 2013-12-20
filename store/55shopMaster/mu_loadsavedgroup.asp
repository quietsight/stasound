<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->    
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim rstemp, conntemp, query, AddrList

idgroup=getUserInput(request("idgroup"),0)
listid=getUserInput(request("listid"),0)

if idgroup="" or listid="" then
	response.redirect "mu_newsWizStep1.asp"
end if

session("CP_NW_ListID")=listid

AddrList=""

call opendb()
query="SELECT pcMailUpSavedGroups_Data FROM pcMailUpSavedGroups WHERE pcMailUpSavedGroups_ID=" & idgroup & ";"
set rs=connTemp.execute(query)

if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "mu_newsWizStep1.asp"
else
	AddrList=rs("pcMailUpSavedGroups_Data")
	set rs=nothing
	AList=split(AddrList,vbcrlf)

	For dem=lbound(AList) to ubound(AList)
		For dem1=dem+1 to ubound(AList)
			if AList(dem)>AList(dem1) then
				TempStr=AList(dem1)
				AList(dem1)=AList(dem)
				AList(dem)=TempStr
			end if
		next
	next

	session("AddrList")=AList
	session("AddrCount")=ubound(AList)
end if
call closedb()
response.redirect "mu_newsWizStep2.asp?from=1"
%>