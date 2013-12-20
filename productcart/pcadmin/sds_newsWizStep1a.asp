<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<%
dim rs, conntemp, query, AddrList

AddrList=""

if request("pagetype")="1" then
	pcv_PageType=1
	pcv_Title="Drop-Shipper"
	pcv_Table="pcDropShipper"
else
	pcv_PageType=0
	pcv_Title="Supplier"
	pcv_Table="pcSupplier"
end if

IF request("action")="add" THEN
	pcv_SDSList=request("sdslist")
	if (trim(pcv_SDSList)="") then
		response.redirect "menu.asp"
	end if
	pcArr=split(pcv_SDSList,",")
	call opendb()
	For i=lbound(pcArr) to ubound(pcArr)
	if trim(pcArr(i)<>"") then
		if pcv_PageType="0" then
			query="SELECT pcSupplier_Email As Email FROM pcSuppliers WHERE pcSupplier_ID=" & pcArr(i) & " ORDER BY pcSupplier_Email ASC;"
		else
			pcArr1=split(pcArr(i),"_")
			if pcArr1(1)="1" then
				query="SELECT pcSupplier_Email As Email FROM pcSuppliers WHERE pcSupplier_ID=" & pcArr1(0) & " AND pcSupplier_IsDropShipper=" & pcArr1(1) & " ORDER BY pcSupplier_Email ASC;"
			else
				query="SELECT pcDropShipper_Email As Email FROM pcDropShippers WHERE pcDropShipper_ID=" & pcArr1(0) & " ORDER BY pcDropShipper_Email ASC;"
			end if
		end if
		set rs=connTemp.execute(query)
		if not rs.eof then
			AddrList=AddrList & rs("Email") & "**"
		end if
		set rs=nothing
	end if
	Next
	call closedb()

	'Sort e-mail address list
	if AddrList<>"" then
		AList=split(AddrList,"**")
		session("AddrList")=AList
		session("AddrCount")=ubound(AList)
	else

		dim BList(1)
		session("AddrList")=BList
		session("AddrCount")=0
	end if
	response.redirect "newsWizStep2.asp?from=1&pagetype=" & pcv_PageType
END IF
%>