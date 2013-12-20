<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->

<% 
	dim connTemp, query, rs, qry_ID
 
	pcvAction=request("action")
	if lcase(pcvAction)<>"archive" and lcase(pcvAction)<>"unarchive" then
		response.redirect "msgb.asp?msg=" & Server.URLEncode("Form action not valid. The system cannot archive or unarchive the selected orders.")
	end if
	
	TmpStr=""
	Count=request("count")
	if (Count="") or (Count="0") then
		response.redirect "msgb.asp?msg=" & Server.URLEncode("Number of orders not a valid number.")
	end if

	For k=1 to Count
		if (request("check" & k)="1") and (request("idord" & k)<>"") then
			TmpStr=TmpStr & request("idord" & k) & "***"
		end if
	Next
	
	if session("CP_OrdSrcPages")>"0" then
		pagepre=request("curpage")
		For i=1 to Clng(session("CP_OrdSrcPages"))
		if i<>clng(pagepre) then
			tmpArr=split(session("CP_OrdSrcPage"&i),",")
			For m=0 to ubound(tmpArr)
				if tmpArr(m)<>"" then
					TmpStr=TmpStr & tmpArr(m) & "***"
				end if
			Next
		end if
		Next
	end if

	if TmpStr="" then
		response.redirect "msgb.asp?msg=" & Server.URLEncode("The order array is empty.")
	end if
	
	call openDb()
	A=split(TmpStr,"***")
	For k=lbound(A) to ubound(A)
		IF validNum(A(k)) Then
			qry_ID=A(k)
			if pcvAction="archive" then
				query="UPDATE orders SET pcOrd_Archived=1 WHERE idOrder=" & qry_ID & ";"
			else
				query="UPDATE orders SET pcOrd_Archived=0 WHERE idOrder=" & qry_ID & ";"
			end if			
			Set rs=Server.CreateObject("ADODB.Recordset")
			Set rs=connTemp.execute(query)
		End If
	Next
	set rs=nothing
	call closedb()
	
	if pcvAction="archive" then
		response.redirect "resultsAdvancedAll.asp?s=1&B1=View+All&dd=1&msg=" & ubound(A) & server.URLEncode(" order(s) successfully archived.")
	else
		response.redirect "resultsAdvancedAll.asp?s=1&B1=View+All&dd=1&msg=" & ubound(A) & server.URLEncode(" order(s) successfully unarchived.")
	end if
%>