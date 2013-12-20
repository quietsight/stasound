<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
prdType=request("prdType")
Main

Sub Main
	Dim intListID, strListName, strMode
'--- Initialize required variables ---
	strMode=""
    
    Dim frmHideInactiveProducts 
    frmHideInactiveProducts = Request("frmHideInactiveProducts")
    
	Dim reqstr, reqidproduct, reqidcategory
	reqstr=Request("reqstr")
	reqidproduct=Request("idproduct")
	reqidcategory=Request("idcategory")
	If reqstr="" then
		reqstr=Request("reqstr")
		reqidproduct=Request("reqidproduct")
		reqidcategory=Request("reqidcategory")
	End If
	
	If Request("delete") <> "" Then
		DeleteList(Request("delete"))
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	ElseIf Request("btnCreate") <> "" AND Request("ListName") <> "" Then
		CreateList(Replace(Request("ListName"), "'", "''"))
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End If

	If Request("lid") <> "" Then
		intListID=Request("lid")
	Else
		Response.Write("No lid")
		Exit Sub
	End If
	
	If Request("CopyTo")<>"" OR Request("MoveTo")<>"" then
		session("cp_cmt_prdlist")=request("Address")
		session("cp_cmt_idcat")=intListID
		If Request("CopyTo")<>"" then
			session("cp_cmt_action")=1
		Else
			session("cp_cmt_action")=2
		End if
		response.redirect "CopyMovePrds.asp"
	End if

	If Request("mode") <> "" Then
		strMode=Request("mode")
	Else
		Response.Write("No Mode")
		Exit Sub
	End If
'---

	Dim rs, connTemp, strSql

	set connTemp=server.createobject("adodb.connection")
 	connTemp.Open scDSN 
	Dim intAddressID, strMsg
	If strMode="view" Then
		if request("UpdateOrder")<>"" then
			A=split(Request("POrder"),", ")
			B=split(Request("Listidproduct"),", ")
			
			For k=lbound(B) to ubound(B)
				Pidpro=B(k)
				POrder=A(k)
				if POrder<>"" then
					if not IsNumeric(POrder) then
					POrder=0
					end if
				else
					POrder=0
				end if
			
				strSQL="Update categories_products set POrder=" & POrder & " WHERE idProduct=" & Pidpro & " AND idCategory=" & intListID
				connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
				if prdType=3 then
					strMsg="Product order%20has%20been%20updated.&s=1"
				else
					strMsg="Product order%20has%20been%20updated.&s=1"
				end if
			Next
		elseif request("ResetOrder")<>"" then
			strSQL="Update categories_products set POrder=0 WHERE idCategory=" & intListID
			connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
			strMsg="Product%20order%20has%20been%20reset.&s=1"
		else		
			For Each intAddressID in Request("Address")
				strSQL="DELETE FROM categories_products WHERE idProduct=" & intAddressID & " AND idCategory=" & intListID
				connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
				if prdType=3 then
					strMsg="BTO Item(s)%20Removed%20From%20Category&s=1"
				else
					strMsg="Product(s)%20Removed%20From%20Category&s=1"
				end if
			Next
		end if	
	Else
		if (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
			strSQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" & id & ", " & intListID & ")"
			connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
			if prdType=3 then
				strMsg="BTO Item(s)%20Added%20to%20Category&s=1"
			else
				strMsg="Product(s)%20Added%20to%20Category&s=1"
			end if
			end if
		Next
		end if
	End If

	connTemp.Close
	Set connTemp=Nothing
	Response.Redirect("editCategories.asp?prdType="&prdType&"&lid="&intListID&"&mode=view&msg="&strMsg&"&reqstr="&reqstr&"&reqidproduct="&reqidproduct&"&reqidcategory="&reqidcategory&"&frmHideInactiveProducts="&frmHideInactiveProducts)
End Sub

Sub DeleteList(intListID)
  Dim connTemp, strSQL

  call openDb()
	connTemp.BeginTrans
	strSQL="DELETE FROM categories WHERE idCategory=" & intListID
	connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
	strSQL="DELETE FROM pcDFCats WHERE pcFCat_IDCategory=" & intListID
	connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
	strSQL="DELETE FROM categories_products WHERE idCategory=" & intListID
	connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
	If Err Then
		connTemp.RollbackTrans
	Else
		connTemp.CommitTrans
	End If
	connTemp.Close
	Set connTemp=Nothing
End Sub

%>