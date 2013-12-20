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

   Dim intListID, strListName, strMode

	intListID=Request.Form("lid")
	if not validNum(intListID) then
		Response.Redirect "manageOptions.asp"
	end if
	Dim rs, connTemp, strSql
	set connTemp=server.createobject("adodb.connection")
 	connTemp.Open scDSN 
	Dim intAddressID, strMsg	
	For Each intAddressID in Request.Form("optionDescrip")
		
		strSQL="Delete From options_OptionsGroups WHERE idProduct=" & intAddressID & " AND idOptionGroup=" & intListID
		connTemp.Execute strSQL, , adCmdText + adExecuteNoRecords
		
		contgo=0
		'// Check if all options have been removed.
		strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="& intAddressID &" AND idoptionGroup="& intListID &";"
		set rstemp=conntemp.execute(strSQL)							
		if rstemp.eof then
			'// It is NOT related
			contgo=1
		end if			
		'// If all Options have been removed then delete the corrisponding record in pcProductOptions
		if contgo=1 then				
			strSQL="DELETE FROM pcProductsOptions WHERE idproduct="& intAddressID &" AND idoptionGroup="& intListID &";"
			set rstemp=conntemp.execute(strSQL)
		end if	
	Next
	connTemp.Close
	Set connTemp=Nothing
	set rstemp=nothing
	Response.Redirect("editGrpOptions.asp?idOptionGroup="&intListID&"&mode=view&msg=" & strMsg)
%>