<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove XML Partner" %>
<% section="layout"%>
<%PmAdmin=19%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<%
pidPartner=trim(request("idPartner"))
tmpIDLog=trim(request("idxml"))

If Not IsNumeric(pidPartner) then
	response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

If Not IsNumeric(tmpIDLog) then
	response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

if request("action")<>"del" then
	response.redirect "menu.asp"
end if

dim conntemp, query, rs

Dim xmlRequestID_value
Dim xmlRequestKey_value
Dim xmlRequestType_value
Dim	xmlBackup_value
Dim xmlUndo_value

Call CheckUndoRequestTags()
Call RunUndoRequest()

Sub CheckUndoRequestTags()
Dim query,rs
	
	call opendb()

	query="SELECT pcXMLLogs.pcXL_id,pcXMLLogs.pcXL_RequestKey,pcXMLLogs.pcXL_RequestType,pcXMLLogs.pcXL_BackupFile,pcXMLLogs.pcXL_Undo FROM pcXMLPartners INNER JOIN pcXMLLogs ON pcXMLPartners.pcXP_ID=pcXMLLogs.pcXP_ID WHERE pcXMLLogs.pcXL_id=" & tmpIDLog & " AND pcXMLPartners.pcXP_ID=" & pidPartner & ";"
	set rs=connTemp.execute(query)
	
	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "viewXMLPartnerLogs.asp?msg=1&idPartner=" & pidPartner
	else
		xmlRequestID_value=rs("pcXL_id")
		xmlRequestKey_value=rs("pcXL_RequestKey")
		xmlRequestType_value=rs("pcXL_RequestType")
		xmlBackup_value=rs("pcXL_BackupFile")
		xmlUndo_value=rs("pcXL_Undo")
		set rs=nothing
		call closedb()
	end if
	
	if (cint(xmlRequestType_value)<9) OR (cint(xmlRequestType_value)>12) then
		response.redirect "viewXMLPartnerLogs.asp?msg=2&idPartner=" & pidPartner
	end if
	
	if xmlUndo_value="1" then
		response.redirect "viewXMLPartnerLogs.asp?msg=3&idPartner=" & pidPartner
	end if
	
	if IsNull(xmlBackup_value) or xmlBackup_value="" then
		response.redirect "viewXMLPartnerLogs.asp?msg=4&idPartner=" & pidPartner
	end if
End Sub

Sub RunUndoRequest()
	dim rs, query, TempLine, tmpID, TempStr
	on error resume next	

	Set fso = server.CreateObject("Scripting.FileSystemObject")
	findit = Server.MapPath("../xml/logs/" & xmlBackup_value)

	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	if Err.number>0 then
		Set f=nothing
		Set fso=nothing
		response.redirect "viewXMLPartnerLogs.asp?msg=4&idPartner=" & pidPartner
	end if

	DO WHILE not f.AtEndofStream
		TempLine=f.Readline
	
		IF TempLine<>"" then
			TempStr=split(TempLine,chr(9))
			tmpID=TempStr(1)
			
			Select Case Ucase(TempStr(0))
			Case "DELPRD":
				call opendb()
				query="DELETE FROM Products WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM DProducts WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM pcGC WHERE pcGC_IDProduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM pcProductsOptions WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM options_optionsGroups WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM categories_products WHERE idproduct=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELCUST":
				call opendb()
				query="DELETE FROM Customers WHERE idcustomer=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="DELETE FROM pcCustomerFieldsValues WHERE idcustomer=" & tmpID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELCATPRD":
				call opendb()
				query="DELETE FROM categories_products WHERE idproduct=" & tmpID & " AND idcategory=" & TempStr(2) & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELPRDGRP":
				call opendb()
				query="DELETE FROM pcProductsOptions WHERE idproduct=" & tmpID & " AND idOptionGroup=" & TempStr(2) & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "DELPRDOPT":
				call opendb()
				query="DELETE FROM options_optionsGroups WHERE idproduct=" & tmpID & " AND idOptionGroup=" & TempStr(2) & " AND idOption=" & TempStr(3) & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
			Case "UPDPRD":
				Call XMLUpdRecord("Products","idproduct",tmpID,TempLine)
			Case "UPDCUST":
				Call XMLUpdRecord("Customers","idcustomer",tmpID,TempLine)
			Case "UPDPRDGRP":
				Call XMLUpdRecord("pcProductsOptions","pcProdOpt_ID",tmpID,TempLine)
			Case "UPDPRDOPT":
				Call XMLUpdRecord("options_optionsGroups","idoptoptgrp",tmpID,TempLine)
			Case "ADDDP":
				Call XMLAddRecord("DProducts","idproduct",tmpIDvalue,TempLine)
			Case "ADDGC":
				Call XMLAddRecord("pcGC","pcGC_IDProduct",tmpIDvalue,TempLine)
			End Select
		END IF 'TempLine<>""
	
	LOOP
	
	f.close
	Set fso=nothing
	
	call opendb()
	query="UPDATE pcXMLLogs SET pcXL_Undo=1 WHERE pcXL_id=" & xmlRequestID_value & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closedb()
	
	response.redirect "viewXMLPartnerLogs.asp?msg=5&idPartner=" & pidPartner

End Sub

Sub XMLAddRecord(tmpTable,tmpIDName,tmpIDvalue,tmpValueStr)
Dim query,rstemp,rs,query2,PreRecord1,PreRecord2,k

	call opendb()
	query2=split(tmpValueStr,chr(9))
	
	query="SELECT * FROM " & tmpTable & ";"
	set rstemp=conntemp.execute(query)
	
	IF not rstemp.eof THEN
	
		query="DELETE FROM " & tmpTable & " WHERE " & tmpIDName & "=" & tmpIDvalue & ";"
		set rs=conntemp.execute(query)
		set rs=nothing
			
		PreRecord1=""
		PreRecord2=""
	
		iCols = rstemp.Fields.Count
	    for k=1 to iCols-1
	    	if query2(k)<>"##" then
	    	if k=1 then
	    		PreRecord1=PreRecord1 & "(" & Rstemp.Fields.Item(k).Name 
	    		PreRecord2=PreRecord2 & "(" & query2(k)
	    	else
	    		PreRecord1=PreRecord1 & "," & Rstemp.Fields.Item(k).Name
	    		PreRecord2=PreRecord2 & "," & query2(k)
	    	end if
	    	end if
	    next
	
	    PreRecord1=PreRecord1 & ")"
	    PreRecord2=PreRecord2 & ")"
	    
		query="INSERT INTO " & tmpTable & " " & PreRecord1 & " VALUES " & PreRecord2 & ";"
		query=replace(query,"DuLTVDu",vbcrlf)
		query=replace(query,"##","' '")
		set rstemp=connTemp.execute(query)
	END IF
	set rstemp=nothing
	call closedb()
	
End Sub

Sub XMLUpdRecord(tmpTable,tmpIDName,tmpIDvalue,tmpValueStr)
Dim query,rstemp,query2,PreRecord1,k

	call opendb()
	query2=split(tmpValueStr,chr(9))
	
	query="SELECT * FROM " & tmpTable & " WHERE " & tmpIDName & "=" & tmpIDvalue & ";"
	set rstemp=conntemp.execute(query)
			
	IF not rstemp.eof THEN
		PreRecord1=""
		iCols = rstemp.Fields.Count
	    for k=1 to iCols-1
	    	if query2(k+1)<>"##" then
	    	if k=1 then
	    		PreRecord1=PreRecord1 & Rstemp.Fields.Item(k).Name & "=" & query2(k+1)
	    	else
	    		PreRecord1=PreRecord1 & "," & Rstemp.Fields.Item(k).Name & "=" & query2(k+1)
	    	end if
	    	end if
	    next
		query="UPDATE " & tmpTable & " SET " & PreRecord1 & " WHERE " & tmpIDName & "=" & tmpIDvalue & ";"
		query=replace(query,"DuLTVDu",vbcrlf)
		set rstemp=connTemp.execute(query)
	END IF
	set rstemp=nothing
	call closedb()
	
End Sub
%>