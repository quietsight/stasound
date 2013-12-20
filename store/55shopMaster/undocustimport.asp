<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<%
on error resume next

dim rstemp, conntemp, query

call opendb()

msg=""
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	
	findit = Server.MapPath("importlogs/custlogs.txt")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	MyTest=1
	if Err.number>0 then
	msg="Can not open Customers Logs File.<br>"
	Err.number=0
	MyTest=0
	Err.Description=0
	end if
	IF MyTest=1 THEN
	ImportType=f.Readline
	PreIDCustomer=f.Readline
	
	
	if IsNumeric(PreIDCustomer) then
	query="delete from Customers where idcustomer>" & PreIDCustomer
	set rstemp=connTemp.execute(query)
	end if
	
	IF ucase(ImportType)="UPDATE" THEN
	
	Do While not f.AtEndofStream
	TempLine=f.Readline
	
	IF TempLine<>"" then
	TempStr=split(TempLine,"****")
	IdCustomer=TempStr(0)
	
	query2=split(TempStr(1),"@@@@@")
	
			query="select * from Customers where idCustomer=" & IdCustomer
			set rstemp=conntemp.execute(query)
			
			IF not rstemp.eof THEN

			PreRecord1=""
						
			iCols = rstemp.Fields.Count
		    for dd=1 to iCols-1
		    	if trim(query2(dd-1))<>"" then
			    	if UCase(query2(dd-1))="TRUE" then
			    		query2(dd-1)="-1"
			    	else
			    		if UCase(query2(dd-1))="FALSE" then
			    			query2(dd-1)="0"
			    		end if
			    	end if
		    	end if
		    	if dd=1 then
		    		PreRecord1=PreRecord1 & Rstemp.Fields.Item(dd).Name & "=" & query2(dd-1)
		    	else
		    		PreRecord1=PreRecord1 & "," & Rstemp.Fields.Item(dd).Name & "=" & query2(dd-1)
		    	end if
		    next
			query="update Customers Set " & PreRecord1 & " where idCustomer=" & IdCustomer
			'response.write query
			set rstemp=connTemp.execute(query)
			
			call updCustEditedDate(IdCustomer)
			
			END IF
	END IF 'TempLine<>""
	
	loop
	
	END IF 'It is UPDATE
	
	f.close
	END IF 'MyTest=1
	
	

	Set f = fso.GetFile(Server.MapPath("importlogs/custlogs.txt"))
	f.Delete
	Set fso = nothing
	
set rstemp = nothing
call closedb()

if msg<>"" then
response.redirect "custindex_import_help.asp?r=0&msg="&msg
else
response.redirect "custindex_import_help.asp?s=1&msg=The last customer import was undone successfully!"
end if
%>