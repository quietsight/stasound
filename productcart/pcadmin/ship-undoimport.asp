<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=9%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
   
<%

dim rstemp, conntemp, query

call opendb()

msg=""
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	
	findit = Server.MapPath("importlogs/ship-prologs.txt")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	MyTest=1
	if Err.number>0 then
	msg="Can not open Data Logs File.<br>"
	Err.number=0
	MyTest=0
	Err.Description=0
	end if
	IF MyTest=1 THEN
		ImportType=f.Readline
	
		IF ucase(ImportType)="UPDATE" THEN
			Do While not f.AtEndofStream
				TempLine=f.Readline
				IF TempLine<>"" then
					TempStr=split(TempLine,"****")
					IdOrder=TempStr(0)
					MyAction=TempStr(1)

					if MyAction="Ord" then
					query2=split(TempStr(2),"@@@@@")
	
					query="select * from Orders where idorder=" & IdOrder
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
					query="update Orders Set " & PreRecord1 & " where IDOrder=" & IDOrder
					set rstemp=connTemp.execute(query)

					END IF
					end if 'MyAction="Pro"
				END	IF
			loop
		END IF
		f.close
	END IF

	Set f1 = fso.GetFile(Server.MapPath("importlogs/ship-prologs.txt"))
	f1.Delete
	Set fso = nothing
	
set rstemp = nothing
call closedb()

if msg<>"" then
	response.redirect "ship-index_import_help.asp?r=0&msg="&msg
else
	response.redirect "ship-index_import_help.asp?s=1&msg=The last import was undone successfully!"
end if
%>