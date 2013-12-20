<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->   
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
	
	findit = Server.MapPath("importlogs/catlogs.txt")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	MyTest=1
	if Err.number>0 then
	msg="Cannot open Categories Logs File.<br>"
	Err.number=0
	MyTest=0
	Err.Description=0
	end if
	IF MyTest=1 THEN
	Do While not f.AtEndofStream
	
	TempLine=f.Readline
	
	IF TempLine<>"" then
	
	TempStr=split(TempLine,"****")
	MyAction=TempStr(0)
	
	if MyAction="Del" then
	query2=split(TempStr(1),"@@@@@")
	query="delete from Categories_Products where idproduct=" & query2(0) & " and idcategory=" & query2(1)
	set rstemp=connTemp.execute(query)
	end if
	
	if MyAction="Add" then
	query2=split(TempStr(1),"@@@@@")
	query="insert into Categories_Products (IDProduct,IDCategory) values (" & query2(0) & "," & query2(1) & ")"
	set rstemp=connTemp.execute(query)
	end if
	
	END IF
	
	loop
	f.close
	END IF

	findit = Server.MapPath("importlogs/prologs.txt")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	MyTest=1
	if Err.number>0 then
	msg="Can not open Products Logs File.<br>"
	Err.number=0
	MyTest=0
	Err.Description=0
	end if
	IF MyTest=1 THEN
	ImportType=f.Readline
	PreIDProduct=f.Readline
	PreIDCat=f.Readline
	PreIDBrand=f.Readline
	
	if IsNumeric(PreIDCat) then
	query="delete from Categories_Products where idcategory>" & PreIDCat
	set rstemp=connTemp.execute(query)
	query="select * from Categories where idcategory>" & PreIDCat & " order by idcategory desc"
	set rstemp=connTemp.execute(query)
	do while not rstemp.eof
	query="delete from Categories where idcategory=" & rstemp("idcategory")
	set rstemp1=connTemp.execute(query)
	rstemp.MoveNext
	loop
	end if
	
	if IsNumeric(PreIDProduct) then
	query="delete from DProducts where idproduct>" & PreIDProduct
	set rstemp=connTemp.execute(query)
	query="delete from Products where idproduct>" & PreIDProduct
	set rstemp=connTemp.execute(query)
	query="delete from options_optionsGroups where idproduct>" & PreIDProduct
	set rstemp=connTemp.execute(query)
	end if
	
	if IsNumeric(PreIDBrand) then
	query="delete from Brands where idBrand>" & PreIDBrand
	set rstemp=connTemp.execute(query)
	end if	
	
	IF ucase(ImportType)="UPDATE" THEN
	Do While not f.AtEndofStream
	TempLine=f.Readline
	IF TempLine<>"" then
	TempStr=split(TempLine,"****")
	IdProduct=TempStr(0)
	MyAction=TempStr(1)
	if MyAction="DelDownPro" then
	query="delete from DProducts where idproduct=" & IdProduct
	set rstemp=connTemp.execute(query)
	end if
	if MyAction="DownPro" then
	query2=split(TempStr(2),"@@@@@")
	
			query="select * from Dproducts where idproduct=" & IdProduct
			set rstemp=conntemp.execute(query)
			
			IF not rstemp.eof THEN

			PreRecord1=""
						
			iCols = rstemp.Fields.Count
		    for dd=1 to iCols-1
		    	if dd=1 then
		    		PreRecord1=PreRecord1 & Rstemp.Fields.Item(dd).Name & "=" & query2(dd-1)
		    	else
		    		PreRecord1=PreRecord1 & "," & Rstemp.Fields.Item(dd).Name & "=" & query2(dd-1)
		    	end if
		    next
			query="update DProducts Set " & PreRecord1 & " where idproduct=" & IdProduct
			set rstemp=connTemp.execute(query)	    
			END IF
	end if
	
	if MyAction="Pro" then
	query2=split(TempStr(2),"@@@@@")
	
			query="select * from Products where idproduct=" & IdProduct
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
			query="update Products Set " & PreRecord1 & " where idproduct=" & IdProduct
			set rstemp=connTemp.execute(query)
			
			call updPrdEditedDate(IdProduct)
			
			END IF
	end if
	END IF
	loop
	END IF
	f.close
	END IF
	
	

	Set f = fso.GetFile(Server.MapPath("importlogs/catlogs.txt"))
	f.Delete
	Set f1 = fso.GetFile(Server.MapPath("importlogs/prologs.txt"))
	f1.Delete
	Set fso = nothing
	
set rstemp = nothing
call closedb()

if msg<>"" then
response.redirect "index_import_help.asp?r=0&msg="&msg
else
response.redirect "index_import_help.asp?s=1&msg=The last import was undone successfully!"
end if
%>