<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<%
on error resume next
dim rstemp, conntemp, query

call openDb()

if request("action")="update" then

query="delete from States"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)

DataFile = "defaultstates.txt"
	findit = Server.MapPath(Datafile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(findit, 1)
	
	Do While not f.AtEndofStream
	
		DataLine=f.Readline
		if trim(DataLine)<>"" then
		DataLine=trim(DataLine)
		TempLine=split(DataLine,",")
		TempL1=replace(TempLine(0),"'","''")
		TempL2=replace(TempLine(1),"'","''")
		TempL3=replace(TempLine(2),"'","''")	
		query="INSERT INTO States (StateName,StateCode,pcCountryCode) values ('" & TempL1 & "','" & TempL2 & "','" & TempL3 & "')"
		Set rstemp=connTemp.execute(query)
		end if
		
		'// Update the Country Table Relationship for each state/ province the user imports.
		query="update countries set pcSubDivisionID=1 where countryCode='" & TempL3 & "' "
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		set rs=nothing
	
	Loop
	
	f.Close
	Set f = nothing
	Set fso = nothing
	call closeDb()
	response.redirect "manageStates.asp"
	
end if	
%>