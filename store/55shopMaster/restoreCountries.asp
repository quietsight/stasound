<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
on error resume next
dim rstemp, conntemp, mysql

if request("action")="update" then
call openDb()
mySQL="delete from Countries"
set rstemp=connTemp.execute(mySQL)
set rstemp=nothing
call clseDb()

DataFile = "defaultcountries.txt"
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
	mysql="insert into Countries (CountryName,CountryCode) values ('" & TempL1 & "','" & TempL2 & "')"
	Set rstemp=connTemp.execute(mySQL)
	end if
	
		'// Check if this country is still a DropDown
		query="select pcCountryCode from States where pcCountryCode='" & TempL2 & "' "
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if rs.eof then
			query="update countries set pcSubDivisionID=2 where countryCode='" & TempL2 & "' "
			set rs2=Server.CreateObject("ADODB.Recordset")
			set rs2=connTemp.execute(query)
			set rs2=nothing
		else
			query="update countries set pcSubDivisionID=1 where countryCode='" & TempL2 & "' "
			set rs2=Server.CreateObject("ADODB.Recordset")
			set rs2=connTemp.execute(query)
			set rs2=nothing
		end if
	
	Loop
	
	f.Close
	Set f = nothing
	Set fso = nothing
	
	response.redirect "manageCountries.asp"
	
end if	
%>