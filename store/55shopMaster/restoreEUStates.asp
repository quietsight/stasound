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

if request("action")="update" then

	dim rstemp, conntemp, mysql
	call openDb()

	mySQL="DELETE FROM pcVATCountries;"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(mySQL)
	set rstemp=nothing
	
		DataFile = "defaultEUstates.txt"
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
				mysql="INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('" & TempL2 & "','" & TempL1 & "')"
				Set rstemp=connTemp.execute(mySQL)
			end if
		
		Loop
		
		f.Close
		Set f = nothing
		Set fso = nothing
		
		call closedb()
		response.redirect "manageEUStates.asp"
		
end if	
%>