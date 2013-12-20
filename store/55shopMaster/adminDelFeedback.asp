<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->


<%
on error resume next
Dim rstemp, connTemp, mySQL
call openDB()

Dim lngIDOrder,lngIDFeedback

lngIDOrder=getUserInput(request("IDOrder"),0)
lngIDFeedback=getUserInput(request("IDFeedback"),0)

mySQL="select * from pcComments where pcComm_IDFeedback=" & lngIDFeedback & " and pcComm_IDParent=0 and pcComm_IDOrder=" & lngIDOrder 
set rstemp=connTemp.execute(mySQL)
 
if rstemp.eof then
	response.redirect "adminviewfeedback.asp?IDOrder=" & lngIDOrder & "&IDFeedback=" & lngIDFeedback & "&r=1&msg=This feedback was not found or you don't have permission to delete it."
end if

mySQL="delete from pcComments where pcComm_IDFeedback=" & lngIDFeedback & "and pcComm_IDParent=0 and pcComm_IDOrder=" & lngIDOrder
set rstemp=connTemp.execute(mySQL)

mySQL="delete from pcComments where pcComm_IDParent=" & lngIDFeedback & "and pcComm_IDOrder=" & lngIDOrder
set rstemp=connTemp.execute(mySQL)

'Delete uploaded files
MySQL="Select * from pcUploadFiles where pcUpld_IDFeedback=" & lngIDFeedback
set rstemp=connTemp.execute(mySQL)

do while not rstemp.eof
	
	Dim strFilename,strQfilePath,strfindit,fso,f
	
	strFilename=rstemp("pcUpld_Filename")
	
	if strFilename<>"" then
		strQfilePath="../pc/Library/" & strFilename
	   	strfindit = Server.MapPath(strQfilePath)
		Set fso = server.CreateObject("Scripting.FileSystemObject")
		Set f = fso.GetFile(strfindit)
		f.Delete
		Set fso = nothing
		Set f = nothing
		Err.number=0
		Err.Description=""
	end if
	rstemp.MoveNext
loop

MySQL="delete from pcUploadFiles where pcUpld_IDFeedback=" & lngIDFeedback
set rstemp=connTemp.execute(mySQL)
set rstemp=nothing
call closedb()

response.redirect "adminviewallmsgs.asp?IDOrder=" & lngIDOrder
%>