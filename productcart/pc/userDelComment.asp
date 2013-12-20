<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!-- #Include File="checkdate.asp" -->

<%
on error resume next
Dim rstemp, connTemp, mySQL
call openDB()

Dim lngIDOrder,lngIDFeedback,lngIDComment

lngIDOrder=Clng(getUserInput(request("IDOrder"),0))-clng(scpre)
lngIDFeedback=getUserInput(request("IDFeedback"),0)
lngIDComment=getUserInput(request("IDComment"),0)

mySQL="select * from pcComments where pcComm_IDFeedback=" & lngIDComment & " and pcComm_IDParent=" & lngIDFeedback & " and pcComm_IDOrder=" & lngIDOrder & " and pcComm_IDUser=" & session("IDCustomer")

set rstemp=connTemp.execute(mySQL)
 
if rstemp.eof then
	call closedb()
	response.redirect "userviewfeedback.asp?IDOrder=" & scpre+clng(lngIDOrder) & "&IDFeedback=" & lngIDFeedback & "&r=1&msg=This comment was not found or you don't have permission to delete it."
end if

mySQL="delete from pcComments where pcComm_IDFeedback=" & lngIDComment & "and pcComm_IDParent=" & lngIDFeedback & " and pcComm_IDOrder=" & lngIDOrder 
set rstemp=connTemp.execute(mySQL)

Dim strFilename,fso,f,strQfilePath,strfindit

'Delete uploaded files
MySQL="Select * from pcUploadFiles where pcUpld_IDFeedback=" & lngIDComment
set rstemp=connTemp.execute(mySQL)

do while not rstemp.eof

	strFilename=rstemp("pcUpld_Filename")

	if strFilename<>"" then
		strQfilePath="Library/" & strFilename
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

MySQL="delete from pcUploadFiles where pcUpld_IDFeedback=" & lngIDComment
set rstemp=connTemp.execute(mySQL)

call closedb()

response.redirect "userviewfeedback.asp?IDOrder=" & scpre+clng(lngIDOrder) & "&IDFeedback=" & lngIDFeedback & "&msg=1"
%>