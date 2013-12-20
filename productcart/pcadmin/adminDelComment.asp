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
Dim rstemp, connTemp, query
call openDB()
Dim lngIDOrder,lngIDFeedback,lngIDComment

lngIDOrder=getUserInput(request("IDOrder"),0)
lngIDFeedback=getUserInput(request("IDFeedback"),0)
lngIDComment=getUserInput(request("IDComment"),0)

query="select * from pcComments where pcComm_IDFeedback=" & lngIDComment & " and pcComm_IDParent=" & lngIDFeedback & " and pcComm_IDOrder=" & lngIDOrder
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
 
if rstemp.eof then
	set rstemp=nothing
	call closedb()
	response.redirect "adminviewfeedback.asp?IDOrder=" & lngIDOrder & "&IDFeedback=" & lngIDFeedback & "&r=1&msg=This comment was not found or you don't have permission to delete it."
end if

query="delete from pcComments where pcComm_IDFeedback=" & lngIDComment & "and pcComm_IDParent=" & lngIDFeedback & " and pcComm_IDOrder=" & lngIDOrder 
set rstemp=connTemp.execute(query)

'Delete uploaded files
query="Select * from pcUploadFiles where pcUpld_IDFeedback=" & lngIDComment
set rstemp=connTemp.execute(query)

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

query="delete from pcUploadFiles where pcUpld_IDFeedback=" & lngIDComment
set rstemp=connTemp.execute(query)

set rstemp=nothing
call closedb()
response.redirect "adminviewfeedback.asp?IDOrder=" & lngIDOrder & "&IDFeedback=" & lngIDFeedback & "&s=1&msg=1"
%>