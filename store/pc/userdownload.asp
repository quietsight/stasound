<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->

<%
on error resume next
Dim rstemp, connTemp, mySQL
Dim lngIDFile,intintTest,lngIDFeedback,strFileName,TempStr
call openDB()

lngIDFile=getUserInput(request("IDFile"),0)

intintTest=0
	
MySQL="select * from pcUploadFiles where pcUpld_IDFile=" & lngIDFile
set rstemp=connTemp.execute(mySQL)

if rstemp.eof then
	intTest=1
else
	lngIDFeedback=rstemp("pcUpld_IDFeedback")
	strFileName=rstemp("pcUpld_FileName")
end if
	
if intTest=0 then
	MySQL="select * from pcComments where pcComm_IDFeedback=" & lngIDFeedback
	set rstemp=connTemp.execute(mySQL)
	if rstemp.eof then
		intTest=1
	else
		lngIDOrder=rstemp("pcComm_IDOrder")
	end if
end if
	
if intTest=0 then
	TempStr=" and IDCustomer=" & session("IDCustomer")
	MySQL="Select * from Orders where IDOrder=" & lngIDOrder & TempStr
	set rstemp=connTemp.execute(mySQL)

	if rstemp.eof then
		intTest=1
	end if
end if
 
if intTest=1 then
	call closedb()
	response.redirect "about:blank"
end if

MySQL="select * from pcUploadFiles where pcUpld_IDFile=" & lngIDFile
set rstemp=connTemp.execute(mySQL)

lngIDFeedback=rstemp("pcUpld_IDFeedback")
strFileName=rstemp("pcUpld_FileName")

Dim strFileName1,strFileN

'Downloadable file name

strFileN="Library/" & strFileName
strFileName1=mid(strFileName,instr(strFileName,"_")+1,len(strFileName))

call closedb()

response.redirect strFileN

%>