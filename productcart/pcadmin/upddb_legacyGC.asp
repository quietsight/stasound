<%@LANGUAGE="VBSCRIPT"%>
<% 
On Error Resume Next
PmAdmin=19
pageTitle = "ProductCart v4.7 - Database Update" 
Section = "" 
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% 
dim conntemp, query, rs

call openDb()

query="SELECT pcOrd_GcCode,pcOrd_GcUsed,idOrder FROM ORders WHERE pcOrd_GcCode<>'';"
set rs=connTemp.execute(query)

if not rs.eof then
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
	For i=0 to intCount
	
	query="SELECT Products.Description FROM  Products INNER JOIN pcGCOrdered ON pcGCOrdered.pcGO_IDProduct=Products.idproduct WHERE Products.pcprod_GC=1 AND pcGCOrdered.pcGO_GcCode LIKE '" & pcArr(0,i) & "'"
	set rstemp=connTemp.execute(query)
	pName=""
	if not rstemp.eof then
		pName=rstemp("Description")
	end if
	set rstemp=nothing
	
	query="UPDATE Orders SET pcOrd_GCDetails='" & pcArr(0,i) & "|s|" & pName & "|s|" & pcArr(1,i) & "',pcOrd_GCAmount=" & pcArr(1,i) & ",pcOrd_GcCode='',pcOrd_GcUsed=0 WHERE pcOrd_GcCode LIKE '" & pcArr(0,i) & "' AND IdOrder=" & pcArr(2,i) & ";"
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
	
	Next
end if
set rs=nothing

response.Redirect "upddb_v47_complete.asp"
%>
<!--#include file="AdminFooter.asp"-->
