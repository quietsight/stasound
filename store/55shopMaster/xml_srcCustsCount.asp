<%PmAdmin=0%><!--#include file="adminv.asp"--><%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="inc_srcCustQuery.asp"-->
<%
totalrecords=0
Dim connTemp
call opendb()
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
if not rs.eof then
	totalrecords=clng(rs.RecordCount)
end if
set rs=nothing
call closedb()
%>
<count><%=totalrecords%></count>