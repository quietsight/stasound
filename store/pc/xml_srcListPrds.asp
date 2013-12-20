<%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<%

'*****************************************************
'* BEGIN: Check HTTP Referer
'*****************************************************

strPath=Request.ServerVariables("PATH_INFO")
dim iCnt, strPath,strPathInfo
iCnt=0
do while iCnt<1
	if mid(strPath,len(strPath),1)="/" then
		iCnt=iCnt+1
	end if
	if iCnt<1 then
		strPath=mid(strPath,1,len(strPath)-1)
	end if
loop
if Ucase(Request.ServerVariables("HTTPS"))="OFF" then
	strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
else
	strPathInfo="https://" & Request.ServerVariables("HTTP_HOST") & strPath
end if
				
if Right(strPathInfo,1)="/" then
else
	strPathInfo=strPathInfo & "/"
end if

strRefferer=Request.ServerVariables("HTTP_REFERER")

'*****************************************************
'* END: Check HTTP Referer
'*****************************************************

'*****************************************************
'* BEGIN: Check Query
'*****************************************************

Dim connTemp
call opendb()

pidcategory=getUserInput(request("idcategory"),10)

tmpTest=0

if not validNum(pidcategory) then
	tmpTest=1
end if

'*****************************************************
'* END: Check Query
'*****************************************************
			
if ((session("store_useAjax")<>"") AND (Instr(ucase(strRefferer),ucase(strPathInfo))=0)) OR (tmpTest=1) then%>
<bcontent>nothing</bcontent>
<%response.end
end if%><!--#include file="pcStartSession.asp" --><!--#include file="../includes/settings.asp"--><!--#include file="../includes/storeconstants.asp"--><!--#include file="../includes/stringfunctions.asp"--><!--#include file="../includes/opendb.asp"--><!--#include file="../includes/adovbs.inc"--><!--#include file="inc_srcPrdQuery.asp"--><%totalrecords=0
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

Dim tmpCount
tmpCount=0
Dim tmpList
tmpList=""
if not rs.eof then
	tmpList="<table border=0 width=100% class=mainbox><tr><td align=left><ul>"
	Do While (not rs.eof) and (tmpCount<10)
		tmpcount=tmpcount+1
		tmpList=tmpList & "<li>" & rs("description") & "</li>"
		rs.MoveNext
	Loop
	if tmpCount>=10 and (not rs.eof) then
		tmpList=tmpList & "<li>.......</li>"
	end if
	tmpList=tmpList & "</ul></td></tr></table>"
end if
set rs=nothing
call closedb()%><bcontent><%=Server.HTMLEncode(tmpList)%></bcontent>