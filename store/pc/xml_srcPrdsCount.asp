<!--#include file="../includes/SearchConstants.asp"-->
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

if pidcategory<>"0" AND tmpTest=0 then
	if (schideCategory = "1") OR (SRCH_SUBS = "1") then		
		Dim TmpCatList
		TmpCatList=""
		call pcs_GetSubCats(pIdCategory) '// get sub cats
		TmpCatList = pidcategory&TmpCatList
	end if
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

if not rs.eof then
	totalrecords=clng(rs.RecordCount)
end if
set rs=nothing
call closedb()%><count><%=totalrecords%></count>