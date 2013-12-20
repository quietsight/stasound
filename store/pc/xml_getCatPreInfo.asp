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

pIDCategory=getUserInput(request("idcategory"),0)

tmpTest=0

if pIDCategory="" then
	tmpTest=1
else
	if pIDCategory="0" then
		tmpTest=1
	else
		if not IsNumeric(pIDCategory) then
			tmpTest=1
		end if
	end if
end if

'*****************************************************
'* END: Check Query
'*****************************************************
			
if ((session("store_useAjax")<>"") AND (Instr(ucase(strRefferer),ucase(strPathInfo))=0)) OR (tmpTest=1) then%>
<bcontent>nothing</bcontent>
<%response.end
end if%><!--#include file="pcStartSession.asp" --><!--#include file="../includes/settings.asp"--><!--#include file="../includes/storeconstants.asp"--><!--#include file="../includes/opendb.asp"--><!--#include file="../includes/languages.asp"--><!--#include file="../includes/currencyformatinc.asp"--><!--#include file="../includes/shipFromSettings.asp"--><!--#include file="../includes/taxsettings.asp"--><!--#include file="../includes/languages_ship.asp"--><!--#include file="../includes/adovbs.inc"--><!--#include file="../includes/stringfunctions.asp"--><% 
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN  

Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

dim query, conntemp, rs, pIDCategory

pIDCategory=getUserInput(request("idcategory"),0)

'--> open database connection
call opendb()

Dim tmpCatArr,tmpCatCount,TmpCatList

tmpCatCount=-1
TmpCatList=""

query="SELECT idcategory,idParentCategory FROM categories WHERE iBTOhide=0"
if session("CustomerType")="0" OR session("CustomerType")="" then
	query=query & " AND pccats_RetailHide=0"
end if
set rs=conntemp.execute(query)

if not rs.eof then
	tmpCatArr=rs.getRows()
	tmpCatCount=ubound(tmpCatArr,2)
end if
set rs=nothing


Sub GetSubCatList(tmpIDParent)
	Dim i
	if tmpIDParent<>"1" then
	For i=0 to tmpCatCount
		if clng(tmpCatArr(1,i))=clng(tmpIDParent) then
			if TmpCatList<>"" then
				TmpCatList=TmpCatList & ","
			end if
			TmpCatList=TmpCatList & tmpCatArr(0,i)
			call GetSubCatList(tmpCatArr(0,i))
		end if
	Next
	end if
End Sub


' --> gets product details from db
query="SELECT categoryDesc FROM categories WHERE idcategory=" &pIDCategory& " AND iBTOhide=0"
if session("CustomerType")="0" OR session("CustomerType")="" then
	query=query & " AND pccats_RetailHide=0"
end if
set rs=conntemp.execute(query)

catDescription=""

if not rs.eof then
	catDescription=rs("categoryDesc")
end if
set rs=nothing

catSubCat5=""

query="SELECT TOP 5 categoryDesc FROM categories WHERE idParentCategory=" &pIDCategory& " AND iBTOhide=0"
if session("CustomerType")="0" OR session("CustomerType")="" then
	query=query & " AND pccats_RetailHide=0"
end if
query=query & " ORDER BY priority, categoryDesc ASC;"
set rs=conntemp.execute(query)
if not rs.eof then
	tmpArr=rs.getRows()
	i=0
	do while i<5 AND i<=ubound(tmpArr,2)
		catSubCat5=catSubCat5 & "<li>" & tmpArr(0,i) & "</li>"
		i=i+1
	loop
	if i=5 then catSubCat5=catSubCat5 & "<li>...</li>"
end if
set rs=nothing

if catSubCat5<>"" then
	catSubCat5="<ul>" & catSubCat5 & "</ul>"
	TmpCatList=""
	call GetSubCatList(pIDCategory)
	if TmpCatList<>"" then
		query="SELECT DISTINCT TOP 1 idproduct FROM categories_products WHERE idcategory IN (" & TmpCatList & ");"
		set rs=connTemp.execute(query)
		if rs.eof then
			catSubCat5=catSubCat5 & "<br>" & dictLanguage.Item(Session("language")&"_xmlcat_2")
		end if
		set rs=nothing
	end if		
	catSubCat5="<div class='mainSubCatBox'><span class='mainSubCatBoxTitle'>" & dictLanguage.Item(Session("language")&"_xmlcat_1") & "</span><br>" & catSubCat5 & "</div>"
end if

catPrd5=""

query="SELECT TOP 5 products.description FROM products INNER JOIN categories_products ON products.idproduct=categories_products.idproduct WHERE categories_products.idCategory=" &pIDCategory& " AND products.active<>0 AND products.removed=0 ORDER BY categories_products.POrder, products.description ASC;"
set rs=conntemp.execute(query)
if not rs.eof then
	tmpArr=rs.getRows()
	i=0
	do while i<5 AND i<=ubound(tmpArr,2)
		catPrd5=catPrd5 & "<li>" & tmpArr(0,i) & "</li>"
		i=i+1
	loop
	if i=5 then catPrd5=catPrd5 & "<li>...</li>"
end if
set rs=nothing

if catPrd5<>"" then
	catPrd5="<div class='mainPrdBox'><span class='mainPrdBoxTitle'>" & dictLanguage.Item(Session("language")&"_xmlcat_3") & "</span><br><ul>" & catPrd5 & "</ul></div>"
end if

if catDescription<>"" then
	if catSubCat5<>"" AND catPrd5<>"" then
		catPrd5="<br>" & catPrd5
	end if
	if catSubCat5="" AND catPrd5="" then
		catPrd5=dictLanguage.Item(Session("language")&"_xmlcat_2")
	end if
	tmpList=catDescription & "|||" & "<table class=""mainbox""><td>" & catSubCat5 & catPrd5 & "</td></tr></table>"
else
	tmpList="nothing"
end if

set rs=nothing
set rstemp=nothing
Set RSlayout = nothing
Set rsIconObj = nothing
Set conlayout=nothing
call closedb()%><bcontent><%=Server.HTMLEncode(tmpList)%></bcontent>