<% pageTitle = "Reverse Category Import Wizard - Export Results" %>
<% section = "products"
Server.ScriptTimeout = 5400%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, rstemp, rs, connTemp
Dim pcArr,i,tmp_query,intCount,pcv_HaveRecords
Dim fs,A,strFile
Dim fseparator
Dim valueArr

pcv_intExportSize = session("cp_ExportSize")
fseparator=session("cp_revCatImport_cseparator")
if session("cp_revCatImport_catlist")="" then
	response.redirect "ReverseCatImport_step1.asp"
end if

call opendb()

tmp_HeadLine=""
tmp_DataLine=""
pcv_HaveRecords=0

Function RmvHTMLWhiteSpace(tmpValue)
Dim tmp1,re,colMatch,objMatch
	tmp1=tmpValue
	Set re = New RegExp

	With re
	  .Pattern = "(\r\n[\s]+)"
	  .Global = True
	End With 

	Set colMatch = re.Execute(tmp1)
	For each objMatch in colMatch
		tmp1=replace(tmp1,objMatch.Value," ")
	Next
	RmvHTMLWhiteSpace=tmp1
End Function

Function GenFileName()
	dim fname
	fname="File-"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime)) & "-"
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function


'***** Generate HeadLine *****
if request("C1")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category ID""" & fseparator
end if
if request("C2")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category Name""" & fseparator
end if
if request("C3")="1" then
	tmp_HeadLine=tmp_HeadLine & """Small Image""" & fseparator
end if
if request("C4")="1" then
	tmp_HeadLine=tmp_HeadLine & """Large Image""" & fseparator
end if
if request("C5")="1" then
	tmp_HeadLine=tmp_HeadLine & """Parent Category Name""" & fseparator
end if
if request("C6")="1" then
	tmp_HeadLine=tmp_HeadLine & """Parent Category ID""" & fseparator
end if
if request("C7")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category Short Description""" & fseparator
end if
if request("C8")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category Long Description""" & fseparator
end if
if request("C9")="1" then
	tmp_HeadLine=tmp_HeadLine & """Hide Category Description""" & fseparator
end if
if request("C10")="1" then
	tmp_HeadLine=tmp_HeadLine & """Display Sub-Categories""" & fseparator
end if
if request("C11")="1" then
	tmp_HeadLine=tmp_HeadLine & """Sub-Categories per Row""" & fseparator
end if
if request("C12")="1" then
	tmp_HeadLine=tmp_HeadLine & """Sub-Category Rows per Page""" & fseparator
end if
if request("C13")="1" then
	tmp_HeadLine=tmp_HeadLine & """Display Products""" & fseparator
end if
if request("C14")="1" then
	tmp_HeadLine=tmp_HeadLine & """Products per Row""" & fseparator
end if
if request("C15")="1" then
	tmp_HeadLine=tmp_HeadLine & """Product Rows per Page""" & fseparator
end if
if request("C16")="1" then
	tmp_HeadLine=tmp_HeadLine & """Hide category""" & fseparator
end if
if request("C17")="1" then
	tmp_HeadLine=tmp_HeadLine & """Hide category from retail customers""" & fseparator
end if
if request("C18")="1" then
	tmp_HeadLine=tmp_HeadLine & """Product Details Page Display Option""" & fseparator
end if
if request("C19")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category Meta Tags - Title""" & fseparator
end if
if request("C20")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category Meta Tags - Description""" & fseparator
end if
if request("C21")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category Meta Tags - Keywords""" & fseparator
end if
if request("C22")="1" then
	tmp_HeadLine=tmp_HeadLine & """Featured Sub-Category Name""" & fseparator
end if
if request("C23")="1" then
	tmp_HeadLine=tmp_HeadLine & """Featured Sub-Category ID""" & fseparator
end if
if request("C24")="1" then
	tmp_HeadLine=tmp_HeadLine & """Use Featured Sub-Category Image""" & fseparator
end if
if request("C25")="1" then
	tmp_HeadLine=tmp_HeadLine & """Category Order""" & fseparator
end if
'***** End of Generate HeadLine *****

'***** Create SQL Query *****
query="SELECT idCategory,categoryDesc,idParentCategory,priority,image,largeimage,iBTOhide,HideDesc,pccats_RetailHide,pcCats_SubCategoryView,pcCats_CategoryColumns,pcCats_CategoryRows,pcCats_PageStyle,pcCats_ProductColumns,pcCats_ProductRows,pcCats_FeaturedCategory,pcCats_FeaturedCategoryImage,pcCats_DisplayLayout,pcCats_MetaTitle,pcCats_MetaDesc,pcCats_MetaKeywords,SDesc,LDesc FROM categories WHERE idParentCategory>-1 "
'Last: 46

if session("cp_revCatImport_catlist")<>"ALL" then
	pcArr=split(session("cp_revCatImport_catlist"),",")
	tmp_query=""
	For i=lbound(pcArr) to ubound(pcArr)
		if trim(pcArr(i))<>"" then
			if tmp_query<>"" then
				tmp_query=tmp_query & ","
			end if
			tmp_query=tmp_query & trim(pcArr(i))
		end if
	Next
	if tmp_query<>"" then
		tmp_query=" AND (idcategory IN (" & tmp_query & "))"
	end if
	query=query & tmp_query
end if
'***** End of Create SQL Query *****

if session("cp_revCatImport_pagecurrent")<>"" then
	Set rs=Server.CreateObject("ADODB.Recordset")

	iPageSize=pcv_intExportSize
	
	rs.CacheSize=iPageSize
	rs.PageSize=iPageSize
	
	rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

	rs.AbsolutePage=session("cp_revCatImport_pagecurrent")

else
	Set rs=Server.CreateObject("ADODB.Recordset")

	iPageSize=pcv_intExportSize
	
	rs.CacheSize=iPageSize
	rs.PageSize=iPageSize
	
	rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

	rs.AbsolutePage=1
	
end if

maxRecords=100000

if not rs.eof then
	if session("cp_revCatImport_pagecurrent")<>"" then
		pcArr=rs.GetRows(iPageSize)
	else
		pcArr=rs.GetRows(maxRecords)
	end if
	intCount=ubound(pcArr,2)
	pcv_HaveRecords=1
end if
set rs=nothing

IF pcv_HaveRecords=1 THEN
	
	For i=0 to intCount

'***** Generate Data Lines *****

if request("C1")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(0,i) & """" & fseparator
end if
if request("C2")="1" then
	tmp_CatName=pcArr(1,i)
	if tmp_CatName<>"" then
		tmp_CatName=replace(tmp_CatName,"""","""""")
	end if
	tmp_DataLine=tmp_DataLine & """" & tmp_CatName & """" & fseparator
end if
if request("C3")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(4,i) & """" & fseparator
end if
if request("C4")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(5,i) & """" & fseparator
end if
if request("C5")="1" then
	query="SELECT categoryDesc FROM categories WHERE idcategory=" & pcArr(2,i) & ";"
	set rsQ=connTemp.execute(query)
	tmp_DataLine=tmp_DataLine & """" & replace(rsQ("categoryDesc"),"""","""""") & """" & fseparator
	set rsQ=nothing
end if
if request("C6")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(2,i) & """" & fseparator
end if
if request("C7")="1" then
	tmp_CatSDesc=pcArr(21,i)
	if tmp_CatSDesc<>"" then
		if fseparator="," then
			tmp_CatSDesc=RmvHTMLWhiteSpace(replace(tmp_CatSDesc,"""",""""""))
			tmp_CatSDesc=replace(tmp_CatSDesc,"&quot;","""""")
		else
			tmp_CatSDesc=RmvHTMLWhiteSpace(replace(tmp_CatSDesc,"&quot;",""""))
		end if
		tmp_CatSDesc=replace(tmp_CatSDesc,vbCrLf,"")
		tmp_CatSDesc=replace(tmp_CatSDesc,vbCr,"")
		tmp_CatSDesc=replace(tmp_CatSDesc,vbLf,"")
	end if
	tmp_DataLine=tmp_DataLine & """" & tmp_CatSDesc & """" & fseparator
end if
if request("C8")="1" then
	tmp_CatLDesc=pcArr(22,i)
	if tmp_CatLDesc<>"" then
		if fseparator="," then
			tmp_CatLDesc=RmvHTMLWhiteSpace(replace(tmp_CatLDesc,"""",""""""))
			tmp_CatLDesc=replace(tmp_CatLDesc,"&quot;","""""")
		else
			tmp_CatLDesc=RmvHTMLWhiteSpace(replace(tmp_CatLDesc,"&quot;",""""))
		end if
		tmp_CatLDesc=replace(tmp_CatLDesc,vbCrLf,"")
		tmp_CatLDesc=replace(tmp_CatLDesc,vbCr,"")
		tmp_CatLDesc=replace(tmp_CatLDesc,vbLf,"")
	end if
	tmp_DataLine=tmp_DataLine & """" & tmp_CatLDesc & """" & fseparator
end if
if request("C9")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(7,i) & """" & fseparator
end if
if request("C10")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(9,i) & """" & fseparator
end if
if request("C11")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(10,i) & """" & fseparator
end if
if request("C12")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(11,i) & """" & fseparator
end if
if request("C13")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(12,i) & """" & fseparator
end if
if request("C14")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(13,i) & """" & fseparator
end if
if request("C15")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(14,i) & """" & fseparator
end if
if request("C16")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(6,i) & """" & fseparator
end if
if request("C17")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(8,i) & """" & fseparator
end if
if request("C18")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(17,i) & """" & fseparator
end if
if request("C19")="1" then
	tmp_MetaTitle=pcArr(18,i)
	if tmp_MetaTitle<>"" then
		tmp_MetaTitle=replace(tmp_MetaTitle,"""","""""")
	end if
	tmp_DataLine=tmp_DataLine & """" & tmp_MetaTitle & """" & fseparator
end if
if request("C20")="1" then
	tmp_MetaDesc=pcArr(19,i)
	if tmp_MetaDesc<>"" then
		if fseparator="," then
			tmp_MetaDesc=RmvHTMLWhiteSpace(replace(tmp_MetaDesc,"""",""""""))
			tmp_MetaDesc=replace(tmp_MetaDesc,"&quot;","""""")
		else
			tmp_MetaDesc=RmvHTMLWhiteSpace(replace(tmp_MetaDesc,"&quot;",""""))
		end if
		tmp_MetaDesc=replace(tmp_MetaDesc,vbCrLf,"")
		tmp_MetaDesc=replace(tmp_MetaDesc,vbCr,"")
		tmp_MetaDesc=replace(tmp_MetaDesc,vbLf,"")
	end if
	tmp_DataLine=tmp_DataLine & """" & tmp_MetaDesc & """" & fseparator
end if
if request("C21")="1" then
	tmp_MetaKey=pcArr(20,i)
	if tmp_MetaKey<>"" then
		if fseparator="," then
			tmp_MetaKey=RmvHTMLWhiteSpace(replace(tmp_MetaKey,"""",""""""))
			tmp_MetaKey=replace(tmp_MetaKey,"&quot;","""""")
		else
			tmp_MetaKey=RmvHTMLWhiteSpace(replace(tmp_MetaKey,"&quot;",""""))
		end if
		tmp_MetaKey=replace(tmp_MetaKey,vbCrLf,"")
		tmp_MetaKey=replace(tmp_MetaKey,vbCr,"")
		tmp_MetaKey=replace(tmp_MetaKey,vbLf,"")
	end if
	tmp_DataLine=tmp_DataLine & """" & tmp_MetaKey & """" & fseparator
end if
if request("C22")="1" then
	if pcArr(15,i)>"0" then
	query="SELECT categoryDesc FROM categories WHERE idcategory=" & pcArr(15,i) & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		tmpSubName=rsQ("categoryDesc")
	else
		tmpSubName=""
	end if
	set rsQ=nothing
	if tmpSubName<>"" then
		tmpSubName=replace(tmpSubName,"""","""""")
	end if
	else
		tmpSubName=""
	end if
	tmp_DataLine=tmp_DataLine & """" & tmpSubName & """" & fseparator
end if
if request("C23")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(15,i) & """" & fseparator
end if
if request("C24")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(16,i) & """" & fseparator
end if
if request("C25")="1" then
	tmp_DataLine=tmp_DataLine & """" & pcArr(3,i) & """" & fseparator
end if

'***** End of Generate Data Lines *****	

	tmp_DataLine=tmp_DataLine & VBCrLf
	
	Next
	
	strFile=GenFileName()
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	if session("cp_revCatImport_fseparator")="0" then
		tmpext=".csv"
	else
		tmpext=".txt"
	end if
	Set A=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & tmpext,True)
	A.Write(tmp_HeadLine & VBCrLf & tmp_DataLine)
	A.Close
	Set A=Nothing
	Set fs=Nothing	
	
END IF 'Have category records

call closedb()

%>
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcSpacer"></td>
</tr>
<%IF pcv_HaveRecords=0 THEN%>
<tr>
	<td colspan="2">
		<div class="pcCPmessage">
			No Categories found!
		</div>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<td colspan="2">
		<input type="button" name="back" value="Start Again" onclick="javasccript:location='ReverseCatImport_step1.asp';" class="ibtnGrey">
	</td>
</tr>
<%ELSE%>
<tr>
	<td colspan="2">
		<div class="pcCPmessageSuccess">
			Categories were exported successfully!
		</div>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td align="center">
		<p><b>Download your file.</b></p>
		<p style="padding-top:10px;"><a href="<%=strFile & tmpext%>"><img src="images/DownLoad.gif"></a></p>
		<p style="padding-top:10px;">To ensure that your file downloads correctly, right click on the icon above and choose &quot;<b>Save Target As...</b>&quot;. If the browser attempts to save the file with a *.htm extension, change the file name so that it uses the extension *.txt.</p>
		<p style="padding-top:10px;">Also note that to open a *.txt file in MS Excel you should first start Excel, and then open the file by selecting &quot;File > Open&quot;. This way your will use the <u>MS Excel Text Import Wizard</u>, where you can specify the custom separator used in the file, if any.
		</p>
	</td>
</tr>
<%END IF%>
<% 
session("cp_revCatImport_catlist")=""
session("cp_revCatImport_pagecurrent")=""
session("cp_revCatImport_cseparator")=""
session("cp_ExportSize")=""
%>
</table>
<!--#include file="AdminFooter.asp"-->