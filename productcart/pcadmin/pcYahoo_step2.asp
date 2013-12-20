<% pageTitle = "Export to Yahoo! Search Marketing" %>
<% section = "specials"
Server.ScriptTimeout = 5400%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, rstemp, rs, connTemp
Dim pcArr,i,tmp_query,intCount,pcv_HaveRecords
Dim fs,A,strFile,pcv_idcat

if session("cp_exportYahoo_prdlist")="" then
	response.redirect "pcYahoo_step1.asp"
end if

call opendb()

tmp_HeadLine=""
tmp_DataLine=""
pcv_HaveRecords=0

Function GenFileName()
	dim fname
	fname="data"
	GenFileName=fname
End Function

Function RmvInvalidChrs(tmpStr)
Dim i,strAlpha,tmp1,tmp2,j
	strAlpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz 0123456789~!@#$%^&*()_+|\{[}]:;""'-=,<.>/`?"
	tmp2=""
	if not IsNull(tmpStr) and tmpStr<>"" then
	j=len(tmpStr)
	For i=1 to j
   		tmp1=Mid(tmpStr,i,1)
   		if Instr(1,strAlpha,tmp1)>0 then
   			tmp2=tmp2 & tmp1
   		End if
   	Next
   	tmp2=replace(tmp2,"&amp;","&")
   	tmp2=replace(tmp2,"&gt;",">")
   	tmp2=replace(tmp2,"&lt;","<")
	tmp2=replace(tmp2,"&quot;","""")
	end if
 	RmvInvalidChrs=tmp2
End Function

Function getCATInfor(pcv_IDProduct)
Dim query,rs1,tmp_Cat,i,tmp1,tmp2,tmp3
	query="SELECT categories.idcategory,categories.pccats_BreadCrumbs FROM categories INNER JOIN categories_products ON categories.idCategory=categories_products.idCategory WHERE categories_products.idProduct=" & pcv_IDProduct & ";"
	set rs1=connTemp.execute(query)
	tmp_Cat=""
	if not rs1.eof then
		pcv_idcat=rs1("idcategory")
		tmp1=rs1("pccats_BreadCrumbs")
		if tmp1<>"" then
			tmp2=split(tmp1,"|,|")
			For i=lbound(tmp2) to ubound(tmp2)
				if tmp2(i)<>"" then
					tmp3=split(tmp2(i),"||")
					if tmp3(1)<>"" then
						if tmp_Cat<>"" then
							tmp_Cat=tmp_Cat & " > "
						end if
						tmp_Cat=tmp_Cat & tmp3(1)
					end if
				end if
			Next
		end if
	end if
	set rs1=nothing
	if tmp_Cat<>"" then
		tmp_Cat=RmvInvalidChrs(ClearHTMLTags2(tmp_Cat,0))
	end if
	getCATInfor=tmp_Cat
End Function

scSiteURL=scStoreURL
if Right(scSiteURL,1)<>"/" then
scSiteURL=scSiteURL & "/"
end if

scSiteURL=scSiteURL & scPcFolder & "/"

'***** Generate HeadLine *****
StrResults=""
StrResults=StrResults & "code" & chr(9) & "name" & chr(9) & "description" & chr(9) & "price" & chr(9) & "product-url" & chr(9) & "merchant-site-category" & chr(9) & "medium" & chr(9) & "image-url" & vbcrlf
'***** End of Generate HeadLine *****

'***** Create SQL Query *****
query="SELECT products.idproduct,products.sku,products.description,products.price,products.imageUrl,smallImageUrl,largeImageURL,products.details,products.sDesc FROM Products WHERE products.removed=0 AND active<>0 AND configonly=0 "

if session("cp_exportYahoo_prdlist")<>"ALL" then
	pcArr=split(session("cp_exportYahoo_prdlist"),",")
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
		tmp_query=" AND products.idproduct IN (" & tmp_query & ")"
	end if
	query=query & tmp_query
end if
'***** End of Create SQL Query *****

if session("cp_exportYahoo_pagecurrent")<>"" then
	Set rs=Server.CreateObject("ADODB.Recordset")

	iPageSize=200
	
	rs.CacheSize=iPageSize
	rs.PageSize=iPageSize
	
	rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

	rs.AbsolutePage=session("cp_exportYahoo_pagecurrent")

else
	set rs=connTemp.execute(query)
end if

if not rs.eof then
	if session("cp_exportYahoo_pagecurrent")<>"" then
		pcArr=rs.GetRows(iPageSize)
	else
		pcArr=rs.GetRows()
	end if
	intCount=ubound(pcArr,2)
	pcv_HaveRecords=1
end if
set rs=nothing

IF pcv_HaveRecords=1 THEN
	For i=0 to intCount
	
	'***** Generate Data Lines *****
	pcv_idcat=""
	pcv_sku=pcArr(1,i)
	pcv_name=RmvInvalidChrs(ClearHTMLTags2(pcArr(2,i),0))
	pcv_desc=pcArr(8,i)
	if trim(pcv_desc)="" then
		pcv_desc=pcArr(7,i)
	end if
	if trim(pcv_desc)="" or trim(pcv_desc)="no information" then
		pcv_desc=pcv_name
	end if
	pcv_desc=RmvInvalidChrs(ClearHTMLTags2(pcv_desc,0))
	pcv_price=Round(pcArr(3,i),2)
	pcv_image=pcArr(4,i)
	if trim(pcv_image)="" then
		pcv_image=pcArr(5,i)
	end if
	if trim(pcv_image)="" then
		pcv_image=pcArr(6,i)
	end if
	if trim(pcv_image)<>"" then
		pcv_image=scSiteURL & "pc/catalog/" & pcv_image
	end if
	pcv_cats=getCATInfor(pcArr(0,i))

	'// SEO Links
	'// Build Product Link
	if scSeoURLs<>1 then
		pcv_prdurl=scSiteURL & "pc/viewPrd.asp?idproduct=" & pcArr(0,i) & "&idcategory=" & pcv_idcat
	else
		pcStrPrdLink=pcv_name & "-" & pcv_idcat & "p" & pcArr(0,i) & ".htm"
		pcStrPrdLink=removeChars(pcStrPrdLink)
		pcv_prdurl=scSiteURL & "pc/" & pcStrPrdLink
	end if
	'//

	'***** End of Generate Data Lines *****	

	StrResults=StrResults & pcv_sku & chr(9) & pcv_name & chr(9) & pcv_desc & chr(9) & pcv_price & chr(9) & pcv_prdurl & chr(9) & pcv_cats & chr(9) & "" & chr(9) & pcv_image & vbcrlf
	
	Next
	
	strFile=GenFileName()
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set A=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".txt",True)
	A.Write(StrResults)
	A.Close
	Set A=Nothing
	Set fs=Nothing	
	
END IF 'Have product records

call closedb()

%>
<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<table border="0" cellpadding="0" cellspacing="0" width="60%">
		<tr>
			<td width="16%" align="center"><img border="0" src="images/step1.gif"></td>
			<td width="84%"><font color="#A8A8A8">Locate products</font></td>
		</tr>
		<tr>
			<td width="16%" align="center"><img border="0" src="images/step2a.gif"></td>
			<td width="84%"><b>Export results</b></td>
		</tr>
	</table>
	</td>
</tr>
<%IF pcv_HaveRecords=0 THEN%>
<tr>
	<td colspan="2">
		<div class="pcCPmessage">
			No Products found!
		</div>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input type="button" name="back" value="Back to Main Menu" onclick="javasccript:location='menu.asp';" class="ibtnGrey">
	</td>
</tr>
<%ELSE%>
<tr>
	<td colspan="2">
		<div class="pcCPmessageSuccess">
			Products were exported successfully!
		</div>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td align="center">
		<p><b>Download your file.</b></p>
		<p style="padding-top:10px;"><a href="<%=strFile & ".txt"%>"><img src="images/DownLoad.gif"></a></p>
		<p style="padding-top:10px;">To ensure that your file downloads correctly, right click on the icon above<br>
			and choose &quot;<b>Save Target As...</b>&quot; from your menu.
		</p>
	</td>
</tr>
<%END IF%>
<% session("cp_exportYahoo_prdlist")=""
session("cp_exportYahoo_pagecurrent")=""
%>
</table>
<!--#include file="AdminFooter.asp"-->