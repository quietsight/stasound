<% pageTitle = "Export to NexTag Wizard" %>
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
<!--#include file="../includes/shipFromSettings.asp" -->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, rstemp, rs, connTemp
Dim pcArr,i,tmp_query,intCount,pcv_HaveRecords
Dim fs,A,strFile,pcv_idcat
pcv_idcat=""

if session("cp_exportNextTag_prdlist")="" then
	response.redirect "pcNextTag_step1.asp"
end if

call opendb()

tmp_HeadLine=""
tmp_DataLine=""
pcv_HaveRecords=0

Function GenFileName()
	dim fname
	fname="nexttagdata"
	GenFileName=fname
End Function

Function RmvDescInvalidChrs(tmpStr)
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
	tmp2=replace(tmp2,VBCrLf," ")
	tmp2=replace(tmp2,VBCr,"")
	tmp2=replace(tmp2,VBLf,"")
	tmp2=replace(tmp2,"&amp;","&")
	tmp2=replace(tmp2,"&gt;",">")
	tmp2=replace(tmp2,"&lt;","<")
	tmp2=replace(tmp2,"""","""""")
	tmp2=replace(tmp2,"&quot;","""""")
	end if
 	RmvDescInvalidChrs=tmp2
End Function

Function RmvNameInvalidChrs(tmpStr)
Dim i,strAlpha,tmp1,tmp2,j
	strAlpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz 0123456789#()_+\{[}]:;""'-=,<.>/`?"
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
   	tmp2=replace(tmp2,"""","""""")
	tmp2=replace(tmp2,"&quot;","""""")
	end if
 	RmvNameInvalidChrs=tmp2
End Function

Function getCATInfor(pcv_IDProduct)
Dim query,rs1,tmp_Cat,i,tmp1,tmp2,tmp3
	query="SELECT categories.idcategory,categories.categoryDesc FROM categories INNER JOIN categories_products ON categories.idCategory=categories_products.idCategory WHERE categories_products.idProduct=" & pcv_IDProduct & ";"
	set rs1=connTemp.execute(query)
	tmp_Cat=""
	if not rs1.eof then
		pcv_idcat=rs1("idcategory")
		tmp_Cat=rs1("categoryDesc")
	end if
	set rs1=nothing
	if tmp_Cat<>"" then
		tmp_Cat=RmvNameInvalidChrs(ClearHTMLTags2(tmp_Cat,0))
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
StrResults=StrResults & "Manufacturer,Manufacturer Part #,Product Name,Product Description,Click-Out URL,Price,Merchant Product Cat.,Image URL,Stock Status,Weight" & vbcrlf
'***** End of Generate HeadLine *****

'***** Create SQL Query *****
query="SELECT products.idproduct,products.sku,products.description,products.price,products.imageUrl,products.smallImageUrl,products.largeImageURL,products.idbrand,products.stock,products.nostock,products.pcProd_BackOrder,products.weight,products.details,products.sDesc FROM Products WHERE products.removed=0 AND products.active<>0 AND configonly=0 "

if session("cp_exportNextTag_prdlist")<>"ALL" then
	pcArr=split(session("cp_exportNextTag_prdlist"),",")
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

if session("cp_exportNextTag_pagecurrent")<>"" then
	Set rs=Server.CreateObject("ADODB.Recordset")

	iPageSize=200
	
	rs.CacheSize=iPageSize
	rs.PageSize=iPageSize
	
	rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

	rs.AbsolutePage=session("cp_exportNextTag_pagecurrent")

else
	set rs=connTemp.execute(query)
end if

if not rs.eof then
	if session("cp_exportNextTag_pagecurrent")<>"" then
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
	pcv_name=RmvNameInvalidChrs(ClearHTMLTags2(pcArr(2,i),0))
	pcv_desc=pcArr(13,i)
	if trim(pcv_desc)="" then
		pcv_desc=pcArr(12,i)
	end if
	if trim(pcv_desc)="" or trim(pcv_desc)="no information" then
		pcv_desc=pcv_name
	end if
	pcv_desc=RmvDescInvalidChrs(ClearHTMLTags2(pcv_desc,0))
	pcv_price=Round(pcArr(3,i),2)
	pcv_image=pcArr(4,i)
	if trim(pcv_image)="" then
		pcv_image=pcArr(5,i)
	end if
	if trim(pcv_image)="" then
		pcv_image=pcArr(6,i)
	end if
	if pcv_image="" then
		pcv_image="no_image.gif"
	end if
	pcv_image=scSiteURL & "pc/catalog/" & pcv_image
	pcv_cats=RmvNameInvalidChrs(ClearHTMLTags2(getCATInfor(pcArr(0,i)),0))
	pcv_brand=""
	if trim(pcArr(7,i))<>"0" then
		query="SELECT brandName FROM Brands WHERE idbrand=" & pcArr(7,i) & ";"
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			pcv_brand=RmvNameInvalidChrs(ClearHTMLTags2(rs1("BrandName"),0))
		end if
		set rs1=nothing
	end if
	pcv_stockstatus="No"
	if clng(pcArr(8,i))>0 OR clng(pcArr(9,i))<>0 OR clng(pcArr(10,i))<>0 then
		pcv_stockstatus="Yes"
	end if
	pcv_prdurl=scSiteURL & "pc/viewPrd.asp?idproduct=" & pcArr(0,i) & "&idcategory=" & pcv_idcat
	pcv_weight=0
	If scShipFromWeightUnit="KGS" then
	else
		pcv_weight=pcArr(11,i)
		pcv_weight=Round(cdbl(pcv_weight)/16,1)
	end if

	'***** End of Generate Data Lines *****	

	StrResults=StrResults & """" & pcv_brand & """,""" & pcv_sku & """,""" & pcv_name & """,""" & pcv_desc & """," & pcv_prdurl & "," & pcv_price & ",""" & pcv_cats & """," & pcv_image & "," & pcv_stockstatus & "," & pcv_weight & vbcrlf
	
	Next
	
	strFile=GenFileName()
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set A=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
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
		<p><strong>NOTE</strong>: NexTag requires that you specify a &quot;Manufacturer&quot; for all your products. This is the &quot;Brand&quot; name in ProductCart. If you do not have this information in your Product File, please contact <a href="mailto:sellersupport@nextag.com">sellersupport@nextag.com</a> for assistance.</p>
		<p style="padding-top:10px;"><b>Download your file.</b></p>
		<p style="padding-top:10px;"><a href="<%=strFile & ".csv"%>"><img src="images/DownLoad.gif"></a></p>
		<p style="padding-top:10px;">To ensure that your file downloads correctly, right click on the icon above<br>
			and choose &quot;<b>Save Target As...</b>&quot; from your menu.
		</p>
	</td>
</tr>
<%END IF%>
<% session("cp_exportNextTag_prdlist")=""
session("cp_exportNextTag_pagecurrent")=""
%>
</table>
<!--#include file="AdminFooter.asp"-->