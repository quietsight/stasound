<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Products" %>
<% section="" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<% 
response.Buffer=true
Response.Expires=0
	
'on error resume next 
dim mySQL, conntemp, rstemp

call openDb()
' Choose the records to display
pid=request.form("fpid")
psku=request.form("fpsku")	
pname=request.form("fpname")
psdesc=request.form("fpsdesc")	
pdesc=request.form("fpdesc")
pprice=request.form("fpprice")	
plprice=request.form("fplprice")
pwprice=request.form("fpwprice")	
ptype=request.form("fptype")
pimg=request.form("fpimg")	
ptimg=request.form("fptimg")
pdimg=request.form("fpdimg")	
pweight=request.form("fpweight")
pstock=request.form("fpstock")
pLIN=request.form("fpLIN")
pbrand=request.form("fpbrand")
pactive=request.form("fpactive")	
psavings=request.form("fpsavings")
pSpecial=request.form("fpspecial")	
pnotax=request.form("fpnotax")
pnoship=request.form("fpnoship")	
pnosale=request.form("fpnosale")
pnosalecopy=request.form("fpnosalecopy")	
poversize=request.form("fpoversize")
pcv_catinfo=request.form("catinfor")
pRemoved=request.form("fpRemoved")

'// Find out if deleted products should be included
pIncludeDeleted=request.Form("includeDeleted")
	if pIncludeDeleted="" or not IsNumeric(pIncludeDeleted) then
		pIncludeDeleted=0
	end if
	if pIncludeDeleted=0 then
		strSQL2=" WHERE removed=0"
		elseif pIncludeDeleted=2 then
		strSQL2=" WHERE removed=-1"
		else
		strSQL2=""
	end if

strSQL="SELECT idproduct,sku,description,sdesc,details,price,listprice,btoBprice,serviceSpec, configOnly,weight,stock,pcProd_ReorderLevel,imageURL,smallImageUrl,largeImageURL,idbrand,active,listHidden,hotDeal,notax,noshipping,formQuantity,emailtext,OverSizeSpec,removed FROM products" & strSQL2 & " order by idproduct"
set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.Open strSQL, conntemp, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<html>
<head>
<title>Custom Product Export</title>
        <style>
		h1 {
			font-family: Arial, Helvetica, sans-serif;
			font-size: 16px;
			font-weight: bold;
		}
		
		table.productExport {
			padding: 0;
			margin: 0;
		}
		
		table.productExport td {
			font-family: Arial, Helvetica, sans-serif;
			font-size: 11px;
			padding: 3px;
			border-right: 1px solid #CCC;
			border-bottom: 1px solid #CCC;
		}
		
		table.productExport th {
			font-family: Arial, Helvetica, sans-serif;
			font-size: 12px;
			padding: 3px;
			font-weight: bold;
			text-align: left;
			background-color: #f5f5f5;
			border-right: 1px solid #CCC;
			border-bottom: 1px solid #CCC;
		}
		</style>
</head>
<body>
<% dim strReturnAs
strReturnAs=request.Form("ReturnAS")
select case strReturnAS
	case "CSV"
		CreateCSVFile()
	case "HTML"
		GenHTML()
	case "XLS"
		CreateXlsFile()
end select		
   

Function GenFileName()
	dim fname
	fname="File"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime))
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

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

Function GenHTML()%>
<h1>Custom Product Export</h1>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=2 width="100%" class="productExport">
	<tr>
	<%	If pid="1" then %>
		<th>Product ID</th>
	
	<% End If
	If psku="1" then %>
		<th>SKU</th>
	
	<% End If
	If pname="1" then %>
		<th>Product Name</th>
	<% End If
	If psdesc="1" then %>
		<th>Short Description</th>
	<% End If
	If pdesc="1" then %>
		<th>Description</th>
	<% End If
	If pprice="1" then %><th>Price</th>
	<% End If
	If plprice="1" then %>
		<th>List Price</th>
	<% End If
	If pwprice="1" then %><th>Wholesale Price</th>
	<% End If
	If ptype="1" then %>
		<th>Product Type</th>
	<% End If
	If pweight="1" then %><th>Weight</th>
	<% End If
	If pstock="1" then %>
		<th>Stock</th>
	<% End If
	If pLIN="1" then %>
		<th>LIN Amount</th>
	<% End If
	If pimg="1" then %><th>General Image</th>
	<% End If
	If ptimg="1" then %>
		<th>Thumbnail Image</th>
	<% End If
	If pdimg="1" then %><th>Details view Image</th>
	<% End If
	If pbrand="1" then %>
		<th>Brand ID</th>
	<% End If
	If pactive="1" then %><th>Active</th>
	<% End If
	If psavings="1" then %>
		<th>Show savings</th>
	<% End If
	If pSpecial="1" then %><th>Special</th>
	<% End If
	If pnotax="1" then %>
		<th>No taxable</th>
	<% End If
	If pnoship="1" then %><th>No shipping charge</th>
	<% End If
	If pnosale="1" then %>
		<th>Not for sale</th>
	<% End If
	If pnosalecopy="1" then %><th>Not for sale copy</th>
	<% End If
	If poversize="1" then %>
		<th>Oversize</th>
	<% End If
	If pcv_catinfo="1" then %>
		<th>Categories</th>
	<% End If
	If pRemoved="1" then %>
		<th>Removed/Deleted</th>
	<% End If %>
	<%if request("CSearchFields")="1" then
	if not rstemp.eof then
		pcArr=rstemp.getRows()
		intCount=ubound(pcArr,2)
		tmpPrd=""
		For i=0 to intCount
			if tmpPrd<>"" then
				tmpPrd=tmpPrd & ","
			end if
			tmpPrd=tmpPrd & pcArr(0,i)
		Next
		tmpPrd="(" & tmpPrd & ")"
		query="SELECT DISTINCT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idProduct IN " & tmpPrd & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			tmpArr=rsQ.getRows()
			SearchFieldCount=ubound(tmpArr,2)+1
		end if
		set rsQ=nothing
		if SearchFieldCount>0 then
			For i=0 to SearchFieldCount-1%>
			<th><%=tmpArr(1,i)%></th>
			<%Next
			ReDim valueArr(ubound(tmpArr,2))
		end if
		rstemp.MoveFirst
	end if
	end if%>
</TR>
<%if(rstemp.BOF=True and rstemp.EOF=True) then%>
	<tr>
		<td valign="top" nowrap>Product database is empty</td>
	</tr>
<% else
	rstemp.MoveFirst
	Do While Not rstemp.EOF
		pcv_idproduct=rstemp("idproduct")%>
		<TR>
			<%If pid="1" then %>
				<td><%=rstemp("idproduct")%>&nbsp;</td>
			<% End If
			If psku="1" then %><td><%=rstemp("sku")%>&nbsp;</td>
			<% End If
			If pname="1" then %>
			<td><%=rstemp("description")%>&nbsp;</td>
			<% End If
			If psdesc="1" then %>
            <td><%=rstemp("sdesc")%>&nbsp;</td>
			<% End If
			If pdesc="1" then %>
			<td><%=rstemp("details")%>&nbsp;</td>
			<% End If
			If pprice="1" then %>
			<td><%=rstemp("price")%>&nbsp;</td>
			<% End If
			If plprice="1" then %>
			<td><%=rstemp("listprice")%>&nbsp;</td>
			<% End If
			If pwprice="1" then %><td><%=rstemp("btoBprice")%>&nbsp;</td>
			<% End If
			If ptype="1" then
				PTType=""
				if rstemp("serviceSpec")=-1 then
				PTType="BTO"
				else
					if rstemp("configOnly")=-1 then
					PTType="ITEM"
					end if
				end if %>
				<td><%=PTType%>&nbsp;</td>
			<% End If
			If pweight="1" then %><td><%=rstemp("weight")%>&nbsp;</td>
			<% End If
			If pstock="1" then %>
			<td><%=rstemp("stock")%>&nbsp;</td>
			<% End If
			If pLIN="1" then %>
			<td><%=rstemp("pcProd_ReorderLevel")%>&nbsp;</td>
			<% End If
			If pimg="1" then %><td><%=rstemp("imageURL")%>&nbsp;</td>
			<% End If
			If ptimg="1" then %>
			<td><%=rstemp("smallImageUrl")%>&nbsp;</td>
			<% End If
			If pdimg="1" then %><td><%=rstemp("largeImageURL")%>&nbsp;</td>
			<% End If
			If pbrand="1" then %>
			<td><%=rstemp("idbrand")%>&nbsp;</td>
			<% End If
			If pactive="1" then %><td><%=rstemp("active")%>&nbsp;</td>
			<% End If
			If psavings="1" then %>
			<td><%=rstemp("listHidden")%>&nbsp;</td>
			<% End If
			If pSpecial="1" then %><td><%=rstemp("hotDeal")%>&nbsp;</td>
			<% End If
			If pnotax="1" then %>
			<td><%=rstemp("notax")%>&nbsp;</td>
			<% End If
			If pnoship="1" then %><td><%=rstemp("noshipping")%>&nbsp;</td>
			<% End If
			If pnosale="1" then %>
			<td><%=rstemp("formQuantity")%>&nbsp;</td>
			<% End If
			If pnosalecopy="1" then %><td><%=rstemp("emailtext")%>&nbsp;</td>
			<% End If
			If poversize="1" then %>
			<td><%=rstemp("OverSizeSpec")%>&nbsp;</td>
			<% End If
			If pcv_catinfo="1" then
			pcv_tmp1=""
			query="SELECT DISTINCT categories.categoryDesc FROM categories INNER JOIN categories_products ON categories.idcategory=categories_products.idcategory WHERE categories_products.idproduct=" & pcv_idproduct & ";"
			set rs1=connTemp.execute(query)
			if not rs1.eof then
				tmp_Arr=rs1.getRows()
				intCount=ubound(tmp_Arr,2)
				For j=0 to intCount
					if pcv_tmp1<>"" then
						pcv_tmp1=pcv_tmp1 & "**"
					end if
					pcv_tmp1=pcv_tmp1 & tmp_Arr(0,j)
				Next
			end if
			set rs1=nothing%>
			<td><%=pcv_tmp1%>&nbsp;</td>
			<% End If
			If pRemoved="1" then %><td><%=rstemp("removed")%>&nbsp;</td>
			<% End If%>
			<%if request("CSearchFields")="1" AND SearchFieldCount>0 then
			For k=0 to SearchFieldCount-1
				valueArr(k)=""
			Next
			query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pcv_idProduct & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpValue=rsQ.getRows()
				set rsQ=nothing
				intCount1=ubound(tmpValue,2)
				For k=0 to intCount1
					For m=0 to SearchFieldCount-1
						if clng(tmpValue(0,k))=clng(tmpArr(0,m)) then
							valueArr(m)=tmpValue(3,k)
							exit for
						end if
					Next
				Next
			end if
			set rsQ=nothing
			For k=0 to SearchFieldCount-1%>
				<td><%=valueArr(k)%>&nbsp;</td>
			<%Next
			end if%>
			</TR>
			<% rstemp.movenext
		loop
	End if%>
</TABLE>
<p align="left"><font face="Verdana" size="1"><b>Exported on: <%=now()%></b></font></p>	
<%End Function
	Function CreateCSVFile()
		strFile=GenFileName()   
		Set fs=Server.CreateObject("Scripting.FileSystemObject")
		Set a=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
		If Not rstemp.EOF Then
		
			set StringBuilderObj = new StringBuilder
		If pid="1" then
			StringBuilderObj.append chr(34) & "Product ID" & chr(34) & ","
		End If
		If psku="1" Then
			StringBuilderObj.append chr(34) & "SKU" & chr(34) & ","
		End If
		If pname="1" then
			StringBuilderObj.append chr(34) & "Product Name" & chr(34) & ","
		End If
		If psdesc="1" Then
			StringBuilderObj.append chr(34) & "Short Description" & chr(34) & ","
		End If
		If pdesc="1" then
			StringBuilderObj.append chr(34) & "Description" & chr(34) & ","
		End If
		If pprice="1" Then
			StringBuilderObj.append chr(34) & "Price" & chr(34) & ","
		End If
		If plprice="1" then
			StringBuilderObj.append chr(34) & "List Price" & chr(34) & ","
		End If
		If pwprice="1" Then
			StringBuilderObj.append chr(34) & "Wholesale Price" & chr(34) & ","
		End If
		If ptype="1" then
			StringBuilderObj.append chr(34) & "Product Type" & chr(34) & ","
		End If
		If pweight="1" Then
			StringBuilderObj.append chr(34) & "Weight" & chr(34) & ","
		End If
		If pstock="1" then
			StringBuilderObj.append chr(34) & "Stock" & chr(34) & ","
		End If
		If pLIN="1" then
			StringBuilderObj.append chr(34) & "LIN Amount" & chr(34) & ","
		End If
		If pimg="1" Then
			StringBuilderObj.append chr(34) & "General Image" & chr(34) & ","
		End If
		If ptimg="1" then
			StringBuilderObj.append chr(34) & "Thumbnail Image" & chr(34) & ","
		End If
		If pdimg="1" Then
			StringBuilderObj.append chr(34) & "Details view Image" & chr(34) & ","
		End If
		If pbrand="1" then
			StringBuilderObj.append chr(34) & "Brand ID" & chr(34) & ","
		End If
		If pactive="1" Then
			StringBuilderObj.append chr(34) & "Active" & chr(34) & ","
		End If
		If psavings="1" then
			StringBuilderObj.append chr(34) & "Show savings" & chr(34) & ","
		End If
		If pSpecial="1" Then
			StringBuilderObj.append chr(34) & "Special" & chr(34) & ","
		End If
		If pnotax="1" then
			StringBuilderObj.append chr(34) & "No taxable" & chr(34) & ","
		End If
		If pnoship="1" Then
			StringBuilderObj.append chr(34) & "No shipping charge" & chr(34) & ","
		End If
		If pnosale="1" then
			StringBuilderObj.append chr(34) & "Not for sale" & chr(34) & ","
		End If
		If pnosalecopy="1" Then
			StringBuilderObj.append chr(34) & "Not for sale copy" & chr(34) & ","
		End If
		If poversize="1" then
			StringBuilderObj.append chr(34) & "Oversize" & chr(34) & ","
		End If
		If pcv_catinfo="1" then
			StringBuilderObj.append chr(34) & "Categories" & chr(34) & ","
		End If
		If pRemoved="1" then
			StringBuilderObj.append chr(34) & "Removed/Deleted" & chr(34) & ","
		End If
		if request("CSearchFields")="1" then
		if not rstemp.eof then
			pcArr=rstemp.getRows()
			intCount=ubound(pcArr,2)
			tmpPrd=""
			For i=0 to intCount
				if tmpPrd<>"" then
					tmpPrd=tmpPrd & ","
				end if
				tmpPrd=tmpPrd & pcArr(0,i)
			Next
			tmpPrd="(" & tmpPrd & ")"
			query="SELECT DISTINCT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idProduct IN " & tmpPrd & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpArr=rsQ.getRows()
				SearchFieldCount=ubound(tmpArr,2)+1
			end if
			set rsQ=nothing
			if SearchFieldCount>0 then
				For i=0 to SearchFieldCount-1
					StringBuilderObj.append chr(34) & tmpArr(1,i) & chr(34) & ","
				Next
				ReDim valueArr(ubound(tmpArr,2))
			end if
			rstemp.MoveFirst
		end if
		end if
		a.WriteLine(StringBuilderObj.toString())
		set StringBuilderObj = nothing
		
		Do Until rstemp.EOF
			pcv_idproduct=rstemp("idproduct")
			
			set StringBuilderObj = new StringBuilder
			If pid="1" then
				StringBuilderObj.append rstemp("idproduct") & ","
			End If
			If psku="1" Then
				StringBuilderObj.append chr(34) & rstemp("sku") & chr(34) & ","
			End If
			If pname="1" then
				StringBuilderObj.append chr(34) & rstemp("description") & chr(34) & ","
			End If
			If psdesc="1" Then
				tmp_PrdSDesc=rstemp("sdesc")
				if trim(tmp_PrdSDesc)<>"" then
					tmp_PrdSDesc=RmvHTMLWhiteSpace(replace(tmp_PrdSDesc,"""",""""""))
					tmp_PrdSDesc=replace(tmp_PrdSDesc,"&quot;","""""")
					tmp_PrdSDesc=replace(tmp_PrdSDesc,vbCrLf,"")
					tmp_PrdSDesc=replace(tmp_PrdSDesc,vbCr,"")
					tmp_PrdSDesc=replace(tmp_PrdSDesc,vbLf,"")
				end if
				StringBuilderObj.append chr(34) & tmp_PrdSDesc & chr(34) & ","
			End If
			If pdesc="1" then
				tmp_PrdDesc=rstemp("details")
				if trim(tmp_PrdDesc)<>"" then
					tmp_PrdDesc=RmvHTMLWhiteSpace(replace(tmp_PrdDesc,"""",""""""))
					tmp_PrdDesc=replace(tmp_PrdDesc,"&quot;","""""")
					tmp_PrdDesc=replace(tmp_PrdDesc,vbCrLf,"")
					tmp_PrdDesc=replace(tmp_PrdDesc,vbCr,"")
					tmp_PrdDesc=replace(tmp_PrdDesc,vbLf,"")
				end if
				StringBuilderObj.append chr(34) & tmp_PrdDesc & chr(34) & ","
			End If
			If pprice="1" Then
				StringBuilderObj.append rstemp("price") & ","
			End If
			If plprice="1" then
				StringBuilderObj.append rstemp("listprice") & ","
			End If
			If pwprice="1" Then
				StringBuilderObj.append rstemp("btobprice") & ","
			End If
			If ptype="1" then
				if rstemp("serviceSpec")=-1 then
					StringBuilderObj.append chr(34) & "BTO" & chr(34) & ","
				else
					if rstemp("configOnly")=-1 then
						StringBuilderObj.append chr(34) & "ITEM" & chr(34) & ","
					else
						StringBuilderObj.append chr(34) & "" & chr(34) & ","
					end if
				end if		
			End If
			If pweight="1" Then
				StringBuilderObj.append rstemp("weight") & ","
			End If
			If pstock="1" then
				StringBuilderObj.append rstemp("stock") & ","
			End If
			If pLIN="1" then
				StringBuilderObj.append rstemp("pcProd_ReorderLevel") & ","
			End If
			If pimg="1" Then
				StringBuilderObj.append chr(34) & rstemp("imageUrl") & chr(34) & ","
			End If
			If ptimg="1" then
				StringBuilderObj.append chr(34) & rstemp("smallimageUrl") & chr(34) & ","
			End If
			If pdimg="1" Then
				StringBuilderObj.append chr(34) & rstemp("largeImageURL") & chr(34) & ","
			End If
			If pbrand="1" then
				StringBuilderObj.append rstemp("idbrand") & ","
			End If
			If pactive="1" Then
				StringBuilderObj.append rstemp("active") & ","
			End If
			If psavings="1" then
				StringBuilderObj.append rstemp("listhidden") & ","
			End If
			If pSpecial="1" Then
				StringBuilderObj.append rstemp("hotDeal") & ","
			End If
			If pnotax="1" then
				StringBuilderObj.append rstemp("notax") & ","
			End If
			If pnoship="1" Then
				StringBuilderObj.append rstemp("noshipping") & ","
			End If
			If pnosale="1" then
				StringBuilderObj.append rstemp("formQuantity") & ","
			End If
			If pnosalecopy="1" Then
				StringBuilderObj.append chr(34) & rstemp("emailText") & chr(34) & ","
			End If
			If poversize="1" then
				StringBuilderObj.append chr(34) & rstemp("OverSizeSpec") & chr(34) & ","
			End If
			If pcv_catinfo="1" then
				pcv_tmp1=""
				query="SELECT DISTINCT categories.categoryDesc FROM categories INNER JOIN categories_products ON categories.idcategory=categories_products.idcategory WHERE categories_products.idproduct=" & pcv_idproduct & ";"
				set rs1=connTemp.execute(query)
				if not rs1.eof then
					tmp_Arr=rs1.getRows()
					intCount=ubound(tmp_Arr,2)
					For j=0 to intCount
						if pcv_tmp1<>"" then
							pcv_tmp1=pcv_tmp1 & "**"
						end if
						pcv_tmp1=pcv_tmp1 & tmp_Arr(0,j)
					Next
				end if
				set rs1=nothing
				StringBuilderObj.append chr(34) & pcv_tmp1 & chr(34) & ","
			End if
			If pRemoved="1" then
				StringBuilderObj.append chr(34) & rstemp("removed") & chr(34) & ","
			End If
			if request("CSearchFields")="1" AND SearchFieldCount>0 then
			For k=0 to SearchFieldCount-1
				valueArr(k)=""
			Next
			query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pcv_idProduct & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpValue=rsQ.getRows()
				set rsQ=nothing
				intCount1=ubound(tmpValue,2)
				For k=0 to intCount1
					For m=0 to SearchFieldCount-1
						if clng(tmpValue(0,k))=clng(tmpArr(0,m)) then
							valueArr(m)=tmpValue(3,k)
							exit for
						end if
					Next
				Next
			end if
			set rsQ=nothing
			For k=0 to SearchFieldCount-1
				StringBuilderObj.append chr(34) & valueArr(k) & chr(34) & ","
			Next
			end if
			a.Writeline(StringBuilderObj.toString())
			set StringBuilderObj = nothing
			rstemp.MoveNext
		Loop
	End If
	a.Close
	Set fs=Nothing
	response.redirect "getFile.asp?file="& strFile &"&Type=csv"	
End Function
		
Function CreateXlsFile()
	Dim xlWorkSheet					' Excel Worksheet object
	Dim xlApplication 
	Set xlApplication=CreateObject("Excel.Application")
	xlApplication.Visible=False
	xlApplication.Workbooks.Add
	Set xlWorksheet=xlApplication.Worksheets(1)
	t=0
	If pid="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Product ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If psku="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="SKU"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pname="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Product Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If psdesc="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Short Description"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pdesc="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Description"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pprice="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Price"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If plprice="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="List Price"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pwprice="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Wholesale Price"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If ptype="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Product Type"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pweight="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Weight"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pstock="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Stock"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pLIN="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="LIN Amount"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pimg="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="General Image"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If ptimg="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Thumbnail Image"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pdimg="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Details view Image"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pbrand="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Brand ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pactive="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Active"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If psavings="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Show savings"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pSpecial="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Special"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pnotax="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="No taxable"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pnoship="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="No shipping charge"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pnosale="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Not for sale"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pnosalecopy="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Not for sale copy"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If poversize="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Oversize"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pRemoved="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Removed/Deleted"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	if request("CSearchFields")="1" then
		if not rstemp.eof then
			pcArr=rstemp.getRows()
			intCount=ubound(pcArr,2)
			tmpPrd=""
			For i=0 to intCount
				if tmpPrd<>"" then
					tmpPrd=tmpPrd & ","
				end if
				tmpPrd=tmpPrd & pcArr(0,i)
			Next
			tmpPrd="(" & tmpPrd & ")"
			query="SELECT DISTINCT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idProduct IN " & tmpPrd & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpArr=rsQ.getRows()
				SearchFieldCount=ubound(tmpArr,2)+1
			end if
			set rsQ=nothing
			if SearchFieldCount>0 then
				For i=0 to SearchFieldCount-1
					t=t+1
					xlWorksheet.Cells(1,t).Value=tmpArr(1,i)
					xlWorksheet.Cells(1,t).Interior.ColorIndex=6
				Next
				ReDim valueArr(ubound(tmpArr,2))
			end if
			rstemp.MoveFirst
		end if
	end if
	iRow=2
	If Not rstemp.EOF Then
		Do Until rstemp.EOF
			t=0
			If pid="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("idproduct")
			End If
			If psku="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("sku")
			End If
			If pname="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("description")
			End If
			If psdesc="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("sdesc")
			End If
			If pdesc="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("details")
			End If
			If pprice="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("price")
			End If
			If plprice="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("listprice")
			End If
			If pwprice="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("btobprice")
			End If
			If ptype="1" then
				t=t+1
				if rstemp("ServiceSpec")=-1 then
					xlWorksheet.Cells(iRow,i + t).Value="'" & "BTO"
				else
					if rstemp("ConfigOnly")=-1 then
						xlWorksheet.Cells(iRow,i + t).Value="'" & "ITEM"
					else
						xlWorksheet.Cells(iRow,i + t).Value="'"
					end if
				end if	
			End If
			If pweight="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("weight")
			End If
			If pstock="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("stock")
			End If
			If pLIN="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("pcProd_ReorderLevel")
			End If
			If pimg="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("imageURL")
			End If
			If ptimg="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("smallImageURL")
			End If
			If pdimg="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("largeImageURL")
			End If
			If pbrand="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("idbrand")
			End If
			If pactive="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("active")
			End If
			If psavings="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("listhidden")
			End If
			If pSpecial="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("hotDeal")
			End If
			If pnotax="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("notax")
			End If
			If pnoship="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("noshipping")
			End If
			If pnosale="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=rstemp("formQuantity")
			End If
			If pnosalecopy="1" Then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("emailtext")
			End If
			If poversize="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("OverSizeSpec")
			End If
			If pRemoved="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & rstemp("removed")
			End If
			if request("CSearchFields")="1" AND SearchFieldCount>0 then
			For k=0 to SearchFieldCount-1
				valueArr(k)=""
			Next
			query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pcv_idProduct & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpValue=rsQ.getRows()
				set rsQ=nothing
				intCount1=ubound(tmpValue,2)
				For k=0 to intCount1
					For m=0 to SearchFieldCount-1
						if clng(tmpValue(0,k))=clng(tmpArr(0,m)) then
							valueArr(m)=tmpValue(3,k)
							exit for
						end if
					Next
				Next
			end if
			set rsQ=nothing
			For k=0 to SearchFieldCount-1
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value="'" & valueArr(k)
			Next
			end if
			iRow=iRow + 1
			rstemp.MoveNext
		Loop
	End If
	strFile=GenFileName()
	xlWorksheet.SaveAs Server.MapPath(".") & "\" & strFile & ".xls"
	xlApplication.Quit												' Close the Workbook
	Set xlWorksheet=Nothing
	Set xlApplication=Nothing
	response.redirect "getFile.asp?file="& strFile &"&Type=xls"
End Function
%>
</body>
</html>