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
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<% 
response.Buffer=true
Response.Expires=0

dim query, conntemp, rstemp

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

query="SELECT idproduct,sku,description,sdesc,details,price,listprice,btoBprice,serviceSpec, configOnly,weight,stock,pcProd_ReorderLevel,imageURL,smallImageUrl,largeImageURL,idbrand,active,listHidden,hotDeal,notax,noshipping,formQuantity,emailtext,OverSizeSpec,removed FROM products order by idproduct"
set rstemp=Server.CreateObject("ADODB.Recordset")     
rstemp.Open query, conntemp, adOpenForwardOnly, adLockReadOnly, adCmdText
IF rstemp.eof then
	set rstemp=nothing
	call closedb()
	%>
	<!--#include file="AdminHeader.asp"-->
	<table class="pcCPcontent">
	<tr>
		<td>
			<div class="pcCPmessage">
				Your search did not return any results.
			</div>
			<p>&nbsp;</p>
			<p>
				<input type=button value=" Back " onclick="javascript:history.back()" class="ibtnGrey">
			</p>
		</td>
	</tr>
	</table>
	<!--#include file="AdminFooter.asp"-->
	<%response.end
ELSE
		HTMLResult=""
		set StringBuilderObj = new StringBuilder
		If pid="1" then
			StringBuilderObj.append "<td><b>" & "Product ID"& "</b></td>"
		End If
		If psku="1" Then
			StringBuilderObj.append "<td><b>" & "SKU"& "</b></td>"
		End If
		If pname="1" then
			StringBuilderObj.append "<td><b>" & "Product Name"& "</b></td>"
		End If
		If psdesc="1" Then
			StringBuilderObj.append "<td><b>" & "Short Description"& "</b></td>"
		End If
		If pdesc="1" then
			StringBuilderObj.append "<td><b>" & "Description"& "</b></td>"
		End If
		If pprice="1" Then
			StringBuilderObj.append "<td><b>" & "Price"& "</b></td>"
		End If
		If plprice="1" then
			StringBuilderObj.append "<td><b>" & "List Price"& "</b></td>"
		End If
		If pwprice="1" Then
			StringBuilderObj.append "<td><b>" & "Wholesale Price"& "</b></td>"
		End If
		If ptype="1" then
			StringBuilderObj.append "<td><b>" & "Product Type"& "</b></td>"
		End If
		If pweight="1" Then
			StringBuilderObj.append "<td><b>" & "Weight"& "</b></td>"
		End If
		If pstock="1" then
			StringBuilderObj.append "<td><b>" & "Stock"& "</b></td>"
		End If
		If pLIN="1" then
			StringBuilderObj.append "<td><b>" & "LIN Amount"& "</b></td>"
		End If
		If pimg="1" Then
			StringBuilderObj.append "<td><b>" & "General Image"& "</b></td>"
		End If
		If ptimg="1" then
			StringBuilderObj.append "<td><b>" & "Thumbnail Image"& "</b></td>"
		End If
		If pdimg="1" Then
			StringBuilderObj.append "<td><b>" & "Details view Image"& "</b></td>"
		End If
		If pbrand="1" then
			StringBuilderObj.append "<td><b>" & "Brand ID"& "</b></td>"
		End If
		If pactive="1" Then
			StringBuilderObj.append "<td><b>" & "Active"& "</b></td>"
		End If
		If psavings="1" then
			StringBuilderObj.append "<td><b>" & "Show savings"& "</b></td>"
		End If
		If pSpecial="1" Then
			StringBuilderObj.append "<td><b>" & "Special"& "</b></td>"
		End If
		If pnotax="1" then
			StringBuilderObj.append "<td><b>" & "No taxable"& "</b></td>"
		End If
		If pnoship="1" Then
			StringBuilderObj.append "<td><b>" & "No shipping charge"& "</b></td>"
		End If
		If pnosale="1" then
			StringBuilderObj.append "<td><b>" & "Not for sale"& "</b></td>"
		End If
		If pnosalecopy="1" Then
			StringBuilderObj.append "<td><b>" & "Not for sale copy"& "</b></td>"
		End If
		If poversize="1" then
			StringBuilderObj.append "<td><b>" & "Oversize"& "</b></td>"
		End If
		If pcv_catinfo="1" then
			StringBuilderObj.append "<td><b>" & "Categories"& "</b></td>"
		End If
		If pRemoved="1" then
			StringBuilderObj.append "<td><b>" & "Removed/Deleted"& "</b></td>"
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
					StringBuilderObj.append "<td><b>" & tmpArr(1,i) & "</b></td>"
				Next
				ReDim valueArr(ubound(tmpArr,2))
			end if
			rstemp.MoveFirst
		end if
		end if
		HTMLResult="<table><tr>" & StringBuilderObj.toString() & "</tr>"
		set StringBuilderObj = nothing
		
		Do Until rstemp.EOF
			pcv_idproduct=rstemp("idproduct")
			set StringBuilderObj = new StringBuilder
			If pid="1" then
				StringBuilderObj.append "<td>" & rstemp("idproduct")& "</td>"
			End If
			If psku="1" Then
				StringBuilderObj.append "<td>" & rstemp("sku")& "</td>"
			End If
			If pname="1" then
				StringBuilderObj.append "<td>" & rstemp("description")& "</td>"
			End If
			If psdesc="1" Then
				StringBuilderObj.append "<td>" & rstemp("sdesc")& "</td>"
			End If
			If pdesc="1" then
				StringBuilderObj.append "<td>" & rstemp("details")& "</td>"
			End If
			If pprice="1" Then
				StringBuilderObj.append "<td>" & rstemp("price")& "</td>"
			End If
			If plprice="1" then
				StringBuilderObj.append "<td>" & rstemp("listprice")& "</td>"
			End If
			If pwprice="1" Then
				StringBuilderObj.append "<td>" & rstemp("btobprice")& "</td>"
			End If
			If ptype="1" then
				if rstemp("serviceSpec")=-1 then
					StringBuilderObj.append "<td>" & "BTO"& "</td>"
				else
					if rstemp("configOnly")=-1 then
						StringBuilderObj.append "<td>" & "ITEM"& "</td>"
					else
						StringBuilderObj.append "<td>" & ""& "</td>"
					end if
				end if		
			End If
			If pweight="1" Then
				StringBuilderObj.append "<td>" & rstemp("weight")& "</td>"
			End If
			If pstock="1" then
				StringBuilderObj.append "<td>" & rstemp("stock")& "</td>"
			End If
			If pLIN="1" then
				StringBuilderObj.append "<td>" & rstemp("pcProd_ReorderLevel")& "</td>"
			End If
			If pimg="1" Then
				StringBuilderObj.append "<td>" & rstemp("imageUrl")& "</td>"
			End If
			If ptimg="1" then
				StringBuilderObj.append "<td>" & rstemp("smallimageUrl")& "</td>"
			End If
			If pdimg="1" Then
				StringBuilderObj.append "<td>" & rstemp("largeImageURL")& "</td>"
			End If
			If pbrand="1" then
				StringBuilderObj.append "<td>" & rstemp("idbrand")& "</td>"
			End If
			If pactive="1" Then
				StringBuilderObj.append "<td>" & rstemp("active")& "</td>"
			End If
			If psavings="1" then
				StringBuilderObj.append "<td>" & rstemp("listhidden")& "</td>"
			End If
			If pSpecial="1" Then
				StringBuilderObj.append "<td>" & rstemp("hotDeal")& "</td>"
			End If
			If pnotax="1" then
				StringBuilderObj.append "<td>" & rstemp("notax")& "</td>"
			End If
			If pnoship="1" Then
				StringBuilderObj.append "<td>" & rstemp("noshipping")& "</td>"
			End If
			If pnosale="1" then
				StringBuilderObj.append "<td>" & rstemp("formQuantity")& "</td>"
			End If
			If pnosalecopy="1" Then
				StringBuilderObj.append "<td>" & rstemp("emailText")& "</td>"
			End If
			If poversize="1" then
				StringBuilderObj.append "<td>" & rstemp("OverSizeSpec")& "</td>"
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
							pcv_tmp1=pcv_tmp1 & "||"
						end if
						pcv_tmp1=pcv_tmp1 & tmp_Arr(0,j)
					Next
				end if
				set rs1=nothing
				StringBuilderObj.append "<td>" & pcv_tmp1& "</td>"
			End if
			If pRemoved="1" then
				StringBuilderObj.append "<td>" & rstemp("removed")& "</td>"
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
				StringBuilderObj.append "<td>" & valueArr(k) & "</td>"
			Next
			end if
			HTMLResult=HTMLResult & "<tr>" & StringBuilderObj.toString() & "</tr>"
			set StringBuilderObj = nothing
			rstemp.MoveNext
		Loop
set rstemp=nothing
HTMLResult=HTMLResult & "</table>"
END IF
closedb()
%>
<% 
Response.ContentType = "application/vnd.ms-excel"
%>
<%=HTMLResult%>