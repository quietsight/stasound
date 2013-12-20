<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<%
Dim strQ, intProductId, ConnTemp, strServerTransfer, strPrdDetails, strCatDetails, strOther404, strQProductIdCount, strRestOfQueryString, pcIntValidPath
'====================================
'== Set page location ===============
'====================================

strCatDetails = "viewCategories.asp"

'====================================
'== You should not have to edit	  ===
'== any code after this point	  ===
'====================================

' Name of product details page
strPrdDetails = "viewPrd.asp" ' Location of product details page

' Name of content page
strCntDetails = "viewcontent.asp" ' Location of content page

' Name of parent content page
strCntParent = "viewpages.asp" ' Location of parent content page

' Redirection to standard 404 error page is peformed by 404b.asp.
' Open and edit 404b.asp to change the name of the default 404 error page.
strOther404 = "404b.asp"

' Get the Page Name
strQ = Request.ServerVariables("QUERY_STRING")
'// Troubleshooting: show page address
'Response.write strQ
'Response.End()

'Check for valid path
pcIntValidPath=0
if instr(strQ,"/pc/")>0 then pcIntValidPath=1

Function SEOcheckAff(tmpQ)
Dim tmpStr1,tmpStr2,k,tmp1,tmp2
	tmp1=Cint(1)
	if Instr(tmpQ,"?")>0 then
	tmp2=split(tmpQ,"?")
	if tmp2(1)<>"" then
		tmpStr1=split(tmp2(1),"&")
		For k=lbound(tmpStr1) to ubound(tmpStr1)
			if tmpStr1(k)<>"" then
				if Instr(Ucase(tmpStr1(k)),"IDAFFILIATE")>0 then
					tmpStr2=split(tmpStr1(k),"=")
					if tmpStr2(1)<>"" then
						if IsNumeric(tmpStr2(1)) then
							tmp1=Clng(tmpStr2(1))
						end if
					end if
				end if
			end if
		Next
	end if
	end if
	SEOcheckAff=tmp1
End Function

' Find the Product, Category, or Page ID
nIndex = InStrRev(strQ,"/")
If (nIndex>0) Then
	' Look for affiliate ID, set special session variable
	session("strSEOAffiliate")=SEOcheckAff(strQ)

	' Remove last character added on refresh (BTO configuration page)
	if InStrRev(strQ,"=",-1,1) then
	 strQCount=len(strQ)
	 strQtemp=left(strQ,strQCount-1)
	 strQ=strQtemp
	end if
	strQProductId = split(strQ,"?")
	strQProductIdCount = ubound(strQProductId)
	intProductId = Right(strQProductId(0),Len(strQProductId(0))-nIndex)
		nIndex2 = InStrRev(intProductId,"-",-1,1)
		If (nIndex2>0) Then
			intProductId = Right(intProductId,Len(intProductId)-nIndex2)
		end if
	if strQProductIdCount > 0 then
		strRestOfQueryString=strQProductId(1)
		session("strSeoQueryString")=strRestOfQueryString
		else
		session("strSeoQueryString")=""
	end if
else
	Server.Transfer(strOther404)
End If

' Detect whether this is an htm page
If Instr(LCase(intProductId),".htm") = 0 Then
	Server.Transfer(strOther404)
End If

' Detect whether this is a product, a category, or a content page
If Instr(LCase(intProductId),"c") <> 0 Then

	' START - CATEGORY PAGE

		intProductId=replace(intProductId,"c","")
		' Trim Off .htm from category ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If

		' Look Up the category Id In the Database
		if isNumeric(intProductid)=True then
				call openDb()
				query = "SELECT idCategory FROM categories WHERE idCategory = " & intProductId
				set rs = Server.CreateObject("ADODB.Recordset")
				Set rs = ConnTemp.Execute(query)
				If (Not rs.EOF) Then
					strServerTransfer = 1
					else
					strServerTransfer = 0
				End If
				set rs = nothing
				call closeDb()
			else
				strServerTransfer = 0
		end if

		' Go to the new page
		If strServerTransfer=1  and pcIntValidPath=1 then
			session("idCategoryRedirect") = intProductId
			session("idCategoryRedirectSF") = intProductId
			Server.Transfer(strCatDetails)
			else
			Server.Transfer(strOther404)
		end if

	' END - CATEGORY PAGE

elseif Instr(LCase(intProductId),"d") <> 0 then ' "d" stands for "document"

	' START - CONTENT PAGE

		intProductId=replace(intProductId,"d","")
		' Trim Off .htm from content page ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If

		' Look Up the content page Id In the Database
		if isNumeric(intProductid)=True then
				call openDb()
				query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_IDPage = "  & intProductId
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=ConnTemp.execute(query)
				If (Not rs.EOF) Then
					strServerTransfer = 1
					else
					strServerTransfer = 0
				End If
				set rs = nothing
				call closeDb()
			else
				strServerTransfer = 0
		end if

		' Go to the content page
		If strServerTransfer=1  and pcIntValidPath=1 then
			session("idContentPageRedirect") = intProductId
			Server.Transfer(strCntDetails)
			else
			Server.Transfer(strOther404)
		end if

	' END - CONTENT PAGE

elseif Instr(LCase(intProductId),"e") <> 0 then ' This handles a parent content page

	' START - PARENT CONTENT PAGE

		intProductId=replace(intProductId,"e","")
		' Trim Off .htm from content page ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If

		' Look Up the content page Id In the Database
		if isNumeric(intProductid)=True then
				call openDb()
				query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_IDPage = "  & intProductId
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=ConnTemp.execute(query)
				If (Not rs.EOF) Then
					strServerTransfer = 1
					else
					strServerTransfer = 0
				End If
				set rs = nothing
				call closeDb()
			else
				strServerTransfer = 0
		end if

		' Go to the content page
		If strServerTransfer=1  and pcIntValidPath=1 then
			session("idParentContentPageRedirect") = intProductId
			Server.Transfer(strCntParent)
			else
			Server.Transfer(strOther404)
		end if

	' END - PARENT CONTENT PAGE

else
	' This is a product page
		If Instr(LCase(intProductId),"p") <> 0 Then
			'This product is with a category
			strPrdCatArry=split(intProductId,"p")
			intTempCatId=strPrdCatArry(0)
			intProductId=strPrdCatArry(1)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		end if
		' Trim Off .htm from Product ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
		End If
		'// Troubleshooting: show Product ID
		'Response.write intProductId
		'Response.End()

		' Look Up the Product Id In the Database
		tmpHiddenCat=0
		if isNumeric(intProductid)=True then
				call openDb()
				query = "SELECT idProduct FROM products WHERE idProduct = " & intProductId
				set rs = Server.CreateObject("ADODB.Recordset")
				Set rs = ConnTemp.Execute(query)
				If (Not rs.EOF) Then
					strServerTransfer = 1
					set rs=nothing
					query="SELECT categories.idcategory FROM categories INNER JOIN categories_products ON categories.idcategory=categories_products.idcategory WHERE categories_products.idProduct=" & intProductId & " AND categories.iBTOhide=0;"
					set rs = Server.CreateObject("ADODB.Recordset")
					Set rs = ConnTemp.Execute(query)
					if rs.eof then
						session("intTempCatId")=0
						tmpHiddenCat=1
					end if
					set rs=nothing
				else
					strServerTransfer = 0
				End If
				set rs = nothing
				call closeDb()
			else
			strServerTransfer = 0
		end if

		' Go to the new page
		If strServerTransfer=1 and pcIntValidPath=1 then
			session("idProductRedirect") = intProductId
			if tmpHiddenCat=0 then
				session("intTempCatId") = intTempCatId
			end if
			session("MobileURL")=strPrdDetails
			Server.Transfer(strPrdDetails)
			else
			Server.Transfer(strOther404)
		end if

end if ' End category vs product link
%>