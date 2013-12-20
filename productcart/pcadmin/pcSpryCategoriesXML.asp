<%@ CodePage=65001 %>
<% Option Explicit %>
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/pcSpryXMLFuntions.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%
Dim conn, rs, sSql
Dim nDisplayRecs
Dim nStartRec, nStopRec, nTotalRecs, nRecCount, nRecActual, nTotalRecs2
Dim x_idCategory
Dim x_idParentCategory
Dim x_categoryDesc
Dim x_pcCats_BreadCrumbs
Dim XMLDoc, XMLRoot, XMLRow, Output, sXMLEncoding
Dim query, pcv_intTotalRecords
Dim tmpParent
Dim pcv_IdRootCategory, pcArray_NextTier
Dim intCount,pcArrP,rstemp1
Dim pcv_ExistingCats

On Error Resume Next
if request("mode")=2 then
	nRecActual = 0
else
	nRecActual = 1
end if

pcv_IdRootCategory=request("idRootCategory")
if NOT isNumeric(pcv_IdRootCategory) or pcv_IdRootCategory="" then
	pcv_IdRootCategory=1
end if

pcv_ExistingCats=request("strCats")
if pcv_ExistingCats="" then
	pcv_ExistingCats="0,"
end if

' Output all records by default
nDisplayRecs = -1

' Get MSXML object
Set XMLDoc = pcv_GetMSXML()
If Not IsObject(XMLDoc) Then
	Response.Write "MSXML 3 or later not installed"
	Response.End
End If

' Create and append the root element
Set XMLRoot = XMLDoc.createElement("categories")
XMLDoc.appendChild XMLRoot

dim pcv_cats1, pcv_cats3, pcv_cats6, pcv_cats4
dim pcv_BC 
dim pcv_tmpParent
dim rsCPObj
dim pcv_SkipNode

'/////////////////////////////////////////////////////////////////////
'// START: Generate Category List Recursively
'/////////////////////////////////////////////////////////////////////
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open scDSN

'// Get a List of Categories
pcs_GetParentList()

'// Build Category List
pcs_RecursiveCategoryList pcv_IdRootCategory,"1",pcv_ExistingCats

conn.Close
Set conn = Nothing
'/////////////////////////////////////////////////////////////////////
'// END: Generate Category List Recursively
'/////////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////////
'// START: OUTPUT XML
'/////////////////////////////////////////////////////////////////////

Output = XMLDoc.xml

'// Clean up
Set XMLDoc = Nothing
Set XMLRoot = Nothing
Set XMLRow = Nothing
If Err.Number > 0 Then Output = "<error>" & pcv_GenXMLError() & "</error>"

'// Print XML
Response.ContentType = "text/xml"
Response.Write "<?xml version='1.0' encoding='iso-8859-1' ?>" & Output
Response.End

'/////////////////////////////////////////////////////////////////////
'// END: OUTPUT XML
'/////////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////////
'// START: Utility Methods
'/////////////////////////////////////////////////////////////////////

'// Recursive Category Listing
Sub pcs_RecursiveCategoryList(idRootCat, LoopMode, myCats)
	
	Dim pcv_cats3
	nTotalRecs=0
	query = "SELECT idCategory,categoryDesc,idParentCategory,pccats_BreadCrumbs FROM [categories]"
	if isNumeric(idRootCat) then
		query=query&" WHERE idParentCategory="&idRootCat
	end if
	if LoopMode="1" then
		query=query&" OR idCategory="&idRootCat
	end if
	Set rs = Server.CreateObject("ADODB.Recordset")
	set rs=conn.execute(query)		
	IF NOT rs.EOF THEN		
		pcArray_NextTier=rs.GetRows()
		nTotalRecs=1
		pcv_cats3=pcArray_NextTier
	END IF
	Set rs = nothing
	
	If nTotalRecs>0 Then	
		dim tmp_A
		tmp_A=split(myCats,",")

		dim k		
		For k=lbound(pcv_cats3,2) to ubound(pcv_cats3,2)
			
			' Get field values
			x_idCategory = pcv_cats3(0,k)
			If NOT IsNull(x_idCategory) Then x_idCategory = CLng(x_idCategory)
			x_categoryDesc = pcv_cats3(1,k)
			If NOT IsNull(x_categoryDesc) Then x_categoryDesc = CStr(x_categoryDesc)
			x_idParentCategory = pcv_cats3(2,k)
			If NOT IsNull(x_idParentCategory) Then x_idParentCategory = CLng(x_idParentCategory)
			x_pcCats_BreadCrumbs = pcv_cats3(3,k)
			If NOT IsNull(x_pcCats_BreadCrumbs) Then x_pcCats_BreadCrumbs = CStr(x_pcCats_BreadCrumbs)
			
			if nRecActual = 0 then
				' Create XML nodes
				Set XMLRow = XMLDoc.createElement("category")
				XMLRoot.appendChild XMLRow
				Call pcv_AddNode(XMLRow, "idCategory", 0, 3)
				Call pcv_AddNode(XMLRow, "categoryDesc", " Any ", 3)
				Call pcv_AddNode(XMLRow, "idParentCategory", 0, 3)
				Call pcv_AddNode(XMLRow, "pcCats_BreadCrumbs", "", 3)
				Call pcv_AddNode(XMLRow, "pcSelected", "", 3)
			End If
		
			nRecActual = nRecActual + 1
			if request("CP")="1" then
				pcv_SkipNode="1"
				query = "SELECT * FROM categories_products WHERE idCategory="&x_idCategory&";"
				Set rsCPObj = Server.CreateObject("ADODB.Recordset")
				set rsCPObj=conn.execute(query)		
				IF NOT rsCPObj.EOF THEN	
					pcv_SkipNode="0"
				End if	
			Else
				pcv_SkipNode="0"
			End if
			
			If request("CP")="3" then
				pcv_SkipNode="1"
				query = "SELECT categories_products.idProduct, categories_products.idCategory, categories_products.POrder FROM categories_products WHERE  (categories_products.idCategory = "&x_idCategory&") AND categories_products.idProduct IN (select options_optionsGroups.idProduct from options_optionsGroups);"
				Set rsCPObj = Server.CreateObject("ADODB.Recordset")
				set rsCPObj=conn.execute(query)		
				IF NOT rsCPObj.EOF THEN	
					pcv_SkipNode="0"
				End if	
			End if

			if pcv_SkipNode="0" then
				' Create XML nodes
				Set XMLRow = XMLDoc.createElement("category")
				XMLRoot.appendChild XMLRow
				
				if instr(x_categoryDesc, "&") then
					x_categoryDesc=replace(x_categoryDesc, "&", "and")
					x_categoryDesc=replace(x_categoryDesc, "andamp;", "and")
				end if
				if instr(x_categoryDesc, """") then
					x_categoryDesc=replace(x_categoryDesc, """", "`")
				end if
		
				Dim tmp2, x_pcSelected, l
				tmp2=0
				For l=lbound(tmp_A) to ubound(tmp_A)
					if trim(tmp_A(l))<>"" then
					if clng(tmp_A(l))=clng(pcv_cats3(0,k)) then
						tmp2=1
						exit for
					end if
					end if
				Next
				if tmp2=1 then
					x_pcSelected=" selected"
				else
					x_pcSelected = ""
				end if
		
				Call pcv_AddNode(XMLRow, "idCategory", x_idCategory, 3)
				Call pcv_AddNode(XMLRow, "categoryDesc", x_categoryDesc, 3)
				Call pcv_AddNode(XMLRow, "idParentCategory", x_idParentCategory, 3)
		
				pcv_BC=pcf_catGetParent(x_idCategory,x_idParentCategory,x_pcCats_BreadCrumbs)
		
				if instr(pcv_BC, "&") then
					pcv_BC=replace(pcv_BC, "&", "and")
					pcv_BC=replace(pcv_BC, "andamp;", "and")
				end if
				if instr(pcv_BC, """") then
					pcv_BC=replace(pcv_BC, """", "&quot;")
				end if
		
				Call pcv_AddNode(XMLRow, "pcCats_BreadCrumbs", pcv_BC, 3)
				Call pcv_AddNode(XMLRow, "pcSelected", x_pcSelected, 3)
			End If
			'// Break the Infinite Loop by excluding the Parent Node
			If (cInt(x_idParentCategory) <> cInt(x_idCategory)) AND (cInt(x_idCategory) <> cInt(pcv_IdRootCategory)) Then
				pcs_RecursiveCategoryList x_idCategory,"0","0,"
			End If
			
		Next		
		
	End If	
	
End Sub

Function pcf_catGetParent(pcv_idcategory,pcv_parentCategory,pcv_BreadCrumbs)	
	Dim tmp_ParentText,tmp_C,tmp_D,p
	tmp_ParentText=""
	if isNULL(pcv_BreadCrumbs) then
		pcv_BreadCrumbs=""
	end if
	IF trim(pcv_BreadCrumbs)<>"" THEN
		tmp_C=split(pcv_BreadCrumbs,"|,|")
		For p=lbound(tmp_C) to ubound(tmp_C)
			if trim(tmp_C(p))<>"" then
			tmp_D=split(tmp_C(p),"||")
			if (clng(tmp_D(0))<>clng(pcv_idcategory)) AND (clng(tmp_D(0))<>"1") then
				if tmp_ParentText="" then
					tmp_ParentText="["
				else
					tmp_ParentText=tmp_ParentText & "/"
				end if
				tmp_ParentText=tmp_ParentText & tmp_D(1)
			end if
			end if
		Next
		if tmp_ParentText<>"" then
			tmp_ParentText=tmp_ParentText & "]"
		else
			tmp_ParentText=""
		end if
	ELSE
		pcv_tmpParent=""
		tmpParent=""
		if pcv_parentCategory="1" then
		else 
			pcv_tmpParent = pcf_FindParent(pcv_parentCategory)
			if pcv_tmpParent<>"" then
				pcv_tmpParent="[" & pcv_tmpParent & "]"
			end if
		end if		
		if pcv_tmpParent="" then
			pcv_tmpParent=""
		end if
		tmp_ParentText=pcv_tmpParent
	END IF
	pcf_catGetParent=tmp_ParentText
End Function

Function pcf_FindParent(idCat)
	Dim k
	if clng(idCat)<>1 then
	For k=0 to intCount
		if (clng(pcArrP(0,k))=clng(idCat)) and (clng(pcArrP(0,k))<>1)	then
			if tmpParent<>"" then
			tmpParent="/" & tmpParent
			end if
			tmpParent=pcArrP(1,k) & tmpParent
			pcf_FindParent(pcArrP(2,k))
			exit for
		end if
	Next
	pcf_FindParent=tmpParent
	end if
End function

Sub pcs_GetParentList()
	query="SELECT categories.idcategory,categories.categorydesc,categories.idParentCategory FROM categories ORDER BY categories.idcategory asc;"
	set rstemp1=conn.execute(query)
	if not rstemp1.eof then
		pcArrP=rstemp1.getRows()
		intCount=ubound(pcArrP,2)
	end if
	set rstemp1=nothing
End Sub

'/////////////////////////////////////////////////////////////////////
'// END: Utility Methods
'/////////////////////////////////////////////////////////////////////
%>