<%@ CodePage=1252 LCID=1033 %>
<% Option Explicit %>
<!--#include file="../includes/storeconstants.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%
'//Modes
'1 = AddTaxPerPrd.asp
'3 = ModPrdOpta1.asp


Dim x_Db_Conn_Str
x_Db_Conn_Str = scDSN

' XML Encoding
Const EW_XML_ENCODING = "iso-8859-1"

' Common constants
Const MSXML_NOT_INSTALLED = "MSXML 3 or later not installed"
Const EW_IS_MSACCESS = False
Const EW_DB_QUOTE_START = "["

' Project related
Const EW_PROJECT_NAME = "CategoryProductsPcV4" ' Project Name

' Request parameters
Const EW_TABLE_REC_PER_PAGE = "pagesize"
Const EW_TABLE_PAGE_NUMBER = "pageno"
Const EW_TABLE_START_REC = "start"
Const EW_TABLE_BASIC_SEARCH = "search"
Const EW_TABLE_BASIC_SEARCH_TYPE = "searchtype"
Const EW_TABLE_SORT = "sort"
Const EW_TABLE_SORT_ORDER = "sortorder"

' Table level constants
Const EW_ROOT_TAG_NAME = "root"
Const EW_ROW_TAG_NAME = "row"
Const EW_SQL_SELECT = "SELECT DISTINCT categories_products.idProduct, categories_products.idCategory, products.description FROM (categories_products INNER JOIN products ON categories_products.idProduct = products.idProduct)"
Const EW_SQL_WHERE = ""
Const EW_SQL_GROUPBY = ""
Const EW_SQL_HAVING = ""
Const EW_SQL_ORDERBY = ""

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"

Dim conn, rs, sSql
Dim nDisplayRecs
Dim nStartRec, nStopRec, nTotalRecs, nRecCount, nRecActual
Dim sSrchAdvanced
Dim psearch
Dim psearchtype
Dim sSrchBasic
Dim sSrchWhere
Dim sDbWhere
Dim sOrderBy
Dim x_idProduct
Dim x_idCategory
Dim x_Description
Dim XMLDoc, XMLRoot, XMLRow, XMLField, Output, sXMLEncoding

if request("mode")="2" then
	nRecActual = 0
else
	nRecActual = 1
end if

' Output all records by default
nDisplayRecs = -1

' Set up records per page dynamically
SetUpDisplayRecs()

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open x_Db_Conn_Str

' Get Search Criteria for Advanced Search
SetUpAdvancedSearch()

' Get Search Criteria for Basic Search
SetUpBasicSearch()

' Build Search Criteria
If sSrchAdvanced <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchAdvanced & ")"
End If
If sSrchBasic <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchBasic & ")"
End If

' Build Filter condition
sDbWhere = "removed = 0"

if request("mode")="1" then
	If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
	sDbWhere = sDbWhere & "categories_products.idProduct NOT IN (select taxPrd.idProduct from taxPrd)"
end if

dim sbDbSelect
sbDbSelect = EW_SQL_SELECT

if request("mode")="3" then
	sbDbSelect = sbDbSelect & " INNER JOIN options_optionsGroups ON categories_products.idProduct = options_optionsGroups.idProduct"
end if

If sSrchWhere <> "" Then
	If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
	sDbWhere = sDbWhere & "(" & sSrchWhere & ")"
End If

' Set Up Sorting Order
sOrderBy = ""
SetUpSortOrder()

' Set up SQL
sSql = ew_BuildSql(sbDbSelect, EW_SQL_WHERE, EW_SQL_GROUPBY, EW_SQL_HAVING, EW_SQL_ORDERBY, sDbWhere, "products.description")

'Response.Write sSql: Response.End ' Uncomment to show SQL for debugging
' Set up Record Set

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open sSql, conn, 1, 2
nTotalRecs = rs.RecordCount
If nDisplayRecs <= 0 Then ' Display All Records
	nDisplayRecs = nTotalRecs
End If
nStartRec = 1
SetUpStartRec() ' Set Up Start Record Position

' Get MSXML object
Set XMLDoc = ew_GetMSXML()
If Not IsObject(XMLDoc) Then
	Response.Write "MSXML 3 or later not installed"
	Response.End
End If

' Create and append the root element
Set XMLRoot = XMLDoc.createElement(EW_ROOT_TAG_NAME)
XMLDoc.appendChild XMLRoot
On Error Resume Next
If nTotalRecs > 0 Then

	' Set the last record to display
	nStopRec = nStartRec + nDisplayRecs - 1
	nRecCount = nStartRec - 1
	If Not rs.Eof Then
		rs.MoveFirst
		rs.Move nStartRec - 1 ' Move to first record directly
	End If
	
	Do While (Not rs.Eof) And (nRecCount < nStopRec)
		if nRecActual = 0 then
			' Create XML nodes
			Set XMLRow = XMLDoc.createElement(EW_ROW_TAG_NAME)
			XMLRoot.appendChild XMLRow
			Set XMLField = ew_AddNode(XMLRow, "idProduct", 0, 3, 0, "")
			Set XMLField = ew_AddNode(XMLRow, "idCategory", 0, 3, 0, "")
			Set XMLField = ew_AddNode(XMLRow, "description", "Any", 3, 0, "")
		End If
	
		nRecCount = nRecCount + 1
		If CLng(nRecCount) >= CLng(nStartRec) Then
			nRecActual = nRecActual + 1

			' Get field values
			x_idProduct = ew_Conv(rs("idProduct"), 3)
			x_idCategory = ew_Conv(rs("idCategory"), 3)
			x_Description = ew_Conv(rs("description"), 203)

			' Create XML nodes
			Set XMLRow = XMLDoc.createElement(EW_ROW_TAG_NAME)
			XMLRoot.appendChild XMLRow
			Set XMLField = ew_AddNode(XMLRow, "idProduct", x_idProduct, 3, 0, "")
			Set XMLField = ew_AddNode(XMLRow, "idCategory", x_idCategory, 3, 0, "")
			Set XMLField = ew_AddNode(XMLRow, "description", x_Description, 3, 0, "")
		End If
		rs.MoveNext
	Loop
End If

' Close recordset and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
Output = XMLDoc.xml

' Clean up
Set XMLDoc = Nothing
Set XMLRoot = Nothing
Set XMLRow = Nothing
If Err.Number > 0 Then Output = "<error>" & ew_ReportError() & "</error>"

' Output XML
Response.ContentType = "text/xml"
sXMLEncoding = ""
If EW_XML_ENCODING <> "" Then sXMLEncoding = " encoding=""" & EW_XML_ENCODING & """"
Response.Write "<?xml version=""1.0""" & sXMLEncoding & "?>" & Output
Response.End

'-------------------------------------------------------------------------------
' Function SetUpDisplayRecs
' - Set up Number of Records displayed per page based on RecPerPage
' - Variables setup: nDisplayRecs

Sub SetUpDisplayRecs()
	Dim sWrk
	sWrk = Request(EW_TABLE_REC_PER_PAGE)
	If sWrk <> "" Then
		If IsNumeric(sWrk) And CInt(sWrk) <> 0 Then
			nDisplayRecs = CLng(sWrk)
		Else
			nDisplayRecs = -1 ' Display all records	
		End If
	Else
			nDisplayRecs = -1 ' Display all records
	End If

	' Start record
	sWrk = Request(EW_TABLE_START_REC)
	If sWrk <> "" Then
		If IsNumeric(sWrk) Then
			nStartRec = CLng(sWrk)
		Else
			nStartRec = 1 ' Non-numeric, Default
		End If
	Else
		nStartRec = 1 ' Default
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpAdvancedSearch
' - Set up Advanced Search parameter based on Request parameters
' - Variables setup: sSrchAdvanced

Sub SetUpAdvancedSearch()
	Call AdvancedSearchSQL("idProduct", "[idProduct]", 1, 0)
	Call AdvancedSearchSQL("idCategory", "[idCategory]", 1, 0)
	Call AdvancedSearchSQL("Description", "[Description]", 1, 0)
End Sub

' Check if search operator is allowed
Function IsValidOpr(Opr, FldType)
	IsValidOpr = (Opr = "=" Or Opr = "<" Or Opr = "<=" Or _
		Opr = ">" Or Opr = ">=" Or Opr = "<>")
	If FldType = 3 Then
		IsValidOpr = IsValidOpr Or Opr = "LIKE" Or Opr = "NOT LIKE"
	End If
End Function

' Build SQL for Advanced Search
Sub AdvancedSearchSQL(FldVar, FldExp, FldType, FldDtFormat)
	Dim FldVal, FldOpr, FldCond, FldVal2, FldOpr2, sSrchStr, IsValidValue
	sSrchStr = ""
	FldVal = Request("x_" & FldVar)
	FldOpr = Request("z_" & FldVar)
	FldCond = Request("v_" & FldVar)
	FldVal2 = Request("y_" & FldVar)
	FldOpr2 = Request("w_" & FldVar)
	FldOpr = UCase(Trim(FldOpr))
	If FldOpr = "" Then FldOpr = "="
	If FldOpr = "BETWEEN" Then
		IsValidValue = (FldType <> 1) Or _
			(FldType = 1 And IsNumeric(FldVal) And IsNumeric(FldVal2))
		If FldVal <> "" And FldVal2 <> "" And IsValidValue Then
			If FldType = 2 Then
				FldVal = ew_UnFormatDateTime(FldVal, FldDtFormat)
				FldVal2 = ew_UnFormatDateTime(FldVal2, FldDtFormat)
			End If
			sSrchStr = sSrchStr & FldExp & " BETWEEN " & ew_QuotedValue(FldVal, FldType) & _
				" AND " & ew_QuotedValue(FldVal2, FldType)
		End If
	ElseIf FldOpr = "IS NULL" Or FldOpr = "IS NOT NULL" Then
		sSrchStr = sSrchStr & FldExp & " " & FldOpr
	Else
		IsValidValue = (FldType <> 1) Or _
			(FldType = 1 And IsNumeric(FldVal))
		If FldVal <> "" And IsValidValue And IsValidOpr(FldOpr, FldType) Then
			If FldType = 2 Then FldVal = ew_UnFormatDateTime(FldVal, FldDtFormat)
			sSrchStr = sSrchStr & FldExp & FldOpr & " " & ew_QuotedValue(FldVal, FldType)
		End If
		FldOpr2 = UCase(Trim(FldOpr2))
		If FldOpr2 = "" Then FldOpr2 = "="
		IsValidValue = (FldType <> 1) Or _
			(FldType = 1 And IsNumeric(FldVal2))
		If FldVal2 <> "" And IsValidValue And IsValidOpr(FldOpr2, FldType) Then
			If FldType = 2 Then	FldVal2 = ew_UnFormatDateTime(FldVal2, FldDtFormat)
			If sSrchStr <> "" Then sSrchStr = sSrchStr & " " & ew_IIf(FldCond="OR", "OR", "AND") & " "
			sSrchStr = sSrchStr & FldExp & FldOpr2 & " " & ew_QuotedValue(FldVal2, FldType)
		End If
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If
End Sub

'-------------------------------------------------------------------------------
' Function BasicSearchSQL
' - Build WHERE clause for a keyword

Function BasicSearchSQL(Keyword)
	Dim sKeyword
	sKeyword = ew_AdjustSql(Keyword)
	BasicSearchSQL = ""
	If Right(BasicSearchSQL, 4) = " OR " Then BasicSearchSQL = Left(BasicSearchSQL, Len(BasicSearchSQL)-4)
End Function

'-------------------------------------------------------------------------------
' Function SetUpBasicSearch
' - Set up Basic Search parameter based on pSearch & pSearchType
' - Variables setup: sSrchBasic

Sub SetUpBasicSearch()
	Dim arKeyword, sKeyword
	psearch = Request(EW_TABLE_BASIC_SEARCH)
	psearchtype = Request(EW_TABLE_BASIC_SEARCH_TYPE)
	psearchtype = UCase(Trim(psearchtype))
	If psearch <> "" Then
		If psearchtype = "AND" Or psearchtype = "OR" Then
			While InStr(psearch, "  ") > 0
				psearch = Replace(psearch, "  ", " ")
			Wend
			arKeyword = Split(Trim(psearch), " ")
			For Each sKeyword In arKeyword
				sSrchBasic = sSrchBasic & "(" & BasicSearchSQL(sKeyword) & ") " & psearchtype & " "
			Next
		Else
			sSrchBasic = BasicSearchSQL(psearch)
		End If
	End If
	If Right(sSrchBasic, 4) = " OR " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-4)
	If Right(sSrchBasic, 5) = " AND " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-5)
End Sub

'-------------------------------------------------------------------------------
' Function SetUpSortOrder

Sub SetUpSortOrder()
	Dim sOrder, sSortField
	Dim dFld, sSort, arSort, sSortOrder, arSortOrder
	Dim str, i

	' Check for an Order parameter
	If Request(EW_TABLE_SORT).Count > 0 Then
		sSort = Request(EW_TABLE_SORT)
		arSort = Split(sSort, ",")
		sSortOrder = Request(EW_TABLE_SORT_ORDER)
		arSortOrder = Split(sSortOrder, ",")
		str = ""

		' Sortable fields
		Set dFld = Server.CreateObject("Scripting.Dictionary")
		dFld.Add "idProduct", "[idProduct]"
		dFld.Add "idCategory", "[idCategory]"
		dFld.Add "Description", "[Description]"

	' Build the ORDER BY clause
	For i = 0 to UBound(arSort)
		If dFld.Exists(arSort(i)) Then
			sSortField = dFld(arSort(i))
			If str <> "" Then str = str & ", "
			str = str & sSortField
			If IsArray(arSortOrder) Then
				If i <= UBound(arSortOrder) Then
			 		If UCase(arSortOrder(i)) = "DESC" Then
						str = str & " DESC"
					End If
				End If
			End if
		End If
	Next
	Set dFld = Nothing
	End If
	sOrderBy = str
	If sOrderBy = "" Then
		If EW_SQL_ORDERBY <> "" Then
			sOrderBy = EW_SQL_ORDERBY
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpStartRec
' - Set up Starting Record parameters
' - Variables setup: nStartRec

Sub SetUpStartRec()
	Dim nPageNo
	If nDisplayRecs = 0 Then Exit Sub

	' Check for a START parameter
	If Request(EW_TABLE_START_REC).Count > 0 Then
		nStartRec = Request(EW_TABLE_START_REC)
	ElseIf Request(EW_TABLE_PAGE_NUMBER).Count > 0 Then
		nPageNo = Request(EW_TABLE_PAGE_NUMBER)
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
					nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
		Else
			nStartRec = 1 'Default
		End If
	Else
		nStartRec = 1 'Default
	End If

	' Check if correct start record counter
	If Not IsNumeric(nStartRec) Or nStartRec = "" Then ' Avoid invalid start record counter
		nStartRec = 1 ' Reset start record counter
	ElseIf CLng(nStartRec) > CLng(nTotalRecs) Then ' Avoid starting record > total records
		nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 ' point to last page first record	
	End If
End Sub
%>
<%
' Get MSXML object
Function ew_GetMSXML()
	Dim i
	For i = 3 to 6
		On Error Resume Next
		Set ew_GetMSXML = Server.CreateObject("Msxml2.DOMDocument." & CStr(i) & ".0")
		If Err.Number = 0 Then Exit For
	Next
End Function

' Function to format date time
' ANamedFormat = 0-8, where 0-4 same as VBScript
' 5 = "yyyy-mm-dd" (where "-" = date separator)
' 6 = "mm-dd-yyyy" (where "-" = date separator)
' 7 = "dd-mm-yyyy" (where "-" = date separator)
' 8 = Short Date & " " & Short Time
' 9 = "yyyymmdd HH:MM:SS"
' 10 = "mmddyyyy HH:MM:SS"
' 11 = "ddmmyyyy HH:MM:SS"
' 12 = RFC822 format

' Unformat date time based on format type
Function ew_UnFormatDateTime(ADate, ANamedFormat)
	Dim arDateTime, arDate
	ADate = Trim(ADate & "")
	While Instr(ADate, "  ") > 0
		ADate = Replace(ADate, "  ", " ")
	Wend
	arDateTime = Split(ADate, " ")
	If UBound(arDateTime) < 0 Then
		ew_UnFormatDateTime = ADate
		Exit Function
	End If
	If ANamedFormat = 0 And IsDate(ADate) Then
		ew_UnFormatDateTime = Year(arDateTime(0)) & "/" & Month(arDateTime(0)) & "/" & Day(arDateTime(0))
		If UBound(arDateTime) > 0 Then
			ew_UnFormatDateTime = ew_UnFormatDateTime & " " & arDateTime(1)
		End If
	Else
		arDate = Split(arDateTime(0), "/")
		If UBound(arDate) = 2 Then
			ew_UnFormatDateTime = arDateTime(0)
			If ANamedFormat = 6 Or ANamedFormat = 10 Then ' mmddyyyy
				If Len(arDate(0)) <= 2 And Len(arDate(1)) <= 2 And Len(arDate(2)) <= 4 Then
					ew_UnFormatDateTime = arDate(2) & "/" & arDate(0) & "/" & arDate(1)
				End If
			ElseIf (ANamedFormat = 7 Or ANamedFormat = 11) Then ' ddmmyyyy
				If Len(arDate(0)) <= 2 And Len(arDate(1)) <= 2 And Len(arDate(2)) <= 4 Then
					ew_UnFormatDateTime = arDate(2) & "/" & arDate(1) & "/" & arDate(0)
				End If
			ElseIf ANamedFormat = 5 Or ANamedFormat = 9 Then ' yyyymmdd
				If Len(arDate(0)) <= 4 And Len(arDate(1)) <= 2 And Len(arDate(2)) <= 2 Then
					ew_UnFormatDateTime = arDate(0) & "/" & arDate(1) & "/" & arDate(2)
				End If
			End If
			If UBound(arDateTime) > 0 Then
				If IsDate(arDateTime(1)) Then ' Is time
					ew_UnFormatDateTime = ew_UnFormatDateTime & " " & arDateTime(1)
				End If
			End If
		Else
			ew_UnFormatDateTime = ADate
		End If
	End If
End Function

' IIf function
Function ew_IIf(cond, v1, v2)
	On Error Resume Next
	If CBool(cond) Then
		ew_IIf = v1
	Else
		ew_IIf = v2
	End If
End Function

' Convert different data type value
Function ew_Conv(v, t)

	Select Case t
		' adBigInt/adUnsignedBigInt/adInteger/adUnsignedInt
		Case 20, 21, 3, 19
			If IsNull(v) Then	ew_Conv = Null	Else ew_Conv = CLng(v)
		' adSmallInt/adTinyInt/adUnsignedTinyInt/adUnsignedSmallInt
		Case 2, 16, 17, 18
			If IsNull(v) Then	ew_Conv = Null	Else ew_Conv = CInt(v)
		' adSingle
		Case 4
			If IsNull(v) Then	ew_Conv = Null	Else ew_Conv = CSng(v)
		' adDouble/adNumeric
		Case 5, 131
			If IsNull(v) Then	ew_Conv = Null	Else ew_Conv = CDbl(v)
		' adCurrency
		Case 6
			If IsNull(v) Then	ew_Conv = Null	Else ew_Conv = CCur(v)
		' Const adBinary/adVarBinary/adLongVarBinary
		Case 128, 204, 205
			ew_Conv = "Binary" ' Not supported
		Case Else
			If IsNull(v) Then	ew_Conv = Null	Else ew_Conv = CStr(v)
	End Select

End Function

' Get SQL field value quote char
Function ew_QuotedValue(Value, FldType)

	Select Case FldType
	
	Case 3
		ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"

	Case 5
		If EW_IS_MSACCESS Then
			ew_QuotedValue = "{guid " & ew_AdjustSql(Value) & "}"
		Else
			ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
		End If

	Case 2
		If EW_IS_MSACCESS Then
				ew_QuotedValue = "#" & ew_AdjustSql(Value) & "#"
		Else
				ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
	End If

	Case Else
		ew_QuotedValue = Value

	End Select
    
End Function

' Encode name for XML tag name
Function ew_AddNode(ToNode, Name, Value, NodeType, NullType, NullValue)
	Dim Element, sValue
	
	sValue = Null
	Set Element = Nothing

	If IsNull(Value) Then
		Select Case NullType
			Case 0
				sValue = ""
			Case 1 ' Skip
			Case 2
				sValue = NullValue
			Case 3
				If NodeType = 3 Or NodeType = 4 Then
					Set Element = XMLDoc.createElement(Name)
					Element.SetAttribute "xsi:nil", "true"
					ToNode.appendChild Element
					Exit Function
				End If
		End Select
	Else
		sValue = CStr(Value)
	End If

	If Not IsNull(sValue) Then
		If NodeType = 3 Then
			Set Element = XMLDoc.createElement(Name)
			ToNode.appendChild Element
			Element.appendChild XMLDoc.createTextNode(sValue)
		ElseIf NodeType = 4 Then
			Set Element = XMLDoc.createElement(Name)
			ToNode.appendChild Element
			Element.appendChild XMLDoc.createCDATASection(sValue)
		ElseIf NodeType = 2 Then
			ToNode.SetAttribute Name, sValue
		End If
	End If
	
	If IsObject(Element) Then Set ew_AddNode = Element

End Function

' Function to build SQL
Function ew_BuildSql(sSelect, sWhere, sGroupBy, sHaving, sOrderBy, sFilter, sSort)

	Dim sSql, sDbWhere, sDbOrderBy

	sDbWhere = sWhere
	If sDbWhere <> "" Then
		sDbWhere = "(" & sDbWhere & ")"
	End If
	If sFilter <> "" Then
		If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
		sDbWhere = sDbWhere & "(" & sFilter & ")"
	End If	
	sDbOrderBy = sOrderBy
	If sSort <> "" Then
		sDbOrderBy = sSort
	End If
	sSql = sSelect
	If sDbWhere <> "" Then
		sSql = sSql & " WHERE " & sDbWhere
	End If
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If
	If sDbOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sDbOrderBy
	End If

	ew_BuildSql = sSql

End Function

' Report error
Function ew_ReportError()
	ew_ReportError = "Error Number: " & Err.Number & "; " & _
		"Description: " & Err.Description & "; Source: " & Err.Source
End Function
%>