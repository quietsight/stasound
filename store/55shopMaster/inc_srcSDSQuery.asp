<%
iPageSize=getUserInput(request("resultCnt"),10)
if iPageSize="" then
	iPageSize=10
end if
if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=server.HTMLEncode(request("iPageCurrent"))
end if

src_PageType=request("src_PageType")
if src_PageType="" then
	src_PageType="0"
end if
if src_PageType="0" then
	pcv_Table="pcSupplier"
else
	pcv_Table="pcDropShipper"
end if

Function CreateQuery(Desc,keynum)
Dim m
Dim tmpStr,keywordArray,keylink,keydesc

	tmpStr=""

	Select Case keynum
		Case 1: keydesc=pcv_Table & "_FirstName"
		Case 2: keydesc=pcv_Table & "_LastName"
		Case 3: keydesc=pcv_Table & "_Company"
		Case 4: keydesc=pcv_Table & "_Email"
		Case 5: keydesc=pcv_Table & "_Phone"
	End Select

	if Instr(Desc," AND ")>0 then
		keywordArray=split(Desc," AND ")
		keylink=" AND "
	else
	if Instr(Desc,",")>0 then
		keywordArray=split(Desc,",")
		keylink=" OR "
	else
		if Instr(Desc," OR ")>0 then
			keywordArray=split(Desc," OR ")
			keylink=" OR "
		else
			keywordArray=split(Desc,"***")
			keylink=" OR "
		end if
	end if
	end if

			
	For m=lbound(keywordArray) to ubound(keywordArray)
	if trim(keywordArray(m))<>"" then
		if tmpStr<>"" then
		tmpStr=tmpStr & keylink
		end if
		tmpStr=tmpStr & "(" & keydesc & " like '%"&trim(keywordArray(m))&"%')"
	end if
	Next
	
	if tmpStr<>"" then
		tmpStr="(" & tmpStr & ")"
	else
		tmpStr="(" & keydesc & " like '%"&Desc&"%')"
	end if

CreateQuery=tmpStr
End Function

strORD=getUserInput(request("order"),4)
if NOT isNumeric(strORD) then
	strORD=1
end if

if strORD<>"" then
	Select Case StrORD
		Case "1": strORD1=pcv_Table & "s." & pcv_Table & "_LastName ASC"
		Case "2": strORD1=pcv_Table & "s." & pcv_Table & "_ID ASC"
		Case "3": strORD1=pcv_Table & "s." & pcv_Table & "_ID DESC"
		Case "4": strORD1=pcv_Table & "s." & pcv_Table & "_LastName ASC"
		Case "5": strORD1=pcv_Table & "s." & pcv_Table & "_LastName DESC"
	End Select
Else
	strORD="1"
	strORD1="customers.lastname ASC"
End If

' create sql statement
	query1=""
	query2=""
	if request("key1")<>"" then
		tmpKey=request("key1")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,1)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key2")<>"" then
		tmpKey=request("key2")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,2)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key3")<>"" then
		tmpKey=request("key3")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,3)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key4")<>"" then
		tmpKey=request("key4")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,4)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key5")<>"" then
		tmpKey=request("key5")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,5)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query="SELECT " & pcv_Table & "_ID," & pcv_Table & "_FirstName," & pcv_Table & "_LastName," & pcv_Table & "_Company," & pcv_Table & "_Phone," & pcv_Table & "_Email"
	if src_PageType="1" then
		query=query & ",0 As IsDropShipper"
	else
		query=query&",pcSupplier_IsDropShipper As IsDropShipper"
	end if
	query=query & " FROM " & pcv_Table & "s " & session("srcSDS_from")
	if query1<>"" then
		query=query & " WHERE " & query1
		if session("srcSDS_where")<>"" then
			query=query & " AND " & session("srcSDS_where")
		end if
	else
		if session("srcSDS_where")<>"" then
			query=query & " WHERE " & session("srcSDS_where")
		end if
	end if
	if src_PageType="0" then
		query=query & "UNION (SELECT pcSupplier_ID,pcSupplier_FirstName,pcSupplier_LastName,pcSupplier_Company,pcSupplier_Phone,pcSupplier_Email"
		query=query & ",pcSupplier_IsDropShipper"
		query=query & " FROM pcSuppliers WHERE pcSupplier_IsDropShipper=1"
		if query1<>"" then
			query=query & " AND " & replace(query1,"pcDropShipper","pcSupplier")
		end if
		query=query & ")"
	end if
	
	query=query&" ORDER BY "& strORD1

%>
