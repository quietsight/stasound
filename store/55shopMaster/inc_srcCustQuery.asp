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

Function CreateQuery(Desc,keynum)
Dim m
Dim tmpStr,keywordArray,keylink,keydesc

	tmpStr=""

	Select Case keynum
		Case 1: keydesc="[Name]"
		Case 2: keydesc="LastName"
		Case 3: keydesc="customerCompany"
		Case 4: keydesc="email"
		Case 5: keydesc="city"
		Case 6: keydesc="phone"
		Case 7: keydesc="stateCode"
		Case 8: keydesc="state"
		Case 9: keydesc="zip"
		Case 10: keydesc="CountryCode"
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
		Case "1": strORD1="customers.lastname ASC"
		Case "2": strORD1="customers.idcustomer ASC"
		Case "3": strORD1="customers.idcustomer DESC"
		Case "4": strORD1="customers.lastname ASC"
		Case "5": strORD1="customers.lastname DESC"
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
	
	query2=""
	if request("key6")<>"" then
		tmpKey=request("key6")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,6)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key7")<>"" then
		tmpKey=request("key7")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,7)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key8")<>"" then
		tmpKey=request("key8")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,8)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key9")<>"" then
		tmpKey=request("key9")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,9)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key10")<>"" then
		tmpKey=request("key10")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,10)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		if request("key11")="1" then
			query1=query1 & "( NOT (" & query2 & "))"
		else
			query1=query1 & query2
		end if
	end if
	
	query="SELECT DISTINCT LastName,[name],customerCompany,phone,customerType,idcustomer,email FROM customers " & session("srcCust_from")
	if query1<>"" then
		query=query & " WHERE " & query1
		if session("srcCust_where")<>"" then
			query=query & " AND " & session("srcCust_where")
		end if
	else
		if session("srcCust_where")<>"" then
			query=query & " WHERE " & session("srcCust_where")
		end if
	end if
	query=query&" ORDER BY "& strORD1

%>
