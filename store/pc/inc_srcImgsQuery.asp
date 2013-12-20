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

dim submit,submit2
query=""
submit=request("Submit")
submit2=request("Submit2")
if submit2<>"" then
    query=session("srcImg_query")
end if

Function CreateQuery(Desc,keynum)
    Dim tmpStr,keydesc

	tmpStr=""

	Select Case keynum
		Case 1: keydesc="pcImgDir_Name"
		Case 2: keydesc="pcImgDir_Type"
		Case 3: keydesc="pcImgDir_Size"
		Case 4: keydesc="pcImgDir_DateUploaded"
	End Select

	if keydesc="pcImgDir_DateUploaded" then
	    if scDB="Access" then
	        tmpStr="(" & keydesc & " >= #"&Desc&"#)"
	    else
	        tmpStr="(" & keydesc & " >= '"&Desc&"')"
	    end if
	elseif keydesc="pcImgDir_Size" then
		tmpStr="(" & keydesc & " <= "&Desc&")"
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
		Case "1": strORD1="pcImageDirectory.pcImgDir_ID ASC"
		Case "2": strORD1="pcImageDirectory.pcImgDir_Name ASC"
		Case "3": strORD1="pcImageDirectory.pcImgDir_Name DESC"
		Case "4": strORD1="pcImageDirectory.pcImgDir_Type ASC"
		Case "5": strORD1="pcImageDirectory.pcImgDir_Type DESC"
		Case "6": strORD1="pcImageDirectory.pcImgDir_Size ASC"
		Case "7": strORD1="pcImageDirectory.pcImgDir_Size DESC"
		Case "8": strORD1="pcImageDirectory.pcImgDir_DateUploaded ASC"
		Case "9": strORD1="pcImageDirectory.pcImgDir_DateUploaded DESC"
	End Select
Else
	strORD="1"
	strORD1="pcImageDirectory.pcImgDir_ID ASC"
End If

' create sql statement
	query1=""
	query2=""
	if request("key1")<>"" then
		tmpKey=trim(request("key1"))
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
		tmpKey=trim(request("key2"))
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
		tmpKey=trim(request("key3"))
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
		tmpKey=trim(request("key4"))
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,4)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	

    if query="" then
	    query="SELECT pcImgDir_ID,pcImgDir_Name,pcImgDir_Type,pcImgDir_Size,pcImgDir_DateUploaded from pcImageDirectory " & session("srcImg_from")
	    if query1<>"" then
		    query=query & " WHERE " & query1
		    if session("srcImg_where")<>"" then
			    query=query & " AND " & session("srcImg_where")
		    end if
	    else
		    if session("srcImg_where")<>"" then
			    query=query & " WHERE " & session("srcImg_where")
		    end if
	    end if
	    query=query&" ORDER BY "& strORD1
	    
	    if submit<>"" then
            session("srcImg_query")=query
        end if
    end if
    
%>
