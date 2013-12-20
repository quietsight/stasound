<%
'//************ START - Update Product Dates *******************
Sub updPrdCreatedDate(tmpID)
Dim queryQ,rsQ,tmpQ
Dim dtTodaysDate
	tmpQ="idproduct=" & tmpID
	
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if scDB="SQL" then
		queryQ="UPDATE Products SET pcprod_EnteredOn='" & dtTodaysDate & "' WHERE " & tmpQ
	else
		queryQ="UPDATE Products SET pcprod_EnteredOn=#" & dtTodaysDate & "# WHERE " & tmpQ
	end if
	set rsQ=Server.CreateObject("ADODB.Recordset")
	set rsQ=connTemp.execute(queryQ)
	set rsQ=nothing
End Sub

Sub updPrdEditedDate(tmpID)
Dim queryQ,rsQ,tmpQ
Dim dtTodaysDate	
	tmpQ="idproduct=" & tmpID
	
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if scDB="SQL" then
		queryQ="UPDATE Products SET pcProd_EditedDate='" & dtTodaysDate & "' WHERE " & tmpQ
	else
		queryQ="UPDATE Products SET pcProd_EditedDate=#" & dtTodaysDate & "# WHERE " & tmpQ
	end if
	set rsQ=Server.CreateObject("ADODB.Recordset")
	set rsQ=connTemp.execute(queryQ)
	set rsQ=nothing
End Sub
'//************ END - Update Product Dates *********************

'//************ START - Update Category Dates ******************
Sub updCatCreatedDate(tmpID,tmpQuery)
Dim queryQ,rsQ,tmpQ
Dim dtTodaysDate	
	tmpQ=tmpQuery
	if tmpQ="" then
		tmpQ="idcategory=" & tmpID
	end if
	
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if scDB="SQL" then
		queryQ="UPDATE Categories SET pcCats_CreatedDate='" & dtTodaysDate & "' WHERE " & tmpQ
	else
		queryQ="UPDATE Categories SET pcCats_CreatedDate=#" & dtTodaysDate & "# WHERE " & tmpQ
	end if
	set rsQ=Server.CreateObject("ADODB.Recordset")
	set rsQ=connTemp.execute(queryQ)
	set rsQ=nothing
End Sub

Sub updCatEditedDate(tmpID,tmpQuery)
Dim queryQ,rsQ,tmpQ
Dim dtTodaysDate	
	tmpQ=tmpQuery
	if tmpQ="" then
		tmpQ="idcategory=" & tmpID
	end if
	
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if scDB="SQL" then
		queryQ="UPDATE Categories SET pcCats_EditedDate='" & dtTodaysDate & "' WHERE " & tmpQ
	else
		queryQ="UPDATE Categories SET pcCats_EditedDate=#" & dtTodaysDate & "# WHERE " & tmpQ
	end if
	set rsQ=Server.CreateObject("ADODB.Recordset")
	set rsQ=connTemp.execute(queryQ)
	set rsQ=nothing
End Sub
'//************ END - Update Category Dates *******************

'//************ START - Update Customer Dates ******************
Sub updCustCreatedDate(tmpID)
Dim queryQ,rsQ,tmpQ
Dim dtTodaysDate	
	tmpQ="idcustomer=" & tmpID
	
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if scDB="SQL" then
		queryQ="UPDATE Customers SET pcCust_DateCreated='" & dtTodaysDate & "' WHERE " & tmpQ
	else
		queryQ="UPDATE Customers SET pcCust_DateCreated=#" & dtTodaysDate & "# WHERE " & tmpQ
	end if
	set rsQ=Server.CreateObject("ADODB.Recordset")
	set rsQ=connTemp.execute(queryQ)
	set rsQ=nothing
End Sub

Sub updCustEditedDate(tmpID)
Dim queryQ,rsQ,tmpQ
Dim dtTodaysDate
	tmpQ="idcustomer=" & tmpID
	
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	if scDB="SQL" then
		queryQ="UPDATE Customers SET pcCust_EditedDate='" & dtTodaysDate & "' WHERE " & tmpQ
	else
		queryQ="UPDATE Customers SET pcCust_EditedDate=#" & dtTodaysDate & "# WHERE " & tmpQ
	end if
	set rsQ=Server.CreateObject("ADODB.Recordset")
	set rsQ=connTemp.execute(queryQ)
	set rsQ=nothing
End Sub
'//************ END - Update Category Dates *******************
%>