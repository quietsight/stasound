<%
Function createGuid()
	Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
	tg = TypeLib.Guid
	createGuid = left(tg, len(tg)-2)
	createGuid = replace(createGuid,"{","")
	createGuid = replace(createGuid,"}","")
	Set TypeLib = Nothing
End Function

pcCartArray=Session("pcCartSession")
pcCartIndex=Session("pcCartIndex")

tmpCartHasPrds=0

for f=1 to pcCartIndex
	if pcCartArray(f,10)=0 then
		tmpCartHasPrds=1
		exit for
	end if
next

pcv_HasNewCart=0
pcv_MoveToReg=0

IF tmpCartHasPrds=1 THEN
	HasSavedCart=0
	IDSC=0
	pcv_strSaveCart=getUserInput(Request("SaveCart"),1)
	pcv_SavedCartName=getUserInput(Request("SavedCartName"),100)
	If pcv_strSaveCart="1" Then
		tmpGUID=""
	Else
		tmpGUID=getUserInput(Request.Cookies("SavedCartGUID"),0)
	End If
	IF tmpGUID<>"" THEN
		if session("IDCustomer")<>"" AND session("IDCustomer")<>"0" then
			tmpIDCust=session("IDCustomer")
		else
			tmpIDCust=0
		end if
		if SaveCustLogin=1 then
			query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
		else
			query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "' AND IDCustomer=" & tmpIDCust & ";"
		end if
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			IDSC=rsQ("SavedCartID")
			HasSavedCart=1
		end if
		set rsQ=nothing
		if HasSavedCart=1 then
			if tmpIDCust>"0" then
				query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "' AND IDCustomer=0;"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					pcv_MoveToReg=1
				end if
			end if
			dtTodaysDate=Date()
			if SQL_Format="1" then
				dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
			else
				dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
			end if
			if session("IDCustomer")<>"" AND session("IDCustomer")<>"0" then
				tmpIDCust=session("IDCustomer")
			else
				tmpIDCust=0
			end if
			if scDB="SQL" then
				query="UPDATE pcSavedCarts SET SavedCartDate='" & dtTodaysDate & "',IDCustomer=" & tmpIDCust & " WHERE SavedCartID=" & IDSC & ";"
			else
				query="UPDATE pcSavedCarts SET SavedCartDate=#" & dtTodaysDate & "#,IDCustomer=" & tmpIDCust & " WHERE SavedCartID=" & IDSC & ";"
			end if
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
	END IF
	IF HasSavedCart=0 THEN
		Response.Cookies("SavedCartGUID")=""
		NewGuid=0
		Do While NewGuid=0
			tmpGUID=createGuid()
			query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
			set rsQ=connTemp.execute(query)
			if rsQ.eof then
				NewGuid=1
			end if
			set rsQ=nothing
		Loop
		
		'// This is the default name for the shopping cart.
		'// Customers can rename it from "CustSavedCarts.asp"
		tmpSaveName=dictLanguage.Item(Session("language") & "_CustSavedCarts_7") & " " & Now()
		
		if session("IDCustomer")<>"" AND session("IDCustomer")<>"0" then
			tmpIDCust=session("IDCustomer")
		else
			tmpIDCust=0
		end if
		
		dtTodaysDate=Date()
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
		end if
		
		if len(pcv_SavedCartName)>0 then
			tmpSaveName=pcv_SavedCartName
		end if
		
		if scDB="SQL" then
			query="INSERT INTO pcSavedCarts (SavedCartGUID,SavedCartDate,SavedCartName,IDCustomer) VALUES ('" & tmpGUID & "','" & dtTodaysDate & "','" & tmpSaveName & "'," & tmpIDCust & ");"
		else
			query="INSERT INTO pcSavedCarts (SavedCartGUID,SavedCartDate,SavedCartName,IDCustomer) VALUES ('" & tmpGUID & "',#" & dtTodaysDate & "#,'" & tmpSaveName & "'," & tmpIDCust & ");"
		end if
		
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			IDSC=rsQ("SavedCartID")
			Response.Cookies("SavedCartGUID")=tmpGUID
			Response.Cookies("SavedCartGUID").Expires=Date()+365
			HasSavedCart=1
			pcv_HasNewCart=1
		end if
		set rsQ=nothing
	END IF
	IF HasSavedCart=1 THEN
	
		query="SELECT pcSCStatID FROM pcSavedCartStatistics WHERE pcSCMonth=" & Month(Date()) & " AND pcSCYear=" & Year(Date()) & ";"
		set rsQ=connTemp.execute(query)
		SCnewmonth=0
		PreviousMonthID=0
		if rsQ.eof then
			SCnewmonth=1
			query="SELECT pcSCStatID FROM pcSavedCartStatistics ORDER BY pcSCStatID DESC;"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				PreviousMonthID=rsQ("pcSCStatID")
			end if
			set rsQ=nothing
		end if
		if SCnewmonth=1 then
			if PreviousMonthID>"0" then
				query="SELECT TOP 10 idProduct,pcSPS_SavedTimes FROM pcSavedPrdStats WHERE pcSPS_SavedTimes>0 ORDER BY pcSPS_SavedTimes DESC;"
				set rsQ=connTemp.execute(query)
				SCPrds=""
				if not rsQ.eof then
					iCount=0
					Do while (not rsQ.eof) AND (iCount<10)
						iCount=iCount+1
						SCPrds=SCPrds & rsQ("idproduct") & "|*|" & rsQ("pcSPS_SavedTimes") & "|$|"
						rsQ.MoveNext
					Loop
				end if
				set rsQ=nothing
				
				query="UPDATE pcSavedCartStatistics SET pcSCTopPrds='" & SCPrds & "' WHERE pcSCStatID=" & PreviousMonthID & ";"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
			end if
		
			query="DELETE FROM pcSavedPrdStats;"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
	
		query="SELECT SCArray0 FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			tmpArr=rsQ.getRows()
			set rsQ=nothing
			tmpintCount=ubound(tmpArr,2)
			For k=0 to tmpintCount
				query="SELECT pcSPS_SavedTimes FROM pcSavedPrdStats WHERE idproduct=" & tmpArr(0,k) & ";"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("pcSPS_SavedTimes")>"0" then
						query="UPDATE pcSavedPrdStats SET pcSPS_SavedTimes=pcSPS_SavedTimes-1 WHERE idproduct=" & tmpArr(0,k) & ";"
						set rsQ=connTemp.execute(query)
						set rsQ=nothing
					end if
				end if
				set rsQ=nothing
			Next
		end if
		set rsQ=nothing
	
		query="DELETE FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		
		Function fixstring(x)
			If Not isNULL(x) Then
			fixstring=replace(x,"'","''")
			End If
		End function
		
		for f=1 to pcCartIndex
		if pcCartArray(f,10)=0 then
		if pcCartArray(f,30)="" OR IsNull(pcCartArray(f,30)) then
			pcCartArray(f,30)="0"
		end if
		if pcCartArray(f,31)="" OR IsNull(pcCartArray(f,31)) then
			pcCartArray(f,31)="0"
		end if
		'SB S
		query="INSERT INTO pcSavedCartArray (SavedCartID, SCArray0, SCArray1, SCArray2, SCArray3, SCArray4, SCArray5, SCArray6, SCArray7, SCArray8, SCArray9, SCArray10, SCArray11, SCArray12, SCArray13, SCArray14, SCArray15, SCArray16, SCArray17, SCArray18, SCArray19, SCArray20, SCArray21, SCArray22, SCArray23, SCArray24, SCArray25, SCArray26,SCArray27, SCArray28, SCArray29,SCArray30, SCArray31, SCArray32,SCArray33, SCArray34, SCArray35, SCArray36, SCArray37, SCArray38, SCArray39, SCArray40, SCArray41, SCArray42, SCArray43, SCArray44, SCArray45) "
		query=query& "VALUES ("& IDSC &","
		query=query& "'"& fixstring(pcCartArray(f,0)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,1)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,2)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,3)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,4)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,5)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,6)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,7)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,8)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,9)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,10)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,11)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,12)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,13)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,14)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,15)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,16)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,17)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,18)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,19)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,20)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,21)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,22)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,23)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,24)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,25)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,26)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,27)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,28)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,29)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,30)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,31)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,32)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,33)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,34)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,35)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,36)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,37)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,38)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,39)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,40)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,41)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,42)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,43)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,44)) &"', "
		query=query& "'"& fixstring(pcCartArray(f,45)) &"');"	
		'SB E        
		set rsQ=server.CreateObject("ADODB.RecordSet")
		set rsQ=conntemp.execute(query)
		set rsQ=nothing
		end if
		next		
		
		for f=1 to pcCartIndex
			if pcCartArray(f,10)=0 then
				query="SELECT idproduct FROM pcSavedPrdStats WHERE idproduct=" & pcCartArray(f,0) & ";"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					query="UPDATE pcSavedPrdStats SET pcSPS_SavedTimes=pcSPS_SavedTimes+1 WHERE idproduct=" & pcCartArray(f,0) & ";"
					set rsQ=connTemp.execute(query)
					set rsQ=nothing
				else
					query="INSERT INTO pcSavedPrdStats (idproduct,pcSPS_SavedTimes) VALUES (" & pcCartArray(f,0) & ",1);"
					set rsQ=connTemp.execute(query)
					set rsQ=nothing
				end if
			end if
		next
				
		if SCnewmonth=1 then
			query="INSERT INTO pcSavedCartStatistics (pcSCMonth,pcSCYear,pcSCTotals,pcSCTopPrds,pcSCAnonymous) VALUES (" & Month(Date()) & "," & Year(Date()) & ",1,'" & SCPrds & "',1);"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		else
			if pcv_HasNewCart=0 AND pcv_MoveToReg=1 then
				pcv_MoveToReg=-1
			else
				if pcv_HasNewCart=1 AND pcv_MoveToReg=0 then
					pcv_MoveToReg=1
				end if
			end if
			query="UPDATE pcSavedCartStatistics SET pcSCTotals=pcSCTotals+" & pcv_HasNewCart & ",pcSCAnonymous=pcSCAnonymous+" & pcv_MoveToReg & " WHERE pcSCMonth=" & Month(Date()) & " AND pcSCYear=" & Year(Date()) & ";"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
	END IF
ELSE 'Cart is empty, remove saved records
	tmpGUID=getUserInput(Request.Cookies("SavedCartGUID"),0)
	if session("IDCustomer")<>"" AND session("IDCustomer")<>"0" then
		tmpIDCust=session("IDCustomer")
	else
		tmpIDCust=0
	end if
	IF tmpGUID<>"" THEN
		query="SELECT IDCustomer FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			if clng(tmpIDCust)<>clng(rsQ("IDCustomer")) then
				tmpGUID=""
				Response.Cookies("SavedCartGUID")=""
			end if
		end if
		set rsQ=nothing
	END IF
	IF tmpGUID<>"" THEN
		query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			IDSC=rsQ("SavedCartID")
			query="SELECT SCArray0 FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpArr=rsQ.getRows()
				intC=ubound(tmpArr,2)
				For k=0 to intC
				query="SELECT pcSPS_SavedTimes FROM pcSavedPrdStats WHERE idproduct=" & tmpArr(0,k) & ";"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					if rsQ("pcSPS_SavedTimes")>"0" then
						query="UPDATE pcSavedPrdStats SET pcSPS_SavedTimes=pcSPS_SavedTimes-1 WHERE idproduct=" & tmpArr(0,k) & ";"
						set rsQ=connTemp.execute(query)
						set rsQ=nothing
					end if
				end if
				set rsQ=nothing
				Next
			end if
			set rsQ=nothing
			query="DELETE FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
			query="DELETE FROM pcSavedCarts WHERE SavedCartID=" & IDSC & ";"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
		set rsQ=nothing
		Response.Cookies("SavedCartGUID")=""
	END IF
END IF
If pcv_strSaveCart="1" Then
	response.Clear()
	response.ContentType = "text/html"
	response.Write("OK")
	response.End()
End If
%>
