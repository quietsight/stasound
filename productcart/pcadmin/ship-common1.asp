<%		
		ship_order=0
		ship_ship=0
		ship_shipdate=""
		ship_sendmail=0
		ship_shipmethod=""
		ship_tracking=""

		ship_order=trim(CSVRecord(orderid))
		if ship_order<>"" then
			ship_order=clng(ship_order)-scpre
		end if
		
		if shipid<>-1 then
		ship_ship=trim(CSVRecord(shipid))
		end if
		
		if sendmailid<>-1 then
		ship_sendmail=trim(CSVRecord(sendmailid))
		end if
		
		if shipdateid<>-1 then
		ship_shipdate=trim(CSVRecord(shipdateid))
		end if
		
		if ship_shipdate<>"" then
		else
		ship_shipdate=Date()
		end if
		
		if shipmethodid<>-1 then
		ship_shipmethod=trim(CSVRecord(shipmethodid))
		end if
		
		if trackingid<>-1 then
		ship_tracking=trim(CSVRecord(trackingid))
		end if
		
		if ship_ship="" then 
		ship_ship="0"
		end if

		if ship_sendmail="" then
		ship_sendmail="0"
		end if
		
		if ship_ship="0" then 
			ship_sendmail="0"
		end if
		
		if ship_shipdate<>"" then
		ship_shipdate=CDate(ship_shipdate)
		if SQL_Format="1" then
			ship_shipdate=Day(ship_shipdate)&"/"&Month(ship_shipdate)&"/"&Year(ship_shipdate)
		else
			ship_shipdate=Month(ship_shipdate)&"/"&Day(ship_shipdate)&"/"&Year(ship_shipdate)
        end if
		end if
		
		if ship_order="" then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not have an Order ID." & vbcrlf
		RecordError=true
		end if

		if isNumeric(ship_ship)=false then
		ErrorsReport=ErrorsReport & "Order " & (scpre+ship_order) & ": The 'Ship' Field is not a number." & vbcrlf
		RecordError=true
		end if
		if isNumeric(ship_sendmail)=false then
		ErrorsReport=ErrorsReport & "Order " & (scpre+ship_order) & ": The 'Send Mail' Field is not a number." & vbcrlf
		RecordError=true
		end if
		
		if ship_ship>"1" then
		ship_ship="1"
		end if
		
		if ship_sendmail>"1" then 
			ship_sendmail="1"
		end if
		
		if RecordError=false then
			query="SELECT idorder FROM Orders WHERE idorder=" & ship_order & " AND orderstatus=4;"
			set rsTest=connTemp.execute(query)
			if not rsTest.eof then
				ErrorsReport=ErrorsReport & "Order " & (scpre+ship_order) & " has already been shipped." & vbcrlf
				RecordError=true
			end if
			set rsTest=nothing
		end if
%>