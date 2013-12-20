<%
Function CheckDate(PreDate)
	if hour(PreDate)+minute(PreDate)+second(PreDate)>0 then
		if scDateFrmt="DD/MM/YY" then
			CheckDate=day(PreDate) & "/" & month(PreDate) & "/" & year(PreDate) & " " & PadNum(hour(PreDate),2) & ":" & PadNum(minute(PreDate),2) & ":" & PadNum(second(PreDate),2)
		else
			CheckDate=month(PreDate) & "/" & day(PreDate) & "/" & year(PreDate) & " " & PadNum(hour(PreDate),2) & ":" & PadNum(minute(PreDate),2) & ":" & PadNum(second(PreDate),2)
		end if
	else
		if scDateFrmt="DD/MM/YY" then
			CheckDate=day(PreDate) & "/" & month(PreDate) & "/" & year(PreDate) 
		else
			CheckDate=month(PreDate) & "/" & day(PreDate) & "/" & year(PreDate)
		end if
	end if
End Function

Function CheckDateSQL(PreDate)
	if hour(PreDate)+minute(PreDate)+second(PreDate)>0 then
		if SQL_Format="1" then
			CheckDateSQL=day(PreDate) & "/" & month(PreDate) & "/" & year(PreDate) & " " & PadNum(hour(PreDate),2) & ":" & PadNum(minute(PreDate),2) & ":" & PadNum(second(PreDate),2)
		else
			CheckDateSQL=month(PreDate) & "/" & day(PreDate) & "/" & year(PreDate) & " " & PadNum(hour(PreDate),2) & ":" & PadNum(minute(PreDate),2) & ":" & PadNum(second(PreDate),2)
		end if
	else
		if SQL_Format="1" then
			CheckDateSQL=day(PreDate) & "/" & month(PreDate) & "/" & year(PreDate) 
		else
			CheckDateSQL=month(PreDate) & "/" & day(PreDate) & "/" & year(PreDate)
		end if
	end if
End Function

Function PadNum(n, total) 
	PadNum=Right(String(total,"0") & n, total) 
End Function 
%>