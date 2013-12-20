<%
 
' Details: date functions

function ShowDateFrmt(datestring)
if scDateFrmt="DD/MM/YY" then
 ShowDateFrmt=day(datestring)&"/"&month(datestring)&"/"&year(datestring)
else
 ShowDateFrmt=month(datestring)&"/"&day(datestring)&"/"&year(datestring)
end if
end function

function ShowMonthFrmt (datestring)
    Dim aDay, aMonth, aYear
    aDay   	= Day(datestring)
    aMonth 	= Monthname(Month(datestring),True)
    aYear 	= Year(datestring)    
    ShowMonthFrmt  = aDay & "-" & aMonth & "-" & aYear        
end Function

Function GetDateGUIDatabase(DBInputName, parseString)
	dim DBInputArray, dtInputDbM, dtInputDbD, dtInputDbY
	
	DBInputArray=split(DBInputName,"/")
	
	if SQL_Format="1" then
		'DD/MM/YYYY
		dtInputDbD=DBInputArray(0)
		dtInputDbM=DBInputArray(1)
		dtInputDbY=DBInputArray(2)
	else
		'MM/DD/YYYY
		dtInputDbM=DBInputArray(0)
		dtInputDbD=DBInputArray(1)
		dtInputDbY=DBInputArray(2)
	end if
	
	if parseString=1 then
		if scDateFrmt="DD/MM/YY" then
			GetDateGUIDatabase=dtInputDbD&"/"&dtInputDbM&"/"&dtInputDbY
		else
			GetDateGUIDatabase=dtInputDbM&"/"&dtInputDbD&"/"&dtInputDbY
		end if
	else
		GetDateGUIDatabase = dtInputDbM&", "&dtInputDbD&", "&dtInputDbY
	end if
End Function

%>