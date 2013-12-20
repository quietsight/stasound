<% 
on error resume next

'function to open your database connection
function openDB()
 
 on error resume next
 set connTemp=server.createobject("adodb.connection")

 'Open your connection 
 connTemp.Open scDSN  
 
 if err.number <> 0 then
	response.redirect "dbError.asp"
	response.End()
 end if

end function


'function to close your database connection
function closeDB()
 on error resume next
 connTemp.close
 set connTemp=nothing
end function

%>