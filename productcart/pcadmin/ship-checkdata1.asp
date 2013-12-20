<%
query="SELECT IDOrder FROM Orders WHERE IDOrder=" & ship_order & ";"
set rs=connTemp.execute(query)
if rs.eof then
	ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": The Order ID #" & ship_order & " does not exist in the database." & vbcrlf
	RecordError=true
end if
%>
