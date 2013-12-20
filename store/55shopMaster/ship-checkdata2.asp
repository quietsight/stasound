<%
query="select IDOrder from Orders where IDOrder=" & ship_order & ";"
set rs=connTemp.execute(query)
if rs.eof then
	ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": The Order ID #" & ship_order & " does not exist in the database." & vbcrlf
	RecordError=true
end if
%>
