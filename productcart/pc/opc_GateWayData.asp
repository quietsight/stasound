<%
'// Save Customer's IP Address for OPC
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

'save only the first 15 characters in case this is returned as a list of IP addresses
pcCustIpAddress = left(pcCustIpAddress,15)

call opendb()

query="UPDATE orders SET pcOrd_CustomerIP='"&pcCustIpAddress&"' WHERE orders.idOrder="&pIdOrder&";"
set rs=server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)	
if err.number<>0 then
	call LogErrorToDatabase()
end if
set rs=nothing

call closedb()
%>
