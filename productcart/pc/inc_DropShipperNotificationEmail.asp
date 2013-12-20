<%'Start SDBA
'Send Order Notification to Drop=Shippers
if pcv_DropShipperID<>"" and pcv_DropShipperID<>"0" then
	if pcv_IsSupplier="" then
		pcv_IsSupplier=0
	end if
	tmpStr=" AND ProductsOrdered.pcDropShipper_ID=" & pcv_DropShipperID & " AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & pcv_IsSupplier
else
	tmpStr=" AND ProductsOrdered.pcDropShipper_ID>0"
end if

query="SELECT DISTINCT ProductsOrdered.pcDropShipper_ID,pcDropShippersSuppliers.pcDS_IsDropShipper FROM pcDropShippersSuppliers INNER JOIN ProductsOrdered ON pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qry_ID & tmpStr
set rsQ=Server.CreateObject("ADODB.Recordset")
set rsQ=connTemp.execute(query)
do while not rsQ.eof
	pcv_DropShipperID=rsQ("pcDropShipper_ID")
	pcv_IsSupplier=rsQ("pcDS_IsDropShipper")
	%>
	<!--#include file="inc_GenDropShipperNotification.asp"-->
	<%rsQ.MoveNext
loop
set rsQ=nothing

'End SDBA%>