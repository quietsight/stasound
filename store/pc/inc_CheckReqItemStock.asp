<%
on error goto 0
Dim pcv_OOS
pcv_OOS=0
'IF pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 THEN
IF pserviceSpec<>0 AND scOutofStockPurchase=-1 THEN
	queryQ="SELECT configProductCategory FROM configSpec_Products WHERE specProduct=" & pIdProduct & " AND requiredCategory<>0;"
	set rsQ=connTemp.execute(queryQ)
	
	if not rsQ.eof then
		tmpArrQ=rsQ.getRows()
		set rsQ=nothing
		intCountQ=ubound(tmpArrQ,2)
		
		For iQ=0 to intCountQ
			queryQ="SELECT products.idProduct FROM Products INNER JOIN configSpec_Products ON Products.idProduct=configSpec_Products.configProduct WHERE configSpec_Products.configProductCategory=" & tmpArrQ(0,iQ) & " AND configSpec_Products.specProduct=" & pIdProduct & " AND ((products.stock>0) OR (products.noStock<>0) OR (products.pcProd_BackOrder<>0)) AND products.removed=0 AND products.active<>0;"
			set rsQ=connTemp.execute(queryQ)
			if rsQ.eof then
				set rsQ=nothing
				pcv_OOS=1
			end if
			set rsQ=nothing
			if pcv_OOS=1 then
				pStock=0
				pNoStock=0
				pcv_intBackOrder=0
				exit for
			end if
		Next
	end if
	set rsQ=nothing
END IF
%>