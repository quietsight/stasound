<%'Start SDBA
	pcv_BackOrderStr=""
	'query="select Products.Description,ProductsOrdered.quantity,Products.pcProd_ShipNDays FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idOrder="& qry_ID & " AND ProductsOrdered.pcPrdOrd_BackOrder=1;"
	query="select Products.Description,ProductsOrdered.quantity,Products.pcProd_ShipNDays FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idOrder="& qry_ID & " AND (((Products.pcProd_BackOrder=1) AND (Products.stock<0)) OR (ProductsOrdered.pcPrdOrd_BackOrder=1));"
	pidorder=qry_ID
	set rsQ=connTemp.execute(query)
	
	if not rsQ.eof then
	
		pcv_BackOrderStr=pcv_BackOrderStr & "======================================================================" & vbcrlf
		pcv_BackOrderStr=pcv_BackOrderStr & ship_dictLanguage.Item(Session("language")&"_admconfirm_msg_1") & vbcrlf
		pcv_BackOrderStr=pcv_BackOrderStr & ship_dictLanguage.Item(Session("language")&"_admconfirm_msg_2") & vbcrlf & vbcrlf

		do while not rsQ.eof
			pcv_PrdName=rsQ("Description")
			pcv_PrdQty=rsQ("quantity")
			pcv_ShipNDays=rsQ("pcProd_ShipNDays")
			if IsNull(pcv_ShipNDays) or pcv_ShipNDays="" then
				pcv_ShipNDays=0
			end if
			
			pcv_BackOrderStr=pcv_BackOrderStr & rsQ("Description") & " - Qty:" & rsQ("quantity")
			if pcv_ShipNDays>"0" then
			pcv_BackOrderStr=pcv_BackOrderStr & " - " & dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_ShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b")
			end if
			pcv_BackOrderStr=pcv_BackOrderStr & vbcrlf & vbcrlf
			rsQ.MoveNext
		loop
		set rsQ=nothing
		pcv_BackOrderStr=pcv_BackOrderStr & vbcrlf & "======================================================================" & vbcrlf & vbcrlf
	end if
	set rsQ=nothing

'End SDBA%>

