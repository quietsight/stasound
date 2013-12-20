<%'Start SDBA
'Send Order Notification to Drop=Shippers
query="SELECT DISTINCT ProductsOrdered.pcDropShipper_ID,pcDropShippersSuppliers.pcDS_IsDropShipper FROM pcDropShippersSuppliers INNER JOIN ProductsOrdered ON pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qry_ID & " AND ProductsOrdered.pcDropShipper_ID>0"
set rsQ=connTemp.execute(query)

do while not rsQ.eof
	pcv_DropShipperID=rsQ("pcDropShipper_ID")
	pcv_IsSupplier=rsQ("pcDS_IsDropShipper")
	if IsNull(pcv_IsSupplier) or pcv_IsSupplier="" then
		pcv_IsSupplier=0
	end if
	
	if pcv_IsSupplier=0 then
		query="SELECT pcDropShipper_FirstName AS B,pcDropShipper_Lastname AS C,pcDropShipper_Company AS A,pcDropShipper_Email AS D,pcDropShipper_NoticeEmail AS E,pcDropShipper_NoticeType AS F,pcDropShipper_NoticeMsg AS G,pcDropShipper_Notifymanually AS H FROM pcDropShippers WHERE pcDropShipper_ID=" & pcv_DropShipperID
	else
		query="SELECT pcSupplier_FirstName AS B,pcSupplier_LastName AS C,pcSupplier_Company AS A,pcSupplier_Email AS D,pcSupplier_NoticeEmail AS E,pcSupplier_NoticeType AS F,pcSupplier_NoticeMsg AS G,pcSupplier_Notifymanually AS H FROM pcSuppliers WHERE pcSupplier_ID=" & pcv_DropShipperID & " AND pcSupplier_IsDropShipper=1"
	end if
	set rsQ1=connTemp.execute(query)
	
	if not rsQ1.eof then
		pcv_DS_Company=rsQ1("A")
		pcv_DS_Name="(" & rsQ1("B") & " " & rsQ1("C") & ")"
		pcv_DS_Email=rsQ1("D")
		pcv_DS_NEmail=rsQ1("E")
		if IsNull(pcv_DS_NEmail) or pcv_DS_NEmail="" then
			pcv_DS_NEmail=pcv_DS_Email
		end if
		pcv_DS_NoticeType=rsQ1("F")
		if IsNull(pcv_DS_NoticeType) or pcv_DS_NoticeType="" then
			pcv_DS_NoticeType=0
		end if
		pcv_DS_NoticeMsg=rsQ1("G")
		pcv_DS_NoticeM=rsQ1("H")
		if IsNull(pcv_DS_NoticeM) or pcv_DS_NoticeM="" then
			pcv_DS_NoticeM=0
		end if
		set rsQ1=nothing
		
		pcv_DropShipperMsg=ship_dictLanguage.Item(Session("language")&"_sds_notifycanceldorder_2")
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<br>", vbCrlf)
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<STORE_NAME>",scCompanyName)
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<ORDER_ID>",(scpre + int(qry_ID)))
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<DROP_SHIPPER_COMPANY>",pcv_DS_Company)
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<DROP_SHIPPER_NAME>",pcv_DS_Name)
		pcv_DropShipperMsg=pcv_DropShipperMsg & vbcrlf & scCompanyName & vbcrlf
		
		pcv_DropShipperSbj=ship_dictLanguage.Item(Session("language")&"_sds_notifycanceldorder_1")
		pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<br>", vbCrlf)
		pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<STORE_NAME>",scCompanyName)
		pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<ORDER_ID>",(scpre + int(qry_ID)))
		pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<DROP_SHIPPER_COMPANY>",pcv_DS_Company)
		pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<DROP_SHIPPER_NAME>",pcv_DS_Name)
		
		'// Don't send if the order was still pending
		if porigstatus > 2 then
			call sendmail (scCompanyName, scEmail, pcv_DS_NEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg, "&quot;", chr(34)))
		end if
		
	end if
	set rsQ1=nothing
		
	rsQ.MoveNext
loop

set rsQ=nothing

'End SDBA%>