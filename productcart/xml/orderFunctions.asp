<%

Sub CheckSrcOrdersTags()
Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	Set fNode=iRoot.selectSingleNode(cm_filters_name)
	if fNode is Nothing then
		exit Sub
	end if
	if fNode.Text="" then
		exit Sub
	end if
	Set ChildNodes = fNode.childNodes
	
	srcOrderID_value=0
	srcCustomerID_value=0
	srcPricingCatID_value=0
	srcOrderStatus_value=0
	srcPaymentStatus_value=0
	srcPaymentType_value=""
	srcShippingType_value=0
	srcStateCode_value=""
	srcDiscountCode_value=""
	srcPrdOrderedID_value=0
		
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		
		Select Case tmpNodeName
			Case srcCustomerID_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcCustomerID_ex=1
					srcCustomerID_value=tmpNodeValue
				end if
			Case srcCustomerType_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcCustomerType_ex=1
					srcCustomerType_value=tmpNodeValue
				end if
			Case srcPricingCatID_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcPricingCatID_ex=1
					srcPricingCatID_value=tmpNodeValue
				end if
			Case srcOrderStatus_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcOrderStatus_ex=1
					srcOrderStatus_value=tmpNodeValue
				end if
			Case srcPaymentStatus_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcPaymentStatus_ex=1
					srcPaymentStatus_value=tmpNodeValue
				end if
			Case srcPaymentType_name:
				call CheckValidXMLTag(strNode,1,5,"")
				if tmpNodeValue<>"" then
					srcPaymentType_ex=1
					srcPaymentType_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcStateCode_name:
				call CheckValidXMLTag(strNode,1,5,"")
				if tmpNodeValue<>"" then
					srcStateCode_ex=1
					srcStateCode_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcCountryCode_name:
				call CheckValidXMLTag(strNode,1,5,"")
				if tmpNodeValue<>"" then
					srcCountryCode_ex=1
					srcCountryCode_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcDiscountCode_name:
				call CheckValidXMLTag(strNode,1,5,"")
				if tmpNodeValue<>"" then
					srcDiscountCode_ex=1
					srcDiscountCode_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcPrdOrderedID_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcPrdOrderedID_ex=1
					srcPrdOrderedID_value=tmpNodeValue
				end if
			Case srcFromDate_name:
				call CheckValidXMLTag(strNode,0,4,"")
				if tmpNodeValue<>"" then
					srcFromDate_ex=1
					srcFromDate_value=ConvertFromXMLDate(tmpNodeValue)
				end if
			Case srcToDate_name:
				tmpValue1=0
				if CheckExistTag(cm_filters_name & "/" & srcFromDate_name) then
					tmpValue1=iRoot.selectSingleNode(cm_filters_name & "/" & srcFromDate_name).Text
				end if
				call CheckValidXMLTag(strNode,0,4,tmpValue1)
				if tmpNodeValue<>"" then
					srcToDate_ex=1
					srcToDate_value=ConvertFromXMLDate(tmpNodeValue)
				end if
			Case srcHideExported_name:
				if cm_ExportAdmin="1" then
					call CheckValidXMLTag(strNode,1,1,"")
					if tmpNodeValue<>"" then
						srcHideExported_ex=1
						srcHideExported_value=tmpNodeValue
						if srcHideExported_value>1 then
							srcHideExported_value=1
						end if
					end if
				end if
			Case Else:
				call XMLcreateError(105,cm_errorStr_105 & tmpNodeName)
				call returnXML()
		End Select
	Next
End Sub

Sub CheckNewOrdersTags()
Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	Set fNode=iRoot.selectSingleNode(cm_filters_name)
	if fNode is Nothing then
		exit Sub
	end if
	if fNode.Text="" then
		exit Sub
	end if
	Set ChildNodes = fNode.childNodes
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		
		Select Case tmpNodeName
			Case srcFromDate_name:
				call CheckValidXMLTag(strNode,0,4,"")
				if tmpNodeValue<>"" then
					srcFromDate_ex=1
					srcFromDate_value=ConvertFromXMLDate(tmpNodeValue)
				end if
			Case Else:
				call XMLcreateError(105,cm_errorStr_105 & tmpNodeName)
				call returnXML()
		End Select
	Next
End Sub

Sub CheckGetOrderDetailsTags()
Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	
	Call CheckRequiredXMLTag(ordID_name)
	Set strNode=iRoot.selectSingleNode(ordID_name)
	call CheckValidXMLTag(strNode,1,1,"")
	ordID_ex=1
	ordID_value=tmpNode.Text

	Set rNode=iRoot.selectSingleNode(cm_requests_name)
	if rNode is Nothing then
		Call SetDefaultOrderDetailsTags()
		exit Sub
	else
		if rNode.Text="" then
			Call SetDefaultOrderDetailsTags()
			exit Sub
		end if
	end if
	Set ChildNodes = rNode.childNodes
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		if	tmpNodeName=cm_request_name then
			Select Case tmpNodeValue
				Case cm_requestDefault_name:
					Call SetDefaultOrderDetailsTags()
				Case cm_requestAll_name:
					Call SetAllOrderDetailsTags()
				Case ordName_name:
					ordName_ex=1
				Case ordDate_name:
					ordDate_ex=1
				Case custID_name:
					custID_ex=1
				Case ordTotal_name:
					ordTotal_ex=1
				Case ordProcessedDate_name:
					ordProcessedDate_ex=1
				Case ordCustDetails_name:
					ordCustDetails_ex=1
				Case ordShippingAddress_name:
					ordShippingAddress_ex=1
				Case ordShipDetails_name:
					ordShipDetails_ex=1
				Case ordDeliveryDate_name:
					ordDeliveryDate_ex=1
				Case ordStatus_name:
					ordStatus_ex=1
				Case ordPaymentStatus_name:
					ordPaymentStatus_ex=1
				Case ordPayDetails_name:
					ordPayDetails_ex=1
				Case ordAffiliate_name:
					ordAffiliate_ex=1
				Case ordRP_name:
					ordRP_ex=1
				Case ordAccruedRP_name:
					ordAccruedRP_ex=1
				Case ordReferrer_name:
					ordReferrer_ex=1
				Case ordRMACredit_name:
					ordRMACredit_ex=1
				Case ordTaxAmount_name:
					ordTaxAmount_ex=1
				Case ordVAT_name:
					ordVAT_ex=1
				Case ordTaxDetails_name:
					ordTaxDetails_ex=1
				Case ordDiscountDetails_name:
					ordDiscountDetails_ex=1
				Case GiftCertificate_name:
					GiftCertificate_ex=1
				Case GiftCertificateUsed_name:
					GiftCertificateUsed_ex=1
				Case ordCatDiscounts_name:
					ordCatDiscounts_ex=1
				Case ordCustomerComments_name:
					ordCustomerComments_ex=1
				Case ordAdminComments_name:
					ordAdminComments_ex=1
				Case ordReturnDate_name:
					ordReturnDate_ex=1
				Case ordReturnReason_name:
					ordReturnReason_ex=1
				Case ordShoppingCart_name:
					ordShoppingCart_ex=1
					prdID_ex=1
				Case prdID_name:
					prdID_ex=1
					ordShoppingCart_ex=1
				Case prdSKU_name:
					prdSKU_ex=1
					ordShoppingCart_ex=1
				Case prdName_name:
					prdName_ex=1
					ordShoppingCart_ex=1
				Case prdUnitPrice_name:
					prdUnitPrice_ex=1
					ordShoppingCart_ex=1
				Case prdQuantity_name:
					prdQuantity_ex=1
					ordShoppingCart_ex=1
				Case prdBTOConfig_name:
					prdBTOConfig_ex=1
					ordShoppingCart_ex=1
				Case prdOption_name:
					prdOption_ex=1
					ordShoppingCart_ex=1
				Case prdQtyDiscounts_name:
					prdQtyDiscounts_ex=1
					ordShoppingCart_ex=1
				Case prdItemDiscounts_name:
					prdItemDiscounts_ex=1
					ordShoppingCart_ex=1
				Case prdGiftWrapping_name:
					prdGiftWrapping_ex=1
					ordShoppingCart_ex=1
				Case packageID_name:
					packageID_ex=1
					ordShoppingCart_ex=1
				Case prdTotalPrice_name:
					prdTotalPrice_ex=1
					ordShoppingCart_ex=1
				Case Else:
					call XMLcreateError(106,cm_errorStr_106 & tmpNodeValue)
					call returnXML()
			End Select
		else
			call XMLcreateError(106,cm_errorStr_106 & tmpNodeName)
			call returnXML()
		end if
	Next
	If ordShoppingCart_ex=1 then
		prdID_ex=1
	End if
End Sub

Sub SetDefaultOrderDetailsTags()
	ordName_ex=1
	ordDate_ex=1
	custID_ex=1
	ordTotal_ex=1
	ordProcessedDate_ex=1
	ordCustDetails_ex=0
	ordShippingAddress_ex=1
	ordShipDetails_ex=1
	ordDeliveryDate_ex=1
	ordStatus_ex=1
	ordPaymentStatus_ex=1
	ordPayDetails_ex=1
	ordAffiliate_ex=0
	ordRP_ex=0
	ordAccruedRP_ex=0
	ordReferrer_ex=0
	ordRMACredit_ex=0
	ordTaxAmount_ex=1
	ordVAT_ex=1
	ordTaxDetails_ex=0
	ordDiscountDetails_ex=1
	GiftCertificate_ex=1
	GiftCertificateUsed_ex=1
	ordCatDiscounts_ex=1
	ordCustomerComments_ex=0
	ordAdminComments_ex=0
	ordReturnDate_ex=0
	ordReturnReason_ex=0
	ordShoppingCart_ex=1
	prdID_ex=1
	prdSKU_ex=1
	prdName_ex=1
	prdUnitPrice_ex=1
	prdQuantity_ex=1
	prdBTOConfig_ex=0
	prdOption_ex=0
	prdQtyDiscounts_ex=1
	prdItemDiscounts_ex=0
	prdGiftWrapping_ex=0
	packageID_ex=0
	prdTotalPrice_ex=1
End Sub

Sub SetAllOrderDetailsTags()
	ordName_ex=1
	ordDate_ex=1
	custID_ex=1
	ordTotal_ex=1
	ordProcessedDate_ex=1
	ordCustDetails_ex=1
	ordShippingAddress_ex=1
	ordShipDetails_ex=1
	ordDeliveryDate_ex=1
	ordStatus_ex=1
	ordPaymentStatus_ex=1
	ordPayDetails_ex=1
	ordAffiliate_ex=1
	ordRP_ex=1
	ordAccruedRP_ex=1
	ordReferrer_ex=1
	ordRMACredit_ex=1
	ordTaxAmount_ex=1
	ordVAT_ex=1
	ordTaxDetails_ex=1
	ordDiscountDetails_ex=1
	GiftCertificate_ex=1
	GiftCertificateUsed_ex=1
	ordCatDiscounts_ex=1
	ordCustomerComments_ex=1
	ordAdminComments_ex=1
	ordReturnDate_ex=1
	ordReturnReason_ex=1
	ordShoppingCart_ex=1
	prdID_ex=1
	prdSKU_ex=1
	prdName_ex=1
	prdUnitPrice_ex=1
	prdQuantity_ex=1
	prdBTOConfig_ex=1
	prdOption_ex=1
	prdQtyDiscounts_ex=1
	prdItemDiscounts_ex=1
	prdGiftWrapping_ex=1
	packageID_ex=1
	prdTotalPrice_ex=1
End Sub

Function GenSrcOrdersQuery()
Dim query,query1,query2,tmpquery
Dim strORD1,tmpKey,tmpFromDate,tmpToDate

	strORD1="orders.idorder ASC"

	' create sql statement
	query1=""
	query2=""
	
	if srcCustomerID_ex=1 then
		query2=query2 & " AND (orders.idCustomer=" & srcCustomerID_value & ")"
	end if
	
	if srcCustomerType_ex=1 then
		query1=query1 & ",customers"
		query2=query2 & " AND ((customers.idcustomer=orders.idcustomer) AND (customers.customerType=" & srcCustomerType_value & "))"
	end if
	
	if srcPricingCatID_ex=1 then
		if srcCustomerType_ex<>1 then
			query1=query1 & ",customers"
		end if
		tmpquery=" AND (customers.idCustomerCategory=" & srcPricingCatID_value & ")"
		if srcCustomerType_ex<>1 then
			tmpquery=" AND ((customers.idcustomer=orders.idcustomer) " & tmpquery & ")"
		end if
		query2=query2 & tmpquery
	end if
	
	if srcOrderStatus_ex=1 then
		query2=query2 & " AND (orders.orderStatus=" & srcOrderStatus_value & ")"
	end if
	
	if srcPaymentStatus_ex=1 then
		query2=query2 & " AND (orders.pcOrd_PaymentStatus=" & srcPaymentStatus_value & ")"
	end if
	
	if srcPaymentType_ex=1 then
		query2=query2 & " AND (orders.paymentDetails LIKE '%" & srcPaymentType_value & "%')"
	end if
	
	if srcStateCode_ex=1 then
		query2=query2 & " AND (orders.stateCode LIKE '" & srcStateCode_value & "')"
	end if
		
	if srcCountryCode_ex=1 then
		query2=query2 & " AND (orders.countryCode LIKE '" & srcCountryCode_value & "')"
	end if
	
	if srcDiscountCode_ex=1 then
		query2=query2 & " AND (orders.discountDetails LIKE '%" & srcDiscountCode_value & "%')"
	end if
	
	if srcPrdOrderedID_ex=1 then
		query1=query1 & ",ProductsOrdered"
		query2=query2 & " AND ((ProductsOrdered.idorder=orders.idorder) AND (ProductsOrdered.idProduct=" & srcPrdOrderedID_value & "))"
	end if
	
	If srcFromDate_ex=1 then
		tmpFromDate=srcFromDate_Value
		if SQL_Format="1" then
			tmpFromDate=Day(tmpFromDate)&"/"&Month(tmpFromDate)&"/"&Year(tmpFromDate)
		else
			tmpFromDate=Month(tmpFromDate)&"/"&Day(tmpFromDate)&"/"&Year(tmpFromDate)
		end if
		query2=query2 & " AND "
		if scDB="Access" then
			query2=query2 & " (orders.orderDate>=#" & tmpFromDate & "#) "
		else
			query2=query2 & " (orders.orderDate>='" & tmpFromDate & "') "
		end if
	End if
	
	If srcToDate_ex=1 then
		tmpToDate=CDate(srcToDate_Value)
		if SQL_Format="1" then
			tmpToDate=Day(tmpToDate)&"/"&Month(tmpToDate)&"/"&Year(tmpToDate)
		else
			tmpToDate=Month(tmpToDate)&"/"&Day(tmpToDate)&"/"&Year(tmpToDate)
		end if
		query2=query2 & " AND "
		if scDB="Access" then
			query2=query2 & " (orders.orderDate<=#" & tmpToDate & "#) "
		else
			query2=query2 & " (orders.orderDate<='" & tmpToDate & "') "
		end if
	End if
	
	if cm_ExportAdmin="1" AND srcHideExported_value="1" then
		query2=query2 & " AND (Orders.idorder NOT IN (SELECT DISTINCT pcXEL_ExportedID FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=2)) "
	end if
	
	query="SELECT DISTINCT orders.idorder FROM orders " & query1 & " WHERE orders.orderStatus>1 " & query2 & " ORDER BY "& strORD1
	
	GenSrcOrdersQuery=query
	
End Function

Sub RunSrcOrders()
	Dim query,rs1,resultCount,pcArr
	Dim requestKey,i,strNode
	on error resume next
	
	query=GenSrcOrdersQuery()
	call opendb()
	set rs1=connTemp.execute(query)
	resultCount=0
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		pcArr=rs1.getRows()
		resultCount=ubound(pcArr,2)+1
	end if
	set rs1=nothing
	call closedb()
	
	IF cm_LogTurnOn=1 THEN
		requestKey=CreateRequestRecord(pcv_PartnerID,2,0,0,0,resultCount,0,0)
		cm_requestKey_value=requestKey
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
	tmpNode.Text=cm_SuccessCode
	oRoot.appendChild(tmpNode)
	
	Set tmpNode=oXML.createNode(1,cm_resultCount_name,"")
	tmpNode.Text=resultCount
	oRoot.appendChild(tmpNode)
	
	if resultCount>0 then
	
		Set tmpNode=oXML.createNode(1,cm_orders,"")
		oRoot.appendChild(tmpNode)
	
		For i=0 to resultCount-1
			Set strNode=oXML.createNode(1,ordID_name,"")
			strNode.Text=pcArr(0,i)
			tmpNode.appendChild(strNode)
		Next
	
	end if
	
End Sub

Function GenNewOrdersQuery()
	Dim strSQL, query, tmpFromDate
	Dim rs1,tmpLastID
	on error resume next

	tmpLastID=0
	
	call opendb()
	
	query="SELECT pcXL_LastID FROM pcXMLLogs WHERE pcXP_id=" & pcv_PartnerID & " AND pcXL_RequestType=8 ORDER BY pcXL_LastID DESC;"
	set rs1=connTemp.execute(query)

	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		tmpLastID=rs1("pcXL_LastID")
	end if
	set rs1=nothing
	
	call closedb()
	
	strSQL=""
	
	IF Clng(tmpLastID)>0 THEN
		strSQL=strSQL & " orders.idorder>" & tmpLastID
	ELSE
		If srcFromDate_ex=0 then
			srcFromDate_ex=1
			srcFromDate_Value=Date()-7
		End if

		tmpFromDate=srcFromDate_Value
		if SQL_Format="1" then
			tmpFromDate=Day(tmpFromDate)&"/"&Month(tmpFromDate)&"/"&Year(tmpFromDate)
		else
			tmpFromDate=Month(tmpFromDate)&"/"&Day(tmpFromDate)&"/"&Year(tmpFromDate)
		end if
		if scDB="Access" then
			strSQL=strSQL & " orders.orderDate>=#" & tmpFromDate & "# "
		else
			strSQL=strSQL & " orders.orderDate>='" & tmpFromDate & "' "
		end if
	END IF
	
	query="SELECT orders.idorder FROM Orders WHERE " & strSQL & " AND orders.orderStatus>1 ORDER BY orders.idorder ASC;"
	
	GenNewOrdersQuery=query
	
End Function

Sub RunNewOrders()
	Dim query,rs1,resultCount,pcArr
	Dim requestKey,i,strNode,tmpLastID
	on error resume next
	
	query=GenNewOrdersQuery()
	call opendb()
	set rs1=connTemp.execute(query)
	resultCount=0
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		pcArr=rs1.getRows()
		resultCount=ubound(pcArr,2)+1
	end if
	set rs1=nothing
	
	tmpLastID=0
	query="SELECT orders.idorder FROM Orders ORDER BY orders.idorder DESC;"
	set rs1=connTemp.execute(query)
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		tmpLastID=rs1("idorder")
	end if
	set rs1=nothing
	
	call closedb()
	
	IF cm_LogTurnOn=1 THEN
		requestKey=CreateRequestRecord(pcv_PartnerID,8,0,0,0,resultCount,tmpLastID,0)
		cm_requestKey_value=requestKey
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
	tmpNode.Text=cm_SuccessCode
	oRoot.appendChild(tmpNode)
	
	Set tmpNode=oXML.createNode(1,cm_resultCount_name,"")
	tmpNode.Text=resultCount
	oRoot.appendChild(tmpNode)
	
	if resultCount>0 then
	
		Set tmpNode=oXML.createNode(1,cm_orders,"")
		oRoot.appendChild(tmpNode)
	
		For i=0 to resultCount-1
			Set strNode=oXML.createNode(1,ordID_name,"")
			strNode.Text=pcArr(0,i)
			tmpNode.appendChild(strNode)
		Next
	
	end if
	
End Sub

Sub XMLgetOrdCustDetails(parentNode,tmpIDCustomer)
Dim query,rs1,attNode,pcArr,intCount,i
	
	call opendb()
	
	query="SELECT name,lastName,email,customerCompany,phone,fax,address,address2,city,stateCode,state,zip,countryCode FROM Customers WHERE idcustomer=" & tmpIDCustomer & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
					
			Set attNode=oXML.createNode(1,ordCustDetails_name,"")
			parentNode.appendChild(attNode)
		
			Call XMLCreateNode(attNode,custFirstName_name,New_HTMLEncode(trim(pcArr(0,i))))
			Call XMLCreateNode(attNode,custLastName_name,New_HTMLEncode(trim(pcArr(1,i))))
			Call XMLCreateNode(attNode,custEmail_name,New_HTMLEncode(trim(pcArr(2,i))))
			Call XMLCreateNode(attNode,custCompany_name,New_HTMLEncode(trim(pcArr(3,i))))
			Call XMLCreateNode(attNode,custPhone_name,New_HTMLEncode(trim(pcArr(4,i))))
			Call XMLCreateNode(attNode,custFax_name,New_HTMLEncode(trim(pcArr(5,i))))
			Call XMLCreateNode(attNode,custAddress_name,New_HTMLEncode(trim(pcArr(6,i))))
			Call XMLCreateNode(attNode,custAddress2_name,New_HTMLEncode(trim(pcArr(7,i))))
			Call XMLCreateNode(attNode,custCity_name,New_HTMLEncode(trim(pcArr(8,i))))
			Call XMLCreateNode(attNode,custStateCode_name,New_HTMLEncode(trim(pcArr(9,i))))
			Call XMLCreateNode(attNode,custProvince_name,New_HTMLEncode(trim(pcArr(10,i))))
			Call XMLCreateNode(attNode,custZip_name,New_HTMLEncode(trim(pcArr(11,i))))
			Call XMLCreateNode(attNode,custCountryCode_name,New_HTMLEncode(trim(pcArr(12,i))))

		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub XMLgetOrdPkgDetails(parentNode,tmpIDOrder)
Dim query,rs1,attNode,pcArr,intCount,i,tmpCreatedDate
	
	call opendb()
	
	query="SELECT pcPackageInfo_ID,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_Comments,pcPackageInfo_TrackingNumber FROM pcPackageInfo WHERE idorder=" & tmpIDOrder & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
					
			Set attNode=oXML.createNode(1,packageInfo_name,"")
			parentNode.appendChild(attNode)
		
			Call XMLCreateNode(attNode,pkgID_name,pcArr(0,i))
			Call XMLCreateNode(attNode,pkgShipMethod_name,New_HTMLEncode(trim(pcArr(1,i))))
			tmpCreatedDate=trim(pcArr(2,i))
			if tmpCreatedDate<>"" then
				tmpCreatedDate=ConvertToXMLDate(tmpCreatedDate)
			end if
			Call XMLCreateNode(attNode,pkgShipDate_name,tmpCreatedDate)
			Call XMLCreateNode(attNode,pkgComment_name,New_HTMLEncode(trim(pcArr(3,i))))
			Call XMLCreateNode(attNode,pkgTrackingNumber_name,trim(pcArr(4,i)))

		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub XMLgetOrdAffDetails(parentNode,tmpIDAff,tmpPayValue)
Dim query,rs1,attNode,pcArr,intCount,i
	
	call opendb()
	
	query="SELECT affiliateName FROM affiliates WHERE idAffiliate=" & tmpIDAff & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
					
			Set attNode=oXML.createNode(1,ordAffiliate_name,"")
			parentNode.appendChild(attNode)
		
			Call XMLCreateNode(attNode,affID_name,tmpIDAff)
			Call XMLCreateNode(attNode,affName_name,New_HTMLEncode(trim(pcArr(0,i))))
			Call XMLCreateNode(attNode,affPayment_name,tmpPayValue)

		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub XMLgetOrdRefDetails(parentNode,tmpIDRef)
Dim query,rs1,attNode,pcArr,intCount,i
	
	call opendb()
	
	query="SELECT [Name] FROM Referrer WHERE IdRefer=" & tmpIDRef & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
					
			Set attNode=oXML.createNode(1,ordReferrer_name,"")
			parentNode.appendChild(attNode)
		
			Call XMLCreateNode(attNode,refID_name,tmpIDRef)
			Call XMLCreateNode(attNode,refName_name,New_HTMLEncode(trim(pcArr(0,i))))

		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub XMLgetOrdGWDetails(parentNode,tmpIDGW,tmpPriceValue)
	Dim query,rs1,attNode,pcArr,intCount,i
	On Error Resume Next
	
	query="SELECT pcGW_OptName FROM pcGWOptions WHERE pcGW_IDOpt=" & tmpIDGW & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
					
			Set attNode=oXML.createNode(1,prdGiftWrapping_name,"")
			parentNode.appendChild(attNode)
		
			Call XMLCreateNode(attNode,gwID_name,tmpIDGW)
			Call XMLCreateNode(attNode,gwName_name,New_HTMLEncode(trim(pcArr(0,i))))
			Call XMLCreateNode(attNode,gwPrice_name,tmpPriceValue)

		Next
	end if
	set rs1=nothing
	
End Sub

Sub XMLgetOrdBTOConfDetails(parentNode,tmpIDConfig,pQuantity)
Dim query,rs1,attNode,pcArr,intCount,i,subNode
Dim tmpAddPrice,stringProducts,stringValues,stringCategories,stringQuantity,stringPrice
Dim ArrProduct,ArrValue,ArrCategory,ArrQuantity,ArrPrice,pidproduct,pdescription,psku,pidcategory,pcategoryDesc,UPrice
	
	if tmpIDConfig<>"0" then 
		query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & tmpIDConfig & ";"
		set rsConfigObj=server.CreateObject("ADODB.RecordSet")
		set rsConfigObj=connTemp.execute(query)
		
		Set attNode=oXML.createNode(1,prdBTOConfig_name,"")
		parentNode.appendChild(attNode)
			
		stringProducts=trim(rsConfigObj("stringProducts"))
		stringValues=trim(rsConfigObj("stringValues"))
		stringCategories=trim(rsConfigObj("stringCategories"))
		stringQuantity=trim(rsConfigObj("stringQuantity"))
		stringPrice=trim(rsConfigObj("stringPrice"))
		set rsConfigObj=nothing
		
		ArrProduct=Split(stringProducts, ",")
		ArrValue=Split(stringValues, ",")
		ArrCategory=Split(stringCategories, ",")
		ArrQuantity=Split(stringQuantity, ",")
		ArrPrice=Split(stringPrice, ",")

		for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
			Set subNode=oXML.createNode(1,btoItem_name,"")
			attNode.appendChild(subNode)
		
			query="SELECT products.idproduct,products.description, products.sku,categories.idcategory,categories.categoryDesc FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
			set rsConfigObj=server.CreateObject("ADODB.RecordSet")
			set rsConfigObj=connTemp.execute(query)
	
			pidproduct=trim(rsConfigObj("idproduct"))
			pdescription=trim(rsConfigObj("description"))
			psku=trim(rsConfigObj("sku"))
			pidcategory=trim(rsConfigObj("idcategory"))
			pcategoryDesc=trim(rsConfigObj("categoryDesc"))
			set rsConfigObj=nothing
	
			Call XMLCreateNode(subNode,itemID_name,pidproduct)
			Call XMLCreateNode(subNode,itemName_name,New_HTMLEncode(pdescription))
			Call XMLCreateNode(subNode,itemSKU_name,New_HTMLEncode(psku))
			Call XMLCreateNode(subNode,itemCategoryID_name,pidcategory)
			Call XMLCreateNode(subNode,itemCategoryName_name,New_HTMLEncode(pcategoryDesc))
			Call XMLCreateNode(subNode,itemQuantity_name,ArrQuantity(i))
			if (ArrQuantity(i)-1)>=0 then
				UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
			else
				UPrice=0
			end if
			tmpAddPrice=(ArrValue(i)+UPrice)*pQuantity
			Call XMLCreateNode(subNode,itemAddPrice_name,tmpAddPrice)
		next
	end if
	
End Sub

Sub XMLgetOrdPrdsDetails(parentNode,tmpIDOrder)
Dim query,rs1,attNode,pcArr,intCount,i,subNode,subNode1,attNode1
Dim tmpOptNameArr,tmpOptPriceArr,j,tmpTotalValue
	
	call opendb()
	
	query="SELECT Products.idProduct,Products.sku,Products.description,ProductsOrdered.unitPrice,ProductsOrdered.quantity,ProductsOrdered.idconfigSession,ProductsOrdered.pcPrdOrd_OptionsArray,ProductsOrdered.pcPrdOrd_OptionsPriceArray,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts,pcPO_GWOpt,pcPO_GWPrice,pcPackageInfo_ID,ProductsOrdered.pcPrdOrd_SelectedOptions FROM Products INNER JOIN ProductsOrdered ON Products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & tmpIDOrder & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		
		Set attNode1=oXML.createNode(1,ordShoppingCart_name,"")
		parentNode.appendChild(attNode1)
		
		Set attNode=oXML.createNode(1,products_name,"")
		attNode1.appendChild(attNode)
		
		For i=0 to intCount
			Set subNode=oXML.createNode(1,cm_product,"")
			attNode.appendChild(subNode)

			if prdID_ex=1 then
				Call XMLCreateNode(subNode,prdID_name,pcArr(0,i))
			end if
			if prdSKU_ex=1 then
				Call XMLCreateNode(subNode,prdSKU_name,New_HTMLEncode(trim(pcArr(1,i))))
			end if
			if prdName_ex=1 then
				Call XMLCreateNode(subNode,prdName_name,New_HTMLEncode(trim(pcArr(2,i))))
			end if
			if prdUnitPrice_ex=1 then
				Call XMLCreateNode(subNode,prdUnitPrice_name,trim(pcArr(3,i)))
			end if
			if prdQuantity_ex=1 then
				Call XMLCreateNode(subNode,prdQuantity_name,trim(pcArr(4,i)))
			end if
			if prdBTOConfig_ex=1 then
				if pcArr(5,i)>"0" then
					Call XMLgetOrdBTOConfDetails(subNode,pcArr(5,i),pcArr(4,i))
				end if
			end if
			if prdOption_ex=1 then
				if trim(pcArr(13,i))<>"" AND trim(pcArr(13,i))<>"NULL" then
					tmpOptIDArr=split(trim(pcArr(13,i)),"|")
					tmpOptNameArr=split(trim(pcArr(6,i)),"|")
					tmpOptPriceArr=split(trim(pcArr(7,i)),"|")
					For j=lbound(tmpOptNameArr) to ubound(tmpOptNameArr)
						Set subNode1=oXML.createNode(1,prdOption_name,"")
						subNode.appendChild(subNode1)
						Call XMLCreateNode(subNode1,ordOptID_name,New_HTMLEncode(tmpOptIDArr(j)))
						Call XMLCreateNode(subNode1,optName_name,New_HTMLEncode(tmpOptNameArr(j)))
						Call XMLCreateNode(subNode1,optPrice_name,tmpOptPriceArr(j))
					next
				end if
			end if
			if prdQtyDiscounts_ex=1 then
				Call XMLCreateNode(subNode,prdQtyDiscounts_name,trim(pcArr(8,i)))
			end if
			if prdItemDiscounts_ex=1 then
				Call XMLCreateNode(subNode,prdItemDiscounts_name,trim(pcArr(9,i)))
			end if
			if prdGiftWrapping_ex=1 then
				Call XMLgetOrdGWDetails(subNode,trim(pcArr(10,i)),trim(pcArr(11,i)))
			end if
			if packageID_ex=1 then
				Call XMLCreateNode(subNode,packageID_name,trim(pcArr(12,i)))
			end if
			if prdTotalPrice_ex=1 then
				pcArr(3,i)=pcf_SanitizeNULL(pcArr(3,i),0)
				pcArr(4,i)=pcf_SanitizeNULL(pcArr(4,i),0)
				pcArr(9,i)=pcf_SanitizeNULL(pcArr(9,i),0)
				tmpTotalValue=cdbl(pcArr(3,i))*cdbl(pcArr(4,i))-cdbl(pcArr(9,i))
				Call XMLCreateNode(subNode,prdTotalPrice_name,tmpTotalValue)
			end if
		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub


Sub RunGetOrderDetails()
Dim query,rs,orderNode,i,pcArr,pcv_HaveRecords,attNode,subNode,queryQ,rsQ,tmpExportedFlag
	
	call opendb()
	
	query="SELECT idOrder,ord_OrderName,orderDate,idCustomer,total,processDate,ShippingFullName,shippingCompany,shippingAddress,shippingAddress2,shippingCity,shippingStateCode,shippingState,shippingZip,shippingCountryCode,pcOrd_shippingPhone,pcOrd_shippingFax,pcOrd_ShippingEmail,shipmentDetails,SRF,ordShiptype,ordPackageNum,pcOrd_ShipWeight,ord_DeliveryDate,orderStatus,pcOrd_PaymentStatus,paymentDetails,gwAuthCode,gwTransId,paymentCode,idAffiliate,affiliatePay,iRewardPoints,iRewardPointsCustAccrued,IDRefer,rmaCredit,taxAmount,taxDetails,ord_VAT,discountDetails,pcOrd_GcCode,pcOrd_GcUsed,pcOrd_CatDiscounts,comments,adminComments,returnDate,returnReason FROM Orders WHERE idorder=" & ordID_value & ";"
	'Last: 46
	
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcArr=rs.GetRows()
		pcv_HaveRecords=1
	end if
	set rs=nothing
	
	call closedb()
	
	IF pcv_HaveRecords=1 THEN
		i=0

		IF cm_LogTurnOn=1 THEN
			requestKey=CreateRequestRecord(pcv_PartnerID,5,ordID_value,0,0,0,0,0)
			cm_requestKey_value=requestKey
			Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
			tmpNode.Text=requestKey
			oRoot.appendChild(tmpNode)
		END IF
		
		Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
		tmpNode.Text=cm_SuccessCode
		oRoot.appendChild(tmpNode)
		
		if cm_ExportAdmin="1" then
			tmpExportedFlag=0
			call opendb()
			queryQ="SELECT pcXEL_ExportedID FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=2 AND pcXEL_ExportedID=" & ordID_value & ";"
			set rsQ=connTemp.execute(queryQ)
			if not rsQ.eof then
				tmpExportedFlag=1
			else
				queryQ="INSERT INTO pcXMLExportLogs (pcXP_ID,pcXEL_ExportedID,pcXEL_IDType) VALUES (" & pcv_PartnerID & "," & ordID_value & ",2);"
				set rsQ=connTemp.execute(queryQ)
			end if
			set rsQ=nothing
			call closedb()
			Set tmpNode=oXML.createNode(1,cm_ExportedFlag_name,"")
			tmpNode.Text=New_HTMLEncode(tmpExportedFlag)
			oRoot.appendChild(tmpNode)
		end if
		
		Set orderNode=oXML.createNode(1,cm_order,"")
		oRoot.appendChild(orderNode)
	
		Set attNode=oXML.createNode(1,ordID_name,"")
		attNode.Text=New_HTMLEncode(ordID_value)
		orderNode.appendChild(attNode)
		
		if ordName_ex=1 then
			Call XMLCreateNode(orderNode,ordName_name,New_HTMLEncode(trim(pcArr(1,i))))
		end if
		
		if ordDate_ex=1 then
			Call XMLCreateNode(orderNode,ordDate_name,ConvertToXMLDate(trim(pcArr(2,i))))
		end if
		
		if custID_ex=1 then
			Call XMLCreateNode(orderNode,custID_name,New_HTMLEncode(trim(pcArr(3,i))))
		end if
		
		if ordTotal_ex=1 then
			Call XMLCreateNode(orderNode,ordTotal_name,trim(pcArr(4,i)))
		end if
		
		if ordProcessedDate_ex=1 then
			tmpProcessedDate=trim(pcArr(5,i))
			if tmpProcessedDate<>"" then
				tmpProcessedDate=ConvertToXMLDate(tmpProcessedDate)
			end if
			Call XMLCreateNode(orderNode,ordProcessedDate_name,tmpProcessedDate)
		end if
		
		if ordCustDetails_ex=1 then
			Call XMLgetOrdCustDetails(orderNode,trim(pcArr(3,i)))
		end if
		
		if ordShippingAddress_ex=1 then
			Set attNode=oXML.createNode(1,ordShippingAddress_name,"")
			orderNode.appendChild(attNode)
			
			Call XMLCreateNode(attNode,ordShipName_name,New_HTMLEncode(trim(pcArr(6,i))))
			Call XMLCreateNode(attNode,ordShipCompany_name,New_HTMLEncode(trim(pcArr(7,i))))
			Call XMLCreateNode(attNode,ordShipAddress_name,New_HTMLEncode(trim(pcArr(8,i))))
			Call XMLCreateNode(attNode,ordShipAddress2_name,New_HTMLEncode(trim(pcArr(9,i))))
			Call XMLCreateNode(attNode,ordShipCity_name,New_HTMLEncode(trim(pcArr(10,i))))
			Call XMLCreateNode(attNode,ordShipStateCode_name,New_HTMLEncode(trim(pcArr(11,i))))
			Call XMLCreateNode(attNode,ordShipProvince_name,New_HTMLEncode(trim(pcArr(12,i))))
			Call XMLCreateNode(attNode,ordShipZip_name,New_HTMLEncode(trim(pcArr(13,i))))
			Call XMLCreateNode(attNode,ordShipCountryCode_name,New_HTMLEncode(trim(pcArr(14,i))))
			Call XMLCreateNode(attNode,ordShipPhone_name,New_HTMLEncode(trim(pcArr(15,i))))
			Call XMLCreateNode(attNode,ordShipFax_name,New_HTMLEncode(trim(pcArr(16,i))))
			Call XMLCreateNode(attNode,ordShipEmail_name,New_HTMLEncode(trim(pcArr(17,i))))
		end if
		
		if ordShipDetails_ex=1 then
			Set attNode=oXML.createNode(1,ordShipDetails_name,"")
			orderNode.appendChild(attNode)
			
			pshipmentDetails=trim(pcArr(18,i))
			pSRF=trim(pcArr(19,i))
			pOrdShipType=trim(pcArr(20,i))
			pOrdPackageNum=trim(pcArr(21,i))
			tmpShipWeight=trim(pcArr(22,i))
			Postage=0
			serviceHandlingFee=0
			If pSRF="1" then
				pshipmentDetails="Shipping charges to be determined."
			else
				shipping=split(pshipmentDetails,",")
				if ubound(shipping)>1 then
					if NOT isNumeric(trim(shipping(2))) then
						varShip="0"
						pshipmentDetails="No shipping charge (or no shipping required)."
					else
						Shipper=shipping(0)
						Service=shipping(1)
						Postage=trim(shipping(2))
						if ubound(shipping)=>3 then
							serviceHandlingFee=trim(shipping(3))
							if NOT isNumeric(serviceHandlingFee) then
								serviceHandlingFee=0
							end if
						else
							serviceHandlingFee=0
						end if
					end if
				else
					varShip="0"
					pshipmentDetails="No shipping charge (or no shipping required)."
				end if 
			end if
			if pOrdShipType=0 then
				pDisShipType="Residential"
			else
				pDisShipType="Commercial" 
			end if
			if pSRF="1" then
				Call XMLCreateNode(attNode,shipMethod_name,New_HTMLEncode(pshipmentDetails))
			else
				if varShip<>"0" AND Service<>"" then
					Call XMLCreateNode(attNode,shipMethod_name,New_HTMLEncode(Service))
				else
					Call XMLCreateNode(attNode,shipMethod_name,New_HTMLEncode(pshipmentDetails))
				end if 
			end if
			Call XMLCreateNode(attNode,shipType_name,New_HTMLEncode(pDisShipType))
			Call XMLCreateNode(attNode,shipFees_name,Postage)
			Call XMLCreateNode(attNode,handlingFees_name,serviceHandlingFee)
			Call XMLCreateNode(attNode,packageCount_name,pOrdPackageNum)
			
			Set subNode=oXML.createNode(1,shipWeight_name,"")
			attNode.appendChild(subNode)
			
			if scShipFromWeightUnit="KGS" then
				tmp_weight=Int(tmpShipWeight/1000)
				tmp_weight1=tmpShipWeight-(tmp_weight*1000)
							
				Call XMLCreateNode(subNode,Kgs_name,tmp_weight)
				Call XMLCreateNode(subNode,Grams_name,tmp_weight1)
			else
				tmp_weight=Int(tmpShipWeight/16)
				tmp_weight1=tmpShipWeight-(tmp_weight*16)
							
				Call XMLCreateNode(subNode,Pounds_name,tmp_weight)
				Call XMLCreateNode(subNode,Ounces_name,tmp_weight1)
			end if
			
			Call XMLgetOrdPkgDetails(attNode,ordID_value)
		end if
		
		if ordDeliveryDate_ex=1 then
			tmpDeliveryDate=trim(pcArr(23,i))
			if tmpDeliveryDate<>"" then
				tmpDeliveryDate=ConvertToXMLDate(tmpDeliveryDate)
			end if
			Call XMLCreateNode(orderNode,ordDeliveryDate_name,tmpDeliveryDate)
		end if
		
		if ordStatus_ex=1 then
			Call XMLCreateNode(orderNode,ordStatus_name,trim(pcArr(24,i)))
		end if
		
		if ordPaymentStatus_ex=1 then
			Call XMLCreateNode(orderNode,ordPaymentStatus_name,trim(pcArr(25,i)))
		end if
		
		if ordPayDetails_ex=1 then
			Set attNode=oXML.createNode(1,ordPayDetails_name,"")
			orderNode.appendChild(attNode)
			
			ppaymentDetails=trim(trim(pcArr(26,i)))
			payment = split(ppaymentDetails,"||")
			PaymentMethod=payment(0)
			if ubound(payment)>=1 then
				if IsNumeric(payment(1)) then
					PayCharge=payment(1)
				else
					PayCharge=0
				end if
			else
				PayCharge=0
			end If

			Call XMLCreateNode(attNode,paymentMethod_name,New_HTMLEncode(PaymentMethod))
			Call XMLCreateNode(attNode,paymentFees_name,PayCharge)
			Call XMLCreateNode(attNode,authorizationCode_name,trim(pcArr(27,i)))
			Call XMLCreateNode(attNode,transactionID_name,trim(pcArr(28,i)))
			Call XMLCreateNode(attNode,paymentGateway_name,trim(pcArr(29,i)))
		end if
		
		if ordAffiliate_ex=1 then
			Call XMLgetOrdAffDetails(orderNode,trim(pcArr(30,i)),trim(pcArr(31,i)))
		end if
		
		if ordRP_ex=1 then
			Call XMLCreateNode(orderNode,ordRP_name,trim(pcArr(32,i)))
		end if
		
		if ordAccruedRP_ex=1 then
			Call XMLCreateNode(orderNode,ordAccruedRP_name,trim(pcArr(33,i)))
		end if
		
		if ordReferrer_ex=1 then
			Call XMLgetOrdRefDetails(orderNode,trim(pcArr(34,i)))
		end if
		
		if ordRMACredit_ex=1 then
			Call XMLCreateNode(orderNode,ordRMACredit_name,trim(pcArr(35,i)))
		end if
		
		if ordTaxAmount_ex=1 then
			Call XMLCreateNode(orderNode,ordTaxAmount_name,trim(pcArr(36,i)))
		end if
		
		if ordVAT_ex=1 then
			Call XMLCreateNode(orderNode,ordVAT_name,trim(pcArr(38,i)))
		end if
		
		if ordTaxDetails_ex=1 then
			ptaxDetails=trim(trim(pcArr(37,i)))
			Set attNode=oXML.createNode(1,ordTaxDetails_name,"")
			orderNode.appendChild(attNode)

			if ptaxDetails<>"" then
				tmptaxArray=split(ptaxDetails,",")
				For k=0 to (ubound(tmptaxArray)-1)
					tmptaxDesc=split(tmptaxArray(k),"|")
					tmptaxName=tmptaxDesc(0)
					tmpTaxAmount=tmptaxDesc(1)
					Call XMLCreateNode(attNode,taxName_name,New_HTMLEncode(tmptaxName))
					Call XMLCreateNode(attNode,taxAmount_name,tmpTaxAmount)
				next
			end if
		end if		
		
		if ordDiscountDetails_ex=1 then
			pdiscountDetails=trim(trim(pcArr(39,i)))
			if (pdiscountDetails<>"") and (instr(pdiscountDetails,"- ||")>0) then
				Set attNode=oXML.createNode(1,ordDiscountDetails_name,"")
				orderNode.appendChild(attNode)
	
				if instr(pdiscountDetails,",") then
					DiscountDetailsArry=split(pdiscountDetails,",")
					intArryCnt=ubound(DiscountDetailsArry)
				else
					intArryCnt=0
				end if
									
				For k=0 to intArryCnt
				
					
					Set attNode2=oXML.createNode(1,ordDiscountDetail_name,"")
					attNode.appendChild(attNode2)
				
					if intArryCnt=0 then
						pTempDiscountDetails=pdiscountDetails
					else
						pTempDiscountDetails=DiscountDetailsArry(k)
					end if
					if instr(pTempDiscountDetails,"- ||") then
						discounts = split(pTempDiscountDetails,"- ||")
						tmpdiscountName=discounts(0)
						tmpdiscountAmount=discounts(1)
						if tmpdiscountAmount="" or isNULL(tmpdiscountAmount) OR tmpdiscountAmount="0" then
							tmpdiscountAmount=0
						end if
						if not IsNumeric(tmpdiscountAmount) then
							tmpdiscountAmount=0
						end if
					else
						tmpdiscountName=pTempDiscountDetails
						tmpdiscountAmount=0
					end if
					Call XMLCreateNode(attNode2,discountName_name,New_HTMLEncode(tmpdiscountName))
					Call XMLCreateNode(attNode2,discountAmount_name,tmpdiscountAmount)
				Next
			end if
		end if
		
		if GiftCertificate_ex=1 then
			Call XMLCreateNode(orderNode,GiftCertificate_name,New_HTMLEncode(trim(pcArr(40,i))))
		end if
		
		if GiftCertificateUsed_ex=1 then
			Call XMLCreateNode(orderNode,GiftCertificateUsed_name,trim(pcArr(41,i)))
		end if
		
		if ordCatDiscounts_ex=1 then
			Call XMLCreateNode(orderNode,ordCatDiscounts_name,New_HTMLEncode(trim(pcArr(42,i))))
		end if
		
		if ordCustomerComments_ex=1 then
			Call XMLCreateNode(orderNode,ordCustomerComments_name,New_HTMLEncode(trim(pcArr(43,i))))
		end if
		
		if ordAdminComments_ex=1 then
			Call XMLCreateNode(orderNode,ordAdminComments_name,New_HTMLEncode(trim(pcArr(44,i))))
		end if
		
		if ordReturnDate_ex=1 then
			tmpReturnDate=trim(pcArr(45,i))
			if tmpReturnDate<>"" then
				tmpReturnDate=ConvertToXMLDate(tmpReturnDate)
			end if
			Call XMLCreateNode(orderNode,ordReturnDate_name,tmpReturnDate)
		end if
		
		if ordReturnReason_ex=1 then
			tmpordReturnReason=trim(pcArr(46,i))
			if tmpordReturnReason<>"" then
				tmpordReturnReason=New_HTMLEncode(tmpordReturnReason)
			end if
			Call XMLCreateNode(orderNode,ordReturnReason_name,tmpordReturnReason)
		end if
		
		if ordShoppingCart_ex=1 then
			Call XMLgetOrdPrdsDetails(orderNode,ordID_value)
		end if
		
		Set pXML1=Server.CreateObject("MSXML2.DOMDocument"&scXML)
		pXML1.async=false
		pXML1.load(oXML)
		If (pXML1.parseError.errorCode <> 0) Then	
			Set oXML=nothing
			call InitResponseDocument(cm_GetOrderDetailsResponse_name)
			call XMLcreateError(pXML1.parseError.errorCode, pXML1.parseError.reason)
			call returnXML()
		End If
		set pXML1 = nothing
		
	ELSE
		call XMLcreateError(116,cm_errorStr_116b)
		call returnXML()
	END IF 'Have order record
	
End Sub

%>
