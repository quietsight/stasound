<%if IsNull(pcv_IsSupplier) or pcv_IsSupplier="" then
		pcv_IsSupplier=0
	end if
	
	if pcv_IsSupplier=0 then
		query="SELECT pcDropShipper_FirstName AS B,pcDropShipper_Lastname AS C,pcDropShipper_Company AS A,pcDropShipper_Email AS D,pcDropShipper_NoticeEmail AS E,pcDropShipper_NoticeType AS F,pcDropShipper_NoticeMsg AS G,pcDropShipper_Notifymanually AS H FROM pcDropShippers WHERE pcDropShipper_ID=" & pcv_DropShipperID
	else
		query="SELECT pcSupplier_FirstName AS B,pcSupplier_LastName AS C,pcSupplier_Company AS A,pcSupplier_Email AS D,pcSupplier_NoticeEmail AS E,pcSupplier_NoticeType AS F,pcSupplier_NoticeMsg AS G,pcSupplier_Notifymanually AS H FROM pcSuppliers WHERE pcSupplier_ID=" & pcv_DropShipperID & " AND pcSupplier_IsDropShipper=1"
	end if
	set rsQ1=Server.CreateObject("ADODB.Recordset") 
	set rsQ1=connTemp.execute(query)	
	pcv_strShowSection = 0
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
		pcv_strShowSection = -1
	end if
	set rsQ1=nothing
	
	If pcv_strShowSection = -1 Then
		
		If pcv_DS_NoticeM=0 OR pcv_ManualSend=1 Then
		
			'Get Additional Comments
			query="SELECT pcACom_Comments FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=4 AND pcDropShipper_ID=" & pcv_DropShipperID & " AND pcACom_IsSupplier=" & pcv_IsSupplier
			set rsQ1=Server.CreateObject("ADODB.Recordset")
			set rsQ1=connTemp.execute(query)
			pcACom_Comments=""
			if not rsQ1.eof then
				pcv_AdmComments=rsQ1("pcACom_Comments")
			end if
			set rsQ1=nothing
			'End of Get Additional Comments
		
			'Start Create Product List
			pcv_PrdList=""
			query="SELECT Products.idproduct,Products.Description,Products.SKU,ProductsOrdered.quantity,ProductsOrdered.idconfigSession,ProductsOrdered.pcPrdOrd_SelectedOptions,ProductsOrdered.pcPrdOrd_OptionsArray,ProductsOrdered.xfdetails FROM pcDropShippersSuppliers INNER JOIN (Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct) ON (pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & pcv_IsSupplier & ") WHERE ProductsOrdered.idorder=" & qry_ID & " AND ProductsOrdered.pcDropShipper_ID=" & pcv_DropShipperID
			set rsQ1=connTemp.execute(query)
			
			if not rsQ1.eof then
				pcv_PrdList=pcv_PrdList & ship_dictLanguage.Item(Session("language")&"_sds_notifyorder_1") & vbcrlf
				do while not rsQ1.eof					
					pcv_IDProduct=rsQ1("idproduct")
					
					query="UPDATE ProductsOrdered SET pcPrdOrd_SentNotice=1 WHERE idorder=" & qry_ID & " AND idproduct=" & rsQ1("idproduct")
					set rsQ2=Server.CreateObject("ADODB.Recordset")
					set rsQ2=connTemp.execute(query)
					set rsQ2=nothing
					pcv_PrdList=pcv_PrdList & rsQ1("Description") & " (" & rsQ1("SKU") & ") - Qty:" & rsQ1("quantity") & vbcrlf
					
					'BTO ADDON-S
					if scBTO=1 then
						pIdConfigSession=rsQ1("idconfigSession")
						if IsNull(pIdConfigSession) or pIdConfigSession="" then
							pIdConfigSession=0
						end if

						if pIdConfigSession<>"0" then
							query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
							set rsQ2=Server.CreateObject("ADODB.Recordset")
							set rsQ2=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsQ2=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							stringProducts=rsQ2("stringProducts")
							stringValues=rsQ2("stringValues")
							stringCategories=rsQ2("stringCategories")
							stringQuantity=rsQ2("stringQuantity")
							stringPrice=rsQ2("stringPrice")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							ArrQuantity=Split(stringQuantity, ",")
							ArrPrice=Split(stringPrice, ",")
							set rsQ2=nothing
							
							if stringProducts<>"na" then
								for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
									
									query="SELECT categories.categoryDesc,products.sku,products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
									set rsQ2=Server.CreateObject("ADODB.Recordset")
									set rsQ2=conntemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
										set rsQ2=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									pcv_strBtoItemCat=rsQ2("categoryDesc")
									pcv_strBtoItemCat=ClearHTMLTags2(pcv_strBtoItemCat,0)
									pcv_strBtoItemSKU=rsQ2("sku")
									pcv_strBtoItemSKU=ClearHTMLTags2(pcv_strBtoItemSKU,0)
									pcv_strBtoItemName = rsQ2("description")
									pcv_strBtoItemName=ClearHTMLTags2(pcv_strBtoItemName,0)
									set rsQ2=nothing
									
									query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i) &" AND specProduct=" & pcv_IDProduct
									set rsQ2=Server.CreateObject("ADODB.Recordset")
									set rsQ2=conntemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
										set rsQ2=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									dispStr = ""
									if rsQ2("displayQF")=True then
										dispStr = pcv_strBtoItemCat &": "& pcv_strBtoItemName & " (" & pcv_strBtoItemSKU & ") - QTY: " & ArrQuantity(i)
									else
										dispStr = pcv_strBtoItemCat &": "& pcv_strBtoItemName & " (" & pcv_strBtoItemSKU & ")"
									end if
									dispStr = replace(dispStr,"&quot;", chr(34))
									pcv_PrdList=pcv_PrdList & "     " & dispStr & vbCrLf
									set rsQ2=nothing
									
								next
							end if

							'BTO Additional Charges
	
							query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
							set rsQ2=Server.CreateObject("ADODB.Recordset")
							set rsQ2=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsQ2=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if

							stringCProducts=rsQ2("stringCProducts")
							stringCValues=rsQ2("stringCValues")
							stringCCategories=rsQ2("stringCCategories")
							
							set rsQ2=nothing
	
							ArrCProduct=Split(stringCProducts, ",")
							ArrCValue=Split(stringCValues, ",")
							ArrCCategory=Split(stringCCategories, ",")
		
							if ArrCProduct(0)<>"na" then
								pcv_PrdList=pcv_PrdList & vbcrlf
								pcv_PrdList=pcv_PrdList & "     " & dictLanguage.Item(Session("language")&"_adminMail_34") & vbcrlf
	
								for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
									
									query="SELECT categories.categoryDesc,products.description,products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
									set rsQ2=Server.CreateObject("ADODB.Recordset") 
									set rsQ2=conntemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
										set rsQ2=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									dispStr = ClearHTMLTags2(rsQ2("categoryDesc"),0)&": "&ClearHTMLTags2(rsQ2("description"),0) & " (" & ClearHTMLTags2(rsQ2("sku"),0) & ")"
									pcv_PrdList=pcv_PrdList & "     " & dispStr & vbcrlf
									set rsQ2=nothing
								next
							end if
							'BTO Additional Charges
						end if
					end if
					'BTO ADDON-E
					
					'Product Options Array
					pcv_strSelectedOptions = rsQ1("pcPrdOrd_SelectedOptions")
					pcv_strOptionsArray = rsQ1("pcPrdOrd_OptionsArray")
					
					if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
						pcv_strSelectedOptions = ""
					end if
			
					If len(pcv_strSelectedOptions)>0 Then
						tmpArr=split(pcv_strOptionsArray,"|")
						for q=lbound(tmpArr) to ubound(tmpArr)
							pcv_PrdList=pcv_PrdList & "     " & tmpArr(q) & vbcrlf
						next
					End if
					
					'Input fields information
					xfdetails=rsQ1("xfdetails")
					if IsNull(xfdetails) or xfdetails="" then
					else
						xfdetails=replace(xfdetails,"&lt;BR&gt;"," - ")
						xfdetails=replace(xfdetails,"<BR>"," - ")
					
						If len(xfdetails)>3 then
							xfarray=split(xfdetails,"|")
							for q=lbound(xfarray) to ubound(xfarray)
								pcv_PrdList=pcv_PrdList & "     " & xfarray(q) & vbcrlf
							next
						End If
					end if
					pcv_PrdList=pcv_PrdList & vbcrlf
					rsQ1.MoveNext
				loop
				pcv_PrdList=pcv_PrdList & vbcrlf
			end if
			set rsQ1=nothing
			'End of Create Product List
			
			'Start Create Shipping Address
			pcv_ShippAddrStr=""
			
				' get order details
				query="SELECT orders.idcustomer,orders.address,orders.City,orders.StateCode,orders.State,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingState,orders.shippingZip, orders.shippingCountryCode,orders.pcOrd_shippingPhone,orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.orderstatus,orders.iRewardPoints,orders.iRewardValue,orders.iRewardRefId,orders.iRewardPointsRef, orders.iRewardPointsCustAccrued, orders.ordPackageNum, customers.phone, orders.ord_DeliveryDate, orders.ord_VAT, orders.pcOrd_DiscountsUsed, orders.pcOrd_Payer FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" & qry_ID
				set rsQ1=server.CreateObject("ADODB.RecordSet")
				set rsQ1=conntemp.execute(query)
				if NOT rsQ1.eof then
					pidcustomer=rsQ1("idcustomer")
					paddress=rsQ1("address")
					pCity=rsQ1("City")
					pStateCode=rsQ1("StateCode")
					pState=rsQ1("State")
					if isNULL(pStateCode) OR pStateCode="" then
						pStateCode=pState
					end if
					pzip=rsQ1("zip")
					pCountryCode=rsQ1("CountryCode")
					pshippingAddress=rsQ1("shippingAddress")
					pshippingCity=rsQ1("shippingCity")
					pshippingStateCode=rsQ1("shippingStateCode")
					pshippingState=rsQ1("shippingState")
					if isNULL(pshippingStateCode) OR pshippingStateCode="" then
						pshippingStateCode=pshippingState
					end if
					pshippingZip=rsQ1("shippingZip")
					pshippingCountryCode=rsQ1("shippingCountryCode")
					pshippingPhone=rsQ1("pcOrd_shippingPhone")
					pShipmentDetails=rsQ1("ShipmentDetails")
					pPaymentDetails=rsQ1("PaymentDetails")
					pdiscountDetails=rsQ1("discountDetails")
					ptaxAmount=rsQ1("taxAmount")
					ptotal=rsQ1("total")
					pcomments=rsQ1("comments")
					pShippingFullName=rsQ1("ShippingFullName")
					paddress2=rsQ1("address2")
					pShippingCompany=rsQ1("ShippingCompany")
					pShippingAddress2=rsQ1("ShippingAddress2")
					ptaxDetails=rsQ1("taxDetails")
					pCurOrderStatus=rsQ1("orderStatus")
					piRewardPoints=rsQ1("iRewardPoints")
					piRewardValue=rsQ1("iRewardValue")
					piRewardRefId=rsQ1("iRewardRefId")
					piRewardPointsRef=rsQ1("iRewardPointsRef") 
					piRewardPointsCustAccrued=rsQ1("iRewardPointsCustAccrued")
					pOrdPackageNum=rsQ1("ordPackageNum")
					pPhone=rsQ1("phone")
					pord_DeliveryDate=rsQ1("ord_DeliveryDate")
					pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)
					pord_VAT=rsQ1("ord_VAT")
					strPcOrd_DiscountsUsed=rsQ1("pcOrd_DiscountsUsed")
					pcOrd_Payer=rsQ1("pcOrd_Payer")
				end if
				set rsQ1=nothing

				' get idCustomer
				query="SELECT email,name,lastName,idCustomer,customerCompany FROM customers WHERE idcustomer=" &pidcustomer
				set rsQ1=server.CreateObject("ADODB.RecordSet")
				set rsQ1=conntemp.execute(query)
				if NOT rsQ1.eof then
					pEmail=rsQ1("email") 
					pName=rsQ1("name")
					pLName=rsQ1("lastName") 
					pIdCustomer=rsQ1("idCustomer") 
					pCustomerCompany=rsQ1("customerCompany")
				end if
				set rsQ1=nothing
			
			If pcv_DS_NoticeType<>"1" then
				pcv_ShippAddrStr=ship_dictLanguage.Item(Session("language")&"_sds_notifyorder_2") & vbcrlf
				If Trim(pshippingAddress) <> "" Then 'Ship to Customer Shipping Address
					if pShippingFullName<>"" then
						pcv_ShippAddrStr=pcv_ShippAddrStr & pShippingFullName& vbCrLf
					end if
					if pShippingCompany<>"" then
						pcv_ShippAddrStr=pcv_ShippAddrStr & pShippingCompany& vbCrLf
					end if
					pcv_ShippAddrStr=pcv_ShippAddrStr & pshippingAddress & vbCrLf
					if pshippingAddress2<>"" then
						pcv_ShippAddrStr=pcv_ShippAddrStr & pshippingAddress2& vbCrLf
					end if
					pcv_ShippAddrStr=pcv_ShippAddrStr & pshippingCity & ", "
					if pshippingState = "" then
						pcv_ShippAddrStr=pcv_ShippAddrStr & pshippingStateCode & " "
					else
						pcv_ShippAddrStr=pcv_ShippAddrStr & pshippingState & " "
					end if
					pcv_ShippAddrStr=pcv_ShippAddrStr & pshippingZip & vbCrLf
					pcv_ShippAddrStr=pcv_ShippAddrStr & pshippingCountryCode & vbCrLf
					pcv_ShippAddrStr=pcv_ShippAddrStr & trim(pshippingPhone) & vbCrLf
					
				Else 'Ship to Customer Billing Address
				
					pcv_ShippAddrStr=pcv_ShippAddrStr & pName & " " & pLName & vbCrLf
					If Trim(pCustomerCompany) <> "" Then
						pcv_ShippAddrStr=pcv_ShippAddrStr & pCustomerCompany & vbCrLf
					End If
					pcv_ShippAddrStr=pcv_ShippAddrStr & paddress & vbCrLf
					if paddress2<>"" then
						pcv_ShippAddrStr=pcv_ShippAddrStr & paddress2 & vbCrLf	
					end if
					pcv_ShippAddrStr=pcv_ShippAddrStr & pCity & ", "
					if pState = "" then
						pcv_ShippAddrStr=pcv_ShippAddrStr & pStateCode & " "
					else
						pcv_ShippAddrStr=pcv_ShippAddrStr & pState & " "
					end if
					pcv_ShippAddrStr=pcv_ShippAddrStr & pzip & vbCrLf
					pcv_ShippAddrStr=pcv_ShippAddrStr & pCountryCode & vbCrLf
					pcv_ShippAddrStr=pcv_ShippAddrStr & pEmail & vbCrLf& vbCrLf
				End if
				
			Else 'Ship to store address
			
				pcv_ShippAddrStr=ship_dictLanguage.Item(Session("language")&"_sds_notifyorder_2a") & vbcrlf
				pcv_ShippAddrStr=pcv_ShippAddrStr & scCompanyName & VbCrLf
				pcv_ShippAddrStr=pcv_ShippAddrStr & scCompanyAddress & VbCrLf
				pcv_ShippAddrStr=pcv_ShippAddrStr & scCompanyCity & ", " & scCompanyState & " - " & scCompanyZip & VbCrLf
				pcv_ShippAddrStr=pcv_ShippAddrStr & scCompanyCountry & VbCrLf
				pcv_ShippAddrStr=pcv_ShippAddrStr & scFrmEmail & VbCrlf
				
			End if
			pcv_ShippAddrStr=pcv_ShippAddrStr & VbCrLf
			'End of Create Shipping Address
			
			'Create Ship Method
			pcv_ShipMethod=""
			query="SELECT shipmentDetails FROM Orders WHERE idorder=" & qry_ID & ";"
			set rsQ1=Server.CreateObject("ADODB.Recordset")
			set rsQ1=connTemp.execute(query)			
			if not rsQ1.eof then
				pcv_ShipMethod=ship_dictLanguage.Item(Session("language")&"_sds_notifyorder_3") & vbcrlf
				pcA=split(rsQ1("shipmentDetails"),",")
				if ubound(pcA)<2 then
					pcv_ShipMethod=pcv_ShipMethod & pcA(0) & vbcrlf
				else
					if pcA(0)<>"CUSTOM" then
					pcv_ShipMethod=pcv_ShipMethod & pcA(0) & vbcrlf
					end if
					pcv_ShipMethod=pcv_ShipMethod & pcA(1) & vbcrlf
				end if
				pcv_ShipMethod=pcv_ShipMethod & vbcrlf				
			end if
			set rsQ1=nothing
			'End of Ship Method
			
			'Create Link
			strPathInfo=""
			strPath=Request.ServerVariables("PATH_INFO")
			iCnt=0
			do while iCnt<2
				if mid(strPath,len(strPath),1)="/" then
					iCnt=iCnt+1
				end if
				if iCnt<2 then
					strPath=mid(strPath,1,len(strPath)-1)
				end if
			loop
	
			strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
			
			strPathInfo=replace(strPathInfo,"/" & scAdminFolderName,"")
						
			if Right(strPathInfo,1)="/" then
			else
				strPathInfo=strPathInfo & "/"
			end if
			
			strPathInfo=strPathInfo & "pc/sds_viewPastD.asp?idOrder=" & (scpre + int(qry_ID))
			
			'End of Create Link
			pcv_DropShipperMsg=replace(pcv_DS_NoticeMsg,"<br>", vbCrlf)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<STORE_NAME>",scCompanyName)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<TODAY_DATE>",todaydate)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<CUSTOMER_NAME>",pName&" "&pLName)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<ORDER_ID>",(scpre + int(qry_ID)))
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<ORDER_DATE>",todaydate)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<DROP_SHIPPER_COMPANY>",pcv_DS_Company)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<DROP_SHIPPER_NAME>",pcv_DS_Name)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<PRODUCTS>",pcv_PrdList)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<SHIPPING_INFO>",pcv_ShippAddrStr)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<LINK>",strPathInfo)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<SHIPPING_METHOD>",pcv_ShipMethod)
			
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;br&gt;", vbCrlf)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;STORE_NAME&gt;",scCompanyName)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;TODAY_DATE&gt;",todaydate)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;CUSTOMER_NAME&gt;",pName&" "&pLName)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;ORDER_ID&gt;",(scpre + int(qry_ID)))
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;ORDER_DATE&gt;",todaydate)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;DROP_SHIPPER_COMPANY&gt;",pcv_DS_Company)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;DROP_SHIPPER_NAME&gt;",pcv_DS_Name)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;PRODUCTS&gt;",pcv_PrdList)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;SHIPPING_INFO&gt;",pcv_ShippAddrStr)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;LINK&gt;",strPathInfo)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"&lt;SHIPPING_METHOD&gt;",pcv_ShipMethod)
	
			'pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"''",chr(39))
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"//","/")
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"http:/","http://")
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"https:/","https://")
	
			pcA=replace(pcv_DropShipperMsg,"[SUBJECT]","!!!!!")
			pcA=replace(pcA,"[/SUBJECT]","!!!!!")
			pcB=split(pcA,"!!!!!")
			
			pcv_DropShipperSbj=pcB(1)
	
			pcC=replace(pcB(2),"[BODY]","!!!!!")
			pcC=replace(pcC,"[/BODY]","!!!!!")
			pcD=split(pcC,"!!!!!")
	
			pcv_DropShipperMsg=pcD(1)
			
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg, "<CUSTOM_COPY>",pcv_AdmComments)
			pcv_DropShipperMsg=replace(pcv_DropShipperMsg, "&lt;CUSTOM_COPY&gt;",pcv_AdmComments)
			pcv_DropShipperMsg=pcv_DropShipperMsg & vbcrlf & scCompanyName & vbcrlf
			
			call sendmail (scCompanyName, scEmail, pcv_DS_NEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg, "&quot;", chr(34)))
			
		End if '// If pcv_DS_NoticeM=0 OR pcv_ManualSend=1 Then (End of Automatically Send Notification Email)
	End if '// If pcv_strShowSection = -1 Then (Have Drop-Shipper)
%>