<%if IsNull(pcv_PackageID) or pcv_PackageID="" or (not isNumeric(pcv_PackageID)) then
	pcv_PackageID=0
end if

IF pcv_PackageID<>0 THEN

	query="SELECT pcPackageInfo_UPSServiceCode, pcPackageInfo_UPSPackageType,  pcPackageInfo_ShipMethod, pcPackageInfo_ShippedDate, pcPackageInfo_TrackingNumber, pcPackageInfo_Comments, pcPackageInfo_MethodFlag FROM pcPackageInfo WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
	set rsQ1=connTemp.execute(query)
	
	if not rsQ1.eof then
		pcv_PK_UPSServiceCode=rsQ1("pcPackageInfo_UPSServiceCode")
		pcv_PK_UPSPackageType=rsQ1("pcPackageInfo_UPSPackageType")
		pcv_PK_ShipMethod=rsQ1("pcPackageInfo_ShipMethod")
		pcv_PK_ShippedDate=rsQ1("pcPackageInfo_ShippedDate")
		pcv_PK_TrackingNumber=rsQ1("pcPackageInfo_TrackingNumber")
		pcv_PK_Comments=rsQ1("pcPackageInfo_Comments")
		pcv_PK_MethodFlag=rsQ1("pcPackageInfo_MethodFlag")
		if pcv_PK_MethodFlag="2" AND pcv_PK_ShipMethod="" then
			select case pcv_PK_UPSServiceCode
				case "01"
					pcv_PK_ShipMethod="UPS Next Day Air"
				case "02"
					pcv_PK_ShipMethod="UPS 2nd Day Air"
				case "03"
					pcv_PK_ShipMethod="UPS Ground"
				case "07"
					pcv_PK_ShipMethod="UPS Worldwide Express"
				case "08"
					pcv_PK_ShipMethod="UPS Worldwide Expedited"
				case "11"
					pcv_PK_ShipMethod="UPS Standard To Canada"
				case "12"
					pcv_PK_ShipMethod="UPS 3 Day Select"
				case "13"
					pcv_PK_ShipMethod="UPS Next Day Air Saver"
				case "14"
					pcv_PK_ShipMethod="UPS Next Day Air"
				case "54"
					pcv_PK_ShipMethod="UPS Worldwide Express Plus"
				case "59"
					pcv_PK_ShipMethod="UPS 2nd Day Air A.M."
				case "65"
					pcv_PK_ShipMethod="UPS Express Saver"
			end select
		end if	
			
		set rsQ1=nothing
		
		query="SELECT idorder FROM ProductsOrdered WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
		set rsQ1=connTemp.execute(query)
		if not rsQ1.eof then
			qry_ID=rsQ1("idorder")
		end if
		set rsQ1=nothing
		
		'Get Additional Comments
		query="SELECT pcACom_Comments FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID
		set rsQ1=connTemp.execute(query)

		pcv_AdmComments=""
		if not rsQ1.eof then
			pcv_AdmComments=rsQ1("pcACom_Comments")
		end if
		set rsQ1=nothing
		'End of Get Additional Comments
		
		'Start Create Product List
		pcv_PrdList=""
		query="SELECT Products.Description,ProductsOrdered.quantity FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qry_ID & " AND ProductsOrdered.pcPackageInfo_ID=" & pcv_PackageID & " AND ProductsOrdered.pcPrdOrd_Shipped=1;"
		set rsQ1=connTemp.execute(query)
			
		if not rsQ1.eof then
			pcv_PrdList=pcv_PrdList & ship_dictLanguage.Item(Session("language")&"_partship_msg_2") & vbcrlf & vbcrlf
			do while not rsQ1.eof
				pcv_PrdList=pcv_PrdList & rsQ1("Description") & " - Qty:" & rsQ1("quantity") & vbcrlf
				rsQ1.MoveNext
			loop
			pcv_PrdList=pcv_PrdList & vbcrlf
		end if
		set rsQ1=nothing
		'End of Create Product List
			
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
					
		if Right(strPathInfo,1)="/" then
		else
			strPathInfo=strPathInfo & "/"
		end if
			
		strPathInfo=strPathInfo & scAdminFolderName & "/OrdDetails.asp?id=" & qry_ID
		'End of Create Link
		
		if pcv_LastShip="1" then ' This is the last (or only) shipment
			pcv_DropShipperSbj=ship_dictLanguage.Item(Session("language")&"_partship_sbj_9")
			pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<ORDER_ID>",(scpre + int(qry_ID)))	
			pcv_DropShipperMsg=pcv_DropShipperSbj & vbcrlf & vbcrlf
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_8") & vbcrlf & vbcrlf
			pcv_DropShipperMsg1=ship_dictLanguage.Item(Session("language")&"_partship_msg_8a") & vbcrlf
		else
			pcv_DropShipperSbj=ship_dictLanguage.Item(Session("language")&"_partship_sbj_1")
			pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<ORDER_ID>",(scpre + int(qry_ID)))	
			pcv_DropShipperMsg=pcv_DropShipperSbj & vbcrlf & vbcrlf
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_1") & vbcrlf & vbcrlf
			pcv_DropShipperMsg1=ship_dictLanguage.Item(Session("language")&"_partship_msg_1a") & vbcrlf & vbcrlf
		end if
		if pcv_AdmComments<>"" then
			pcv_DropShipperMsg=pcv_DropShipperMsg & vbcrlf & replace(pcv_AdmComments,"''","'") & vbcrlf & vbcrlf
		end if
		pcv_DropShipperMsg=pcv_DropShipperMsg & pcv_PrdList & vbcrlf
		pcv_DropShipperMsg1=pcv_DropShipperMsg1 & pcv_PrdList & vbcrlf
		pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_3") & vbcrlf
		pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_3") & vbcrlf
		if pcv_PK_ShipMethod<>"" then
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_4") & pcv_PK_ShipMethod & vbcrlf
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_4") & pcv_PK_ShipMethod & vbcrlf
		end if
		if pcv_PK_TrackingNumber<>"" then	
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_5") & pcv_PK_TrackingNumber & vbcrlf
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_5") & pcv_PK_TrackingNumber & vbcrlf
		end if
		if not IsNull(pcv_PK_ShippedDate) then
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_6") & ShowDateFrmt(pcv_PK_ShippedDate) & vbcrlf
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_6") & ShowDateFrmt(pcv_PK_ShippedDate) & vbcrlf
		end if
		if pcv_PK_TrackingNumber<>"" then		
			'//  Start: Tracking Link
			pcv_strTempLink=""
			if instr(ucase(pcv_PK_ShipMethod),"UPS:") then
				pcv_strTempLink = scStoreURL & "/" & scPcFolder & "/pc/custUPSTracking.asp?itracknumber=" & pcv_PK_TrackingNumber & vbCrLf & vbCrLf
				pcv_strTempUPSLink=replace(pcv_strTempLink,"//","/")
				pcv_strTempLink=replace(pcv_strTempLink,"http:/","http://")
			elseif instr(ucase(pcv_PK_ShipMethod),"FEDEX:") then
				pcv_strTempLink = "http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=" & pcv_PK_TrackingNumber
			end if	
			if pcv_strTempLink<>"" then
				pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_9") & pcv_strTempLink & vbcrlf
				pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_9") & pcv_strTempLink & vbcrlf
			end if
			'//  End: Tracking Link	
		end if
		if pcv_PK_Comments<>"" then
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_7") & vbcrlf & pcv_PK_Comments & vbcrlf
		end if
		
		if (pcv_PrdList="") and (pcv_PK_Comments<>"") then
			pcv_DropShipperMsg1=ship_dictLanguage.Item(Session("language")&"_partship_msg_7") & vbcrlf & pcv_PK_Comments & vbcrlf
		end if
		
		pcv_DropShipperMsg1=pcv_DropShipperMsg1 & vbcrlf & strPathInfo & vbcrlf
			
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"''",chr(39))
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"//","/")
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"http:/","http://")
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"https:/","https://")
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<ORDER_ID>",(scpre + int(qry_ID)))
		
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"''",chr(39))
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"//","/")
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"http:/","http://")
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"https:/","https://")
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"<ORDER_ID>",(scpre + int(qry_ID)))
		
		if pcv_ResendShip="1" then
		
			query="SELECT pcACom_Comments FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=3;"
			set rsQ1=connTemp.execute(query)
			
			pcv_AdmComments1=""
			if not rsQ1.eof then
				pcv_AdmComments1=rsQ1("pcACom_Comments")
			end if
			set rsQ1=nothing
			
			if pcv_AdmComments1<>"" then
				pcv_DropShipperMsg=replace(pcv_AdmComments1,"''","'") & vbcrlf & "------------------------------------------------" & vbcrlf  & vbcrlf & pcv_DropShipperMsg 
			end if
		end if
			
		if (pcv_SendCust="1") and (pcv_PrdList<>"") then
			query="SELECT Customers.email, Orders.pcOrd_ShippingEmail FROM Customers INNER JOIN Orders ON Customers.idcustomer = Orders.idCustomer WHERE Orders.idOrder=" & qry_ID & ";"
			set rsQ1=connTemp.execute(query)
			pEmail=rsQ1("email")
			pShippingEmail=rsQ1("pcOrd_ShippingEmail")
			
			set rsQ1=nothing
			pcv_DropShipperMsg=pcv_DropShipperMsg & vbcrlf & dictLanguage.Item(Session("language")&"_sendMail_36") & scCompanyName & "." & vbCrLf & vbCrLf
			
			'************************************************************************************************************************
			' START: Shipper Information
			' The e-mail is sent so that the shipment appears to be coming from your store
			' If you want to include the drop-shipper's information instead, uncomment the following 4 commented lines of code
			'************************************************************************************************************************

			'if pcv_UseDropShipperInfo="1" then
			'	call sendmail (pcv_DS_Name, pcv_DS_Email, pEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg, "&quot;", chr(34)))
			'else
				call sendmail (scCompanyName, scEmail, pEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg, "&quot;", chr(34)))
				'//Send email to shipping email if it is different and exist
				if trim(pShippingEmail)<>"" AND trim(pShippingEmail)<>trim(pEmail) then
					call sendmail (scCompanyName, scEmail, pShippingEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg, "&quot;", chr(34)))
				end if
			'end if
			
			'************************************************************************************************************************
			' END: Shipper Information
			'************************************************************************************************************************
			
		end if
		if pcv_SendAdmin="1" then
			if (pcv_PrdList="") and (pcv_PK_Comments<>"") then
				pcv_DropShipperSbj=ship_dictLanguage.Item(Session("language")&"_partship_sbj_1a")
				pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<ORDER_ID>",(scpre + int(qry_ID)))
			end if
			if pcv_UseDropShipperInfo="1" then
				call sendmail (pcv_DS_Name, pcv_DS_Email, scFrmEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg1, "&quot;", chr(34)))
			else
				call sendmail (scCompanyName, scEmail, scFrmEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg1, "&quot;", chr(34)))
			end if
		end if
		
	End if 'Have Package Info
	set rsQ1=nothing
	
END IF%>
