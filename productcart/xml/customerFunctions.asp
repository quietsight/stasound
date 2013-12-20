<%

Sub CheckSrcCustomersTags()
Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	Set fNode=iRoot.selectSingleNode(cm_filters_name)
	if fNode is Nothing then
		exit Sub
	end if
	if fNode.Text="" then
		exit Sub
	end if
	Set ChildNodes = fNode.childNodes
	
	srcIncSuspended_ex=1
	srcIncSuspended_value=0
	srcIncSuspended_ex=1
	srcIncSuspended_value=0
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		
		Select Case tmpNodeName
			Case srcFirstName_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcFirstName_ex=1
					srcFirstName_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcLastName_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcLastName_ex=1
					srcLastName_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcCompany_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcCompany_ex=1
					srcCompany_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcEmail_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcEmail_ex=1
					srcEmail_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcCity_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcCity_ex=1
					srcCity_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcCountryCode_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcCountryCode_ex=1
					srcCountryCode_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcPhone_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcPhone_ex=1
					srcPhone_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcCustomerType_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcCustomerType_ex=1
					srcCustomerType_value=tmpNodeValue
					if srcCustomerType_value>1 then
						srcCustomerType_value=1
					end if
				end if
			Case srcPricingCategory_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcPricingCategory_ex=1
					srcPricingCategory_value=tmpNodeValue
				end if
			Case srcCustomerField_name:
				call CheckValidXMLTag(strNode,0,5,"")
				
				if tmpNodeValue<>"" then
					srcCustomerField_ex=1
					srcCustomerField_value=getUserInput(tmpNodeValue,0)
				end if
			Case srcIncLocked_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcIncLocked_ex=1
					srcIncLocked_value=tmpNodeValue
					if srcIncLocked_value>1 then
						srcIncLocked_value=1
					end if
				end if
			Case srcIncSuspended_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcIncSuspended_ex=1
					srcIncSuspended_value=tmpNodeValue
					if srcIncSuspended_value>1 then
						srcIncSuspended_value=1
					end if
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

Sub CheckNewCustomersTags()
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

Sub CheckGetCustomerDetailsTags()
	Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	
	Call CheckRequiredXMLTag(custID_name)
	Set strNode=iRoot.selectSingleNode(custID_name)
	call CheckValidXMLTag(strNode,1,1,"")
	custID_ex=1
	custID_value=tmpNode.Text

	Set rNode=iRoot.selectSingleNode(cm_requests_name)
	if rNode is Nothing then
		Call SetDefaultCustomerDetailsTags()
		exit Sub
	else
		if rNode.Text="" then
			Call SetDefaultCustomerDetailsTags()
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
					Call SetDefaultCustomerDetailsTags()
				Case cm_requestAll_name:
					Call SetAllCustomerDetailsTags()
				Case custName_name:
					custName_ex=1
				'Case custPassword_name:
					'custPassword_ex=1
				Case custType_name:
					custType_ex=1
				Case custBillingAddress_name:
					custBillingAddress_ex=1
				Case custShippingAddress_name:
					custShippingAddress_ex=1
				Case custRPBalance_name:
					custRPBalance_ex=1
				Case custRPUsed_name:
					custRPUsed_ex=1
				Case custPricingCategory_name:
					custPricingCategory_ex=1
				Case custField_name:
					custField_ex=1
				Case custNewsletter_name:
					custNewsletter_ex=1
				Case custTotalOrders_name:
					custTotalOrders_ex=1
				Case custTotalSales_name:
					custTotalSales_ex=1
				Case custCreatedDate_name:
					custCreatedDate_ex=1
				Case custStatus_name:
					custStatus_ex=1
				Case Else:
					call XMLcreateError(106,cm_errorStr_106 & tmpNodeValue)
					call returnXML()
			End Select
		else
			call XMLcreateError(106,cm_errorStr_106 & tmpNodeName)
			call returnXML()
		end if
	Next
End Sub

Sub SetDefaultCustomerDetailsTags()
	custName_ex=0
	'custPassword_ex=1
	custType_ex=1
	custBillingAddress_ex=1
	custShippingAddress_ex=1
	custRPBalance_ex=0
	custRPUsed_ex=0
	custPricingCategory_ex=1
	custField_ex=0
	custNewsletter_ex=1
	custTotalOrders_ex=0
	custTotalSales_ex=0
	custCreatedDate_ex=1
	custStatus_ex=1
End Sub

Sub SetAllCustomerDetailsTags()
	custName_ex=1
	'custPassword_ex=1
	custType_ex=1
	custBillingAddress_ex=1
	custShippingAddress_ex=1
	custRPBalance_ex=1
	custRPUsed_ex=1
	custPricingCategory_ex=1
	custField_ex=1
	custNewsletter_ex=1
	custTotalOrders_ex=1
	custTotalSales_ex=1
	custCreatedDate_ex=1
	custStatus_ex=1
End Sub

Function CreateCustQuery(Desc,keynum)
Dim m
Dim tmpStr,keywordArray,keylink,keydesc

	tmpStr=""

	Select Case keynum
		Case 1: keydesc="customers.[Name]"
		Case 2: keydesc="customers.LastName"
		Case 3: keydesc="customers.customerCompany"
		Case 4: keydesc="customers.email"
		Case 5: keydesc="customers.city"
		Case 6: keydesc="customers.countryCode"
		Case 7: keydesc="customers.phone"
		Case 8: keydesc="pcCustomerFieldsValues.pcCFV_Value"
	End Select

	if Instr(Desc," AND ")>0 then
		keywordArray=split(Desc," AND ")
		keylink=" AND "
	else
	if Instr(Desc,",")>0 then
		keywordArray=split(Desc,",")
		keylink=" OR "
	else
		if Instr(Desc," OR ")>0 then
			keywordArray=split(Desc," OR ")
			keylink=" OR "
		else
			keywordArray=split(Desc,"***")
			keylink=" OR "
		end if
	end if
	end if

			
	For m=lbound(keywordArray) to ubound(keywordArray)
	if trim(keywordArray(m))<>"" then
		if tmpStr<>"" then
		tmpStr=tmpStr & keylink
		end if
		tmpStr=tmpStr & "(" & keydesc & " like '%"&trim(keywordArray(m))&"%')"
	end if
	Next
	
	if tmpStr<>"" then
		tmpStr="(" & tmpStr & ")"
	else
		tmpStr="(" & keydesc & " like '%"&Desc&"%')"
	end if

	CreateCustQuery=tmpStr
	
End Function

Function GenSrcCustomersQuery()
	Dim query,query1,query2,query3
	Dim strORD1,tmpKey,tmpFromDate,tmpToDate

	strORD1="customers.idcustomer ASC"

	' create sql statement
	query1=""
	query2=""
	query3=""
	
	if srcFirstName_value<>"" then
		tmpKey=srcFirstName_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,1)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	if srcLastName_value<>"" then
		tmpKey=srcLastName_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,2)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	if srcCompany_value<>"" then
		tmpKey=srcCompany_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,3)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	if srcEmail_value<>"" then
		tmpKey=srcEmail_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,4)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	if srcCity_value<>"" then
		tmpKey=srcCity_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,5)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	if srcCountryCode_value<>"" then
		tmpKey=srcCountryCode_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,6)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	if srcPhone_value<>"" then
		tmpKey=srcPhone_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,7)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if

	if srcCustomerType_ex=1 then
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & "customerType=" & srcCustomerType_value
	end if
	
	if srcPricingCategory_ex=1 then
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & "idCustomerCategory=" & srcPricingCategory_value
	end if
	
	if srcIncLocked_value<>"" then
		if srcIncLocked_value=0 then
			if query1<>"" then
				query1=query1 & " AND "
			end if
			query1=query1 & "customerType<=1"
		end if
	end if
	
	if srcIncSuspended_value<>"" then
		if srcIncSuspended_value=0 then
			if query1<>"" then
				query1=query1 & " AND "
			end if
			query1=query1 & "suspend=0"
		end if
	end if
	
	if srcCustomerField_value<>"" then
		tmpKey=srcCustomerField_value
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateCustQuery(tmpKey,8)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & "((pcCustomerFieldsValues.idcustomer=customers.idcustomer) AND (" & query2 & "))"
		query3=",pcCustomerFieldsValues"
	end if
	
	If srcFromDate_ex=1 then
		tmpFromDate=srcFromDate_Value
		if SQL_Format="1" then
			tmpFromDate=Day(tmpFromDate)&"/"&Month(tmpFromDate)&"/"&Year(tmpFromDate)
		else
			tmpFromDate=Month(tmpFromDate)&"/"&Day(tmpFromDate)&"/"&Year(tmpFromDate)
		end if
		if query1<>"" then
			query1=query1 & " AND "
		end if
		if scDB="Access" then
			query1=query1 & " customers.pcCust_DateCreated>=#" & tmpFromDate & "# "
		else
			query1=query1 & " customers.pcCust_DateCreated>='" & tmpFromDate & "' "
		end if
	End if
	
	If srcToDate_ex=1 then
		tmpToDate=CDate(srcToDate_Value)
		if SQL_Format="1" then
			tmpToDate=Day(tmpToDate)&"/"&Month(tmpToDate)&"/"&Year(tmpToDate)
		else
			tmpToDate=Month(tmpToDate)&"/"&Day(tmpToDate)&"/"&Year(tmpToDate)
		end if
		if query1<>"" then
			query1=query1 & " AND "
		end if
		if scDB="Access" then
			query1=query1 & " customers.pcCust_DateCreated<=#" & tmpToDate & "# "
		else
			query1=query1 & " customers.pcCust_DateCreated<='" & tmpToDate & "' "
		end if
	End if
	
	if cm_ExportAdmin="1" AND srcHideExported_value="1" then
		query1=query1 & " AND (Customers.idcustomer NOT IN (SELECT DISTINCT pcXEL_ExportedID FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=1)) "
	end if
	
	query="SELECT DISTINCT customers.idcustomer FROM customers " & query3
	if query1<>"" then
		query=query & " WHERE " & query1
	end if
	
	query=query&" ORDER BY "& strORD1
	
	GenSrcCustomersQuery=query
	
End Function

Sub RunSrcCustomers()
Dim query,rs1,resultCount,pcArr
Dim requestKey,i,strNode
on error resume next
	query=GenSrcCustomersQuery()
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
		requestKey=CreateRequestRecord(pcv_PartnerID,1,0,0,0,resultCount,0,0)
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
	
		Set tmpNode=oXML.createNode(1,cm_customers,"")
		oRoot.appendChild(tmpNode)
	
		For i=0 to resultCount-1
			Set strNode=oXML.createNode(1,custID_name,"")
			strNode.Text=pcArr(0,i)
			tmpNode.appendChild(strNode)
		Next
	
	end if
	
End Sub

Function GenNewCustomersQuery()
Dim strSQL, query, tmpFromDate
Dim rs1,tmpLastID
on error resume next

	tmpLastID=0
	
	call opendb()
	
	query="SELECT pcXL_LastID FROM pcXMLLogs WHERE pcXP_id=" & pcv_PartnerID & " AND pcXL_RequestType=7 ORDER BY pcXL_LastID DESC;"
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
		strSQL=strSQL & " customers.idcustomer>" & tmpLastID
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
			strSQL=strSQL & " customers.pcCust_DateCreated>=#" & tmpFromDate & "# "
		else
			strSQL=strSQL & " customers.pcCust_DateCreated>='" & tmpFromDate & "' "
		end if
	END IF
	
	query="SELECT customers.idcustomer FROM Customers WHERE " & strSQL & " AND customers.customerType<=1 AND customers.suspend=0 ORDER BY customers.idcustomer ASC;"
	
	GenNewCustomersQuery=query
	
End Function

Sub RunNewCustomers()
Dim query,rs1,resultCount,pcArr
Dim requestKey,i,strNode,tmpLastID
on error resume next
	query=GenNewCustomersQuery()
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
	query="SELECT customers.idcustomer FROM Customers ORDER BY customers.idcustomer DESC;"
	set rs1=connTemp.execute(query)
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		tmpLastID=rs1("idcustomer")
	end if
	set rs1=nothing
	
	call closedb()
	
	IF cm_LogTurnOn=1 THEN
		requestKey=CreateRequestRecord(pcv_PartnerID,7,0,0,0,resultCount,tmpLastID,0)
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
	
		Set tmpNode=oXML.createNode(1,cm_customers,"")
		oRoot.appendChild(tmpNode)
	
		For i=0 to resultCount-1
			Set strNode=oXML.createNode(1,custID_name,"")
			strNode.Text=pcArr(0,i)
			tmpNode.appendChild(strNode)
		Next
	
	end if
	
End Sub

Sub XMLgetCustCustomField(parentNode,tmpIDCustomer)
Dim query,rs1,tmpCFName,attNode,pcArr,intCount,i
	
	call opendb()
	
	query="SELECT pcCustomerFields.pcCField_ID,pcCustomerFields.pcCField_Name,pcCustomerFieldsValues.pcCFV_Value FROM pcCustomerFields INNER JOIN pcCustomerFieldsValues ON pcCustomerFields.pcCField_ID=pcCustomerFieldsValues.pcCField_ID WHERE pcCustomerFieldsValues.idcustomer=" & tmpIDCustomer & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
			tmpCFID=trim(pcArr(0,i))
			tmpCFName=trim(pcArr(1,i))
			tmpCFValue=trim(pcArr(2,i))
		
			Set attNode=oXML.createNode(1,custField_name,"")
			parentNode.appendChild(attNode)
		
			Call XMLCreateNode(attNode,fieldID_name,tmpCFID)
			Call XMLCreateNode(attNode,fieldName_name,New_HTMLEncode(tmpCFName))
			Call XMLCreateNode(attNode,fieldValue_name,New_HTMLEncode(tmpCFValue))
		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub


Sub XMLgetCustShipAddress(parentNode,tmpIDCustomer)
Dim query,rs1,tmpCFName,attNode,pcArr,intCount,i
	
	call opendb()
	
	query="SELECT idRecipient,recipient_NickName,recipient_FirstName,recipient_LastName,recipient_Email,recipient_Company,recipient_Address,recipient_Address2,recipient_City,recipient_StateCode,recipient_State,recipient_Zip,recipient_CountryCode,recipient_Phone,recipient_Fax FROM recipients WHERE idcustomer=" & tmpIDCustomer & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArr=rs1.getRows()
		set rs1=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
					
			Set attNode=oXML.createNode(1,custShippingAddress_name,"")
			parentNode.appendChild(attNode)
			
			Call XMLCreateNode(attNode,custShipAddressID_name,New_HTMLEncode(trim(pcArr(0,i))))
			Call XMLCreateNode(attNode,custShipNickName_name,New_HTMLEncode(trim(pcArr(1,i))))
			Call XMLCreateNode(attNode,custShipFirstName_name,New_HTMLEncode(trim(pcArr(2,i))))
			Call XMLCreateNode(attNode,custShipLastName_name,New_HTMLEncode(trim(pcArr(3,i))))
			Call XMLCreateNode(attNode,custShipEmail_name,New_HTMLEncode(trim(pcArr(4,i))))
			Call XMLCreateNode(attNode,custShipCompany_name,New_HTMLEncode(trim(pcArr(5,i))))
			Call XMLCreateNode(attNode,custShipAddress_name,New_HTMLEncode(trim(pcArr(6,i))))
			Call XMLCreateNode(attNode,custShipAddress2_name,New_HTMLEncode(trim(pcArr(7,i))))
			Call XMLCreateNode(attNode,custShipCity_name,New_HTMLEncode(trim(pcArr(8,i))))
			Call XMLCreateNode(attNode,custShipStateCode_name,New_HTMLEncode(trim(pcArr(9,i))))
			Call XMLCreateNode(attNode,custShipProvince_name,New_HTMLEncode(trim(pcArr(10,i))))
			Call XMLCreateNode(attNode,custShipZip_name,New_HTMLEncode(trim(pcArr(11,i))))
			Call XMLCreateNode(attNode,custShipCountryCode_name,New_HTMLEncode(trim(pcArr(12,i))))
			Call XMLCreateNode(attNode,custShipPhone_name,New_HTMLEncode(trim(pcArr(13,i))))
			Call XMLCreateNode(attNode,custShipFax_name,New_HTMLEncode(trim(pcArr(14,i))))

		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub XMLgetCustPricingCat(parentNode,tmpIDPricingCat)
Dim query,rs1,tmpPriceCatName,attNode
	
	call opendb()
	
	query="SELECT pcCC_Name FROM pcCustomerCategories WHERE idCustomerCategory=" & tmpIDPricingCat & ";"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		tmpPriceCatName=trim(rs1("pcCC_Name"))
		set rs1=nothing
		
		Set attNode=oXML.createNode(1,custPricingCategory_name,"")
		parentNode.appendChild(attNode)
		
		Call XMLCreateNode(attNode,pricingCatID_name,tmpIDPricingCat)
		Call XMLCreateNode(attNode,pricingCatName_name,New_HTMLEncode(tmpPriceCatName))
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub RunGetCustomerDetails()
	Dim query,rs,custNode,i,pcArr,pcv_HaveRecords,attNode,subNode,queryQ,rsQ,tmpExportedFlag
	
	call opendb()
	
	query="SELECT name,lastName,email,password,customerType,customerCompany,phone,fax,address,address2,city,stateCode,state,zip,countryCode,shippingCompany,shippingaddress,shippingAddress2,shippingcity,shippingStateCode,shippingState,shippingZip,shippingCountryCode,iRewardPointsAccrued,iRewardPointsUsed,dtRewardsStarted,idCustomerCategory,RecvNews,TotalOrders,TotalSales,pcCust_DateCreated,suspend FROM customers WHERE idcustomer=" & custID_value & ";"
	'Last: 31
	
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
			requestKey=CreateRequestRecord(pcv_PartnerID,4,custID_value,0,0,0,0,0)
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
			queryQ="SELECT pcXEL_ExportedID FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=1 AND pcXEL_ExportedID=" & custID_value & ";"
			set rsQ=connTemp.execute(queryQ)
			if not rsQ.eof then
				tmpExportedFlag=1
			else
				queryQ="INSERT INTO pcXMLExportLogs (pcXP_ID,pcXEL_ExportedID,pcXEL_IDType) VALUES (" & pcv_PartnerID & "," & custID_value & ",1);"
				set rsQ=connTemp.execute(queryQ)
			end if
			set rsQ=nothing
			call closedb()
			Set tmpNode=oXML.createNode(1,cm_ExportedFlag_name,"")
			tmpNode.Text=New_HTMLEncode(tmpExportedFlag)
			oRoot.appendChild(tmpNode)
		end if
		
		Set custNode=oXML.createNode(1,cm_customer,"")
		oRoot.appendChild(custNode)
	
		Set attNode=oXML.createNode(1,custID_name,"")
		attNode.Text=New_HTMLEncode(custID_value)
		custNode.appendChild(attNode)
		
		if custName_ex=1 then
			Call XMLCreateNode(custNode,custName_name,New_HTMLEncode(trim(pcArr(0,i)) & " " & trim(pcArr(1,i))))
		end if
		'if custPassword_ex=1 then
		'	tmpPassword = enDeCrypt(trim(pcArr(3,i)), scCrypPass)
		'	If IsEncodingIssues=1 Then tmpPassword = EncodeExtendedAsciiChars(tmpPassword)				
		'	Call XMLCreateNode(custNode,custPassword_name,New_HTMLEncode(tmpPassword))
		'end if
		if custType_ex=1 then
			tmpCustType=trim(pcArr(4,i))
			if cint(tmpCustType)>1 then
				tmpCustType=0
			end if
			Call XMLCreateNode(custNode,custType_name,tmpCustType)
		end if
		
		if custBillingAddress_ex=1 then
			Set attNode=oXML.createNode(1,custBillingAddress_name,"")
			custNode.appendChild(attNode)
			Call XMLCreateNode(attNode,custFirstName_name,New_HTMLEncode(trim(pcArr(0,i))))
			Call XMLCreateNode(attNode,custLastName_name,New_HTMLEncode(trim(pcArr(1,i))))
			Call XMLCreateNode(attNode,custEmail_name,New_HTMLEncode(trim(pcArr(2,i))))
			Call XMLCreateNode(attNode,custCompany_name,New_HTMLEncode(trim(pcArr(5,i))))
			Call XMLCreateNode(attNode,custPhone_name,New_HTMLEncode(trim(pcArr(6,i))))
			Call XMLCreateNode(attNode,custFax_name,New_HTMLEncode(trim(pcArr(7,i))))
			Call XMLCreateNode(attNode,custAddress_name,New_HTMLEncode(trim(pcArr(8,i))))
			Call XMLCreateNode(attNode,custAddress2_name,New_HTMLEncode(trim(pcArr(9,i))))
			Call XMLCreateNode(attNode,custCity_name,New_HTMLEncode(trim(pcArr(10,i))))
			Call XMLCreateNode(attNode,custStateCode_name,New_HTMLEncode(trim(pcArr(11,i))))
			Call XMLCreateNode(attNode,custProvince_name,New_HTMLEncode(trim(pcArr(12,i))))
			Call XMLCreateNode(attNode,custZip_name,New_HTMLEncode(trim(pcArr(13,i))))
			Call XMLCreateNode(attNode,custCountryCode_name,New_HTMLEncode(trim(pcArr(14,i))))
		end if
		if custShippingAddress_ex=1 then
		
			'START - Default Shipping Address
			Set attNode=oXML.createNode(1,custShippingAddress_name,"")
			custNode.appendChild(attNode)
			Call XMLCreateNode(attNode,custShipAddressID_name,"0")
			Call XMLCreateNode(attNode,custShipNickName_name,"")
			Call XMLCreateNode(attNode,custShipFirstName_name,"")
			Call XMLCreateNode(attNode,custShipLastName_name,"")
			Call XMLCreateNode(attNode,custShipEmail_name,"")
			Call XMLCreateNode(attNode,custShipCompany_name,New_HTMLEncode(trim(pcArr(15,i))))
			Call XMLCreateNode(attNode,custShipAddress_name,New_HTMLEncode(trim(pcArr(16,i))))
			Call XMLCreateNode(attNode,custShipAddress2_name,New_HTMLEncode(trim(pcArr(17,i))))
			Call XMLCreateNode(attNode,custShipCity_name,New_HTMLEncode(trim(pcArr(18,i))))
			Call XMLCreateNode(attNode,custShipStateCode_name,New_HTMLEncode(trim(pcArr(19,i))))
			Call XMLCreateNode(attNode,custShipProvince_name,New_HTMLEncode(trim(pcArr(20,i))))
			Call XMLCreateNode(attNode,custShipZip_name,New_HTMLEncode(trim(pcArr(21,i))))
			Call XMLCreateNode(attNode,custShipCountryCode_name,New_HTMLEncode(trim(pcArr(22,i))))
			Call XMLCreateNode(attNode,custShipPhone_name,"")
			Call XMLCreateNode(attNode,custShipFax_name,"")
			'END - Default Shipping Address
			
			'START - Multiple Shipping Addresses
			
			Call XMLgetCustShipAddress(custNode,custID_value)
			
			'END - Multiple Shipping Addresses
			
		end if
		if custRPBalance_ex=1 then
			Call XMLCreateNode(custNode,custRPBalance_name,trim(pcArr(23,i)))
		end if
		if custRPUsed_ex=1 then
			Call XMLCreateNode(custNode,custRPUsed_name,trim(pcArr(24,i)))
		end if
		if custPricingCategory_ex=1 then
			Call XMLgetCustPricingCat(custNode,trim(pcArr(26,i)))
		end if
		if custField_ex=1 then
			Call XMLgetCustCustomField(custNode,custID_value)
		end if
		if custNewsletter_ex=1 then
			Call XMLCreateNode(custNode,custNewsletter_name,trim(pcArr(27,i)))
		end if
		if custTotalOrders_ex=1 then
			Call XMLCreateNode(custNode,custTotalOrders_name,trim(pcArr(28,i)))
		end if
		if custTotalSales_ex=1 then
			Call XMLCreateNode(custNode,custTotalSales_name,trim(pcArr(29,i)))
		end if
		if custCreatedDate_ex=1 then
			tmpCreatedDate=trim(pcArr(30,i))
			if tmpCreatedDate<>"" then
				tmpCreatedDate=ConvertToXMLDate(tmpCreatedDate)
			end if
			Call XMLCreateNode(custNode,custCreatedDate_name,tmpCreatedDate)
		end if
		if custStatus_ex=1 then
			tmpStatus=trim(pcArr(31,i))
			if tmpStatus=1 then
				tmpStatus=2
			else
				if cint(trim(pcArr(4,i)))=2 then
					tmpStatus=1
				else
					tmpStatus=0
				end if
			end if
			Call XMLCreateNode(custNode,custStatus_name,tmpStatus)
		end if
		
		Set pXML1=Server.CreateObject("MSXML2.DOMDocument"&scXML)
		pXML1.async=false
		pXML1.load(oXML)
		If (pXML1.parseError.errorCode <> 0) Then	
			Set oXML=nothing
			call InitResponseDocument(cm_GetCustomerDetailsResponse_name)
			call XMLcreateError(pXML1.parseError.errorCode, pXML1.parseError.reason)
			call returnXML()
		End If
		set pXML1 = nothing
		
	ELSE
		call XMLcreateError(116,cm_errorStr_116a)
		call returnXML()
	END IF 'Have customer record
	
End Sub

Sub PresetCustValues()
	custType_value=0
	custRPBalance_value=0
	custRPUsed_value=0
	pricingCatID_value=0
	custNewsletter_value=0
	custTotalOrders_value=0
	custTotalSales_value=0
	custStatus_value=0
End Sub

Sub CheckAddUpdCustomer(requestType)

	Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1,subNode,ChildNodes1,attNode,tmpNodeName1,tmpNodeValue1
	
	Call CheckRequiredXMLTag(cm_customer)
	Call PresetCustValues()
	
	if requestType=0 then
		Call CheckRequiredXMLTag(cm_customer & "/" & custBillingAddress_name & "/" & custFirstName_name)
		Call CheckRequiredXMLTag(cm_customer & "/" & custBillingAddress_name & "/" & custLastName_name)
		Call CheckRequiredXMLTag(cm_customer & "/" & custBillingAddress_name & "/" & custEmail_name)
	else
		Call CheckRequiredXMLTag(cm_customer & "/" & ImportField_name)
	end if
	
	Set rNode=iRoot.selectSingleNode(cm_customer)
	Set ChildNodes = rNode.childNodes
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
			Select Case tmpNodeName					
				Case ImportField_name:
					if requestType=1 then
						call CheckValidXMLTag(strNode,1,1,"")
						ImportField_ex=1
						ImportField_value=tmpNodeValue						
						Select Case ImportField_value
							Case 1:
								Call CheckRequiredXMLTag(cm_customer & "/" & custBillingAddress_name & "/" & custEmail_name)
							Case 2:
								Call CheckRequiredXMLTag(cm_customer & "/" & custID_name)
							Case Else
								call XMLcreateError(128, cm_errorStr_128)
								call returnXML()
						End Select						
					end if
				Case custID_name:
					call CheckValidXMLTag(strNode,1,5,"")
					custID_ex=1
					custID_value=getUserInput(tmpNodeValue,0)
				Case custPassword_name:
					call CheckValidXMLTag(strNode,1,5,"")
					custPassword_ex=1
					custPassword_value=getUserInput(tmpNodeValue,0)
				Case custType_name:
					call CheckValidXMLTag(strNode,1,1,"")
					custType_ex=1
					custType_value=tmpNodeValue
				Case custBillingAddress_name:
					If strNode.Text<>"" then
						custBillingAddress_ex=1
						Set ChildNodes1 = strNode.childNodes
						For Each attNode In ChildNodes1
							tmpNodeName1=attNode.nodeName
							tmpNodeValue1=trim(attNode.Text)
							Select Case tmpNodeName1
								Case custFirstName_name:
									call CheckValidXMLTag(attNode,1,5,"")
									custFirstName_ex=1
									custFirstName_value=getUserInput(tmpNodeValue1,0)
								Case custLastName_name:
									call CheckValidXMLTag(attNode,1,5,"")
									custLastName_ex=1
									custLastName_value=getUserInput(tmpNodeValue1,0)
								Case custEmail_name:
									call CheckValidXMLTag(attNode,1,5,"")
									custEmail_ex=1
									custEmail_value=getUserInput(tmpNodeValue1,0)
								Case custCompany_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custCompany_ex=1
									custCompany_value=getUserInput(tmpNodeValue1,0)
								Case custPhone_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custPhone_ex=1
									custPhone_value=getUserInput(tmpNodeValue1,0)
								Case custFax_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custFax_ex=1
									custFax_value=getUserInput(tmpNodeValue1,0)
								Case custAddress_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custAddress_ex=1
									custAddress_value=getUserInput(tmpNodeValue1,0)
								Case custAddress2_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custAddress2_ex=1
									custAddress2_value=getUserInput(tmpNodeValue1,0)
								Case custCity_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custCity_ex=1
									custCity_value=getUserInput(tmpNodeValue1,0)
								Case custStateCode_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custStateCode_ex=1
									custStateCode_value=getUserInput(tmpNodeValue1,0)
								Case custProvince_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custProvince_ex=1
									custProvince_value=getUserInput(tmpNodeValue1,0)
								Case custZip_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custZip_ex=1
									custZip_value=getUserInput(tmpNodeValue1,0)
								Case custCountryCode_name:
									call CheckValidXMLTag(attNode,0,5,"")
									custCountryCode_ex=1
									custCountryCode_value=getUserInput(tmpNodeValue1,0)
							End Select
						Next
					End if
				Case custShippingAddress_name:
					If strNode.Text<>"" then
						custShippingAddress_ex=1
					End if
				Case custRPBalance_name:
					call CheckValidXMLTag(strNode,1,1,"")
					custRPBalance_ex=1
					custRPBalance_value=tmpNodeValue
				Case custRPUsed_name:
					call CheckValidXMLTag(strNode,1,1,"")
					custRPUsed_ex=1
					custRPUsed_value=tmpNodeValue
				Case custPricingCategory_name:
					if tmpNodeValue<>"" then
						if CheckExistTag(cm_customer & "/" & custPricingCategory_name & "/" & pricingCatID_name) then
							Set subNode=iRoot.selectSingleNode(cm_customer & "/" & custPricingCategory_name & "/" & pricingCatID_name)
							call CheckValidXMLTag(subNode,1,1,"")
							pricingCatID_ex=1
							pricingCatID_value=subNode.Text
						end if
						if CheckExistTag(cm_customer & "/" & custPricingCategory_name & "/" & pricingCatName_name) then
							Set subNode=iRoot.selectSingleNode(cm_customer & "/" & custPricingCategory_name & "/" & pricingCatName_name)
							call CheckValidXMLTag(subNode,1,5,"")
							pricingCatName_ex=1
							pricingCatName_value=getUserInput(subNode.Text,0)
						end if
						custPricingCategory_ex=1
					end if
				Case custField_name:
					if tmpNodeValue<>"" then
						if CheckExistTagEx(strNode,fieldName_name) then
							Call CheckRequiredXMLTagEx(strNode,fieldName_name)
						else
							Call CheckRequiredXMLTagEx(strNode,fieldID_name)
						end if
						custField_ex=1
					end if
				Case custNewsletter_name:
					call CheckValidXMLTag(strNode,1,1,"")
					custNewsletter_ex=1
					custNewsletter_value=tmpNodeValue
				Case custTotalOrders_name:
					call CheckValidXMLTag(strNode,1,1,"")
					custTotalOrders_ex=1
					custTotalOrders_value=tmpNodeValue
				Case custTotalSales_name:
					call CheckValidXMLTag(strNode,1,2,"")
					custTotalSales_ex=1
					custTotalSales_value=tmpNodeValue
				Case custCreatedDate_name:
					if tmpNodeValue<>"" then
						if requestType=0 then
							call XMLcreateError(117,cm_errorStr_117 & tmpNodeName)
							call returnXML()
						else
							call CheckValidXMLTag(strNode,0,4,"")
							if tmpNodeValue<>"" then
								custCreatedDate_ex=1
								custCreatedDate_value=ConvertFromXMLDate(tmpNodeValue)
							end if
						end if
					end if
				Case custStatus_name:
					call CheckValidXMLTag(strNode,1,1,"")
					custStatus_ex=1
					custStatus_value=tmpNodeValue
					if cint(custStatus_value)=1 then
						custType_value=2
					end if
				Case Else:
					call XMLcreateError(117,cm_errorStr_117 & tmpNodeName)
					call returnXML()
			End Select
	Next
End Sub

Sub XMLBackUpCustData(BackUpType,tmpIDValue)

	Dim query,rs,tmpTable,tmpField,tmp1,tmp2
	Dim BackupStr1,BackupStr2
	
	BackupStr1=""
	BackupStr2=""

	call opendb()
	
	if custID_value=0 then
		query="SELECT idcustomer FROM Customers WHERE [email] like '" & custEmail_value & "';"
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			custID_value=rstemp("idcustomer")
		end if
		set rstemp=nothing
	end if

	
	tmp2=""
	
	Select Case BackUpType
		Case 0:
			tmpTable="Customers"
			tmpField="idcustomer"
			BackupStr1=BackupStr1 & "UPDCUST" & chr(9) & custID_value
			tmp1=1
		Case 1:
			tmpTable="recipients"
			tmpField="idRecipient"
			BackupStr1=BackupStr1 & "UPDRECI" & chr(9) & tmpIDValue
			tmp1=1
	End Select
	if BackUpType<>"1" then
		query="SELECT * FROM " & tmpTable & " WHERE " & tmpField & "=" & custID_value & tmp2 & ";"
	else
		query="SELECT * FROM " & tmpTable & " WHERE " & tmpField & "=" & tmpIDValue & tmp2 & ";"
	end if
	set rstemp=conntemp.execute(query)
	
	IF not rstemp.eof THEN
		iCols = rstemp.Fields.Count
		For dd=tmp1 to iCols-1
			FType="" & Rstemp.Fields.Item(dd).Type
			if (Ftype="202") or (Ftype="203") or (Ftype="135") then
				PTemp=Rstemp.Fields.Item(dd).Value
				if PTemp<>"" then
					PTemp=replace(PTemp,"'","''")
					PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
				end if
	
				if (scDB="Access") and (Ftype="135") then
					BackupStr2=BackupStr2 & chr(9) & "#" & PTemp & "#"
				else
					BackupStr2=BackupStr2 & chr(9) & "'" & PTemp & "'"
				end if
			else
				PTemp="" & Rstemp.Fields.Item(dd).Value
				if PTemp<>"" then
				else
					PTemp="0"
				end if
				BackupStr2=BackupStr2 & chr(9) & PTemp
			end if
		Next
	END IF
	set rstemp=nothing
	
	if BackupStr2<>"" then
		BackupStr=BackupStr & BackupStr1 & BackupStr2  & vbcrlf
	end if
	
	call closedb()
	
End Sub

Function XMLCheckCustPricingCat(PriceCatID,PriceCatName)
Dim query,rstemp

	call opendb()
	
	if PriceCatID>0 then
		query="SELECT idCustomerCategory FROM pcCustomerCategories WHERE idCustomerCategory=" & PriceCatID & ";"
	else
		query="SELECT idCustomerCategory FROM pcCustomerCategories WHERE pcCC_Name like '" & PriceCatName & "';"
	end if
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		XMLCheckCustPricingCat=0
		call XMLcreateError(121,cm_errorStr_121)
	else
		XMLCheckCustPricingCat=rstemp("idCustomerCategory")
	end if
	
	set rstemp=nothing
	
	call closedb()
	
End Function

Function XMLCheckCustCustomField(cfid,cfname)
Dim tmpCFID,query,rstemp,rstemp1

	call opendb()
	
	if cfid>0 then
		query="SELECT pcCField_ID FROM pcCustomerFields WHERE pcCField_ID=" & cfid & ";"
	else
		query="SELECT pcCField_ID FROM pcCustomerFields WHERE pcCField_Name like '" & cfname & "';"
	end if
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		XMLCheckCustCustomField=0
		call XMLcreateError(122,cm_errorStr_122)
	else
		XMLCheckCustCustomField=rstemp("pcCField_ID")
	end if
	set rstemp=nothing
	
	call closedb()
	
End Function

Sub XMLAddUpdCustCF()
Dim attNode,subNode,ChildNodes,rNode
Dim tmpNodeName,tmpNodeValue
Dim query,rstemp
	
	Set rNode=iRoot.selectNodes(cm_customer & "/" & custField_name)
	For Each attNode In rNode
		If attNode.Text<>"" then
			Set ChildNodes = attNode.childNodes
			fieldName_ex=0
			fieldValue_ex=0
			fieldID_value=0
			fieldName_value=""
			fieldValue_value=""
		
			For Each subNode In ChildNodes
				tmpNodeName=subNode.NodeName
				tmpNodeValue=subNode.Text
				Select Case tmpNodeName
					Case fieldID_name:
						if tmpNodeValue<>"" then
							if IsNumeric(tmpNodeValue) AND tmpNodeValue>"0" then
								fieldID_ex=1
								fieldID_value=tmpNodeValue
							end if
						end if
					Case fieldName_name:
						fieldName_ex=1
						fieldName_value=getUserInput(tmpNodeValue,0)
					Case fieldValue_name:
						fieldValue_ex=1
						fieldValue_value=getUserInput(tmpNodeValue,0)
				End Select
			Next
				
			if (fieldName_ex=1) AND (fieldValue_ex=1) then
				fieldID_value=XMLCheckCustCustomField(fieldID_value,fieldName_value)
			end if
			if cint(fieldID_value)>0 AND (fieldValue_ex=1) then
				call opendb()
				
				query="DELETE FROM pcCustomerFieldsValues WHERE idcustomer=" & custID_value & " AND pcCField_ID=" & fieldID_value & ";"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing
				
				query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & custID_value & "," & fieldID_value & ",'" & fieldValue_value & "');"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing
				
				call closedb()
			end if
		End if
	Next

End Sub

Sub XMLAddUpdCustShipAddr()
Dim attNode,subNode,ChildNodes,rNode
Dim tmpNodeName,tmpNodeValue
Dim query,rstemp
Dim custShipFullName
	
	Set rNode=iRoot.selectNodes(cm_customer & "/" & custShippingAddress_name)
	For Each attNode In rNode
		If attNode.Text<>"" then
			Set ChildNodes = attNode.childNodes
			custShipAddressID_ex=0
			custShipAddressID_value=0
			custShipNickName_value=""
			custShipFirstName_value=""
			custShipLastName_value=""
			custShipEmail_value=""
			custShipCompany_value=""
			custShipAddress_value=""
			custShipAddress2_value=""
			custShipCity_value=""
			custShipStateCode_value=""
			custShipProvince_value=""
			custShipZip_value=""
			custShipCountryCode_value=""
			custShipPhone_value=""
			custShipFax_value=""
			custShipFullName=""
					
			For Each subNode In ChildNodes
				tmpNodeName=subNode.NodeName
				tmpNodeValue=getUserInput(subNode.Text,0)
				Select Case tmpNodeName
					Case custShipAddressID_name:
						if tmpNodeValue<>"" then
							if IsNumeric(tmpNodeValue) AND tmpNodeValue>"0" then
								custShipAddressID_ex=1
								custShipAddressID_value=tmpNodeValue
							end if
						end if
					Case custShipNickName_name:
						custShipNickName_ex=1
						custShipNickName_value=tmpNodeValue
					Case custShipFirstName_name:
						custShipFirstName_ex=1
						custShipFirstName_value=tmpNodeValue
					Case custShipLastName_name:
						custShipLastName_ex=1
						custShipLastName_value=tmpNodeValue
					Case custShipEmail_name:
						custShipEmail_ex=1
						custShipEmail_value=tmpNodeValue
					Case custShipCompany_name:
						custShipCompany_ex=1
						custShipCompany_value=tmpNodeValue
					Case custShipAddress_name:
						custShipAddress_ex=1
						custShipAddress_value=tmpNodeValue
					Case custShipAddress2_name:
						custShipAddress2_ex=1
						custShipAddress2_value=tmpNodeValue
					Case custShipCity_name:
						custShipCity_ex=1
						custShipCity_value=tmpNodeValue
					Case custShipStateCode_name:
						custShipStateCode_ex=1
						custShipStateCode_value=tmpNodeValue
					Case custShipProvince_name:
						custShipProvince_ex=1
						custShipProvince_value=tmpNodeValue
					Case custShipZip_name:
						custShipZip_ex=1
						custShipZip_value=tmpNodeValue
					Case custShipCountryCode_name:
						custShipCountryCode_ex=1
						custShipCountryCode_value=tmpNodeValue
					Case custShipPhone_name:
						custShipPhone_ex=1
						custShipPhone_value=tmpNodeValue
					Case custShipFax_name:
						custShipFax_ex=1
						custShipFax_value=tmpNodeValue
				End Select
			Next
			
			if custShipFirstName_value & custShipLastName_value<>"" then
				custShipFullName=custShipFirstName_value & " " & custShipLastName_value
			end if
			
			call opendb()
			if cint(custShipAddressID_value)>0 AND (custShipAddressID_ex=1) then
				query="SELECT idRecipient FROM recipients WHERE idRecipient=" & custShipAddressID_value & " AND idcustomer=" & custID_value & ";"
				set rstemp=connTemp.execute(query)
				if not rstemp.eof then
					Call XMLBackUpCustData(1,custShipAddressID_value)
					call opendb()
					query="UPDATE recipients SET recipient_FullName='" & custShipFullName & "',recipient_NickName='" & custShipNickName_value & "',recipient_FirstName='" & custShipFirstName_value & "',recipient_LastName='" & custShipLastName_value & "',recipient_Email='" & custShipEmail_value & "',recipient_Company='" & custShipCompany_value & "',recipient_Address='" & custShipAddress_value & "',recipient_Address2='" & custShipAddress2_value & "',recipient_City='" & custShipCity_value & "',recipient_StateCode='" & custShipStateCode_value & "',recipient_State='" & custShipProvince_value & "',recipient_Zip='" & custShipZip_value & "',recipient_CountryCode='" & custShipCountryCode_value & "',recipient_Phone='" & custShipPhone_value & "',recipient_Fax='" & custShipFax_value & "' WHERE idRecipient=" & custShipAddressID_value & " AND idcustomer=" & custID_value & ";"
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				else
					query="INSERT INTO recipients (idcustomer,recipient_FullName,recipient_NickName,recipient_FirstName,recipient_LastName,recipient_Email,recipient_Company,recipient_Address,recipient_Address2,recipient_City,recipient_StateCode,recipient_State,recipient_Zip,recipient_CountryCode,recipient_Phone,recipient_Fax) VALUES (" & custID_value & ",'" & custShipFullName & "','" & custShipNickName_value & "','" & custShipFirstName_value & "','" & custShipLastName_value & "','" & custShipEmail_value & "','" & custShipCompany_value & "','" & custShipAddress_value & "','" & custShipAddress2_value & "','" & custShipCity_value & "','" & custShipStateCode_value & "','" & custShipProvince_value & "','" & custShipZip_value & "','" & custShipCountryCode_value & "','" & custShipPhone_value & "','" & custShipFax_value & "');"
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
					query="SELECT idRecipient FROM recipients WHERE idcustomer=" & custID_value & " ORDER BY idRecipient DESC;"
					set rstemp=connTemp.execute(query)
					if not rstemp.eof then
						custShipAddressID_value=rstemp("idRecipient")
					end if
					set rstemp=nothing
					BackupStr=BackupStr & "DELRECI" & chr(9) & custShipAddressID_value & chr(9) & custID_value & vbcrlf
				end if
				
			else
				if custShipNickName_value="" AND custShipFirstName_value="" AND custShipLastName_value="" then
					query="UPDATE customers SET shippingCompany='" & custShipCompany_value & "',shippingaddress='" & custShipAddress_value & "',shippingAddress2='" & custShipAddress2_value & "',shippingcity='" & custShipCity_value & "',shippingStateCode='" & custShipStateCode_value & "',shippingState='" & custShipProvince_value & "',shippingZip='" & custShipZip_value & "',shippingCountryCode='" & custShipCountryCode_value & "',shippingEmail='" & custShipEmail_value & "',shippingPhone='" & custShipPhone_value & "',shippingFax='" & custShipFax_value & "' WHERE idcustomer=" & custID_value & ";"
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				else
					query="INSERT INTO recipients (idcustomer,recipient_FullName,recipient_NickName,recipient_FirstName,recipient_LastName,recipient_Email,recipient_Company,recipient_Address,recipient_Address2,recipient_City,recipient_StateCode,recipient_State,recipient_Zip,recipient_CountryCode,recipient_Phone,recipient_Fax) VALUES (" & custID_value & ",'" & custShipFullName & "','" & custShipNickName_value & "','" & custShipFirstName_value & "','" & custShipLastName_value & "','" & custShipEmail_value & "','" & custShipCompany_value & "','" & custShipAddress_value & "','" & custShipAddress2_value & "','" & custShipCity_value & "','" & custShipStateCode_value & "','" & custShipProvince_value & "','" & custShipZip_value & "','" & custShipCountryCode_value & "','" & custShipPhone_value & "','" & custShipFax_value & "');"
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
					query="SELECT idRecipient FROM recipients WHERE idcustomer=" & custID_value & " ORDER BY idRecipient DESC;"
					set rstemp=connTemp.execute(query)
					if not rstemp.eof then
						custShipAddressID_value=rstemp("idRecipient")
					end if
					set rstemp=nothing
					BackupStr=BackupStr & "DELRECI" & chr(9) & custShipAddressID_value & chr(9) & custID_value & vbcrlf
				end if
			end if
			call closedb()

		End if
	Next

End Sub

Sub RunAddCustomer()
Dim tmp_Suspend,dtTodaysDate,tmp_str1,tmp_str2
Dim query,rstemp,i

	tmp_Suspend=0
	if custStatus_value=2 then
		tmp_Suspend=1
	end if
	
	if custPricingCategory_ex=1 then
		pricingCatID_value=XMLCheckCustPricingCat(pricingCatID_value,pricingCatName_value)
	end if
	
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	
	if custPassword_value<>"" then
	else
		For i=1 to 16
 		randomize
 		custPassword_value=custPassword_value & Cstr(Fix(rnd*10))
 		Next
	end if
	custPassword_value=enDeCrypt(custPassword_value, scCrypPass)
	
	call opendb()
	
	if scDB="SQL" then
		query="INSERT INTO Customers (name,lastName,customerCompany,phone,email,password,address,zip,stateCode,state,city,countryCode,shippingaddress,shippingcity,shippingStateCode,shippingState,shippingCountryCode,shippingZip,customerType,TotalOrders,TotalSales,iRewardPointsAccrued,iRewardPointsUsed,address2,shippingCompany,shippingAddress2,RecvNews,suspend,idCustomerCategory,fax,pcCust_DateCreated) VALUES ('" & custFirstName_value & "','" & custLastName_value & "','" & custCompany_value & "','" & custPhone_value & "','" & custEmail_value & "','" & custPassword_value & "','" & custAddress_value & "','" & custZip_value & "','" & custStateCode_value & "','" & custProvince_value & "','" & custCity_value & "','" & custCountryCode_value & "','" & custShipAddress_value & "','" & custShipCity_value & "','" & custShipStateCode_value & "','" & custShipProvince_value & "','" & custShipCountryCode_value & "','" & custShipZip_value & "'," & custType_value & "," & custTotalOrders_value & "," & custTotalSales_value & "," & custRPBalance_value & "," & custRPUsed_value & ",'" & custAddress2_value & "','" & custShipCompany_value & "','" & custShipAddress2_value & "'," & custNewsletter_value & "," & tmp_Suspend & "," & pricingCatID_value & ",'" & custFax_value & "','" & dtTodaysDate & "');"
	else
		query="INSERT INTO Customers (name,lastName,customerCompany,phone,email,password,address,zip,stateCode,state,city,countryCode,shippingaddress,shippingcity,shippingStateCode,shippingState,shippingCountryCode,shippingZip,customerType,TotalOrders,TotalSales,iRewardPointsAccrued,iRewardPointsUsed,address2,shippingCompany,shippingAddress2,RecvNews,suspend,idCustomerCategory,fax,pcCust_DateCreated) VALUES ('" & custFirstName_value & "','" & custLastName_value & "','" & custCompany_value & "','" & custPhone_value & "','" & custEmail_value & "','" & custPassword_value & "','" & custAddress_value & "','" & custZip_value & "','" & custStateCode_value & "','" & custProvince_value & "','" & custCity_value & "','" & custCountryCode_value & "','" & custShipAddress_value & "','" & custShipCity_value & "','" & custShipStateCode_value & "','" & custShipProvince_value & "','" & custShipCountryCode_value & "','" & custShipZip_value & "'," & custType_value & "," & custTotalOrders_value & "," & custTotalSales_value & "," & custRPBalance_value & "," & custRPUsed_value & ",'" & custAddress2_value & "','" & custShipCompany_value & "','" & custShipAddress2_value & "'," & custNewsletter_value & "," & tmp_Suspend & "," & pricingCatID_value & ",'" & custFax_value & "',#" & dtTodaysDate & "#);"
	end if
	query=replace(query,chr(34),"&quot;")
	set rstemp=conntemp.execute(query)
	
	query="SELECT idcustomer FROM Customers WHERE [email] like '" & custEmail_value & "' ORDER BY idcustomer DESC;"
	set rstemp=connTemp.execute(query)
	if not rstemp.eof then
		custID_value=rstemp("idcustomer")
	end if
	set rstemp=nothing
		
	call closedb()
	
	if custShippingAddress_ex=1 then
		Call XMLAddUpdCustShipAddr()
	end if
	
	if custField_ex=1 then
		Call XMLAddUpdCustCF()
	end if
	
	BackupStr=BackupStr & "DELCUST" & chr(9) & custID_value & vbcrlf
	
End Sub

Sub RunUpdCustomer()

	Dim tmp_Suspend,tmp1
	Dim query,rstemp,i	
	
	tmp_Suspend=""
	if custStatus_ex=1 then
		if custStatus_value=2 then
			tmp_Suspend=1
		else
			tmp_Suspend=0
		end if
	end if
	
	if custPricingCategory_ex=1 then
		pricingCatID_value=XMLCheckCustPricingCat(pricingCatID_value,pricingCatName_value)
	end if
	
	if custPassword_ex=1 then
		custPassword_value=enDeCrypt(custPassword_value, scCrypPass)
	end if
		
	Call XMLBackUpCustData(0,0)

	tmp1=""
	if custFirstName_ex=1 then
		tmp1=tmp1 & ",name='" & custFirstName_value & "'"
	end if
	if custLastName_ex=1 then
		tmp1=tmp1 & ",lastName='" & custLastName_value & "'"
	end if
	if custCompany_ex=1 then
		tmp1=tmp1 & ",customerCompany='" & custCompany_value & "'"
	end if
	if custPhone_ex=1 then
		tmp1=tmp1 & ",phone='" & custPhone_value & "'"
	end if
	if custEmail_ex=1 then
		tmp1=tmp1 & ",email='" & custEmail_value & "'"
	end if
	if custPassword_ex=1 then
		tmp1=tmp1 & ",password='" & custPassword_value & "'"
	end if
	if custAddress_ex=1 then
		tmp1=tmp1 & ",address='" & custAddress_value & "'"
	end if
	if custZip_ex=1 then
		tmp1=tmp1 & ",zip='" & custZip_value & "'"
	end if
	if custStateCode_ex=1 then
		tmp1=tmp1 & ",stateCode='" & custStateCode_value & "'"
	end if
	if custProvince_ex=1 then
		tmp1=tmp1 & ",state='" & custProvince_value & "'"
	end if
	if custCity_ex=1 then
		tmp1=tmp1 & ",city='" & custCity_value & "'"
	end if
	if custCountryCode_ex=1 then
		tmp1=tmp1 & ",countryCode='" & custCountryCode_value & "'"
	end if
	if custType_ex=1 then
		tmp1=tmp1 & ",customerType=" & custType_value
	end if
	if custTotalOrders_ex=1 then
		tmp1=tmp1 & ",TotalOrders=" & custTotalOrders_value
	end if
	if custTotalSales_ex=1 then
		tmp1=tmp1 & ",TotalSales=" & custTotalSales_value
	end if
	if custRPBalance_ex=1 then
		tmp1=tmp1 & ",iRewardPointsAccrued=" & custRPBalance_value
	end if
	if custRPUsed_ex=1 then
		tmp1=tmp1 & ",iRewardPointsUsed=" & custRPUsed_value
	end if
	if custAddress2_ex=1 then
		tmp1=tmp1 & ",address2='" & custAddress2_value & "'"
	end if
	if custNewsletter_ex=1 then
		tmp1=tmp1 & ",RecvNews=" & custNewsletter_value
	end if
	if tmp_Suspend<>"" then
		tmp1=tmp1 & ",suspend=" & tmp_Suspend
	end if
	if custPricingCategory_ex=1 then
		tmp1=tmp1 & ",idCustomerCategory=" & pricingCatID_value
	end if
	if custFax_ex=1 then
		tmp1=tmp1 & ",fax='" & custFax_value & "'"
	end if
	
	if custCreatedDate_ex=1 then
		dtTodaysDate=custCreatedDate_value
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
		end if
		if scDB="SQL" then
			tmp1=tmp1 & ",pcCust_DateCreated='" & dtTodaysDate & "'"
		else
			tmp1=tmp1 & ",pcCust_DateCreated=#" & dtTodaysDate & "#"
		end if
	end if
	
	if tmp1<>"" then
		tmp1=mid(tmp1,2,len(tmp1))

		call opendb()
	
		query="UPDATE Customers SET " & tmp1 & " where idcustomer=" & custID_value & ";"
		query=replace(query,chr(34),"&quot;")
		set rstemp=conntemp.execute(query)
		set rstemp=nothing
	
		call closedb()
	end if
	
	if custShippingAddress_ex=1 then
		Call XMLAddUpdCustShipAddr()
	end if
	
	if custField_ex=1 then
		Call XMLAddUpdCustCF()
	end if
	
End Sub

Sub RunAddUpdCustomer(requestType)

Dim query,rs,tmpBackUp
Dim requestKey,fso,afi

	BackupStr=""
	
	call opendb()
	if requestType=0 then
		query="SELECT idcustomer FROM Customers WHERE [email] like '" & custEmail_value & "';"
		set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closedb()
			call XMLcreateError(118,cm_errorStr_118e & cm_errorStr_118f & custEmail_value & cm_errorStr_118c)
			call returnXML()
		end if
		set rs=nothing		
	else		
		if ImportField_value = 1 then
			query="SELECT idcustomer FROM Customers WHERE [email] like '" & custEmail_value & "';"
			set rstemp=conntemp.execute(query)
			if rstemp.eof then
				set rstemp=nothing
				call closedb()
				call XMLcreateError(118,cm_errorStr_118e & cm_errorStr_118f & custEmail_value & cm_errorStr_118d)
				call returnXML()
			else
				custID_value=rstemp("idcustomer")
			end if
			set rstemp=nothing
		end if		
		if custID_value > 0 then
			query="SELECT Customers.email FROM Customers WHERE idcustomer=" & custID_value & ";"
			set rs=connTemp.execute(query)
		
			if rs.eof then
				set rs=nothing
				call closedb()
				call XMLcreateError(118,cm_errorStr_118e & cm_errorStr_118a & custID_value & cm_errorStr_118d)
				call returnXML()
			else
				if custEmail_ex<>"1" then
					custEmail_ex=1
					custEmail_value=rs("email")
				end if
			end if
			set rs=nothing

		end if			
	end if	
	
	call closedb()	
	
	if requestType=0 then
		Call RunAddCustomer()
	else
		Call RunUpdCustomer()
	end if
	
	if BackupStr<>"" then
		tmpBackUp=1
	else
		tmpBackUp=0
	end if
	
	IF cm_LogTurnOn=1 THEN
		if requestType=0 then
			requestKey=CreateRequestRecord(pcv_PartnerID,10,custID_value,tmpBackUp,0,0,0,0)
			cm_requestKey_value=requestKey
		else
			requestKey=CreateRequestRecord(pcv_PartnerID,12,custID_value,tmpBackUp,0,0,0,0)
			cm_requestKey_value=requestKey
		end if
	
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	if xmlHaveErrors=0 then
		Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
		tmpNode.Text=cm_SuccessCode
		oRoot.appendChild(tmpNode)
	else
		oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_HalfSuccessCode
	end if

	Set tmpNode=oXML.createNode(1,custID_name,"")
	tmpNode.Text=custID_value
	oRoot.appendChild(tmpNode)
	
	Set tmpNode=oXML.createNode(1,custEmail_name,"")
	tmpNode.Text=custEmail_value
	oRoot.appendChild(tmpNode)
	
	if BackupStr<>"" then
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set afi=fso.CreateTextFile(server.MapPath(".") & "\logs\" & requestKey & ".txt",True)
		afi.Write(BackupStr)
		afi.Close
		Set afi=nothing
		Set fso=nothing
	end if
	
End Sub
%>
