<% 
'File Name: adminNewCustEmail.asp
'File Purpose: This file sends a new customer registration notification to the store admin.
' It also sends a welcome e-mail to the customer.

'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

	'------------------------------------------------------	
	'- Notify admin that the account has been created
	'------------------------------------------------------
	pcStrBillingFirstNameTemp = pcStrBillingFirstName
	pcStrBillingLastNameTemp = pcStrBillingLastName
	pcStrBillingCompanyTemp = pcStrBillingCompany
	pcStrCustomerEmailTemp = pcStrCustomerEmail
	pcStrBillingAddressTemp = pcStrBillingAddress
	pcStrBillingAddress2Temp = pcStrBillingAddress2
	pcStrBillingCityTemp = pcStrBillingCity
	pcStrBillingPostalCodeTemp = pcStrBillingPostalCode

	' Check to see if the store manager requested to be notified on new customer registrations
	if scNoticeNewCust="1" AND pcv_strNoticeNewCust="1" then
		
		' Check for "first name" variable
		if pFirstName = "" then
			pFirstName = pName
		end if
		' Calculate customer number using sccustpre constant
		Dim pcCustomerNumber
		if len(sccustpre)>0 then
			pcCustomerNumber = (sccustpre + int(session("idCustomer")))
		else
			pcCustomerNumber = (int(session("idCustomer")))
		end if

		MsgBody=""
		MsgBody=MsgBody & "A new customer registered with your store. Below are the customer's details:" & VBCRLF
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=MsgBody & "==========================================" & VBCRLF
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=MsgBody & "Customer ID: " & pcCustomerNumber & VBCRLF
		MsgBody=MsgBody & "Customer Name: " & removeSQ(pcStrBillingFirstNameTemp) & " " & removeSQ(pcStrBillingLastNameTemp) & VBCRLF
		MsgBody=MsgBody & "Company: " & removeSQ(pcStrBillingCompanyTemp) & VBCRLF
		MsgBody=MsgBody & "Phone: " & pcStrBillingPhone & VBCRLF
		MsgBody=MsgBody & "E-mail: " & removeSQ(pcStrCustomerEmailTemp) & VBCRLF
		MsgBody=MsgBody & "Address: " & removeSQ(pcStrBillingAddressTemp) & VBCRLF
		if pAddress2<>"" then
			MsgBody=MsgBody & "         " & removeSQ(pcStrBillingAddress2Temp) & VBCRLF
		end if
		MsgBody=MsgBody & "City: " & removeSQ(pcStrBillingCityTemp) & VBCRLF
		MsgBody=MsgBody & "State/Province: " & pcStrBillingProvince & pcStrBillingStateCode & VBCRLF
		MsgBody=MsgBody & "Postal Code: " & removeSQ(pcStrBillingPostalCodeTemp) & VBCRLF
		MsgBody=MsgBody & "Country Code: " & pcStrBillingCountryCode & VBCRLF
		
		'Start Special Customer Fields
		session("pcSFCustFieldsExist")=""
		session("sf_nc_custfields")=""
		call opendb()
		query="SELECT pcCField_ID,pcCField_Name,pcCField_FieldType,pcCField_Value,pcCField_Length,pcCField_Maximum,pcCField_Required,pcCField_PricingCategories,pcCField_ShowOnReg,pcCField_ShowOnCheckout,'',pcCField_Description FROM pcCustomerFields;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.eof then
			session("pcSFCustFieldsExist")="YES"
			session("sf_nc_custfields")=rs.GetRows()
		end if
		set rs=nothing
	
		if session("pcSFCustFieldsExist")="YES" AND Session("idCustomer")<>0 then
			pcArr=session("sf_nc_custfields")
			For k=0 to ubound(pcArr,2)
				pcArr(10,k)=""
				query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & Session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if not rs.eof then
					pcArr(10,k)=rs("pcCFV_Value")
				end if
				set rs=nothing
				if trim(pcArr(10,k))<>"" then
					MsgBody=MsgBody & pcArr(1,k) & ": " & pcArr(10,k) & VBCRLF
				end if
			Next
			session("sf_nc_custfields")=pcArr
		end if
		
		call closedb()
		'End of Special Customer Fields
		
		if (Session("pcSFIDrefer")<>"0") and (Session("pcSFIDrefer")<>"") then
			call opendb()
			query="select [name] from Referrer where IDRefer=" & Session("pcSFIDrefer")
			set rsRef=Server.CreateObject("ADODB.Recordset")	
			set rsRef=connTemp.execute(query)
			if not rsRef.eof then
				MsgBody=MsgBody & "Referred by: " & rsRef("name") & VBCRLF
			end if
			set rsRef=nothing
			call closedb()
		end if
		if Session("pcSFCRecvNews")="1" then
			MsgBody=MsgBody & "Signed up for the store newsletter" & VBCRLF
		end if
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=MsgBody & "==========================================" & VBCRLF
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=Replace(MsgBody,"&quot;",chr(34))
		pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_5")
		call sendmail (scCompanyName, scEmail, scFrmEmail, pcv_strSubject, MsgBody)
	end if
	
	
	'------------------------------------------------------	
	'- Notify customer that the account has been created
	'------------------------------------------------------
	if session("pcSFPassWordExists")<>"NOREG" AND pcv_strNewCustEmail="1" then
		MsgBody=""
		MsgBody=MsgBody & removeSQ(pcStrBillingFirstNameTemp) & " " & removeSQ(pcStrBillingLastNameTemp) & dictLanguage.Item(Session("language")&"_storeEmail_20") & VBCRLF & VBCRLF
		MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_storeEmail_21") & VBCRLF	& VBCRLF
		
		dim tempURL
		tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/custPref.asp"),"//","/")
		tempURL=replace(tempURL,"https:/","https://")
		tempURL=replace(tempURL,"http:/","http://")
			
		MsgBody=MsgBody & tempURL & VBCRLF & VBCRLF
		MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_storeEmail_22")
		MsgBody=MsgBody & "" & VBCRLF
		MsgBody=Replace(MsgBody,"&quot;",chr(34))
		pcv_strSubject = scCompanyName & dictLanguage.Item(Session("language")&"_storeEmail_23")
		call sendmail (scCompanyName, scEmail, pcStrCustomerEmailTemp, pcv_strSubject, MsgBody)
	end if
%>