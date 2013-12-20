<%
IF sendmailid<>-1 THEN
IF ship_sendmail="1" THEN
		pIdOrder=ship_order
		shipVia =ship_shipmethod
		TrackingNum =ship_tracking
		shipDate=ship_shipdate
		
		' Get customer id 
		' Get country code to determine FedEx tracking URL
		myquery="Select orderDate, idcustomer, CountryCode, ShippingCountryCode, paymentCode, pcOrd_ShippingEmail FROM orders WHERE idOrder="& pIdOrder
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(myquery)
		pcv_orderDate=rs("orderDate")
		pIdCustomer=rs("idcustomer")
		pCountryCode=rs("CountryCode")
		pshippingCountryCode=rs("shippingCountryCode")
		paymentCode=rs("paymentCode")
		pShippingEmail=rs("pcOrd_ShippingEmail")
		
		set rs=nothing
		
		if pshippingCountryCode <> "" then
				strFedExCountryCode=pshippingCountryCode
			else
				strFedExCountryCode=pCountryCode
		end if
		
		' Get other customer info
		myquery="Select name,lastname,email,customercompany FROM customers WHERE idcustomer="& pIdCustomer
		Set rsCust=Server.CreateObject("ADODB.Recordset")
		Set rsCust=conntemp.execute(myquery)
		' compile emails
		customerCancelledEmail=Cstr("")
		' Build body of message ...
		
		customerShippedEmail=""
		'Customized message from store owner
		todaydate=showDateFrmt(now())
		personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
		personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
		personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
		personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
		personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",rsCust("name")&" "&rsCust("lastname"))
		personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(pIdOrder)))
		personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pcv_orderDate))
		personalmessage=replace(personalmessage,"//","/")
		personalmessage=replace(personalmessage,"http:/","http://")
		personalmessage=replace(personalmessage,"https:/","https://")
		If scShippedEmail<>"" Then
			customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf & vbCrLf
		end if
		if shipVia <> "" then
		customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_15") &replace(shipVia,"'","''")& vbCrLf
		end if
		if shipDate<>"" then
		customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_16") &ShowDateFrmt(shipDate)& vbCrLf
		end if
		if TrackingNum <> "" then
		customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_17") &TrackingNum& vbCrLf
	' Start tracking URL, if any
		if instr(ucase(shipVia),"UPS") then
			customerShippedEmail=customerShippedEmail & scStoreURL & "/" & scPcFolder & "/pc/custUPSTracking.asp?itracknumber=" & TrackingNum & vbCrLf & vbCrLf
			customerShippedEmail=replace(customerShippedEmail,"//","/")
			customerShippedEmail=replace(customerShippedEmail,"http:/","http://")
			else
				if instr(ucase(shipVia),"FEDEX") then
					if ucase(strFedExCountryCode)="US" then
						customerShippedEmail=customerShippedEmail & "http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=" & TrackingNum & vbCrLf & vbCrLf
						else
						customerShippedEmail=customerShippedEmail & "http://www.fedex.com/Tracking?cntry_code=" & strFedExCountryCode & vbCrLf & vbCrLf
					end if
				end if
		end if
	' End tracking URL, if any
		else
		customerShippedEmail=customerShippedEmail & vbCrLf & vbCrLf
		end if
		CustomerShippedEmail=replace(CustomerShippedEmail,"//","/")
		CustomerShippedEmail=replace(CustomerShippedEmail,"http:/","http://")
		CustomerShippedEmail=replace(CustomerShippedEmail,"https:/","https://")
		CustomerShippedEmail=replace(CustomerShippedEmail,"''",chr(39))
		pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_7")
		strCustEmail=rsCust("email")
		pTmpEmail=rsCust("email")
		call sendmail (scCompanyName, scEmail, pTmpEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
		'//Send email to shipping email if it is different and exist
		if trim(pShippingEmail)<>"" AND trim(pShippingEmail)<>trim(pTmpEmail) then
			call sendmail (scCompanyName, scEmail, pShippingEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
		end if
END IF
END IF
%>
