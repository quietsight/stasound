<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Server.ScriptTimeout = 600 %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="adminHeader.asp" -->
<% 
dim query, conntemp, rs
'how many checkboxes?
checkboxCnt=request.Form("checkboxCnt")
successCnt = 0
call opendb()

'do for each checkbox
dim r, dtShippedDate
for r=1 to checkboxCnt	
	if request.Form("checkOrd"&r)="YES" then		
		pcv_method=request("shipmethod"&r)
		pcv_tracking=request("tracking"&r)	
		pcv_shippedDate=request("shipdate"&r)	
		if not isDate(pcv_shippedDate) OR pcv_shippedDate="" then	
			pcv_shippedDate=Date()
		end if
		'// Reverse "International" Date Format for db entry
		DBInputArray=split(pcv_shippedDate,"/")

		if scDateFrmt = "DD/MM/YY" then
			'DD/MM/YYYY
			dtInputDbD=DBInputArray(0)
			dtInputDbM=DBInputArray(1)
			dtInputDbY=DBInputArray(2)
		else
			'MM/DD/YYYY
			dtInputDbM=DBInputArray(0)
			dtInputDbD=DBInputArray(1)
			dtInputDbY=DBInputArray(2)
		end if

		if pcv_shippedDate<>"" then
			if SQL_Format="1" then
				dtShippedDate=(dtInputDbD&"/"&dtInputDbM&"/"&dtInputDbY)
			else
				dtShippedDate=(dtInputDbM&"/"&dtInputDbD&"/"&dtInputDbY)
			end if
		end if
		
		pCheckEmail=request("checkEmail"&r)
		pOrderStatus=request("orderstatus"&r)
		pcv_PrdList=request("PrdList"&r)	
		pcv_IdOrder=Request("idOrder"&r)  & ""
		if pcv_IdOrder="" then
			pcv_IdOrder=0
		end if
		
		'// Form the Admin Comment with "Fixed" Date
		pcv_AdmComments = "This order was batch shipped on "& pcv_shippedDate &""
		
		'// Insert Shipping Details
		if pcv_shippedDate<>"" then
			if scDB="SQL" then
				query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "','" & dtShippedDate & "','" & pcv_tracking & "','" & pcv_AdmComments & "', 4);"
			else
				query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "',#" & dtShippedDate & "#,'" & pcv_tracking & "','" & pcv_AdmComments & "', 4);"
			end if
		else					
			query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "','" & pcv_tracking & "','" & pcv_AdmComments & "', 4);"
		end if
		set rs=connTemp.execute(query)
		set rs=nothing
		
		'// Look Up the New ID 
		query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_IdOrder & " ORDER by pcPackageInfo_ID DESC;"
		set rs=connTemp.execute(query)
		pcv_PackageID=rs("pcPackageInfo_ID")
		set rs=nothing
		
		qry_ID=pcv_IdOrder	

		'// Associate a Product List
		if trim(pcv_PrdList)<>"" then
			pcA=split(pcv_PrdList,",")
			For i=lbound(pcA) to ubound(pcA)
				if trim(pcA(i)<>"") then
					query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1,pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND idProductOrdered=" & pcA(i) & ";"
					set rs=connTemp.execute(query)
					set rs=nothing
				end if
			Next
		else
			query="UPDATE ProductsOrdered SET pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND pcPrdOrd_Shipped=0 AND pcDropShipper_ID=0;"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if		
		
		pcv_SendCust="1"
		pcv_SendAdmin="0"	
		pcv_LastShip="0"
		
		'// Determine LastShip
		'query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder AND ProductsOrdered.pcPrdOrd_Shipped=0) WHERE Orders.idorder=" & qry_ID & " AND Orders.orderstatus<>4;"
		
		pcv_LastShip="1"
		query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1 WHERE idorder=" & qry_ID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		'// Update the Order Status
		if trim(pcv_PrdList)<>"" then
			if pcv_LastShip="1" then
				query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
			else
				query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & qry_ID & ";"
			end if
			set rs=connTemp.execute(query)
			set rs=nothing
		end if		
		
		query="Select orders.idcustomer, orders.orderdate, orders.CountryCode, orders.shippingCountryCode, Orders.pcOrd_ShippingEmail FROM orders WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		pIdCustomer=rs("idcustomer")
		pOrderDate=rs("orderdate")
		' Get country code to determine FedEx tracking URL
		pCountryCode=rs("CountryCode")
		pshippingCountryCode=rs("shippingCountryCode")
		if pshippingCountryCode <> "" then
				strFedExCountryCode=pshippingCountryCode
			else
				strFedExCountryCode=pCountryCode
		end if
		' End get country code to determine FedEx tracking URL
		pShippingEmail=rs("pcOrd_ShippingEmail")

		query="Select name,lastname,email,customercompany FROM customers WHERE idcustomer="& pIdCustomer
		Set rsCust=Server.CreateObject("ADODB.Recordset")
		Set rsCust=conntemp.execute(query)
		
		
		' compile emails
		customerShippedEmail=Cstr("")
	
		todaysdate = showDateFrmt(now())
		'Customized message from store owner
		personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
		personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
		personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
		personalmessage=replace(personalmessage,"<TODAY_DATE>",todaysdate)
		personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",rsCust("name")&" "&rsCust("lastname"))
		personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(qry_ID)))
		personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pOrderDate))
		personalmessage=replace(personalmessage,"//","/")
		personalmessage=replace(personalmessage,"http:/","http://")
		personalmessage=replace(personalmessage,"https:/","https://")
		
		If scShippedEmail<>"" Then
			customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf
			customerShippedEmail=customerShippedEmail & vbcrlf
		end if
		
		if pcv_method <> "" then
			customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_15") &replace(pcv_method,"'","''")& vbCrLf
		end if   
		
		customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_16") & GetDateGUIDatabase(dtShippedDate,1) & vbCrLf
		
		if pcv_tracking <> "" then
		customerShippedEmail=customerShippedEmail & dictLanguage.Item(Session("language")&"_storeEmail_17") &pcv_tracking& vbCrLf
		' Start tracking URL, if any
			if instr(ucase(pcv_method),"UPS") then
				customerShippedEmail=customerShippedEmail & scStoreURL & "/" & scPcFolder & "/pc/custUPSTracking.asp?itracknumber=" & pcv_tracking & vbCrLf & vbCrLf
				customerShippedEmail=replace(customerShippedEmail,"//","/")
				customerShippedEmail=replace(customerShippedEmail,"http:/","http://")
				else
					if instr(ucase(pcv_method),"FEDEX") then
						if ucase(strFedExCountryCode)="US" then
							customerShippedEmail=customerShippedEmail & "http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=" & pcv_tracking & vbCrLf & vbCrLf
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
		pEmail=rsCust("email")

		if pCheckEmail="YES" then
			pTmpEmail=rsCust("email")
			call sendmail (scCompanyName, scEmail, pTmpEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
			'//Send email to shipping email if it is different and exist
			if trim(pShippingEmail)<>"" AND trim(pShippingEmail)<>trim(pTmpEmail) then
				call sendmail (scCompanyName, scEmail, pShippingEmail, pcv_strSubject, replace(customerShippedEmail, "&quot;", chr(34)))
			end if
		end if

		successCnt=successCnt+1
		successData=successData&"Order Number "& (scpre+int(qry_ID)) &" was shipped successfully.<BR>" 
	end if
	'response.end
Next
call closedb()
%>
<table class="pcCPcontent">
  <tr>
    <td>
    	<div class="pcCPmessageSuccess"><strong><%=successCnt%></strong> records were successfully shipped. <a href="batchshiporders.asp">Ship other orders</a>.</div>
			<% if successData<>"" then %>
				<div style="padding: 20px;"><br><%=successData%><br></div>
			<% end if %>
		</td>
  </tr>
</table>
<!--#include file="adminFooter.asp" -->