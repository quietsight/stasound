<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<%
'***********************************************************************************
' START: RECIEVE NEW ORDER
'***********************************************************************************
Function processNewOrderNotification(domResponseObj)
    on error resume next
	
	Dim xmlMcResults
	
	'// Process <new-order-notification>
    xmlMcResults = createNewOrderResults(domResponseObj)	
    
	'// Respond with <new-order-notification> XML
    Response.write xmlMcResults	
	
	'// Clear All Sessions
	session.Abandon()
End Function
'***********************************************************************************
' END: RECIEVE NEW ORDER
'***********************************************************************************

Sub CreateDownloadInfo1(pIDProduct,pQuantity)
		Dim query,rstemp,pSku,pLicense,pLocalLG,pRemoteLG,k,dd

			query="select sku,License,LocalLG,RemoteLG from Products,DProducts where products.idproduct=" & pIdproduct & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			if not rstemp.eof then
				pSku=rstemp("sku")
				pLicense=rstemp("License")
				pLocalLG=rstemp("LocalLG")
				pRemoteLG=rstemp("RemoteLG")
				set rstemp=nothing
				
				IF (pLicense<>"") and (pLicense="1") THEN
					if pLocalLG<>"" then
						SPath1=Request.ServerVariables("PATH_INFO")
						mycount1=0
						do while mycount1<1
							if mid(SPath1,len(SPath1),1)="/" then
								mycount1=mycount1+1
							end if
							if mycount1<1 then
								SPath1=mid(SPath1,1,len(SPath1)-1)
							end if
						loop
						SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
						if Right(SPathInfo,1)="/" then
							pLocalLG=SPathInfo & "licenses/" & pLocalLG					
						else
							pLocalLG=SPathInfo & "/licenses/" & pLocalLG
						end if
						pLocalLG=replace(pLocalLG,"/pc/","/"&scAdminFolderName&"/")
						L_Action=pLocalLG
					else
						L_Action=pRemoteLG
					end if
					L_postdata=""
					L_postdata=L_postdata&"idorder=" & pIdOrder
					L_postdata=L_postdata&"&orderDate=" & pOrderDate
					L_postdata=L_postdata&"&ProcessDate=" & pProcessDate
					L_postdata=L_postdata&"&idcustomer=" & pIdCustomer
					L_postdata=L_postdata&"&idproduct=" & pIdproduct
					L_postdata=L_postdata&"&quantity=" & pQuantity
					L_postdata=L_postdata&"&sku=" & pSKU

					Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
					srvXmlHttp.open "POST", L_Action, False
					srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					srvXmlHttp.send L_postdata
					result1 = srvXmlHttp.responseText
					AR=split(result1,"<br>")

					rIdOrder=AR(0)
					rIdProduct=AR(1)
					Lic1=split(AR(2),"***")
					Lic2=split(AR(3),"***")
					Lic3=split(AR(4),"***")
					Lic4=split(AR(5),"***")
					Lic5=split(AR(6),"***")
	
					For k=0 to Cint(pQuantity)-1
						if K<=ubound(Lic1) then
							PLic1=Lic1(k)
						else
							PLic1=""
						end if
						if K<=ubound(Lic2) then
							PLic2=Lic2(k)
						else
							PLic2=""
						end if
						if K<=ubound(Lic3) then
							PLic3=Lic3(k)
						else
							PLic3=""
						end if
						if K<=ubound(Lic4) then
							PLic4=Lic4(k)
						else
							PLic4=""
						end if
						if K<=ubound(Lic5) then
							PLic5=Lic5(k)
						else
							PLic5=""
						end if
						if ppStatus=0 then
							query="Insert into DPLicenses (IdOrder,IdProduct,Lic1,Lic2,Lic3,Lic4,Lic5) values (" & rIdOrder & "," & rIdProduct & ",'" & PLic1 & "','" & PLic2 & "','" & PLic3 & "','" & PLic4 & "','" & PLic5 & "')"   
							set rstemp=server.CreateObject("ADODB.RecordSet")
							set rstemp=connTemp.execute(query)
							set rstemp=nothing
						end if
					Next
				END IF
				
				DO
					Tn1=""
						For dd=1 to 24
							Randomize
							myC=Fix(3*Rnd)
							Select Case myC
								Case 0: 
									Randomize
									Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
								Case 1: 
									Randomize
									Tn1=Tn1 & Cstr(Fix(10*Rnd))
								Case 2: 
									Randomize
									Tn1=Tn1 & Chr(Fix(26*Rnd)+97)		
							End Select		
						Next
	
						ReqExist=0
	
						query="select IDOrder from DPRequests where RequestSTR='" & Tn1 & "'" 
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=connTemp.execute(query)
	
						if not rstemp.eof then
							ReqExist=1
						end if
						set rstemp=nothing
				LOOP UNTIL ReqExist=0
	
				if ppStatus=0 then
					pTodaysDate=Date()
					if SQL_Format="1" then
						pTodaysDate=(day(pTodaysDate)&"/"&month(pTodaysDate)&"/"&year(pTodaysDate))
					else
						pTodaysDate=(month(pTodaysDate)&"/"&day(pTodaysDate)&"/"&year(pTodaysDate))
					end if
		
					'Insert Standard & BTO Products Download Requests into DPRequests Table
					if scDB="Access" then
						query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "',#" & pTodaysDate & "#)"   
					else
						query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "','" & pTodaysDate & "')"
					end if
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				end if
			end if
			set rstemp=nothing

	End Sub

'***********************************************************************************
' START: RECIEVE NEW ORDER CODE
'***********************************************************************************
Function createNewOrderResults(domMcCallbackObj)
	on error resume next
	
	'// Clear All Sessions
	session.Abandon()
	
    '// Define the objects used to read the xml <new-order-notification>
    Dim domMcCallbackObjRoot
 	Dim domOrderNumberList
	Dim pcv_strOrderNumber
	Dim Nodes
	Dim Node
	
    Set domMcCallbackObjRoot = domMcCallbackObj.documentElement
	
    '// Order Number
	Set Nodes = domMcCallbackObj.selectNodes("//new-order-notification")	
	For Each Node In Nodes
		pcv_strOrderNumber = Node.selectSingleNode("google-order-number").text
		pcv_strBuyerMarketingPreferences = Node.selectSingleNode("buyer-marketing-preferences/email-allowed").text
		pcv_strOrderAdjustment = Node.selectSingleNode("order-adjustment").text
		pcv_strOrderTotal = Node.selectSingleNode("order-total").text
		pcv_strFulfillmentOrderState = Node.selectSingleNode("fulfillment-order-state").text
		pcv_strFinancialOrderState = Node.selectSingleNode("financial-order-state").text
		pcv_strBuyerID = Node.selectSingleNode("buyer-id").text
		pcv_strTimestamp = Node.selectSingleNode("timestamp").text
	Next	
	
	'// <shopping-cart>
	Set Nodes = domMcCallbackObj.selectNodes("//shopping-cart/merchant-private-data")
	For Each Node In Nodes
		pcv_strMerchantNote = Node.selectSingleNode("merchant-note").text
	Next	
	
	'// Escape the Process if its a Google Invoice
	if pcv_strMerchantNote="" then
		sendNotificationAcknowledgment
		Exit Function
	end if
	
	'// Escape the Process if there is no Order ID
	if pcv_strOrderNumber="" then
		sendNotificationAcknowledgment
		Exit Function
	end if
	
	'// <buyer-shipping-address>
	Set Nodes = domMcCallbackObj.selectNodes("//new-order-notification/buyer-shipping-address")	
	For Each Node In Nodes
		pcv_strShippingContactName = Node.selectSingleNode("contact-name").text
		pcv_strShippingCompanyName = Node.selectSingleNode("company-name").text
		pcv_strShippingEmail = Node.selectSingleNode("email").text
		pcv_strShippingPhone = Node.selectSingleNode("phone").text
		pcv_strShippingFax = Node.selectSingleNode("fax").text
		pcv_strShippingAddress1 = Node.selectSingleNode("address1").text
		pcv_strShippingAddress2 = Node.selectSingleNode("address2").text
		pcv_strShippingCity = Node.selectSingleNode("city").text
		pcv_strShippingRegion = Node.selectSingleNode("region").text
		pcv_strShippingPostalCode = Node.selectSingleNode("postal-code").text
		pcv_strShippingCountryCode = Node.selectSingleNode("country-code").text
	Next
	
	'// <buyer-billing-address>
	Set Nodes = domMcCallbackObj.selectNodes("//new-order-notification/buyer-billing-address")	
	For Each Node In Nodes
		pcv_strBillingContactName = Node.selectSingleNode("contact-name").text
		pcv_strBillingCompanyName = Node.selectSingleNode("company-name").text
		pcv_strBillingEmail = Node.selectSingleNode("email").text
		pcv_strBillingPhone = Node.selectSingleNode("phone").text
		pcv_strBillingFax = Node.selectSingleNode("fax").text
		pcv_strBillingAddress1 = Node.selectSingleNode("address1").text
		pcv_strBillingAddress2 = Node.selectSingleNode("address2").text
		pcv_strBillingCity = Node.selectSingleNode("city").text
		pcv_strBillingRegion = Node.selectSingleNode("region").text
		pcv_strBillingPostalCode = Node.selectSingleNode("postal-code").text
		pcv_strBillingCountryCode = Node.selectSingleNode("country-code").text
	Next
	
	'// <order-adjustment>
	Set Nodes = domMcCallbackObj.selectNodes("//new-order-notification/order-adjustment")	
	For Each Node In Nodes
		pcv_strMerchantCalculationSuccessful = Node.selectSingleNode("merchant-calculation-successful").text
		pcv_strMerchantCalculationTax = Node.selectSingleNode("total-tax").text		
	Next
	
	'// <order-adjustment/shipping>
	Set Nodes = domMcCallbackObj.selectNodes("//new-order-notification/order-adjustment/shipping/merchant-calculated-shipping-adjustment")	
	For Each Node In Nodes
		pcv_strShippingName = Node.selectSingleNode("shipping-name").text
		pcv_strShippingCost = Node.selectSingleNode("shipping-cost").text
	Next
	
	pcv_strCalculatedAmount = 0
	pcv_strTotalCodesUsed = 0
	pcv_strAmount = 0
	pcv_strGoogleDiscountCode = ""
	session("pcSFIdDbSession")=""
	
	'// <gift-certificate-adjustment>
	Set Nodes = domMcCallbackObj.selectNodes("//new-order-notification/order-adjustment/merchant-codes/gift-certificate-adjustment")	
	For Each Node In Nodes
		pcv_strCalculatedAmount = Node.selectSingleNode("applied-amount").text
		pcv_strAmount = Node.selectSingleNode("applied-amount").text
		pcv_strAppliedAmount = Node.selectSingleNode("applied-amount").text
		pcv_strGoogleDiscountCode = Node.selectSingleNode("code").text
		pcv_strGoogleMessage = Node.selectSingleNode("message").text
		pcv_strTotalCodesUsed = pcv_strTotalCodesUsed + 1
	Next
	
	'// <coupon-adjustment>
	Set Nodes = domMcCallbackObj.selectNodes("//new-order-notification/order-adjustment/merchant-codes/coupon-adjustment")	
	For Each Node In Nodes
		pcv_strCalculatedAmount = pcv_strCalculatedAmount + Node.selectSingleNode("calculated-amount").text
		pcv_strAmount = pcv_strAmount & Node.selectSingleNode("calculated-amount").text & ","
		pcv_strAppliedAmount = Node.selectSingleNode("applied-amount").text
		pcv_strGoogleDiscountCode = pcv_strGoogleDiscountCode & Node.selectSingleNode("code").text & ","
		pcv_strGoogleMessage = Node.selectSingleNode("message").text
		pcv_strTotalCodesUsed = pcv_strTotalCodesUsed + 1
	Next
	
	'// <buyer-marketing-preferences>
	if pcv_strBuyerMarketingPreferences="false" then
		pcv_strEmailAllowed=0
	else
		pcv_strEmailAllowed=1
	end if
	
	Session("DiscountCode")=pcv_strGoogleDiscountCode '// string of all discount codes separated by a comma
	Session("DiscountTotal")= pcv_strAmount
	Session("DiscountCodeTotal")= pcv_strCalculatedAmount '// Total of all discounts
	Session("TotalCodesUsed")= pcv_strTotalCodesUsed '// Number of Codes Used
	
	'// Customer Handle          
	If len(pcv_strBillingEmail) > 0 Then		

		'// Merchant Order Reference Number
		pcArray_MerchantNote = split(pcv_strMerchantNote, chr(124))	
		pcStrCustomerRefKey=pcArray_MerchantNote(0)
		session("pcSFIdDbSession")=pcStrCustomerRefKey		
		pcv_intPackageNum=pcArray_MerchantNote(2)
		pcv_intCustomerId=pcArray_MerchantNote(3)
		pcv_strAffiliateID=pcArray_MerchantNote(4)
		
		'// Customer Details	
		If NOT len(pcv_intCustomerId)>0 Then
			pcv_intCustomerId=pcf_CustomerID(pcv_strBillingEmail)
		End If
		pcv_strFirstName = Left(pcv_strBillingContactName,(instr(pcv_strBillingContactName," ")-1))
		pcv_strLastName =  Right(pcv_strBillingContactName,(len(pcv_strBillingContactName)-instr(pcv_strBillingContactName," ")))
		
		pcStrCustomerPassword=enDeCrypt(pcStrCustomerRefKey, scCrypPass)
		pcv_strFirstName=replace(pcv_strFirstName,"'","''")
		pcv_strLastName=replace(pcv_strLastName,"'","''")
		pcv_strBillingCompanyName=replace(pcv_strBillingCompanyName,"'","''")
		pcv_strBillingAddress1=replace(pcv_strBillingAddress1,"'","''")
		pcv_strBillingAddress2=replace(pcv_strBillingAddress2,"'","''")
		pcv_strBillingCity=replace(pcv_strBillingCity,"'","''")
		
		'// Billing - Do not insert a length greater than 4 into the stateCode column
		Dim pcv_BillingStateCode, pcv_BillingState
		pcv_BillingStateCode=pcv_strBillingRegion
		if pcv_BillingStateCode<>"" then
			if len(pcv_BillingStateCode)>4 then
				pcv_BillingStateCode=""
			end if
		end if
		pcv_BillingState=pcv_strBillingRegion
		
		'// Shipping - Do not insert a length greater than 4 into the stateCode column
		Dim pcv_ShippingStateCode, pcv_ShippingState
		pcv_ShippingStateCode=pcv_strShippingRegion
		if pcv_ShippingStateCode<>"" then
			if len(pcv_ShippingStateCode)>4 then
				pcv_ShippingStateCode=""
			end if
		end if
		pcv_ShippingState=pcv_strShippingRegion
		
		if pcv_intCustomerId>0 then
			'// Update the Google User
			query="UPDATE customers SET customers.name='"&pcv_strFirstName&"', customers.lastName='"&pcv_strLastName&"', customers.customerCompany='"&pcv_strBillingCompanyName&"', customers.phone='"&pcv_strBillingPhone&"', customers.email='"&pcv_strBillingEmail&"', customers.address='"&pcv_strBillingAddress1&"', customers.zip='"&pcv_strBillingPostalCode&"', customers.stateCode='"&pcv_BillingStateCode&"', customers.state='"&pcv_BillingState&"', customers.city='"&pcv_strBillingCity&"', customers.countryCode='"&pcv_strBillingCountryCode&"', customers.address2='"&pcv_strBillingAddress2&"', RecvNews="& pcv_strEmailAllowed &", customers.fax='"&pcv_strBillingFax&"'"
			query=query&" WHERE idCustomer="&pcv_intCustomerId&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)	
		else
			'// Insert the Google User	
			query="INSERT INTO customers ([name], lastname, email, [password], customerCompany, phone, address, zip, stateCode, state, city, countryCode, IDRefer, address2, RecvNews, fax, pcCust_DateCreated) VALUES "
			query=query&"('"&pcv_strFirstName&"', '"&pcv_strLastName&"', '"&pcv_strBillingEmail&"', '"&pcStrCustomerPassword&"','"&pcv_strBillingCompanyName&"', "
			query=query&"'"&pcv_strBillingPhone&"', '"&pcv_strBillingAddress1&"', '"&pcv_strBillingPostalCode&"', "
			query=query&"'"&pcv_BillingStateCode&"', '"&pcv_BillingState&"', '"&pcv_strBillingCity&"', '"&pcv_strBillingCountryCode&"', "
			query=query&"'"& 0 &"', '"&pcv_strBillingAddress2&"', "
			query=query&"'"& pcv_strEmailAllowed &"', '"&pcv_strBillingFax&"', "
			query=query&""& pcArray_MerchantNote(1) &");"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)		
			
			pcv_intCustomerId=pcf_CustomerID(pcv_strBillingEmail)
		end if		
	End If
	

	'// Save the Order
	%>
	<!--#include file="pcPay_GoogleCheckout_SaveOrd.asp"-->    
	<%		
		'------------------------------------------------
		'- START: Send confirmation email
		'------------------------------------------------
		' Get order information from the database
		query="SELECT orders.idcustomer,orders.address,orders.City,orders.StateCode,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingZip,orders.shippingCountryCode, orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.iRewardValue,orders.iRewardRefId, orders.iRewardPointsRef,orders.iRewardPointsCustAccrued,customers.phone,ord_DeliveryDate,ord_VAT, pcOrd_CatDiscounts FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" & qry_ID
		Set rsEmailInfo=Server.CreateObject("ADODB.Recordset")
		Set rsEmailInfo=connTemp.execute(query)
			pidcustomer=rsEmailInfo("idcustomer")
			paddress=rsEmailInfo("address")
			pCity=rsEmailInfo("city")
			pStateCode=rsEmailInfo("StateCode")
			pzip=rsEmailInfo("zip")
			pCountryCode=rsEmailInfo("CountryCode")
			pshippingAddress=rsEmailInfo("shippingAddress")
			pshippingCity=rsEmailInfo("shippingCity")
			pshippingStateCode=rsEmailInfo("shippingStateCode")
			pshippingZip=rsEmailInfo("shippingZip")
			pshippingCountryCode=rsEmailInfo("shippingCountryCode")
			pShipmentDetails=rsEmailInfo("ShipmentDetails")
			pPaymentDetails=rsEmailInfo("paymentDetails")
			pdiscountDetails=rsEmailInfo("discountDetails")
			ptaxAmount=rsEmailInfo("taxAmount")
			ptotal=rsEmailInfo("total")
			pcomments=rsEmailInfo("comments")
			pShippingFullName=rsEmailInfo("ShippingFullName")
			paddress2=rsEmailInfo("address2")
			pShippingCompany=rsEmailInfo("ShippingCompany")
			pShippingAddress2=rsEmailInfo("ShippingAddress2")
			ptaxDetails=rsEmailInfo("taxDetails")
			piRewardValue=rsEmailInfo("iRewardValue")
			piRewardRefId=rsEmailInfo("iRewardRefId")
			piRewardPointsRef=rsEmailInfo("iRewardPointsRef")
			piRewardPointsCustAccrued=rsEmailInfo("iRewardPointsCustAccrued")
			pPhone=rsEmailInfo("phone")
			pord_DeliveryDate=rsEmailInfo("ord_DeliveryDate")
			pord_VAT=rsEmailInfo("ord_VAT")
			pcOrd_CatDiscounts=rsEmailInfo("pcOrd_CatDiscounts")
		set rsEmailInfo=nothing
	
		pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)

		
		'// Get customer details for this order
		query="Select name,lastname,customerCompany,email FROM customers WHERE idcustomer="& pIdCustomer
		Set rsCust=Server.CreateObject("ADODB.Recordset")
		Set rsCust=conntemp.execute(query)
			pName=rsCust("name")
			pLName=rsCust("lastname")
			pCustomerCompany=rsCust("customerCompany")
			pEmail=rsCust("email")
		Set rsCust=nothing

		'// Send Order Confirmation email to admin
		%>
		<!--#include file="pcPay_GoogleCheckout_AdminEmail.asp"-->
		<% 
		strNewOrderSubject = dictLanguage.Item(Session("language")&"_storeEmail_9")&(scpre + int(pIdOrder))	
		call sendmail (scCompanyName, scEmail, scFrmEmail, strNewOrderSubject, replace(storeAdminEmail,"&quot;", chr(34)))
		'------------------------------------------------
		'- END: Send confirmation email
		'------------------------------------------------		
	
	'// This section creates an <add-merchant-order-number> request
	attrGoogleOrderNumber = pcv_strOrderNumber
	elemMerchantOrderNumber = (int(pIdOrder)+scpre)
	AddOrderNumberRequest = createAddMerchantOrderNumber(attrGoogleOrderNumber, elemMerchantOrderNumber)	
	AddOrderNumberResponse = SendRequest(AddOrderNumberRequest, requestUrl)
	ProcessXmlData(AddOrderNumberResponse)

	Session("DiscountCode")=""
	Session("DiscountTotal")=""
	Session("DiscountCodeTotal")=""
	Session("TotalCodesUsed")=""
	session("pcSFIdDbSession")=""
	session("idOrderSaved")=""
	Set domMcCallbackObjRoot = Nothing
    Set domOrderNumberList = Nothing
	
End Function

Public Function pcf_CustomerID(BillingEmail)
	query="SELECT customers.idcustomer, customers.email FROM customers WHERE customers.email='"& BillingEmail &"';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if NOT rs.eof then
		pcf_CustomerID=rs("idcustomer")
	else
		pcf_CustomerID=0
	end if
	set rs=nothing
End Function

Public Function pcf_RestoreCartIndex(varCartArray)
	pcf_RestoreCartIndex=UBOUND(varCartArray)
End Function

Public Function FixedField(ByVal Width, ByVal Justify, ByVal Text)
	Select Case True
		Case Width < Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Left(Text, Width)
				Case Justify="R"
					FixedField=Right(Text, Width)
				Case Else
			End Select
									
		Case Width=Len(Text)
			FixedField=Text

		Case Width > Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Text & String(Width - Len(Text), " ")
				Case Justify="R"
					FixedField=String(Width - Len(Text), " ") & Text
				Case Else
			End Select

	End Select

End Function
'***********************************************************************************
' END: RECIEVE NEW ORDER CODE
'***********************************************************************************






'***********************************************************************************
' START: RECIEVE NEW STATUS
'***********************************************************************************
Function processOrderStateChangeNotification(domResponseObj)
	on error resume next
	
	Dim xmlMcResults
	
	'// Process <order-state-change-notification>
    xmlMcResults = createOrderStateChangeNotification(domResponseObj)
	
	'// Respond with <order-state-change-notification> XML
    Response.write xmlMcResults	
	
    '// <order-state-change-notification>
    sendNotificationAcknowledgment

	
End Function
'***********************************************************************************
' END: RECIEVE NEW STATUS
'***********************************************************************************



'***********************************************************************************
' START: RECIEVE NEW STATUS CODE
'***********************************************************************************
Function createOrderStateChangeNotification(domMcCallbackObj)
	on error resume next

    '// Define the objects used to read the xml <order-state-change-notification>
    Dim domMcCallbackObjRoot
 	Dim domOrderNumberList
	Dim pcv_strOrderNumber
	Dim Nodes
	Dim Node

	Set domMcCallbackObjRoot = domMcCallbackObj.documentElement


    Set Nodes = domMcCallbackObj.selectNodes("//order-state-change-notification")	


	For Each Node In Nodes
		pcv_strOrderNumber = Node.selectSingleNode("google-order-number").text		
		pcv_strNewFinancialOrderState = Node.selectSingleNode("new-financial-order-state").text
		pcv_strNewFulfillmentOrderState = Node.selectSingleNode("new-fulfillment-order-state").text
		pcv_strPrevFinancialOrderState = Node.selectSingleNode("previous-financial-order-state").text
		pcv_strPrevFulfillmentOrderState = Node.selectSingleNode("previous-fulfillment-order-state").text
		'// pcv_strReason = Node.selectSingleNode("reason").text
		pcv_strTimestamp = Node.selectSingleNode("timestamp").text
	Next	

	
	'// Find out original status first
	query="SELECT idOrder, orderstatus, paymentCode, pcOrd_Payer FROM orders WHERE pcOrd_GoogleIDOrder='"& pcv_strOrderNumber &"' "
	set rs=server.CreateObject("ADODB.RecordSet")
	Set rs=conntemp.execute(query)	
	if rs.eof then
	'// Escape the Process if its a Google Invoice
		sendNotificationAcknowledgment
		Exit Function
	else
		pidOrder=rs("idOrder")
		porigstatus=rs("orderstatus")
		paymentCode=rs("paymentCode")
		pcOrd_Payer=rs("pcOrd_Payer")
	end if

	'// Synchronize Financial Status with ProductCart
	Select Case pcv_strNewFinancialOrderState	
		Case "REVIEWING": pcv_PayStatusName=0 '// "Pending"		
		Case "CHARGEABLE": pcv_PayStatusName=1 '// "Authorized"		
		Case "CHARGING": pcv_PayStatusName=7 '// "Charging" This is a transitional status and will only be stated for a couple minutes at most
		Case "CHARGED": pcv_PayStatusName=2 '// "Paid"	
		Case "PAYMENT_DECLINED": pcv_PayStatusName=3 '// "Declined"	
		Case "CANCELLED": pcv_PayStatusName=4 '// "Cancelled"	
		Case "CANCELLED_BY_GOOGLE": pcv_PayStatusName=5 '// "Cancelled By Google"
		Case Else: pcv_PayStatusName=0 '// "Pending"
	End Select

	pcv_PaymentStatus = pcv_PayStatusName
	
	'// Update Financial Status
	query="UPDATE Orders SET pcOrd_PaymentStatus=" & pcv_PaymentStatus & " WHERE pcOrd_GoogleIDOrder='"& pcv_strOrderNumber &"' "
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	
	'// Synchronize Fullment Status with ProductCart
	Select Case pcv_strNewFulfillmentOrderState	
		Case "NEW": os=2 '// "Pending"	
		Case "PROCESSING": os=3 '// "Processed"	
		Case "DELIVERED": os=10 '// "Delivered"	aka "Shipped"
		Case "WILL_NOT_DELIVER": os=5 '// "Will Not Deliver" aka "Canceled"	
	End Select
	
	
	'// Update FullFillment Status
	query="UPDATE Orders SET orderstatus=" & os & " WHERE pcOrd_GoogleIDOrder='"& pcv_strOrderNumber &"' "
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	
	
	'// When the status changes to cancelled
	IF pcv_PayStatusName=4 OR pcv_PayStatusName=5 THEN
		
		qry_ID=pidOrder
		
		'// Start SDBA
		pcv_SubmitType=0
			
		'// Find out original status first
		query="SELECT orderstatus,paymentCode FROM orders WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		porigstatus=rs("orderstatus")
		pPaymentCode=rs("paymentCode")		

		
		'// Update orderstatus to 5(canceled) and input form variables
		adminComments=""
		query="UPDATE orders SET orderstatus=5 WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		if pcv_SubmitType=0 then
		Set rs=conntemp.execute(query)
		end if
		
		'// Update startDate
		StartDate=CDate("1/1/2000")
		if scDB="Access" then 
			query="UPDATE DPRequests SET StartDate=#"&StartDate&"# WHERE idOrder="& qry_ID
		else
			query="UPDATE DPRequests SET StartDate='"&StartDate&"' WHERE idOrder="& qry_ID
		end if
		Set rs=Server.CreateObject("ADODB.Recordset")
		if pcv_SubmitType=0 then
		Set rs=conntemp.execute(query)
		end if
		
		'// GGG Add-on start
	
			query="UPDATE pcGCOrdered SET pcGO_Status=0 WHERE pcGO_idOrder="& qry_ID
			Set rs=conntemp.execute(query)
			
			query="Select total,pcOrd_GcCode,pcOrd_GcUsed from orders WHERE idOrder="& qry_ID
			Set rs=conntemp.execute(query)
			
			ototal=rs("total")
			pGiftCode=rs("pcOrd_GcCode")
			pGiftUsed=rs("pcOrd_GcUsed")
			
			if pGiftCode<>"" then
			ototal=cdbl(ototal)+cdbl(pGiftUsed)
			query="update orders set total=" & ototal & ",pcOrd_GcCode='',pcOrd_GcUsed=0 WHERE idOrder="& qry_ID
			Set rs=conntemp.execute(query)
			
			query="select pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
			Set rs=conntemp.execute(query)
			
			pGCAmount=rs("pcGO_Amount")
			pGCStatus=rs("pcGO_Status")
			
			if pGCAmount="0" then
			pGCStatus=1
			end if
			
			pGCAmount=cdbl(pGCAmount)+cdbl(pGiftUsed)
			
			query="update pcGCOrdered set pcGO_Amount=" & pGCAmount & ",pcGO_Status=" & pGCStatus & " WHERE pcGO_GcCode='" & pGiftCode & "'"
			Set rs=conntemp.execute(query)
			end if '// Have Gift Code
			
			'// Increase Remaining Products of Gift Registry
			query="Select pcPO_EPID,quantity from ProductsOrdered WHERE idOrder="& qry_ID
			Set rs=conntemp.execute(query)
			do while not rs.eof
				geID=rs("pcPO_EPID")
				if geID<>"" then
				else
				geID="0"
				end if
				gQty=rs("quantity")
				if gQty<>"" then
				else
				gQty="0"
				end if
				if geID<>"0" then
				query="Update pcEvProducts set pcEP_HQty=pcEP_HQty-" & gQty & " WHERE pcEP_ID="& geID
				Set rs1=conntemp.execute(query)
				end if
				rs.MoveNext
			loop
			
			set rs=nothing
			set rs1=nothing
	
		'// GGG Add-on end
		
		'// Set any Auth or PFP orders to captured
		select case pPaymentCode
			case "PFLink", "PFPro", "PFPRO", "PFLINK"
				query="UPDATE pfporders SET captured=1 WHERE idOrder="& qry_ID
				Set rs=Server.CreateObject("ADODB.Recordset")
				if pcv_SubmitType=0 then
				Set rs=conntemp.execute(query)
				end if
			case "Authorize"
				query="UPDATE authorders SET captured=1 WHERE idOrder="& qry_ID
				Set rs=Server.CreateObject("ADODB.Recordset")
				if pcv_SubmitType=0 then
				Set rs=conntemp.execute(query)
				end if
		end select
		
		'// Update reward pts.
		query="Select * FROM orders WHERE idOrder="& qry_ID
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		pIdCustomer=rs("idcustomer")
	
			piRewardPoints=rs("iRewardPoints")
			piRewardRefId=rs("iRewardRefId")
			piRewardPointsRef=rs("iRewardPointsRef") 
			piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
			'take away points from refferer if any points were awarded. if order was processed
			If porigstatus<>"2" then
				If piRewardRefId>0 AND piRewardPointsRef>0 then
					query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & piRewardRefId
					set rsCust=conntemp.execute(query)
					iAccrued=rsCust("iRewardPointsAccrued") - piRewardPointsRef
					query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & piRewardRefId
					if pcv_SubmitType=0 then
						set rsCust=conntemp.Execute(query)
					end if
				end if 
				'take away accrued points from customer if any points were accrued
				If piRewardPointsCustAccrued>0 then
					query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
					set rsCust=conntemp.execute(query)
					iAccrued=rsCust("iRewardPointsAccrued") - piRewardPointsCustAccrued
					query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & pIdCustomer
					if pcv_SubmitType=0 then
						set rsCust=conntemp.Execute(query)
					end if
				end if
				If piRewardPoints>0 then
					query="SELECT iRewardPointsUsed, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
					set rsCust=conntemp.execute(query)
					iUsed=rsCust("iRewardPointsUsed") - piRewardPoints
					query="UPDATE customers SET iRewardPointsUsed=" & iUsed & " WHERE idCustomer=" & pIdCustomer
					if pcv_SubmitType=0 then
						set rsCust=conntemp.Execute(query)
					end if
				end if
			end if
	
		query="Select name,lastname,email,customercompany FROM customers WHERE idcustomer="& pIdCustomer
		Set rsCust=Server.CreateObject("ADODB.Recordset")
		Set rsCust=conntemp.execute(query)
		if porigstatus="3" or porigstatus="2" then
			query="SELECT idproduct,quantity,idconfigSession FROM ProductsOrdered WHERE ProductsOrdered.idOrder="& qry_ID
			set rsOrderDetails=conntemp.execute(query)
			Do While Not rsOrderDetails.EOF
				pidproduct=rsOrderDetails("idproduct")
				pqty=rsOrderDetails("quantity")
				idconfig=rsOrderDetails("idconfigSession")
				'check if stock is ignored or not
				query="SELECT noStock FROM products WHERE idProduct="&pIdProduct
				set rsStockObj=server.CreateObject("ADODB.RecordSet")
				set rsStockObj=conntemp.execute(query)   
				pNoStock=rsStockObj("noStock")
				set rsStockObj=nothing
				'---------------
				' increase stock 
				'--------------- 
				query="SELECT stock, sales, description FROM products WHERE idProduct="&pidproduct
				set rsStockObj=server.CreateObject("ADODB.RecordSet")
				set rsStockObj=conntemp.execute(query) 
				if pNoStock=0 then
					query="UPDATE products SET stock=stock+"&pqty&" WHERE idProduct="&pidproduct
					if pcv_SubmitType=0 then
					set rsStockObj=conntemp.execute(query)  
					end if					
					'Update BTO Items & Additional Charges stock and sales 
					IF (idconfig<>"") and (idconfig<>"0") then
					query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
					set rs1=conntemp.execute(query)
					stringProducts=rs1("stringProducts")
					stringQuantity=rs1("stringQuantity")
					stringCProducts=rs1("stringCProducts")
					if (stringProducts<>"") and (stringProducts<>"na") then
						PrdArr=split(stringProducts,",")
						QtyArr=split(stringQuantity,",")
						
						for k=lbound(PrdArr) to ubound(PrdArr)
							if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
								query="UPDATE products SET stock=stock+" &QtyArr(k)*pqty&",sales=sales-" &QtyArr(k)*pqty&" WHERE idProduct=" &PrdArr(k)
								if pcv_SubmitType=0 then
								set rs1=conntemp.execute(query)
								end if
							end if
						next
					end if
					if (stringCProducts<>"") and (stringCProducts<>"na") then
						CPrdArr=split(stringCProducts,",")
						
						for k=lbound(CPrdArr) to ubound(CPrdArr)
							if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
								query="UPDATE products SET stock=stock+" &pqty&",sales=sales-" &pqty&" WHERE idProduct=" &CPrdArr(k)
								if pcv_SubmitType=0 then
								set rs1=conntemp.execute(query)
								end if
							end if
						next
					end if
				END IF
				'End Update BTO Items & Additional Charges
	
				end if
				set rsStockObj=nothing 
				'--------------
				' end increase stock
				'-------------- 
				'--------------
				'update sales
				'--------------
				query="UPDATE products SET sales=sales-"&pqty&" WHERE idProduct="&pidproduct
				if pcv_SubmitType=0 then
				set rsSalesObj=conntemp.execute(query)
				end if
				set rsSalesObj=nothing 
				'--------------
				' end update sales
				'--------------
				set rsStockObj=nothing
				rsOrderDetails.MoveNext
			loop
		end if
		
		'// Send email to customer
		customerCancelledEmail=Cstr("")
	
		'// Customized message from store owner
		If scCancelledEmail<>"" Then
			todaydate=showDateFrmt(now())
			personalmessage=replace(scCancelledEmail,"<br>", vbCrlf)
			personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
			personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
			personalmessage=replace(personalmessage,"<TODAY_DATE>",todaydate)
			personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",rsCust("name")&" "&rsCust("lastname"))
			personalmessage=replace(personalmessage,"<ORDER_ID>",(scpre + int(qry_ID)))
			personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(rs("orderDate")))
			personalmessage=replace(personalmessage,"//","/")
			personalmessage=replace(personalmessage,"http:/","http://")
			personalmessage=replace(personalmessage,"https:/","https://")
			customerCancelledEmail=customerCancelledEmail & vbCrLf & personalmessage & vbCrLf
			customerCancelledEmail=replace(customerCancelledEmail,"''",chr(39))
			pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_8")
			'// PAYPALEXP EMAIL
			pEmail=rsCust("email")
			if pcv_strPrevFulfillmentOrderState="NEW"<>"" then
				call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerCancelledEmail, "&quot;", chr(34)))
			else
				call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerCancelledEmail, "&quot;", chr(34)))
			end if
		end if
	END IF	
	
	
	'// When the status changes from Authorized to Paid we complete the following processes
	IF pcv_PayStatusName=2 AND pcv_strPrevFinancialOrderState<>"CHARGED" THEN
	
		pOrderStatus=2
		pCheckEmail="YES"
		qry_ID=pIdOrder	

		'------------------------------------------------
		'- Look for downloadable products
		'------------------------------------------------
		query="select idproduct,idconfigSession from ProductsOrdered WHERE idOrder="&pIdOrder&";"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		DPOrder="0"
		do while not rs.eof
			pTempProductId=rs("idproduct")
			tmpidConfig=rs("idconfigSession")
			query="select downloadable from products where idproduct=" & pTempProductId
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=connTemp.execute(query)
			if not rstemp.eof then
				pdownloadable=rstemp("downloadable")
				if (pdownloadable<>"") and (pdownloadable="1") then
					DPOrder="1"
				end if
			end if
			set rstemp=nothing
			'Find downloadable items in BTO configuration
			if tmpidConfig<>"" AND tmpidConfig>"0" then
				query="SELECT stringProducts,stringQuantity,stringCProducts FROM configSessions WHERE idconfigSession=" & tmpidConfig & ";"
				set rs1=connTemp.execute(query)
				if not rs1.eof then
					stringProducts=rs1("stringProducts")
					stringQuantity=rs1("stringQuantity")
					stringCProducts=rs1("stringCProducts")
					if (stringProducts<>"") and (stringProducts<>"na") then
						PrdArr=split(stringProducts,",")
						QtyArr=split(stringQuantity,",")
						
						for k=lbound(PrdArr) to ubound(PrdArr)
							if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
								query="SELECT idproduct FROM Products WHERE idProduct=" & PrdArr(k) & " AND Downloadable=1;"
								set rs1=conntemp.execute(query)
								if not rs1.eof then
									DPOrder="1"
								end if
								set rs1=nothing
							end if
						next
					end if
					if (stringCProducts<>"") and (stringCProducts<>"na") then
						CPrdArr=split(stringCProducts,",")
						for k=lbound(CPrdArr) to ubound(CPrdArr)
							if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
								query="SELECT idproduct FROM Products WHERE idProduct=" & CPrdArr(k) & " AND Downloadable=1;"
								set rs1=conntemp.execute(query)
								if not rs1.eof then
									DPOrder="1"
								end if
								set rs1=nothing
							end if
						next
					end if
				end if
				set rs1=nothing
			end if
		rs.moveNext
		loop
		set rs=nothing
		
		'------------------------------------------------
		'- Look for gift certificates
		'------------------------------------------------
		query="select idproduct from ProductsOrdered WHERE idOrder="& qry_ID
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		pGCs="0"

		do while not rs.eof
			pTempProductId=rs("idproduct")
			query="select pcprod_GC from products where idproduct=" & pTempProductId & ";"
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=connTemp.execute(query)			
			if not rstemp.eof then
				pGC=rstemp("pcprod_GC")
				if (pGC<>"") and (pGC="1") then
					pGCs="1"
				end if
			end if
			set rstemp=nothing
			rs.moveNext
		loop
		set rs=nothing		


		'------------------------------------------------
		'- Get today's date
		'------------------------------------------------
		Dim pTodaysDate
		pTodaysDate=Date()
		if SQL_Format="1" then
			pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
		else
			pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
		end if
		
		'------------------------------------------------
		'- Update the order information and status
		'------------------------------------------------
		if scDB="Access" then
			query="UPDATE orders SET pcOrd_GCs=" & pGCs & ",DPs=" & DPOrder & ", orderstatus=3, processDate=#"& pTodaysDate &"# WHERE idOrder="&pIdOrder&";"
		else
			query="UPDATE orders SET pcOrd_GCs=" & pGCs & ",DPs=" & DPOrder & ", orderstatus=3, processDate='"& pTodaysDate &"' WHERE idOrder="&pIdOrder&";"
		end if
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		set rs=nothing
	
		'------------------------------------------------
		'- Get customer information
		'------------------------------------------------
		query="select idcustomer,orderdate,processdate from Orders WHERE idOrder="&pIdOrder&";"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=conntemp.execute(query)
		if not rs.eof then
			pIdCustomer=rs("IdCustomer")
			pOrderDate=rs("OrderDate")
			pProcessDate=rs("ProcessDate")
		end if
		Set rs=nothing


		'------------------------------------------------
		'- START: Create licenses for downloadable products
		'------------------------------------------------
		
	IF DPOrder="1" then
		query="select idproduct,quantity,idconfigSession from ProductsOrdered WHERE idOrder="& qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
	
		do while not rs.eof
			pIdProduct=rs("idproduct")
			pQuantity=rs("quantity")
			tmpidConfig=rs("idconfigSession")
			Call CreateDownloadInfo1(pIDProduct,pQuantity)
			'Find downloadable items in BTO configuration
			if tmpidConfig<>"" AND tmpidConfig>"0" then
				query="SELECT stringProducts,stringQuantity,stringCProducts FROM configSessions WHERE idconfigSession=" & tmpidConfig & ";"
				set rs1=connTemp.execute(query)
				if not rs1.eof then
					stringProducts=rs1("stringProducts")
					stringQuantity=rs1("stringQuantity")
					stringCProducts=rs1("stringCProducts")
					if (stringProducts<>"") and (stringProducts<>"na") then
						PrdArr=split(stringProducts,",")
						QtyArr=split(stringQuantity,",")
					
						for k=lbound(PrdArr) to ubound(PrdArr)
							if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
								query="SELECT idproduct FROM Products WHERE idProduct=" & PrdArr(k) & " AND Downloadable=1;"
								set rs1=conntemp.execute(query)
								if not rs1.eof then
									Call CreateDownloadInfo1(PrdArr(k),QtyArr(k)*pQuantity)
								end if
								set rs1=nothing
							end if
						next
					end if
					if (stringCProducts<>"") and (stringCProducts<>"na") then
						CPrdArr=split(stringCProducts,",")
						for k=lbound(CPrdArr) to ubound(CPrdArr)
							if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
								query="SELECT idproduct FROM Products WHERE idProduct=" & CPrdArr(k) & " AND Downloadable=1;"
								set rs1=conntemp.execute(query)
								if not rs1.eof then
									Call CreateDownloadInfo1(CPrdArr(k),1)
								end if
								set rs1=nothing
							end if
						next
					end if
				end if
				set rs1=nothing
			end if
			rs.moveNext
		loop
		set rs=nothing
	END IF
		'------------------------------------------------
		'- END: Create licenses for downloadable products
		'------------------------------------------------


		'------------------------------------------------
		'- START: Create Gift Certificate code
		'------------------------------------------------
		IF pGCs="1" then
			query="select idproduct,quantity from ProductsOrdered WHERE idOrder="& qry_ID
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=connTemp.execute(query)
			DO while not rstemp.eof
				query="select pcGC.pcGC_Exp,pcGC.pcGC_ExpDate,pcGC.pcGC_ExpDays,pcGC.pcGC_CodeGen,pcGC.pcGC_GenFile,products.sku,products.price from pcGC,Products where pcGC.pcGC_idproduct=" & rstemp("idproduct") & " and Products.idproduct=pcGC.pcGC_idproduct and products.pcprod_GC=1"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
		
				if not rs.eof then
					pIdproduct=rstemp("idproduct")
					pQuantity=rstemp("quantity")
					pGCExp=rs("pcGC_Exp")
					pGCExpDate=rs("pcGC_ExpDate")
					pGCExpDay=rs("pcGC_ExpDays")
					pGCGen=rs("pcGC_CodeGen")
					pGCGenFile=rs("pcGC_GenFile")
					pSku=rs("sku")
					pGCAmount=rs("price")
					if pGCGen<>"" then
					else
						pGCGen="0"
					end if
					if (pGCGen=1) and (pGCGenFile="") then
						pGCGen="0"
						pGCGenFile="DefaultGiftCode.asp"
					end if
	
					if (pGCGen="0") or (not (pGCGenFile<>"")) then
						pGCGenFile="DefaultGiftCode.asp"
					end if
					
					if (pGCExp="2") then
						pGCExpDate=Now()+cint(pGCExpDay)
					end if
					
					if (pGCExp="1") and (year(pGCExpDate)=1900) then
						pGCExp="0"
						pGCExpDate="01/01/1900"
					end if
					
					if (pGCExp="2") and (pGCExpDay="0") then
						pGCExp="0"
						pGCExpDate="01/01/1900"
					end if
					
					if SQL_Format="1" then
						pGCExpDate=(day(pGCExpDate)&"/"&month(pGCExpDate)&"/"&year(pGCExpDate))
					else
						pGCExpDate=(month(pGCExpDate)&"/"&day(pGCExpDate)&"/"&year(pGCExpDate))
					end if

					IF (pGCGenFile<>"") THEN

							SPath1=Request.ServerVariables("PATH_INFO")
							mycount1=0
							do while mycount1<1
								if mid(SPath1,len(SPath1),1)="/" then
								mycount1=mycount1+1
								end if
								if mycount1<1 then
								SPath1=mid(SPath1,1,len(SPath1)-1)
								end if
							loop
							SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
							if Right(SPathInfo,1)="/" then
								pGCGenFile=SPathInfo & "licenses/" & pGCGenFile					
							else
								pGCGenFile=SPathInfo & "/licenses/" & pGCGenFile
							end if
							L_Action=pGCGenFile
							
						L_postdata=""
						L_postdata=L_postdata&"idorder=" & pIdOrder
						L_postdata=L_postdata&"&orderDate=" & pOrderDate
						L_postdata=L_postdata&"&ProcessDate=" & pProcessDate
						L_postdata=L_postdata&"&idcustomer=" & pIdCustomer
						L_postdata=L_postdata&"&idproduct=" & pIdproduct
						L_postdata=L_postdata&"&quantity=" & pQuantity
						L_postdata=L_postdata&"&sku=" & pSKU

						For k=1 to Cint(pQuantity)
						
						DO
						
						Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
						srvXmlHttp.open "POST", L_Action, False
						srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
						srvXmlHttp.send L_postdata
						result1 = srvXmlHttp.responseText
						
						RArray = split(result1,"<br>")
						GiftCode= RArray(2)
						
						'If have errors from GiftCode Generator
						IF (IsNumeric(RArray(0))=false) and (IsNumeric(RArray(1))=false) then
						
						Tn1=""
						For w=1 to 6
						Randomize
						myC=Fix(3*Rnd)
						Select Case myC
							Case 0: 
							Randomize
							Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
							Case 1: 
							Randomize
							Tn1=Tn1 & Cstr(Fix(10*Rnd))
							Case 2: 
							Randomize
							Tn1=Tn1 & Chr(Fix(26*Rnd)+97)		
						End Select
						Next
						
						GiftCode=Tn1 & Day(Now()) & Minute(Now()) & Second(Now())
						
						END IF
						
						ReqExist=0
					
						query="select pcGO_IDProduct from pcGCOrdered where pcGO_GcCode='" & GiftCode & "'" 
						set rstemp2=Server.CreateObject("ADODB.Recordset")
						set rstemp2=connTemp.execute(query)					
						if not rstemp2.eof then
							ReqExist=1
						end if
					
						LOOP UNTIL ReqExist=0
						set rstemp2=nothing
						
						'Insert Gift Codes to Database
	
						if scDB="Access" then
							query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "',#" & pGCExpDate & "#," & pGCAmount & ",1)"   
						else
							query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "','" & pGCExpDate & "'," & pGCAmount & ",1)"
						end if
						set rstemp2=Server.CreateObject("ADODB.Recordset")
						set rstemp2=connTemp.execute(query)
						set rstemp2=nothing
						
						Next
		
					END IF
			
				end if
				rstemp.moveNext
			LOOP
			set rstemp=nothing
		END IF
		'------------------------------------------------
		'- END: Create Gift Certificate code
		'------------------------------------------------
		
		'------------------------------------------------
		'- START: Send confirmation email
		'------------------------------------------------
		' Get order information from the database
		query="SELECT orders.idcustomer,orders.address,orders.City,orders.StateCode,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingZip,orders.shippingCountryCode, orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.iRewardValue,orders.iRewardRefId, orders.iRewardPointsRef,orders.iRewardPointsCustAccrued,customers.phone,ord_DeliveryDate,ord_VAT, pcOrd_CatDiscounts FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" & qry_ID
		Set rsEmailInfo=Server.CreateObject("ADODB.Recordset")
		Set rsEmailInfo=connTemp.execute(query)
			pidcustomer=rsEmailInfo("idcustomer")
			paddress=rsEmailInfo("address")
			pCity=rsEmailInfo("city")
			pStateCode=rsEmailInfo("StateCode")
			pzip=rsEmailInfo("zip")
			pCountryCode=rsEmailInfo("CountryCode")
			pshippingAddress=rsEmailInfo("shippingAddress")
			pshippingCity=rsEmailInfo("shippingCity")
			pshippingStateCode=rsEmailInfo("shippingStateCode")
			pshippingZip=rsEmailInfo("shippingZip")
			pshippingCountryCode=rsEmailInfo("shippingCountryCode")
			pShipmentDetails=rsEmailInfo("ShipmentDetails")
			pPaymentDetails=rsEmailInfo("paymentDetails")
			pdiscountDetails=rsEmailInfo("discountDetails")
			ptaxAmount=rsEmailInfo("taxAmount")
			ptotal=rsEmailInfo("total")
			pcomments=rsEmailInfo("comments")
			pShippingFullName=rsEmailInfo("ShippingFullName")
			paddress2=rsEmailInfo("address2")
			pShippingCompany=rsEmailInfo("ShippingCompany")
			pShippingAddress2=rsEmailInfo("ShippingAddress2")
			ptaxDetails=rsEmailInfo("taxDetails")
			piRewardValue=rsEmailInfo("iRewardValue")
			piRewardRefId=rsEmailInfo("iRewardRefId")
			piRewardPointsRef=rsEmailInfo("iRewardPointsRef")
			piRewardPointsCustAccrued=rsEmailInfo("iRewardPointsCustAccrued")
			pPhone=rsEmailInfo("phone")
			pord_DeliveryDate=rsEmailInfo("ord_DeliveryDate")
			pord_VAT=rsEmailInfo("ord_VAT")
			pcOrd_CatDiscounts=rsEmailInfo("pcOrd_CatDiscounts")
		set rsEmailInfo=nothing
	
		pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)

		
		'// Get customer details for this order
		query="Select name,lastname,customerCompany,email FROM customers WHERE idcustomer="& pIdCustomer
		Set rsCust=Server.CreateObject("ADODB.Recordset")
		Set rsCust=conntemp.execute(query)
			pName=rsCust("name")
			pLName=rsCust("lastname")
			pCustomerCompany=rsCust("customerCompany")
			pEmail=rsCust("email")
		Set rsCust=nothing

		'// Send Order Confirmation email to customer, if checked
		if pCheckEmail="YES" then%>
			<!--#include file="pcPay_GoogleCheckout_CustomerEmail.asp"-->
			<% 
			pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_6")			
			call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerEmail, "&quot;", chr(34)))
		end if


		'// Start SDBA		
		pcv_DropShipperID=0
		pcv_IsSupplier=0 
		%> <!--#include file="inc_DropShipperNotificationEmail.asp"--> <%		
		'// End SDBA
		'------------------------------------------------
		'- END: Send confirmation email
		'------------------------------------------------		
		
		
		
		'------------------------------------------------
		'- START: Update Reward Points
		'------------------------------------------------
		If RewardsActive <> 0 then
			'add points from refferer if any points were awarded.
			If piRewardRefId>0 AND piRewardPointsRef>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & piRewardRefId
				Set rsCust=Server.CreateObject("ADODB.Recordset")
				set rsCust=conntemp.execute(query)
				iAccrued=rsCust("iRewardPointsAccrued") + piRewardPointsRef
				set rsCust=nothing
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & piRewardRefId
				set rsCust=server.CreateObject("ADODB.RecordSet")
				set rsCust=conntemp.Execute(query)
				set rsCust=nothing
			end if 
			'add accrued points from customer if any points were accrued
			If piRewardPointsCustAccrued>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
				Set rsCust=Server.CreateObject("ADODB.Recordset")
				set rsCust=conntemp.execute(query)
				iAccrued=rsCust("iRewardPointsAccrued") + piRewardPointsCustAccrued
				set rsCust=nothing
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & pIdCustomer
				set rsCust=server.CreateObject("ADODB.RecordSet")
				set rsCust=conntemp.Execute(query)
				set rsCust=nothing
			End If
		End If 
		'------------------------------------------------
		'- END: Update Reward Points
		'------------------------------------------------
	
		'------------------------------------------------
		'- Create Report on processed orders
		'------------------------------------------------
		successCnt=successCnt+1
		successData=successData&"Order Number "& (int(pIdOrder)+scpre) &" was processed successfully<BR>"
	END IF	 

	Set domMcCallbackObjRoot = Nothing

End Function
Public Function FixedField(ByVal Width, ByVal Justify, ByVal Text)

	Select Case True
		Case Width < Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Left(Text, Width)
				Case Justify="R"
					FixedField=Right(Text, Width)
				Case Else
			End Select
									
		Case Width=Len(Text)
			FixedField=Text

		Case Width > Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Text & String(Width - Len(Text), " ")
				Case Justify="R"
					FixedField=String(Width - Len(Text), " ") & Text
				Case Else
			End Select

	End Select

End Function 
'***********************************************************************************
' END: RECIEVE NEW STATUS CODE
'***********************************************************************************





'***********************************************************************************
' START: RISK MANAGEMENT
'***********************************************************************************
Function processRiskInformationNotification(domResponseObj)
	on error resume next
			
	Dim xmlMcResults
	
	'// Process <risk-information-notification>
    xmlMcResults = createRiskInformationNotification(domResponseObj)
	
	'// Respond with <risk-information-notification> XML
    Response.write xmlMcResults	
	
    '// <risk-information-notification>
    sendNotificationAcknowledgment
	
End Function
'***********************************************************************************
' END: RISK MANAGEMENT
'***********************************************************************************





'***********************************************************************************
' START: RECIEVE NEW STATUS CODE
'***********************************************************************************
Function createRiskInformationNotification(domMcCallbackObj)
	on error resume next

    '// Define the objects used to read the xml <risk-information-notification>
    Dim domMcCallbackObjRoot
 	Dim domOrderNumberList
	Dim pcv_strOrderNumber
	Dim Nodes
	Dim Node

	Set domMcCallbackObjRoot = domMcCallbackObj.documentElement

    Set Nodes = domMcCallbackObj.selectNodes("//risk-information-notification")	

	For Each Node In Nodes	
		pcv_strOrderNumber = Node.selectSingleNode("google-order-number").text			
		pcv_strIpAddress = Node.selectSingleNode("risk-information/ip-address").text
		pcv_strEligibleForProtection = Node.selectSingleNode("risk-information/eligible-for-protection").text
		pcv_strAVSRespond = Node.selectSingleNode("risk-information/avs-response").text
		pcv_strCVNResponse = Node.selectSingleNode("risk-information/cvn-response").text
		pcv_strPartialCCNumber = Node.selectSingleNode("risk-information/partial-cc-number").text
		pcv_strwBuyerAccountAge = Node.selectSingleNode("risk-information/buyer-account-age").text		
	Next	
	
	'// Update Risk Management Status
	'save only the first 15 characters in case this is returned as a list of IP addresses
	pcv_strIpAddress = left(pcv_strIpAddress,15)
	
	query="UPDATE Orders SET pcOrd_CustomerIP='" & pcv_strIpAddress & "', pcOrd_EligibleForProtection='" & pcv_strEligibleForProtection & "', pcOrd_AVSRespond='" & pcv_strAVSRespond & "', pcOrd_CVNResponse='" & pcv_strCVNResponse & "', pcOrd_PartialCCNumber='" & pcv_strPartialCCNumber & "', pcOrd_BuyerAccountAge='" & pcv_strwBuyerAccountAge & "' WHERE pcOrd_GoogleIDOrder='"& pcv_strOrderNumber &"' "
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	
	Set domMcCallbackObjRoot = Nothing

End Function
'***********************************************************************************
' END: RECIEVE NEW STATUS CODE
'***********************************************************************************



Function processChargeAmountNotification(domResponseObj)
    sendNotificationAcknowledgment
End Function

Function processChargebackAmountNotification(domResponseObj)
    sendNotificationAcknowledgment
End Function

Function processRefundAmountNotification(domResponseObj)
    sendNotificationAcknowledgment
End Function


'*******************************************************************************
' The sendNotificationAcknowledgment function responds to a Google Checkout
' notification with a <notification-acknowledgment> message. If you do
' not send a <notification-acknowledgment> in response to a Google Checkout
' notification, Google Checkout will resend the notification multiple times.
'*******************************************************************************
Function sendNotificationAcknowledgment

    ' Respond to the notification with <notification-acknowledgment>
    xmlAcknowledgment = _
      "<?xml version=""1.0"" encoding=""UTF-8""?>" _
      & "<notification-acknowledgment " _
      & "xmlns=""" & strXmlns & """/>"
    response.write xmlAcknowledgment
    
    ' Log <notification-acknowledgment>
    logMessage logFilename, xmlAcknowledgment

End Function



'*******************************************************************************
' The createAddMerchantOrderNumber function creates the XML for the
' <add-merchant-order-number> Order Processing API command. This command
' is used to associate the Google order number with the ID that the
' merchant assigns to the same order.
'
' Inputs:   attrGoogleOrderNumber    A number, assigned by Google Checkout, that
'                                    that uniquely identifies an order.
'           elemMerchantOrderNumber  A string, assigned by the merchant, 
'                                    that uniquely identifies the order.
' Returns:  <add-merchant-order-number> XML
'*******************************************************************************
Function createAddMerchantOrderNumber(attrGoogleOrderNumber, _
    elemMerchantOrderNumber)

    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createAddMerchantOrderNumber()"

    ' Check for missing parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrGoogleOrderNumber", _
        attrGoogleOrderNumber
    checkForError errorType, strFunctionName, "elemMerchantOrderNumber", _
        elemMerchantOrderNumber

    ' Define the objects used to create the <add-merchant-order-number> command
    Dim domAddMerchantOrderNumberObj
    Dim domAddMerchantOrderNumber
    Dim domMerchantOrderNumber

    Set domAddMerchantOrderNumberObj = Server.CreateObject(strMsxmlDomDocument)
    domAddMerchantOrderNumberObj.async = False
    domAddMerchantOrderNumberObj.appendChild( _
        domAddMerchantOrderNumberObj.createProcessingInstruction("xml", _
            strXmlVersionEncoding))

    ' Create the root tag for the Order Processing API command.
    ' Also set the "xmlns" and "google-order-number" attributes
    ' on that element.
    Set domAddMerchantOrderNumber = domAddMerchantOrderNumberObj.appendChild( _
        domAddMerchantOrderNumberObj.createElement("add-merchant-order-number"))

    domAddMerchantOrderNumber.setAttribute "xmlns", strXmlns

    domAddMerchantOrderNumber.setAttribute "google-order-number", _
        attrGoogleOrderNumber

    ' Add the <merchant-order-number> element
    Set domMerchantOrderNumber = domAddMerchantOrderNumber.appendChild( _
        domAddMerchantOrderNumberObj.createElement("merchant-order-number"))

    domMerchantOrderNumber.Text = elemMerchantOrderNumber

    createAddMerchantOrderNumber = domAddMerchantOrderNumberObj.xml
End Function
%>

