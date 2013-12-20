<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../pc/gwUSAePay_xcenums.inc"-->
<!--#include file="inc_GenDownloadInfo.asp"-->
<!--#include file="adminHeader.asp" -->
<% 
Dev_Testmode = 0
on error resume next
dim query, conntemp, rs, rstemp, rstemp2, rsEmailInfo, rsCust, pTempProductId
call opendb()

	query="SELECT pcPay_Uep_SourceKey,pcPay_Uep_TransType,pcPay_Uep_TestMode FROM pcPay_USAePay WHERE pcPay_Uep_Id=1"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	dim pcPay_Uep_SourceKey,pcPay_Uep_TransType,pcPay_Uep_TestMode
	pcPay_Uep_SourceKey=rs("pcPay_Uep_SourceKey")
	pcPay_Uep_TransType=rs("pcPay_Uep_TransType")
	pcPay_Uep_TestMode=rs("pcPay_Uep_TestMode")
	'decrypt
	pcPay_Uep_SourceKey=enDeCrypt(pcPay_Uep_SourceKey, scCrypPass)
	sIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If sIPAddress="" Then 
		sIPAddress = Request.ServerVariables("REMOTE_ADDR")
	end if

call closeDb()

	dim successCnt, successData, failedCnt, failedData
	successCnt=0
	successData="" 
	failedCnt=0
	failedData=""


			
'how many checkboxes?
dim checkboxCnt
checkboxCnt=request.Form("checkboxCnt")

'do for each checkbox
dim r, orderVoid
for r=1 to checkboxCnt
	orderVoid=0
	IF request.Form("checkOrd"&r)="YES" THEN
		pidePayOrder=request.Form("idePayOrder"&r)
		authamount=request.Form("authamount"&r)
		pOrderStatus=request.Form("orderstatus"&r)
		pCheckEmail=request.Form("checkEmail"&r)
		pIdOrder=Request.Form("idOrder"&r)  & ""
		qry_ID=pIdOrder
	
		Dim objXMLHTTP, xml

		'Send the request to the Authorize.NET processor.
		stext="UMcommand=cc:capture"
		stext=stext & "&UMkey=" & pcPay_Uep_SourceKey
		stext=stext & "&UMamount=" & request.Form("curamount"&r)     
		stext=stext & "&UMrefNum =" & Request.Form("RefNum"&r)
		'Send the transaction info as part of the querystring
		set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)

		if Dev_Testmode = 1 Then
			xml.open "POST", "https://sandbox.usaepay.com/gate?"& stext & "", false
		else
			xml.open "POST", "https://www.usaepay.com/gate?"& stext & "", false
		end if

		xml.send ""
		strStatus = xml.Status

		'store the response
		strRetVal = xml.responseText
		Set xml = Nothing

		' PARSE HTTP RESPONSE INTO DICTIONARY OBJECT FOR EASIER ACCESS
		ResponseArray = Split(strRetVal, "&")
		Set objDictResponse = server.createobject("Scripting.Dictionary")
		For each ResponseItem in ResponseArray
			NameValue = Split(ResponseItem, "=")
			objDictResponse.Add NameValue(0), NameValue(1)
		Next

		' Parse the response into local vars
		UMstatus = objDictResponse.Item("UMstatus") 'Status of the transaction. The possible values are: Approved, Declined, Verification and Error.
		UMresult = objDictResponse.Item("UMresult") 'Transaction result - A, D, E, or V (for Verification)
		UMerror = objDictResponse.Item("UMerror") 'Error description if UMstatus is Declined or Error.
		UMerrorcode = objDictResponse.Item("UMerrorcode") 'A numerical error code.

		uep_success=0
		
		'response.write UMstatus&"<BR>"
		'response.write UMresult&"<BR>"
		'response.write UMerror&"<BR>"
		'response.write UMerrorcode&"<BR>"
		'response.end
		If UMstatus = "Approved" Then
			uep_success=1
		End If
		
		'Check USAePay response code 1=approved 2=declined 3=error.
		if uep_success = 1 then
			'order was successful
			call opendb()
			
			'update authorders to captured
			query="UPDATE pcPay_USAePay_Orders SET captured=1 WHERE idePayOrder="&pidePayOrder&";"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
			
			'purge cc data
			query="SELECT ccnum, pcSecurityKeyID FROM pcPay_USAePay_Orders WHERE idePayOrder="&pidePayOrder&";"
			set rstemp=connTemp.execute(query)
			if NOT rstemp.EOF then
				cardnumber=rstemp("ccnum")
				tempSecurityKeyID=rstemp("pcSecurityKeyID")
			end if
			set rstemp=nothing
			
			tempfour=pcf_PurgeCardNumber(cardnumber,tempSecurityKeyID)

			query="UPDATE pcPay_USAePay_Orders SET ccnum='*' WHERE idePayOrder="&pidePayOrder&";"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
			
			'if order has not alread been processed
			IF pOrderStatus="2" THEN
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
					query="select pcprod_GC from products where idproduct=" & pTempProductId
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
				set rs=nothing
				
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
						Call CreateDownloadInfo(pIDProduct,pQuantity)
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
												Call CreateDownloadInfo(PrdArr(k),QtyArr(k)*pQuantity)
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
												Call CreateDownloadInfo(CPrdArr(k),1)
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
				query="SELECT orders.idcustomer,orders.address,orders.City,orders.StateCode,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingZip,orders.shippingCountryCode,orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.iRewardValue,orders.iRewardRefId, orders.iRewardPointsRef,orders.iRewardPointsCustAccrued,customers.phone,ord_DeliveryDate,ord_VAT, pcOrd_CatDiscounts FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" &qry_ID
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
				
				'Get customer details for this order
				query="Select name, lastname, customerCompany, email, pcCust_VATID, pcCust_SSN FROM customers WHERE idcustomer="& pIdCustomer
				Set rsCust=Server.CreateObject("ADODB.Recordset")
				Set rsCust=conntemp.execute(query)
				pName=rsCust("name")
				pLName=rsCust("lastname")
				pCustomerCompany=rsCust("customerCompany")
				pEmail=rsCust("email")
				pVATID=rsCust("pcCust_VATID")
				pSSN=rsCust("pcCust_SSN")
				set rsCust=nothing
				
				'Send Order Confirmation email to customer, if checked
				if pCheckEmail="YES" then%>
					<!--#include file="sendmailCustomerProcessed.asp"-->
					<% 
					pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_2") & " - " & dictLanguage.Item(Session("language")&"_sendMail_1") & (scpre + int(qry_ID))
					call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerEmail, "&quot;", chr(34)))
				end if
				'------------------------------------------------
				'- END: Send confirmation email
				'------------------------------------------------
				
				'Start SDBA
				'Send Order Notification E-mail to Drop-Shippers
				pcv_DropShipperID=0
				pcv_IsSupplier=0 %>
				<!--#include file="../pc/inc_DropShipperNotificationEmail.asp"-->
				<%
				'End SDBA
		
				'------------------------------------------------
				'- START: Update Reward Points
				'------------------------------------------------
				If RewardsActive <> 0 then
					'add points from refferer if any points were awarded.
					If piRewardRefId>0 AND piRewardPointsRef>0 then
						query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & piRewardRefId
						set rsCust=server.CreateObject("ADODB.RecordSet")
						set rsCust=conntemp.execute(query)
						iAccrued=rsCust("iRewardPointsAccrued") + piRewardPointsRef
						set rsCust=nothing
						
						query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & piRewardRefId
						set rsCust=server.CreateObject("ADODB.RecordSet")
						set rsCust=conntemp.execute(query)
						set rsCust=nothing
					end if 
					'add accrued points from customer if any points were accrued
					If piRewardPointsCustAccrued>0 then
						query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
						set rsCust=server.CreateObject("ADODB.RecordSet")
						set rsCust=conntemp.execute(query)
						iAccrued=rsCust("iRewardPointsAccrued") + piRewardPointsCustAccrued
						set rsCust=nothing
						query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & pIdCustomer
						set rsCust=server.CreateObject("ADODB.RecordSet")
						set rsCust=conntemp.execute(query)
						set rsCust=nothing
					End If
				End If 
				'------------------------------------------------
				'- END: Update Reward Points
				'------------------------------------------------

			END IF 'Order had not already been processed
			'Update the payment status to processed
			query="UPDATE orders SET pcOrd_PaymentStatus=2 WHERE idOrder="&qry_ID&";"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
			'------------------------------------------------
			'- Create Report on processed orders
			'------------------------------------------------
			successCnt=successCnt+1
			successData=successData&"Order Number "& (int(pIdOrder)+scpre) &" was processed successfully<BR>"
		else
			failedCnt=failedCnt+1
			failedData=failedData&"Order Number: "& (int(pIdOrder)+scpre) &" was NOT processed.<BR>" 
		end if 'Order was not approved
	END IF 
Next
call closedb()
%>
<table class="pcCPcontent">
  <tr>
    <td><div class="pcCPmessageSuccess"><%=successCnt%> records were successfully processed.</div>
			<% if successData<>"" then %>
				<br><%=successData%><br>
			<% end if %>
			<%if failedCnt>0 then%>
			<hr size="1" noshade>
			<div class="pcCPmessage"><%=failedCnt%> records failed.</div>
				<% if failedData<>"" then %>
					<br><%=failedData %><br>
				<% end if %>
			<%end if%>
		</td>
  </tr>
	<tr>
    <td><p>&nbsp;</p>
    <p><a href="resultsAdvancedAll.asp?B1=View%2BAll&dd=1">Manage Orders</a></p></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<% Public Function FixedField(ByVal Width, ByVal Justify, ByVal Text)

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

End Function %><!--#include file="adminFooter.asp" -->