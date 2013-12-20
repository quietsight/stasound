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
<!--#include file="inc_GenDownloadInfo.asp"-->
<!--#include file="adminHeader.asp" -->
<% on error resume next
dim query, conntemp, rs, rstemp, rstemp2, rsEmailInfo, rsCust, pTempProductId

call opendb()

query="SELECT pcPay_EIG_Type, pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key, pcPay_EIG_Curcode, pcPay_EIG_CVV, pcPay_EIG_TestMode, pcPay_EIG_SaveCards, pcPay_EIG_UseVault FROM pcPay_EIG WHERE pcPay_EIG_ID=1"
set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
x_Username=rs("pcPay_EIG_Username")
x_Username=enDeCrypt(x_Username, pcs_GetSecureKey)
x_Password=rs("pcPay_EIG_Password")
x_Password=enDeCrypt(x_Password, pcs_GetSecureKey)
x_Key=rs("pcPay_EIG_Key")
x_Key=enDeCrypt(x_Key, pcs_GetSecureKey)
x_CVV=rs("pcPay_EIG_CVV")
x_Type=rs("pcPay_EIG_Type")
x_TypeArray=Split(x_Type,"||")
x_TransType=x_TypeArray(0)
x_Curcode=rs("pcPay_EIG_Curcode")
x_TestMode=rs("pcPay_EIG_TestMode")
x_SaveCards=rs("pcPay_EIG_SaveCards")
x_UseVault=rs("pcPay_EIG_UseVault")
set rs=nothing

call closeDb()

dim successCnt, successData, failedCnt, failedData
successCnt=0
successData=""
failedCnt=0
failedData=""

Dim sIPAddress
sIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If sIPAddress="" Then sIPAddress = Request.ServerVariables("REMOTE_ADDR")

'how many checkboxes?
dim checkboxCnt
checkboxCnt=request.Form("EIGcheckboxCnt")

'do for each checkbox
dim r, orderVoid
For r=1 to checkboxCnt
	orderVoid=0
	IF request.Form("checkOrd"&r)="YES" THEN
		'if order totals don't match, void original order and flag for new order
		authamount=request.Form("authamount"&r)
		curamount=request.Form("curamount"&r)
		pidAuthOrder=request.Form("idauthorder"&r)
		pOrderStatus=request.Form("orderstatus"&r)
		pCheckEmail=request.Form("checkEmail"&r)
		x_invoice_num=Request.Form("idOrder"&r)
		x_VaultToken=Request.Form("vaultToken"&r)
		pcv_SecurityKeyID=Request.Form("SecurityKeyID"&r)

		dim pvc_reauth
		pvc_reauth=0

		if cdbl(authamount)<cdbl(curamount) then
			pvc_reauth=1
			%>
			<!--#include file="batch_authvoidcapture_EIG.asp"-->
		<% else

			'// CAPTURE
			strTest = ""
			strTest = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
			strTest = strTest & "<capture>"
			strTest = strTest & "<api-key>" & x_Key & "</api-key>"
			strTest = strTest & "<transaction-id>" & Request.Form("transid"&r) & "</transaction-id>"
			strTest = strTest & "<amount>" & Request.Form("curamount"&r) & "</amount>"
			strTest = strTest & "</capture>"

			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			xml.open "POST", "https://secure.nmi.com/api/v2/three-step", false
			xml.setRequestHeader "Content-Type", "text/xml"
			xml.send strTest
			strStatus = xml.Status
			strRetVal = xml.responseText
			Set xml = Nothing

			strResult = pcf_GetNode(strRetVal, "result", "*")
			strResultText = pcf_GetNode(strRetVal, "result-text", "*")
			strTransactionID = pcf_GetNode(strRetVal, "transaction-id", "*")
			strResultCode = pcf_GetNode(strRetVal, "result-code", "*")
			strAuthorizationCode = pcf_GetNode(strRetVal, "authorization-code", "*")
			pcv_strCustomerVaultID = pcf_GetNode(strRetVal, "customer-vault-id", "*")

			'response.Write(strRetVal & ".<br />")
			'response.Write(strResult & ".<br />")
			'response.Write(strResultText & ".<br />")
			'response.Write(strTransactionID & ".<br />")
			'response.Write(strResultCode & ".<br />")
			'response.Write(authorization-code & ".<br />")
			'response.Write(pcv_strCustomerVaultID & ".<br />")
			'response.End

		end if

		pIdOrder=(int(x_invoice_num)-scpre)
		qry_ID=pIdOrder
		idOrder=pIdOrder

		'// PROCESS RESULTS
		'// 1 = Transaction Approved
		'// 2 = Transaction Declined
		'// 3 = Error in transaction data or system error
		If strResult="1" Then

			call opendb()

			'// Update Authorder Status
			if pvc_reauth=1 then
				query="UPDATE pcPay_EIG_Authorize SET authcode='"& strAuthorizationCode &"', captured=1 WHERE idauthorder="& pidAuthOrder &";"
			else
				query="UPDATE pcPay_EIG_Authorize SET captured=1 WHERE idauthorder="& pidAuthOrder &";"
			end if
			set rstemp=connTemp.execute(query)
			set rstemp=nothing

			on error goto 0
			If len(x_VaultToken)=0 Then '// Purge CC Data

				query="SELECT ccnum, pcSecurityKeyID FROM pcPay_EIG_Authorize WHERE idauthorder="& pidAuthOrder &";"
				set rstemp=connTemp.execute(query)
				if NOT rstemp.EOF then
					cardnumber=rstemp("ccnum")
					tempSecurityKeyID=rstemp("pcSecurityKeyID")
				end if
				set rstemp=nothing

				tempfour=pcf_PurgeCardNumber(cardnumber,tempSecurityKeyID)

				query="UPDATE pcPay_EIG_Authorize SET ccnum='"&tempfour&"' WHERE idauthorder="&pidAuthOrder&";"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing

			Else '// Purge Vault Data

				pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
				x_VaultTokenTmp=enDeCrypt(x_VaultToken, pcv_SecurityPass)

				query="SELECT IsSaved FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_Token='"& x_VaultTokenTmp & "'"
				set rs=Server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if NOT rs.eof then
					pcv_intIsSaved = rs("IsSaved")
					pcv_strCustomerVaultID = x_VaultToken '// vault token exists
				else
					pcv_intIsSaved = 0
					pcv_strCustomerVaultID = "" '// vault token does not exist
				end if
				set rs=nothing



				If pcv_intIsSaved=0 Then

					'// Contact Vault
					strTest = ""
					strTest = strTest & "username=" & x_Username
					strTest = strTest & "&password=" & x_Password
					strTest = strTest & "&customer_vault=delete_customer"
					strTest = strTest & "&customer_vault_id=" & pcv_strCustomerVaultID

					set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
					xml.open "POST", "https://secure.nmi.com/api/transact.php", false
					xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					xml.send strTest
					strStatus = xml.Status
					strRetVal = xml.responseText
					Set xml = Nothing

					If len(pcv_strCustomerVaultID)>0 Then
						query="DELETE FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_Token='"& x_VaultTokenTmp & "'"
						set rs=Server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						set rs=nothing
					End If

				End If

			End If

			if pvc_reauth=1 then
				query="UPDATE orders SET gwAuthCode='"&strAuthorizationCode&"', gwTransID='"&strTransactionID&"' WHERE idOrder="&pIdOrder&";"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing
			end if

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
				query="SELECT orders.idcustomer,orders.address,orders.City,orders.StateCode,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingZip,orders.shippingCountryCode,orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.iRewardValue,orders.iRewardRefId, orders.iRewardPointsRef, orders.iRewardPointsCustAccrued, customers.phone, ord_DeliveryDate, ord_VAT, pcOrd_CatDiscounts FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" &qry_ID
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
			successData=successData&"Order Number "& x_invoice_num &" was updated successfully<BR>"

		else
			failedCnt=failedCnt+1
			failedData=failedData&"Order Number: "& x_invoice_num &" was NOT updated, " & strResultText & "<BR>"
		end if 'Order was not approved
	END IF
Next
call closedb()
%>
<table class="pcCPcontent">
  <tr>
	<td><div class="pcCPmessageSuccess"><%=successCnt%> records were successfully updated.</div>
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
<%
Function pcf_GetNode(responseXML, nodeName, nodeParent)
	Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)
	myXmlDoc.loadXml(responseXML)
	Set Nodes = myXmlDoc.selectnodes(nodeParent)
	For Each Node In Nodes
		pcf_GetNode = pcf_CheckNode(Node,nodeName,"")
	Next
	Set Node = Nothing
	Set Nodes = Nothing
	Set myXmlDoc = Nothing
End Function

Function pcf_CheckNode(Node,tagName,default)
	Dim tmpNode
	Set tmpNode=Node.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		pcf_CheckNode=default
	Else
		pcf_CheckNode=Node.selectSingleNode(tagName).text
	End if
End Function

Function pcf_FixXML(str)
	str=replace(str, "&","and")
	pcf_FixXML=str
End Function

Public Function deformatNVP(nvpstr)
	Dim AndSplitedArray, EqualtoSplitedArray, Index1, Index2, NextIndex
	Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
	AndSplitedArray = Split(nvpstr, "&", -1, 1)
	NextIndex=0
	For Index1 = 0 To UBound(AndSplitedArray)
		EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
		For Index2 = 0 To UBound(EqualtoSplitedArray)
			NextIndex=Index2+1
			NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
			Index2=Index2+1
		Next
	Next
	Set deformatNVP = NvpCollection
End Function
%>