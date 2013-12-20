<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Server.ScriptTimeout = 3600 %>
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
dim query, conntemp, rs, rstemp, rstemp2, rsEmailInfo, rsCust, pTempProductId, pTodaysDate


call opendb()

	query="SELECT v_Partner, v_Vendor, v_User, v_Password,pfl_testmode,pfl_transtype,pfl_CSC FROM verisign_pfp WHERE id=1;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if err.number <> 0 then
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
	end If 
	
	pPARTNER=rs("v_Partner")
	pVENDOR=rs("v_Vendor")
	pUSER=rs("v_User")
	pPWD=rs("v_Password")
	pfl_testmode=rs("pfl_testmode")
	pfl_transtype=rs("pfl_transtype")
	pfl_CSC=rs("pfl_CSC")		
	set rs=nothing

if pfl_testmode="YES" then
	GatewayHost = "https://pilot-payflowpro.paypal.com"
else
	GatewayHost = "https://payflowpro.paypal.com"
end if


call closedb()

failedCnt=0
failedData=""
successCnt=0
successData=""

pTodaysDate=Date()
if SQL_Format="1" then
	pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
else
	pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
end if

Dim pfp_client, parmList, pfp_ctx, pfp_string,objTest
			
'how many checkboxes?
dim checkboxCnt
checkboxCnt=request.Form("checkboxCnt")
'do for each checkbox
dim r

for r=1 to checkboxCnt
	IF request.Form("checkOrd"&r)="YES" THEN
		Dim requestID
		
		' Need to generate a unique id for the request id
		requestID = generateRequestID()

		'if order totals don't match, void original order and flag for new order
		pfpamount=request.Form("pfpamount"&r)
		curamount=request.Form("curamount"&r)
			pfpidorder=request.Form("pfpidorder"&r)
			pOrderStatus=request.Form("orderstatus"&r)
			pCheckEmail=request.Form("checkEmail"&r)
			pIdOrder=Request.Form("idOrder"&r)
			qry_ID=pIdOrder
			idOrder=pIdOrder

		parmList="TRXTYPE=D&TENDER=C"
		parmList=parmList & "&USER=" & pUSER
		parmList=parmList & "&PWD=" & pPWD
		parmList=parmList & "&PARTNER=" & pPARTNER
		parmList=parmList & "&VENDOR=" & pVENDOR
		parmList=parmList & "&COMMENT1=" & idOrder
		parmList=parmList & "&ORIGID=" & request.Form("ORIGID"&r)
		parmList=parmList & "&AMT=" & request.Form("curamount"&r)
		parmList=parmList & "&STREET=" & request.Form("ADDRESS"&r)
		parmList=parmList & "&ZIP=" & request.Form("ZIP"&r)
		'Open Session
		Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objWinHttp.Open "POST", GatewayHost, False
				
		objWinHttp.setRequestHeader "Content-Type", "text/namevalue" ' for XML, use text/xml
		objWinHttp.SetRequestHeader "X-VPS-Timeout", "30"
		objWinHttp.SetRequestHeader "X-VPS-Request-ID", requestID
		'Send Parameter List
		objWinHttp.Send parmList
				
		' Print out the request status:
		'Response.Write "Status: " & objWinHttp.Status & " " & objWinHttp.StatusText & "<br />"
	
		' Get the text of the response.
		transResponse = objWinHttp.ResponseText
				
		' Trash our object now that we are finished with it.
		Set objWinHttp = Nothing

		pfp_pnref = ShowResponse(transResponse, "PNREF")
			session("pnref")=pfp_pnref
		pfp_result = ShowResponse(transResponse, "RESULT")
		pfp_respmsg = ShowResponse(transResponse, "RESPMSG")
		pfp_authcode = ShowResponse(transResponse, "AUTHCODE")
			session("authcode")=pfp_authcode
		pfp_avsaddr=ShowResponse(transResponse, "AVSADDR")
		pfp_avszip=ShowResponse(transResponse, "AVSZIP")
		pfp_comment1=ShowResponse(transResponse, "COMMENT1")
		pfp_comment2=ShowResponse(transResponse, "COMMENT2")
		pfp_amt=ShowResponse(transResponse, "AMT")

		idOrder=(int(idOrder)+scPre)

		'//PAYPAL LOGGING START
		scPPLogging="0"
		
		If scPPLogging = "1" Then
			if PPD="1" then
				pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/PFPLOG.txt")
			else
				pcStrFileName=Server.Mappath ("../includes/PFPLOG.txt")
			end if

			dim strFileName
			dim fs
			dim OutputFile
		
			'Specify directory and file to store silent post information
			strFileName = pcStrFileName
			Set fs = CreateObject("Scripting.FileSystemObject")
			Set OutputFile = fs.OpenTextFile (strFileName, 8, True)
			OutputFile.WriteLine now()
			OutputFile.WriteLine "=============HEADER======================="
			OutputFile.WriteLine """Content-Type"", ""text/namevalue"""
			OutputFile.WriteLine """X-VPS-Timeout"", ""30"""
			OutputFile.WriteLine """X-VPS-Request-ID"", "&requestID
			OutputFile.WriteLine "=============NVP=================="
			OutputFile.WriteLine "Request from ProductCart: " & parmList
			OutputFile.WriteBlankLines(1)
			OutputFile.WriteLine "Response from PayFlow Pro: " & transResponse
			OutputFile.WriteBlankLines(1)
			OutputFile.WriteLine "PNREF: " & pfp_pnref
			OutputFile.WriteLine "RESULT: " & pfp_result
			OutputFile.WriteLine "RESPMSG: " & pfp_respmsg
			OutputFile.WriteLine "AUTHCODE: " & pfp_authcode
			OutputFile.WriteBlankLines(2)
			OutputFile.Close
		End If
		'//PAYPAL LOGGING END
				
		'Check Payflow Pro response code 1=approved 2=declined 3=error. 
		'pfp_result=0 'unREM for testing
		If pfp_result <> "0" Then
			failedCnt=failedCnt+1
			failedData=failedData&"Order Number: "&idOrder&" was NOT processed, " & pfp_respmsg & "<BR>" 
		ElseIf pfp_result=0 Then
			'order was successful
			call opendb()
			
			'update authorders to captured
			query="UPDATE pfporders SET captured=1 WHERE idpfporder="&pfpidorder&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			'purge cc data
			query="SELECT acct, pcSecurityKeyID FROM pfporders WHERE idpfporder="&pfpidorder&";"
			set rstemp=connTemp.execute(query)
			if NOT rstemp.EOF then
				cardnumber=rstemp("acct")
				tempSecurityKeyID=rstemp("pcSecurityKeyID")
			end if
			set rstemp=nothing
			
			tempfour=pcf_PurgeCardNumber(cardnumber,tempSecurityKeyID)

			query="UPDATE pfporders SET acct='"&tempfour&"' WHERE idpfporder="&pfpidorder&";"
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
			successData=successData&"Order Number "& (int(pIdOrder)+scpre) &" was successfully updated<BR>"
			
		
		end if 'Order was not approved
	END IF 
next
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
<% ' VeriSign PayFlow Pro :: Code Block 2 :: Sale & Authorization Transactions :: Database & Redirects Code
FUNCTION pfp_getvalue(nameStr, resultStr)
  startpos=1
  doloop=1
  Do
    cpos=InStr(startpos, resultStr, "&")
    if cpos < 1 then 
        doloop=0
        cpos=Len(resultStr)+1
    end if
    thisPair=(Mid(resultStr, startpos, cpos - startpos ) )
	eqPos=InStr(thisPair, "=")
    name=Left(thisPair, eqPos-1)
	value=Right(thisPair, Len(thisPair) - eqPos )
	if name=nameStr then 
	  pfp_getvalue=value	  
	end if
	startpos=cpos + 1
  Loop while doloop=1
END FUNCTION

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

Function ShowResponse(transResponse, Param)

     curString = transResponse

     Do while Len(curString) <> 0

          if InStr(curString,"&") Then
               varString = Left(curString, InStr(curString , "&" ) -1)
          else
               varString = curString
          end if

          name = Left(varString, InStr(varString, "=" ) -1)
          value = Right(varString, Len(varString) - (Len(name)+1))

          if name = Param Then
               MyValue = value
               Exit Do
          end if

          if Len(curString) <> Len(varString) Then
               curString = Right(curString, Len(curString) - (Len(varString)+1))
          else
               curString = ""
          end if

     Loop
     ShowResponse = MyValue

End Function

Function generateRequestID()
	pcKeys_Key_ID= Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & session("idcustomer")
    generateRequestID = pcKeys_Key_ID
End Function


%>
<!--#include file="adminFooter.asp" -->