<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
Server.ScriptTimeout = 5400
Response.Buffer = False
%>
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

query="SELECT x_Type,x_Login,x_Password,x_Curcode,x_Method,x_AIMType,x_CVV,x_testmode FROM authorizeNet Where id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
	dim x_Login, x_Password, x_CVV, x_Curcode, x_Method, x_AIMType, x_testmode
	x_Login=rs("x_Login")
	'decrypt
	x_Login=enDeCrypt(x_Login, scCrypPass)
	
	x_Password=rs("x_Password")
	'decrypt
	x_Password=enDeCrypt(x_Password, scCrypPass)
	x_CVV=rs("x_CVV")
	x_Curcode=rs("x_Curcode")
	x_Method=rs("x_Method")
	x_AIMType=rs("x_AIMType")
	x_testmode=rs("x_testmode")

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
checkboxCnt=request.Form("checkboxCnt")

'do for each checkbox
dim r, orderVoid
For r=1 to checkboxCnt
	orderVoid=0
	IF request.Form("checkOrd"&r)="YES" THEN
		'if order totals don't match, void original order and flag for new order
		pidAuthOrder=request.Form("idauthorder"&r)
		curamount=request.Form("curamount"&r)
		pCheckEmail=request.Form("checkEmail"&r)
		
		call opendb()
		'===========================================
		query="SELECT authorders.idOrder, authorders.idauthorder, authorders.amount, authorders.paymentmethod, authorders.transtype, authorders.authcode, authorders.ccnum,  authorders.ccexp, authorders.idCustomer, authorders.fname, authorders.lname, authorders.address, authorders.zip, authorders.pcSecurityKeyID, orders.orderDate, orders.orderStatus, orders.gwTransID, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.ShippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.comments, orders.adminComments, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email FROM         customers INNER JOIN authorders INNER JOIN orders ON authorders.idOrder = orders.idOrder ON authorders.idCustomer = customers.idcustomer AND customers.idcustomer = orders.idCustomer WHERE (authorders.idauthorder = "&pidAuthOrder&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)

		amount=rs("amount")
		paymentmethod=rs("paymentmethod")
		transtype=rs("transtype")
		authcode=rs("authcode")
		ccnum=rs("ccnum")
		ccexp=rs("ccexp")
		idCustomer=rs("idCustomer")
		fname=rs("fname")
		lname=rs("lname")
		address=rs("address")
		zip=rs("zip")
		pcv_SecurityKeyID=rs("pcSecurityKeyID")
		orderDate=rs("orderDate")
		orderStatus=rs("orderstatus")
		gwTransId=rs("gwTransId")
		stateCode=rs("stateCode")
		if stateCode="" then
			stateCode=rs("State")
		end if
		City=rs("city")
		countryCode=rs("countryCode")
		shippingAddress=rs("shippingAddress")
		shippingStateCode=rs("shippingStateCode")
		shippingState=rs("shippingState")
		shippingCity=rs("shippingCity")
		shippingCountryCode=rs("shippingCountryCode")
		shippingZip=rs("shippingZip")
		shippingFullName=rs("shippingFullName")
		address2=rs("address2")
		shippingCompany=rs("shippingCompany")
		shippingAddress2=rs("shippingAddress2")
		pcv_custcomments=trim(rs("comments"))
		pcv_admcomments=trim(rs("admincomments"))
		customerName=rs("name") & " " & rs("lastName")
		customerCompany=rs("customerCompany")
			if trim(customerCompany)<>"" then
				customerInfo=customerName & " (" & customerCompany & ")"
				else
				customerInfo=customerName
			end if
		phone=rs("phone")
		email =rs("email")
		
		pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
		ccnum2=enDeCrypt(ccnum, pcv_SecurityPass)
		
		call closedb()
		'===========================================
		authamount=amount
		pOrderStatus=orderstatus
		
		dim pvc_reauth
		pvc_reauth=0
		
		if cdbl(authamount)<>cdbl(curamount) then
			pvc_reauth=1
			%>
			<!--#include file="batch_authvoidcapture.asp"-->
		<% else
		'send pre-auth-capture
			data = "x_login=" & x_Login & ""   
			data = data & "&x_tran_key=" & x_Password & ""
			data = data & "&x_customer_ip=" & sIPAddress & ""
			data = data & "&x_first_name=" & fName  & ""
			data = data & "&x_last_name=" & lname  & ""
			data = data & "&x_company=" & customerCompany  & ""
			data = data & "&x_address=" & address  & ""
			data = data & "&x_city=" & city  & ""
			data = data & "&x_state=" & stateCode  & ""
			data = data & "&x_zip=" & zip & ""
			data = data & "&x_country=" & countryCode  & ""
			data = data & "&x_phone=" & phone  & ""
			data = data & "&x_email=" & email  & ""
			data = data & "&x_card_num=" & ccnum2  & ""
			data = data & "&x_exp_date=" & ccexp & ""
			data = data & "&x_amount=" & curamount  & ""
			data = data & "&x_cust_id="
			data = data & "&x_invoice_num=" & Request.Form("idOrder"&r)  & ""
			data = data & "&x_description="
			data = data & "&x_trans_id=" & gwTransId  & ""
			data = data & "&x_delim_data=True"
			data = data & "&x_delim_char=,"
			data = data & "&x_version=3.1"
			data = data & "&x_method=CC"
			data = data & "&x_type=PRIOR_AUTH_CAPTURE"
			data = data & "&x_relay_response=FALSE"
			if x_testmode="1" then
				data = data & "&x_test_request=True"
			end if  
			data = data & "&x_email_customer=False"

			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			xml.open "POST", "https://secure.authorize.net/gateway/transact.dll?" & data & "", false
			xml.send ""
			authnetStatus = xml.Status
			authnetVal = xml.responseText
			set xml = nothing
		end if
		
		authnetResults = split(authnetVal, ",", -1)
		x_response_code           = authnetResults(0)
		x_response_subcode        = authnetResults(1)
		x_response_reason_code    = authnetResults(2)
		x_response_reason_text    = authnetResults(3)
		x_auth_code               = authnetResults(4)    
		x_avs_code                = authnetResults(5)
		x_trans_id                = authnetResults(6)   
		x_invoice_num             = authnetResults(7)
		x_description             = authnetResults(8)
		x_amount                  = authnetResults(9)
		x_method                  = authnetResults(10)
		x_type                    = authnetResults(11)
		x_cust_id                 = authnetResults(12)
		x_first_name              = authnetResults(13)
		x_last_name               = authnetResults(14)
		x_company                 = authnetResults(15)
		x_address                 = authnetResults(16)
		x_city                    = authnetResults(17)
		x_state                   = authnetResults(18)
		x_zip                     = authnetResults(19)
		x_country                 = authnetResults(20)
		x_phone                   = authnetResults(21)
		x_fax                     = authnetResults(22)
		x_email                   = authnetResults(23)
		x_ship_to_first_name      = authnetResults(24)
		x_ship_to_last_name       = authnetResults(25)
		x_ship_to_company         = authnetResults(26)
		x_ship_to_address         = authnetResults(27)
		x_ship_to_city            = authnetResults(28)
		x_ship_to_state           = authnetResults(29)
		x_ship_to_zip             = authnetResults(30)
		x_ship_to_country         = authnetResults(31)
		x_tax                     = authnetResults(32)
		x_duty                    = authnetResults(33)
		x_freight                 = authnetResults(34)
		x_tax_exempt              = authnetResults(35)
		x_po_num                  = authnetResults(36)
		x_md5_hash                = authnetResults(37)
		x_card_code               = authnetResults(38)

		pIdOrder=(int(x_invoice_num)-scpre)
		qry_ID=pIdOrder
		idOrder=pIdOrder

		
		'Check authorize.net response code 1=approved 2=declined 3=error.
		if x_response_code = 1 then
			'order was successful
			call opendb()
			
			'update authorders to captured
			if pvc_reauth=1 then
				query="UPDATE authorders SET authcode='"&x_auth_code&"', captured=1 WHERE idauthorder="&pidAuthOrder&";"
			else
				query="UPDATE authorders SET captured=1 WHERE idauthorder="&pidAuthOrder&";"
			end if
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
			
			'purge cc data
			query="SELECT ccnum, pcSecurityKeyID FROM authorders WHERE idauthorder="&pidAuthOrder&";"
			set rstemp=connTemp.execute(query)
			if NOT rstemp.EOF then
				cardnumber=rstemp("ccnum")
				tempSecurityKeyID=rstemp("pcSecurityKeyID")
			end if
			set rstemp=nothing
			
			tempfour=pcf_PurgeCardNumber(cardnumber,tempSecurityKeyID)
			
			query="UPDATE authorders SET ccnum='"&tempfour&"' WHERE idauthorder="&pidAuthOrder&";"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
			
			if pvc_reauth=1 then
				query="UPDATE orders SET gwAuthCode='"&x_auth_code&"',gwTransID='"&x_trans_id&"' WHERE idOrder="&pIdOrder&";"
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
			failedData=failedData&"Order Number: "& x_invoice_num &" was NOT updated, " & x_response_reason_text & "<BR>"
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