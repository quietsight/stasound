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
<!--#include file="adminHeader.asp" -->
<% on error resume next
dim query, conntemp, rs, rstemp, rstemp2, rsEmailInfo, rsCust, pTempProductId, pTodaysDate
dim varReply, nStatus, strErrorInfo	

Dim outXml, resp, host
		        
Dim res1, resDesc
' create transaction object	
Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")
' Set this one to false if you don't want any logging
Const fLog = False

' set logging level  
' 0 - no logging
' 1 - errors
' 2 - debug ( errors + addl info )
' 3 - trace ( full debug mode with call tracing )
Const logLvl = 0

' set log file name
' IMPORTANT: this file must have write access rights 
'            for IIS' default IUSR_XXXXXX user account.
'	     Otherwise no logging will take place
logFile = "LINKLOG.Log" ' Change this if you want logging

call opendb()

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT storeName, transType, lp_testmode, lp_cards, CVM, lp_yourpay FROM LinkPoint where id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
lp_storeName=rs("storeName") 
lp_transType=rs("transType")
lp_testmode=rs("lp_testmode")
lp_cards=rs("lp_cards")
lp_CVM=rs("CVM")
lp_yourpay=rs("lp_yourpay")
if lp_CVM<>1 then
	lp_CVM=0
end if

configfile = lp_storeName ' Change this to your store number 
if PPD="1" then
	filename="/"&scPcFolder&"/" & scAdminFolderName
else
	filename="../"&scAdminFolderName
end if
filename = Server.MapPath (filename)
keyfile    = filename &"\" &lp_storeName&".pem" ' Change this to the name and location of your certificate file 

if lp_testmode ="YES" Then
	host = "staging.linkpt.net"
else
	host = "secure.linkpt.net"
End if 
Const port = "1129"

set rs=nothing
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

'how many checkboxes?
dim checkboxCnt
checkboxCnt=request.Form("checkboxCnt")
'do for each checkbox
dim r
for r=1 to checkboxCnt
	IF request.Form("checkOrd"&r)="YES" THEN
		'if order totals don't match, void original order and flag for new order
		lpamount=request.Form("lpamount"&r)
		curamount=request.Form("curamount"&r)
		lpidorder=request.Form("lpidorder"&r)
		pOrderStatus=request.Form("orderstatus"&r)
		pCheckEmail=request.Form("checkEmail"&r)
		pIdOrder=Request.Form("idOrder"&r)
		cardnumber =request.Form("ccnum"&r)
		expmonth =request.Form("ccexpmonth"&r)
		expyear = right(request.Form("ccexpyear"&r),2)
		lpdate = request.Form("lpdate"&r)
		qry_ID=pIdOrder
		idOrder=pIdOrder
				
		if cdbl(lpamount)<>cdbl(curamount) then
			pvc_reauth=1
			%>
			<!--#include file="batch_lpvoidcapture.asp"-->
		<% else
			'*************************************************************************************
			' This is where you would post info to the gateway
			' START
			'*************************************************************************************
			' Create an empty order
			Set order = Server.CreateObject("LpiCom_6_0.LPOrderPart")
			order.setPartName("order")
			' Create an empty part
			Set op = Server.CreateObject("LpiCom_6_0.LPOrderPart")                

			' Build 'orderoptions'
			' For a test, set result to GOOD, DECLINE, or DUPLICATE
			if lp_testmode ="YES" Then
				res=op.put("result", "GOOD")			
			else
				res=op.put("result", "LIVE")			
			End if 
		
			res=op.put("ordertype","POSTAUTH")
			' add 'orderoptions to order
			res=order.addPart("orderoptions", op)
			
			res=op.clear()
			
			res=op.put("oid",pIdOrder)
			
			' add 'merchantinfo to order
			res=order.addPart("transactiondetails", op)

			' Build 'merchantinfo'
			res=op.clear()
			
			res=op.put("configfile",configfile)
			' add 'merchantinfo to order
			res=order.addPart("merchantinfo", op)
			
			' Build 'creditcard'
			res=op.clear()
			res=op.put("cardnumber", cardnumber)
			res=op.put("cardexpmonth", expmonth)
			res=op.put("cardexpyear", expyear)
			
			' add 'creditcard to order
			res=order.addPart("creditcard", op)
			
			' Build 'payment'
			res=op.clear()
			res=op.put("chargetotal", money(curamount))
			' add 'payment to order
			res=order.addPart("payment", op)
        
			if (fLog = True) and ( logLvl > 0 ) Then
			
				resDesc = "ORDID: " & pIdOrder
				res1 = res
				if PPD="1" then
					filename2="/"&scPcFolder&"/includes"
				else
					filename2="../includes"
				end if
				logFile = Server.MapPath(filename2) &"\" & logFile
				  
				'Next call return level of accepted logging in 'res1'
				'On error 'res1' contains negative number
				'You can check 'resDesc' to get error description
				'if any
				
				res = LPTxn.setDbgOpts(logFile,logLvl,resDesc,res1)
			
			End If
        
			' get outgoing XML from 'order' object
			
			
			outXml = order.toXML()
			'response.write keyfile
			'Response.end
			' Call LPTxn
			rsp = LPTxn.send(keyfile, host, port, outXml)
			
			'response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
			
			'Store transaction data on Session and redirect
      
			Set LPTxn = Nothing
			Set order = Nothing
			Set op    = Nothing
			
			
			
			R_Time = ParseTag("r_time", rsp)
			R_Ref = ParseTag("r_ref", rsp)		
			R_Approved = ParseTag("r_approved", rsp)
			R_Code = ParseTag("r_code", rsp)
			R_OrderNum = ParseTag("r_ordernum", rsp)
			R_Message = ParseTag("r_message", rsp)		
			R_Error = ParseTag("r_error", rsp)		
			R_TDate = ParseTag("r_tdate", rsp)
			Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")
		End if 
	 
		idOrder=(int(idOrder)+scPre)

		If R_Approved = "APPROVED" Then
			
			'order was successful
			call opendb()
			
			'update authorders to captured
			query="UPDATE pcPay_LinkPointAPI SET pcPay_LPAPI_Captured=1 WHERE pcPay_LPAPI_ID="&lpidorder&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			'purge cc data
			query="SELECT pcPay_LPAPI_CCNum, pcSecurityKeyID FROM pcPay_LinkPointAPI WHERE pcPay_LPAPI_ID="&lpidorder&";"
			set rstemp=connTemp.execute(query)
			if NOT rstemp.EOF then
				cardnumber=rstemp("pcPay_LPAPI_CCNum")
				tempSecurityKeyID=rstemp("pcSecurityKeyID")
			end if
			set rstemp=nothing
			
			tempfour=pcf_PurgeCardNumber(cardnumber,tempSecurityKeyID)
			
			query="UPDATE pcPay_LinkPointAPI SET pcPay_LPAPI_CCNum='"&tempfour&"' WHERE pcPay_LPAPI_ID="&lpidorder&";"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
			
			'if order has not alread been processed
			IF pOrderStatus="2" THEN
				'------------------------------------------------
				'- Look for downloadable products
				'------------------------------------------------
				query="select idproduct from ProductsOrdered WHERE idOrder="&pIdOrder&";"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
				DPOrder="0"
				do while not rs.eof
					pTempProductId=rs("idproduct")
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
					pTempProductId=rstemp("idproduct")
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
					query="select idproduct,quantity from ProductsOrdered WHERE idOrder="&pIdOrder&";"
					set rstemp=Server.CreateObject("ADODB.Recordset")
					set rstemp=connTemp.execute(query)
					do while not rstemp.eof
						tIdProduct=rstemp("idproduct")
						pQuantity=rstemp("quantity")
						query="select products.idproduct,sku,License,LocalLG,RemoteLG from Products,DProducts where products.idproduct=" & tIdProduct & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
						Set rs=Server.CreateObject("ADODB.Recordset")
						set rs=connTemp.execute(query)
						if not rs.eof then
							pIdproduct=rs("idproduct")
							pSku=rs("sku")
							pLicense=rs("License")
							pLocalLG=rs("LocalLG")
							pRemoteLG=rs("RemoteLG")
							
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
								
								Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
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
									query="Insert into DPLicenses (IdOrder,IdProduct,Lic1,Lic2,Lic3,Lic4,Lic5) values (" & rIdOrder & "," & rIdProduct & ",'" & PLic1 & "','" & PLic2 & "','" & PLic3 & "','" & PLic4 & "','" & PLic5 & "')"
									set rstemp2=Server.CreateObject("ADODB.Recordset")
									set rstemp2=connTemp.execute(query)
									set rstemp2=nothing
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
							set rstemp2=Server.CreateObject("ADODB.Recordset")
							set rstemp2=connTemp.execute(query)
							
							if not rstemp2.eof then
								ReqExist=1
							end if
							
							LOOP UNTIL ReqExist=0
							set rstemp2=nothing
							'Insert Standard & BTO Products Download Requests into DPRequests Table
							
							if scDB="Access" then
								query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "',#" & pTodaysDate & "#)"   
							else
								query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "','" & pTodaysDate & "')"
							end if
							set rstemp2=Server.CreateObject("ADODB.Recordset")
							set rstemp2=connTemp.execute(query)
							set rstemp2=nothing
						end if
						set rs=nothing
						rstemp.moveNext
					loop
					set rstemp=nothing
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
				query="SELECT orders.idcustomer,orders.address,orders.City,orders.StateCode,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingZip,orders.shippingCountryCode,orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.iRewardValue,orders.iRewardRefId, orders.iRewardPointsRef,orders.iRewardPointsCustAccrued,customers.phone,ord_DeliveryDate,ord_VAT FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" &qry_ID
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
				set rsEmailInfo=nothing
				pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)
				
				'Get customer details for this order
				query="Select name,lastname,customerCompany,email FROM customers WHERE idcustomer="& pIdCustomer
				Set rsCust=Server.CreateObject("ADODB.Recordset")
				Set rsCust=conntemp.execute(query)
				pName=rsCust("name")
				pLName=rsCust("lastname")
				pCustomerCompany=rsCust("customerCompany")
				pEmail=rsCust("email")
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
		else
		    failedCnt=failedCnt+1
			failedData=failedData&"Order Number: "&(int(pIdOrder)+scpre)&" was NOT processed, " & pfp_respmsg & "<BR>" 
		
		end if 'Order was not approved
	 END IF '
   
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
<% ' Functions 

     Function LeadingZero(ByVal InpStr) 
        If Len(InpStr) = 1 Then
            LeadingZero =  ("0" & InpStr)
        Else
            LeadingZero = InpStr
        End If
    End Function

 

    

     Function ParseTag( tag ,  rsp ) 
        Dim sb 
        Dim idxSt, idxEnd 'As Integer
        
        rsp = rsp
        
        sb = "<" & tag & ">"
        idxSt = -1
        idxEnd = -1

        idxSt = InStr(rsp,sb)
        If 0 = idxSt Then
            ParseTag = ""
            Exit Function
        End If

        idxSt = idxSt + Len(sb)
        sb = "</" & tag & ">"
        idxEnd = InStr(idxSt, rsp,sb)

        If 0 = idxEnd Then
           ParseTag = ""
           Exit Function
        End If

        ParseTag = Mid(rsp, idxSt, (idxEnd - idxSt))

    End Function

     Sub Cleanup()

        If Not (LPTxn Is Nothing) Then
            Set LPTxn = Nothing
        End If
        If Not (order Is Nothing) Then
            res = order.removeAllParts()
            order = Nothing
        End If
        If Not (op Is Nothing) Then
            res = op.removeAllParts()
            op = Nothing
        End If

    End Sub
%><!--#include file="adminFooter.asp" -->