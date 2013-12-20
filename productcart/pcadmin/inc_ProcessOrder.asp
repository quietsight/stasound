<%	
	'Start SDBA
	if (pcv_SubmitType=0) or (pcv_SubmitType=1) or (pcv_SubmitType=3) then
		query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=1;"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments) VALUES (" & qry_ID & ",1,'" & pcv_AdmComments & "');"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
	else
		query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID& " AND pcACom_ComType=0;"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments) VALUES (" & qry_ID & ",0,'" & pcv_AdmComments & "');"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
	end if
	if pcv_AdmComments<>"" then
		pcv_AdmComments=replace(pcv_AdmComments,"''","'")
	end if
	'End SDBA
	

	query="select idproduct,idconfigSession from ProductsOrdered WHERE idOrder="& qry_ID
	set rs=connTemp.execute(query)
	DPOrder="0"
	pGCs="0"
	do while not rs.eof
		pIdProduct=rs("idproduct")
		tmpidConfig=rs("idconfigSession")
		IF DPOrder="0" THEN
		query="select downloadable from products where idproduct=" & pIdProduct
		set rstemp=connTemp.execute(query)
		if not rstemp.eof then
			pdownloadable=rstemp("downloadable")
			if (pdownloadable<>"") and (pdownloadable="1") then
				DPOrder="1"
			end if
		end if
		set rstemp=nothing
		END IF
		'Find downloadable items in BTO configuration
		if tmpidConfig<>"" AND tmpidConfig>"0" AND DPOrder="0" then
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
		'GGG Add-on start
		if ((request.querystring("Submit4")<>"") or (pcv_SubmitType=3)) AND (pGCs="0") then
			query="SELECT products.pcprod_GC FROM products WHERE idproduct=" & pIdProduct
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			if not rstemp.eof then
				pGC=rstemp("pcprod_GC")
				If isNULL(pGC) Then pGC="0"
				if pGC="1" then
					pGCs="1"
				end if
			end if
			set rstemp=nothing
		end if
		'GGG Add-on end
	rs.moveNext
	loop
	set rs=nothing

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: Submit 4
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if (request.querystring("Submit4")<>"") or (pcv_SubmitType=3) then

	'// Update orderstatus to 3(processed) and input today's date
	Dim pTodaysDate
	pTodaysDate=Date()
	if SQL_Format="1" then
		pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
	else
		pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
	end if
	
	if scDB="Access" then
		query="UPDATE orders SET pcOrd_GCs=" & pGCs & ",DPs=" & DPOrder & ", orderstatus=3, processDate=#"& pTodaysDate &"# WHERE idOrder="& qry_ID
	else
		query="UPDATE orders SET pcOrd_GCs=" & pGCs & ",DPs=" & DPOrder & ", orderstatus=3, processDate='"& pTodaysDate &"' WHERE idOrder="& qry_ID
	end if	
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	Set rs=nothing
	
	query="select idcustomer,orderdate,processdate from Orders WHERE idOrder="& qry_ID
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	if not rs.eof then
		pIdCustomer=rs("IdCustomer")
		pOrderDate=rs("OrderDate")
		pProcessDate=rs("ProcessDate")
	end if	
	Set rs=nothing
	
	pIdOrder=qry_ID
	
	'// Call License Generator for Standard & BTO Products
	IF DPOrder="1" then
		query="select idproduct,quantity,idconfigSession from ProductsOrdered WHERE idOrder="& qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
	
		if not rs.eof then
			tmpArr=rs.getRows()
			intCount=ubound(tmpArr,2)
			set rs=nothing
			For ik=0 to intCount
			pIdProduct=tmpArr(0,ik)
			pQuantity=tmpArr(1,ik)
			tmpidConfig=tmpArr(2,ik)
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
			Next
		end if
		set rs=nothing
	END IF
	
	'GGG Add-on start	
	IF pGCs="1" then
		
		query="SELECT idproduct,quantity FROM ProductsOrdered WHERE idOrder="& qry_ID
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=connTemp.execute(query)
		do while NOT rstemp.eof
			pIdproduct=rstemp("idproduct")
			pQuantity=rstemp("quantity")
				
			query="SELECT pcGC.pcGC_Exp,pcGC.pcGC_ExpDate,pcGC.pcGC_ExpDays,pcGC.pcGC_CodeGen,pcGC.pcGC_GenFile,products.sku,products.price FROM pcGC,Products WHERE pcGC.pcGC_idproduct=" & pIdproduct & " AND Products.idproduct=pcGC.pcGC_idproduct AND products.pcprod_GC=1"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)	
			pcv_strRunSection = 0
			if not rs.eof then
				pGCExp=rs("pcGC_Exp")
				pGCExpDate=rs("pcGC_ExpDate")
				pGCExpDay=rs("pcGC_ExpDays")
				pGCGen=rs("pcGC_CodeGen")
				pGCGenFile=rs("pcGC_GenFile")
				pSku=rs("sku")
				pGCAmount=rs("price")
				pcv_strRunSection = -1
			end if
			set rs = nothing
		
			If pcv_strRunSection = -1 Then
			
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
					
						'///////////////////////////////////////////////////
						'// START: DO
						'///////////////////////////////////////////////////
						DO
						
							Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
							srvXmlHttp.open "POST", L_Action, False
							srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
							srvXmlHttp.send L_postdata
							result1 = srvXmlHttp.responseText
							
							RArray = split(result1,"<br>")
							GiftCode= RArray(2)
							
							'// If have errors from GiftCode Generator
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
						
							query="SELECT pcGO_IDProduct FROM pcGCOrdered WHERE pcGO_GcCode='" & GiftCode & "'" 
							set rsG=Server.CreateObject("ADODB.Recordset")
							set rsG=connTemp.execute(query)					
							if not rsG.eof then
								ReqExist=1
							end if
							set rsG=nothing
						
						LOOP UNTIL ReqExist=0
						'///////////////////////////////////////////////////
						'// END: DO
						'///////////////////////////////////////////////////
						
						'// Insert Gift Codes to Database
						if scDB="Access" then
							query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "',#" & pGCExpDate & "#," & pGCAmount & ",1)"   
						else
							query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "','" & pGCExpDate & "'," & pGCAmount & ",1)"
						end if
						set rsG=Server.CreateObject("ADODB.Recordset")
						set rsG=connTemp.execute(query)				
						set rsG=nothing
					
					Next '// For k=1 to Cint(pQuantity)
	
				END IF '// IF (pGCGenFile<>"") THEN
		
			End If '// If pcv_strRunSection = -1 Then
			
			rstemp.moveNext			
		loop
		set rstemp=nothing
		
	END IF	
	'GGG Add-on end
	
end if 'Process Order - Submit 4
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: Submit 4
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	query="SELECT orders.idcustomer,orders.OrderDate,orders.address,orders.City,orders.StateCode,orders.State,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingState,orders.shippingZip,orders.shippingCountryCode,orders.pcOrd_shippingPhone,orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.iRewardValue,orders.iRewardRefId, orders.iRewardPointsRef,orders.iRewardPointsCustAccrued,orders.ordPackageNum, customers.phone, orders.ord_DeliveryDate, orders.ord_VAT, pcOrd_CatDiscounts FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" &qry_ID
		
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)	
	todaydate=showDateFrmt(now())	
	If NOT rs.EOF Then
		pidcustomer=rs("idcustomer")
		pcv_OrderDate=ShowDateFrmt(rs("OrderDate"))
		paddress=rs("address")
		pCity=rs("city")
		pStateCode=rs("StateCode")
		pState=rs("State")
		pzip=rs("zip")
		pCountryCode=rs("CountryCode")
		pshippingAddress=rs("shippingAddress")
		pshippingCity=rs("shippingCity")
		pshippingStateCode=rs("shippingStateCode")
		pshippingState=rs("shippingState")
		pshippingZip=rs("shippingZip")
		pshippingCountryCode=rs("shippingCountryCode")
		pshippingPhone=rs("pcOrd_shippingPhone")
		pShipmentDetails=rs("ShipmentDetails")
		pPaymentDetails=rs("paymentDetails")
		pdiscountDetails=rs("discountDetails")
		ptaxAmount=rs("taxAmount")
		ptotal=rs("total")
		pcomments=rs("comments")
		pShippingFullName=rs("ShippingFullName")
		paddress2=rs("address2")
		pShippingCompany=rs("ShippingCompany")
		pShippingAddress2=rs("ShippingAddress2")
		ptaxDetails=rs("taxDetails")
		piRewardValue=rs("iRewardValue")
		piRewardRefId=rs("iRewardRefId")
		piRewardPointsRef=rs("iRewardPointsRef")
		piRewardPointsCustAccrued=rs("iRewardPointsCustAccrued")
		pOrdPackageNum=rs("ordPackageNum")
		pPhone=rs("phone")
		pord_DeliveryDate=rs("ord_DeliveryDate")
		pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)
		pord_VAT=rs("ord_VAT")
		pcOrd_CatDiscounts=rs("pcOrd_CatDiscounts")
	End If
	set rs = nothing
	
	
	if (request.querystring("Submit4")<>"") or (pcv_SubmitType=3) then	
		'// Update reward pts.
		'RP ADDON-S
		If RewardsActive <> 0 then			
			'// Add points from refferer if any points were awarded.
			If piRewardRefId>0 AND piRewardPointsRef>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & piRewardRefId
				set rsCust=Server.CreateObject("ADODB.Recordset")
				set rsCust=conntemp.execute(query)
				iAccrued=rsCust("iRewardPointsAccrued") + piRewardPointsRef
				set rsCust=nothing
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & piRewardRefId
				set rsCust=Server.CreateObject("ADODB.Recordset")
				set rsCust=conntemp.Execute(query)
				set rsCust=nothing
			end if 
			
			'// Add accrued points from customer if any points were accrued
			If piRewardPointsCustAccrued>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
				set rsCust=Server.CreateObject("ADODB.Recordset")
				set rsCust=conntemp.execute(query)
				iAccrued=rsCust("iRewardPointsAccrued") + piRewardPointsCustAccrued
				set rsCust=nothing
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & pIdCustomer
				set rsCust=Server.CreateObject("ADODB.Recordset")
				set rsCust=conntemp.Execute(query)
				set rsCust=nothing
			end if
		end if 
		'RP ADDON-E		
	end if 'Process Order - Submit 4
	
	
	query="SELECT name, lastname, customerCompany, email, pcCust_VATID, pcCust_SSN FROM customers WHERE idcustomer="& pIdCustomer
	Set rsCust=Server.CreateObject("ADODB.Recordset")
	Set rsCust=conntemp.execute(query)
	if NOT rsCust.eof then
		pName=rsCust("name")
		pLName=rsCust("lastname")
		pCustomerCompany=rsCust("customerCompany")
		pEmail=rsCust("email")
		pVATID=rsCust("pcCust_VATID")
		pSSN=rsCust("pcCust_SSN")
	end if
	set rsCust=nothing
	%>

	<!--#include file="sendmailCustomerProcessed.asp"-->

	<%
	if (request.QueryString("sendEmailConf")="YES") OR (request.querystring("Submit4A")<>"") OR (request.querystring("Submit4B")<>"") OR (pcv_SubmitType=3) then
		if (pcv_SubmitType=0) or (pcv_SubmitType=1) or (pcv_SubmitType=3) then
			pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_2") & " - " & dictLanguage.Item(Session("language")&"_sendMail_1") & (scpre + int(pIdOrder))
		else
			pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_1") & " - " & dictLanguage.Item(Session("language")&"_sendMail_1") & (scpre + int(pIdOrder))
		end if
		if pCheckEmail="YES" then
			call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerEmail, "&quot;", chr(34)))
		end if
	end if
	
'Start SDBA
'Send Order Notification E-mail to Drop-Shippers
if (pcv_SubmitType=0) or (pcv_SubmitType=3) then
	pcv_DropShipperID=0
	pcv_IsSupplier=0 
	%>	
	<!--#include file="../pc/inc_DropShipperNotificationEmail.asp"-->	
	<%
end if
'End SDBA

if (request.querystring("Submit4")<>"") or (pcv_SubmitType=3) then	
	
	'// If order was by an affiliate, send affiliate email
	query="SELECT orders.idaffiliate, orders.affiliatePay, affiliates.affiliateemail, affiliates.affiliateName FROM orders, affiliates WHERE affiliates.idaffiliate=orders.idaffiliate AND orders.idOrder="& qry_ID
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	
	If NOT rstemp.eof Then

		if rstemp("idaffiliate")<>1 then
			AffiliateEmail=rstemp("affiliateemail")
			AffiliateName=rstemp("affiliateName")
			
			AffiliateOrderEmail=""
			AffiliateOrderEmail=AffiliateOrderEmail & dictLanguage.Item(Session("language")&"_storeEmail_10") & AffiliateName &","
			AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf
			AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf & dictLanguage.Item(Session("language")&"_storeEmail_11") & scCompanyName & dictLanguage.Item(Session("language")&"_storeEmail_12")
			AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf
			AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf & dictLanguage.Item(Session("language")&"_storeEmail_13") & scCurSign&money(rstemp("affiliatePay"))

			AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf
			AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf & dictLanguage.Item(Session("language")&"_storeEmail_14") & scCompanyName
			pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_3")
			call sendmail (scCompanyName, scEmail, AffiliateEmail, pcv_strSubject, AffiliateOrderEmail)
		end if
		
	End If '// If NOT rstemp.eof then
	set rstemp=nothing
	
end if 'Process Order - Submit 4

%>
